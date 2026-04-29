/**
 * Optional path: plain-English Smartsheet work goes through Smartsheet's hosted MCP
 * (https://mcp.smartsheet.com) with Claude choosing tools — see README-MCP.md.
 */
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

function anthropicRequestSignal() {
  const ms = Math.min(600_000, Math.max(30_000, Number(process.env.SMARTSHEET_MCP_ANTHROPIC_TIMEOUT_MS || 120_000) || 120_000));
  if (typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function') {
    return AbortSignal.timeout(ms);
  }
  const c = new AbortController();
  setTimeout(() => c.abort(), ms);
  return c.signal;
}

function normalizeAnthropicInputSchema(schema) {
  if (!schema || typeof schema !== 'object') {
    return { type: 'object', properties: {} };
  }
  const s = { ...schema };
  if (!s.type) s.type = 'object';
  if (s.type !== 'object') {
    return { type: 'object', properties: { value: s } };
  }
  if (!s.properties) s.properties = {};
  return s;
}

function mcpToolsToAnthropic(mcpTools) {
  return (mcpTools || []).map((t) => ({
    name: String(t.name || 'tool').slice(0, 200),
    description: String(t.description || `Smartsheet MCP tool ${t.name || ''}`).slice(0, 8000),
    input_schema: normalizeAnthropicInputSchema(t.inputSchema),
  }));
}

function mcpCallResultToToolResultContent(result) {
  const out = [];
  const content = result?.content;
  if (Array.isArray(content)) {
    for (const block of content) {
      if (block?.type === 'text' && typeof block.text === 'string') {
        out.push({ type: 'text', text: block.text.slice(0, 120_000) });
      }
    }
  }
  if (out.length === 0) {
    let raw = '';
    try {
      raw = JSON.stringify(result ?? {}, null, 2);
    } catch {
      raw = String(result);
    }
    out.push({ type: 'text', text: raw.slice(0, 120_000) });
  }
  return out;
}

function buildSystemPrompt({ sheetId, sheetMapLines }) {
  return [
    'You are Transform Sim, a Slack assistant that edits Smartsheet for employees using the attached Smartsheet MCP tools.',
    'FIRST when the user asks for a yes/no "dropdown" on a row or single cell: state clearly that Smartsheet cannot do one-cell dropdowns — PICKLIST is a column type for every row. Then say what you will do instead (whole column, optional cell value).',
    'Use only the MCP tools to read or change Smartsheet. Prefer the smallest set of tool calls that completes the request.',
    'When a tool needs a sheet id, use the active sheet id below unless the user clearly names a different sheet (then you may need to list or search sheets first if tools allow).',
    `Active Smartsheet sheet id (string): ${sheetId || '(not set — ask the user to configure SMARTSHEET_SHEET_ID or switch sheet in Slack)'}.`,
    sheetMapLines ? `Known sheet name → id map:\n${sheetMapLines}` : '',
    'If the user names a sheet that appears in the map above, assume the active sheet id already targets that sheet (Slack routing). If the name does not match the map, say they should fix SMARTSHEET_SHEET_MAP or switch sheets in Slack first.',
    'Smartsheet PICKLIST / dropdown is a column type: Yes/No applies to every row in that column, not a single cell. If they mention a row with "create dropdown", they may want that cell set to Yes or No after the column change—use tools to update that cell if appropriate.',
    'For "what is in column N row M" style questions, read the cell with tools; never infer from column titles or headers.',
    'For bold/italic on a specific cell: Smartsheet uses a 17-field format string (API); bold is index 2 = 1. Use tools to GET the sheet with include=format, merge bold into the cell format, then PUT the row with that cell format while preserving the cell value.',
    'Complex scoped clears ("clear all cells in column 4 except row 1…") may not map cleanly to MCP tools; if you cannot express it with tools, say so briefly and suggest `smartsheet …` commands or turning off SMARTSHEET_USE_OFFICIAL_MCP for that request.',
    'After tools succeed, reply in concise Slack markdown: what changed, which sheet id you affected, caveats (e.g. column-wide picklists), and avoid dumping raw JSON unless the user asked.',
    'If you cannot complete the request with available tools, say what is missing in one short paragraph.',
  ]
    .filter(Boolean)
    .join('\n\n');
}

/**
 * @param {string} userText
 * @param {{ threadPriorBlock?: string, sheetId: string, anthropicClient: import('@anthropic-ai/sdk').default }} opts
 */
export async function handleSmartsheetPlainEnglishViaOfficialMcp(userText, opts) {
  const { threadPriorBlock = '', sheetId, anthropicClient } = opts;
  const token = String(process.env.SMARTSHEET_API_TOKEN || '').trim();
  if (!token) {
    return 'Smartsheet token is not configured (SMARTSHEET_API_TOKEN).';
  }

  const base = String(process.env.SMARTSHEET_MCP_URL || 'https://mcp.smartsheet.com').replace(/\/+$/, '');
  let mcpUrl;
  try {
    mcpUrl = new URL(base);
  } catch {
    return `Invalid SMARTSHEET_MCP_URL: ${base}`;
  }

  const transport = new StreamableHTTPClientTransport(mcpUrl, {
    requestInit: {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  });

  const mcp = new Client({ name: 'transform-sim-slack', version: '1.0.0' });
  try {
    await mcp.connect(transport);
  } catch (e) {
    const msg = String(e?.message || e);
    return (
      `Could not connect to Smartsheet MCP at ${base} (${msg.slice(0, 500)}). ` +
      'Check SMARTSHEET_API_TOKEN, plan access (Smartsheet MCP requires a supported plan), and region URL if not US.'
    );
  }

  let listed;
  try {
    listed = await mcp.listTools();
  } catch (e) {
    try {
      await transport.close();
    } catch {
      /* ignore */
    }
    return `Connected to MCP but tools/list failed: ${String(e?.message || e).slice(0, 600)}`;
  }

  const tools = mcpToolsToAnthropic(listed.tools || []);
  if (tools.length === 0) {
    try {
      await transport.close();
    } catch {
      /* ignore */
    }
    return 'Smartsheet MCP returned no tools — check server status and your account/plan.';
  }

  const mapRaw = String(process.env.SMARTSHEET_SHEET_MAP || '').trim();
  let sheetMapLines = '';
  if (mapRaw) {
    try {
      const obj = JSON.parse(mapRaw);
      sheetMapLines = Object.entries(obj)
        .map(([k, v]) => `- ${k}: ${v}`)
        .join('\n');
    } catch {
      sheetMapLines = '(SMARTSHEET_SHEET_MAP is not valid JSON)';
    }
  }

  const system = buildSystemPrompt({ sheetId, sheetMapLines });
  const userBody = threadPriorBlock ? `${threadPriorBlock}${userText}` : userText;

  const messages = [{ role: 'user', content: userBody }];
  const model = String(process.env.SMARTSHEET_MCP_CLAUDE_MODEL || 'claude-sonnet-4-20250514');
  const maxIterations = Math.min(24, Math.max(4, Number(process.env.SMARTSHEET_MCP_MAX_TOOL_ROUNDS || 12)));

  let lastText = '';

  try {
    for (let round = 0; round < maxIterations; round++) {
      const resp = await anthropicClient.messages.create(
        {
          model,
          max_tokens: 4096,
          temperature: 0,
          system,
          tools,
          messages,
        },
        { signal: anthropicRequestSignal() }
      );

      const blocks = resp.content || [];
      const textParts = [];
      const toolUses = [];
      for (const b of blocks) {
        if (b.type === 'text' && b.text) textParts.push(b.text);
        if (b.type === 'tool_use') toolUses.push(b);
      }
      lastText = textParts.join('\n').trim() || lastText;

      if (toolUses.length === 0) {
        try {
          await transport.close();
        } catch {
          /* ignore */
        }
        return lastText || '_(no text response)_';
      }

      messages.push({ role: 'assistant', content: blocks });

      const toolResultBlocks = [];
      for (const tu of toolUses) {
        const name = tu.name;
        const input = tu.input && typeof tu.input === 'object' ? tu.input : {};
        try {
          const raw = await mcp.callTool({ name, arguments: input });
          toolResultBlocks.push({
            type: 'tool_result',
            tool_use_id: tu.id,
            content: mcpCallResultToToolResultContent(raw),
          });
        } catch (err) {
          toolResultBlocks.push({
            type: 'tool_result',
            tool_use_id: tu.id,
            is_error: true,
            content: [{ type: 'text', text: String(err?.message || err).slice(0, 8000) }],
          });
        }
      }

      messages.push({ role: 'user', content: toolResultBlocks });
    }

    try {
      await transport.close();
    } catch {
      /* ignore */
    }
    return `${lastText ? `${lastText}\n\n` : ''}_Stopped after ${maxIterations} tool rounds (SMARTSHEET_MCP_MAX_TOOL_ROUNDS)._`;
  } catch (err) {
    try {
      await transport.close();
    } catch {
      /* ignore */
    }
    return `Smartsheet MCP agent error: ${String(err?.message || err).slice(0, 1200)}`;
  }
}
