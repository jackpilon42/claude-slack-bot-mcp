/**
 * Transform Energy proposal pipeline — Slack bot integration (claude-slack-bot-mcp).
 * Word output via docx; optional PDF via LibreOffice headless.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawnSync } = require('child_process');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  Footer,
  PageNumber,
  ShadingType,
  BorderStyle,
  HeadingLevel,
  LevelFormat,
  TableLayoutType,
  convertInchesToTwip,
} = require('docx');
const {
  findProjectFolder,
  listProjectFiles,
  findFileByName,
  readFile,
  readAllFilesInFolder,
  formatFileListForSlack,
  detectFileReference,
  detectProjectNameReference,
  detectProjectFileQuestion,
} = require('./file-ingestor');

function getProposalsOutputDir() {
  return process.env.PROPOSALS_OUTPUT_DIR || path.join(os.homedir(), 'Downloads', 'transform-proposals');
}

const BRAND = {
  primary: '1F4E79',
  mid: '2E75B6',
  light: 'D5E8F0',
  white: 'FFFFFF',
};

const COMPANY = {
  name: 'Transform Energy',
  website: 'transformenergy.com',
  ceo: 'Todd Filbrun',
  phone: '209.606.0191',
  email: 'Todd.filbrun@transformenergy.com',
  license: 'Lic. #1063970 | DIR #2000003638',
  location: 'Escalon, CA — serves California (Central Valley and beyond)',
  services:
    'Commercial Solar PV Installation, Battery Storage & EV Charging, Commercial Roofing (TPO, Metal, FRP), Electrical C-10 (Switchgear, Transformers, LED Lighting), Asset Management & O&M',
};

const PAGE_W = 12240;
const PAGE_H = 15840;
const MARGIN_1IN = convertInchesToTwip(1);
const CONTENT_W = PAGE_W - MARGIN_1IN * 2;

const TRADE_KEYWORDS =
  /\b(roofing|electrical|solar|hvac|concrete|plumbing|drywall|framing|demolition|excavation|paving|metal\s+roof|tpo|epdm|switchgear|transformer|battery|ev\s+charg|photovoltaic|pv)\b/i;
const BID_KEYWORDS =
  /\b(bid\s*date|due\s*date|scope\s+of\s+work|invitation\s+to\s+bid|\bitb\b|\brfp\b|rfq|pre-?bid|addendum|plan\s*holders|prebid)\b/i;
const OPPORTUNITY_URL =
  /\b(?:https?:\/\/)?(?:www\.)?(?:planhub|buildingconnected|constructconnect|planetbids)\.(?:com|net)\/[^\s]*\b/i;
const EXPLICIT_PROPOSAL =
  /\b(?:generate|create|make|write|build|draft)\b[\s\S]{0,80}\b(?:proposal|bid|quote|estimate)\b|\b(?:proposal|bid|quote|estimate)\b[\s\S]{0,80}\b(?:generate|create|make|write|build|draft)\b/i;

const pendingProposalByThread = new Map();
/** @type {Map<string, { folderPath: string, files: string[], baseUserText: string, createdAtMs: number }>} */
const pendingFileIngestByThread = new Map();
const PENDING_TTL_MS = 24 * 60 * 60 * 1000;

function threadKey(channel, threadTs) {
  return `${String(channel || '')}:${String(threadTs || '')}`;
}

function looksLikeExplicitProposalRequest(text) {
  return EXPLICIT_PROPOSAL.test(String(text || ''));
}

function looksLikePastedOpportunity(text) {
  const t = String(text || '');
  if (OPPORTUNITY_URL.test(t)) return true;
  if (BID_KEYWORDS.test(t) && TRADE_KEYWORDS.test(t)) return true;
  return false;
}

function looksLikeProposalRequest(text) {
  return looksLikeExplicitProposalRequest(text) || looksLikePastedOpportunity(text);
}

function looksLikeConfirmation(text) {
  const t = String(text || '')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/^(yes|yep|yeah|approve|approved|ok|okay|sure|go\s*ahead|do\s*it|please|proceed|sounds\s+good|let'?s\s+go)\b/.test(t)) return true;
  if (/^(y|ok)\W*$/i.test(t)) return true;
  return false;
}

function shouldHandleAsProposalPipeline(text, ctx = {}) {
  const t = String(text || '');
  const ch = ctx.channel;
  const ts = ctx.threadTs;
  const key = threadKey(ch, ts);

  const pendingFile = pendingFileIngestByThread.get(key);
  if (pendingFile) {
    if (Date.now() - pendingFile.createdAtMs > PENDING_TTL_MS) {
      pendingFileIngestByThread.delete(key);
    } else if (looksLikeConfirmation(t) || detectFileReference(t)) {
      return true;
    }
  }

  const pending = pendingProposalByThread.get(key);
  if (pending && looksLikeConfirmation(t)) {
    if (Date.now() - pending.createdAtMs > PENDING_TTL_MS) {
      pendingProposalByThread.delete(key);
      return false;
    }
    return true;
  }
  if (detectProjectFileQuestion(t)) return true;
  return looksLikeProposalRequest(t);
}

function extractAssistantText(message) {
  const blocks = message?.content || [];
  return blocks
    .filter((b) => b.type === 'text')
    .map((b) => b.text)
    .join('\n')
    .trim();
}

function parseJsonFromModelText(raw) {
  let s = String(raw || '').trim();
  const fence = /^```(?:json)?\s*([\s\S]*?)```$/im.exec(s);
  if (fence) s = fence[1].trim();
  const start = s.indexOf('{');
  const end = s.lastIndexOf('}');
  if (start >= 0 && end > start) s = s.slice(start, end + 1);
  return JSON.parse(s);
}

async function extractProjectDetails(userText, anthropicClient) {
  const prompt = `You are extracting structured project data from a construction bid opportunity pasted in Slack (may include URLs, ITB text, scope snippets).

Return a single JSON object ONLY (no markdown) with these keys:
projectName, projectAddress, city, state, ownerOrGC, contactName, contactEmail, scopeDescription, tradeCategories (array of strings), estimatedValue (number or null), bidDueDate (ISO date string or null), projectStartDate (ISO or null), estimatedDuration (string or null), squareFootage (number or null), additionalNotes (string), sourceUrl (string or null).

Use null where unknown. tradeCategories: short labels like "Roofing", "Electrical", "Solar PV".

User message:
---
${String(userText).slice(0, 12000)}
---`;

  const msg = await anthropicClient.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }],
  });
  return parseJsonFromModelText(extractAssistantText(msg));
}

async function scoreProjectFit(projectDetails, anthropicClient) {
  const prompt = `Transform Energy (California) services:
${COMPANY.services}

Score how well this opportunity fits Transform Energy (0-100). Return JSON ONLY:
{ "score": number, "recommendation": "BID" | "PASS" | "REVIEW", "matchedServices": string[], "reasoning": string }

Rules: BID if score >= 60. PASS if score < 40. REVIEW if 40-59.

Project JSON:
${JSON.stringify(projectDetails).slice(0, 8000)}`;

  const msg = await anthropicClient.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 800,
    messages: [{ role: 'user', content: prompt }],
  });
  const parsed = parseJsonFromModelText(extractAssistantText(msg));
  let score = Number(parsed.score);
  if (!Number.isFinite(score)) score = 50;
  score = Math.max(0, Math.min(100, score));
  let rec = String(parsed.recommendation || '').toUpperCase();
  if (rec !== 'BID' && rec !== 'PASS' && rec !== 'REVIEW') {
    if (score >= 60) rec = 'BID';
    else if (score < 40) rec = 'PASS';
    else rec = 'REVIEW';
  }
  return {
    score,
    recommendation: rec,
    matchedServices: Array.isArray(parsed.matchedServices) ? parsed.matchedServices : [],
    reasoning: String(parsed.reasoning || ''),
  };
}

async function generateProposalContent(projectDetails, fitResult, anthropicClient) {
  const prompt = `You are preparing proposal content for Transform Energy (commercial solar, roofing, electrical, storage, EV, O&M — California).

Return JSON ONLY (no markdown) with this shape:
{
  "proposalDate": "YYYY-MM-DD",
  "executiveSummaryNarrative": "2-4 sentences",
  "differentiators": ["bullet as plain string", "..."],
  "scopeItems": [{ "item": "short title", "description": "..." }],
  "lineItems": [
    {
      "division": "01",
      "divisionTitle": "General Requirements",
      "items": [{ "num": "01-001", "description": "...", "unit": "LS", "qty": 1, "totalCost": 12345 }]
    }
  ],
  "totalContractPrice": number,
  "costSummaryRows": [{ "category": "Division 07 Roofing", "amount": 12345 }],
  "timeline": {
    "startDate": "YYYY-MM-DD",
    "completionDate": "YYYY-MM-DD",
    "durationWeeks": number,
    "phases": [{ "phase": "Mobilization", "tasks": ["..."], "startWeek": 1, "endWeek": 2 }]
  }
}

Use realistic California commercial construction pricing. Align divisions with CSI MasterFormat (01, 07, 16, 26, etc.). Include division subtotals inside lineItems groups. costSummaryRows should mirror major divisions plus fees; last row conceptually is grand total (you may omit duplicate if totalContractPrice is authoritative).

Fit context: ${JSON.stringify(fitResult).slice(0, 2000)}
Project: ${JSON.stringify(projectDetails).slice(0, 6000)}`;

  const msg = await anthropicClient.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4000,
    messages: [{ role: 'user', content: prompt }],
  });
  const proposalContent = parseJsonFromModelText(extractAssistantText(msg));

  // Timeline validation/correction: keep dates in the future and internally consistent.
  const today = new Date();

  function parseFlexibleDate(str) {
    if (!str) return null;
    const d = new Date(str);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function addWeeks(date, weeks) {
    const d = new Date(date);
    d.setDate(d.getDate() + weeks * 7);
    return d;
  }

  function formatDate(date) {
    return date.toISOString().split('T')[0];
  }

  if (proposalContent.timeline) {
    const t = proposalContent.timeline;
    let startDate = parseFlexibleDate(t.startDate);
    const durationWeeks = Number(t.durationWeeks) > 0 ? Number(t.durationWeeks) : 8;

    // If start date is in the past or missing, anchor to next Monday from today.
    if (!startDate || startDate < today) {
      const nextMonday = new Date(today);
      const day = nextMonday.getDay();
      const daysUntilMonday = day === 0 ? 1 : ((8 - day) % 7 || 7);
      nextMonday.setDate(nextMonday.getDate() + daysUntilMonday);
      startDate = nextMonday;
    }

    t.startDate = formatDate(startDate);
    t.durationWeeks = durationWeeks;
    t.completionDate = formatDate(addWeeks(startDate, durationWeeks));
  }

  return proposalContent;
}

function safeFileSegment(name) {
  const base = String(name || 'project')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .slice(0, 80);
  return base || 'project';
}

function moneyUsd(n) {
  const x = Number(n);
  if (!Number.isFinite(x)) return '$0';
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(x);
}

function arialRun(text, opts = {}) {
  return new TextRun({
    text: String(text ?? ''),
    font: 'Arial',
    size: opts.size ?? 22,
    bold: opts.bold,
    italics: opts.italics,
    color: opts.color,
  });
}

function p(children, paraOpts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...paraOpts,
    children,
  });
}

function headerCell(text, widthTwip) {
  return new TableCell({
    width: { size: widthTwip, type: WidthType.DXA },
    shading: { fill: BRAND.primary, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [arialRun(text, { bold: true, color: BRAND.white, size: 20 })] })],
  });
}

function bodyCell(text, widthTwip, opts = {}) {
  return new TableCell({
    width: { size: widthTwip, type: WidthType.DXA },
    shading: opts.shadingFill
      ? { fill: opts.shadingFill, type: ShadingType.CLEAR }
      : { type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [arialRun(text, { size: opts.size ?? 22, bold: opts.bold })] })],
  });
}

function tableThinBorders() {
  const b = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b };
}

async function buildDocx(projectDetails, proposalContent, outputDir, options = {}) {
  const safeProjectName = safeFileSegment(projectDetails.projectName || 'Transform-Proposal');
  const baseOut = path.resolve(String(outputDir || '').trim() || getProposalsOutputDir());
  const projectDir = options.useOutputDirAsProjectDir ? baseOut : path.join(baseOut, safeProjectName);
  fs.mkdirSync(projectDir, { recursive: true });
  const docxPath = path.join(projectDir, `${safeProjectName}-proposal.docx`);

  const proposalDate =
    proposalContent.proposalDate || new Date().toISOString().slice(0, 10);
  const total = Number(proposalContent.totalContractPrice) || 0;
  const scopeItems = Array.isArray(proposalContent.scopeItems) ? proposalContent.scopeItems : [];
  const lineGroups = Array.isArray(proposalContent.lineItems) ? proposalContent.lineItems : [];
  const costSummaryRows = Array.isArray(proposalContent.costSummaryRows) ? proposalContent.costSummaryRows : [];
  const timeline = proposalContent.timeline || {};
  const phases = Array.isArray(timeline.phases) ? timeline.phases : [];
  const differentiators = Array.isArray(proposalContent.differentiators) ? proposalContent.differentiators : [];

  const children = [];

  const numbering = {
    config: [
      {
        reference: 'teBullet',
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: '-',
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.25) },
              },
            },
          },
        ],
      },
    ],
  };

  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          arialRun(
            `${COMPANY.name} | ${COMPANY.website} | ${COMPANY.license} | Page `,
            { size: 18 }
          ),
          new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18 }),
        ],
      }),
    ],
  });

  // --- 1. Cover ---
  children.push(
    p([arialRun('PROPOSAL', { bold: true, size: 48, color: BRAND.primary })], { alignment: AlignmentType.CENTER, spacing: { after: 200 } })
  );
  children.push(
    p([arialRun(String(projectDetails.projectName || 'Project'), { bold: true, size: 36 })], {
      alignment: AlignmentType.CENTER,
    })
  );
  children.push(
    p([arialRun(`Prepared for: ${projectDetails.ownerOrGC || 'Client'}`, { size: 24 })], { alignment: AlignmentType.CENTER })
  );
  if (projectDetails.contactName || projectDetails.contactEmail) {
    children.push(
      p(
        [
          arialRun(
            `Contact: ${[projectDetails.contactName, projectDetails.contactEmail].filter(Boolean).join(' | ')}`,
            { size: 22 }
          ),
        ],
        { alignment: AlignmentType.CENTER }
      )
    );
  }
  children.push(
    p([arialRun(`Prepared by: ${COMPANY.ceo}, ${COMPANY.name}`, { size: 24 })], { alignment: AlignmentType.CENTER })
  );
  children.push(
    p(
      [
        arialRun(
          `${proposalDate}  |  Total proposal value: ${moneyUsd(total)}  |  ${projectDetails.projectAddress || ''}`,
          { size: 22 }
        ),
      ],
      { alignment: AlignmentType.CENTER }
    )
  );
  if (projectDetails.scopeDescription) {
    children.push(
      p([arialRun(`Scope (summary): ${String(projectDetails.scopeDescription).slice(0, 800)}`, { italics: true })], {
        alignment: AlignmentType.CENTER,
      })
    );
  }

  // --- 2. Executive Summary ---
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [arialRun('Executive Summary', { bold: true, color: BRAND.primary })] }));
  children.push(p([arialRun(proposalContent.executiveSummaryNarrative || 'Executive summary pending.', { size: 22 })]));

  // --- 3. About ---
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [arialRun(`About ${COMPANY.name}`, { bold: true, color: BRAND.primary })] }));
  children.push(
    p([
      arialRun(
        `${COMPANY.name} is headquartered in ${COMPANY.location}. We deliver ${COMPANY.services}.`,
        { size: 22 }
      ),
    ])
  );
  for (const line of differentiators.slice(0, 12)) {
    children.push(
      new Paragraph({
        numbering: { reference: 'teBullet', level: 0 },
        children: [arialRun(String(line), { size: 22 })],
      })
    );
  }

  // --- 4. Scope of Work table ---
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [arialRun('Scope of Work', { bold: true, color: BRAND.primary })] }));
  {
    const w1 = Math.floor(CONTENT_W * 0.28);
    const w2 = CONTENT_W - w1;
    const rows = [
      new TableRow({ children: [headerCell('Scope Item', w1), headerCell('Description', w2)] }),
      ...scopeItems.slice(0, 40).map(
        (row) =>
          new TableRow({
            children: [
              bodyCell(row.item || '', w1),
              bodyCell(row.description || '', w2),
            ],
          })
      ),
    ];
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [w1, w2],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows,
      })
    );
  }

  // --- 5. Detailed Cost Estimate ---
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [arialRun('Detailed Cost Estimate', { bold: true, color: BRAND.primary })],
    })
  );
  {
    const c0 = 500;
    const c1 = CONTENT_W - c0 - 700 - 600 - 1100;
    const c2 = 700;
    const c3 = 600;
    const c4 = 1100;
    const estRows = [
      new TableRow({
        children: [
          headerCell('#', c0),
          headerCell('Description', c1),
          headerCell('Unit', c2),
          headerCell('Qty', c3),
          headerCell('Total Cost', c4),
        ],
      }),
    ];
    for (const grp of lineGroups) {
      const divLabel = `Div ${grp.division || ''} — ${grp.divisionTitle || ''}`;
      estRows.push(
        new TableRow({
          children: [
            new TableCell({
              columnSpan: 5,
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: BRAND.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [arialRun(divLabel, { bold: true, color: BRAND.primary })] })],
            }),
          ],
        })
      );
      let sub = 0;
      const items = Array.isArray(grp.items) ? grp.items : [];
      for (const it of items) {
        const amt = Number(it.totalCost) || 0;
        sub += amt;
        estRows.push(
          new TableRow({
            children: [
              bodyCell(String(it.num || ''), c0),
              bodyCell(String(it.description || ''), c1),
              bodyCell(String(it.unit || ''), c2),
              bodyCell(String(it.qty ?? ''), c3),
              bodyCell(moneyUsd(amt), c4, { bold: true }),
            ],
          })
        );
      }
      estRows.push(
        new TableRow({
          children: [
            new TableCell({
              columnSpan: 4,
              width: { size: c0 + c1 + c2 + c3, type: WidthType.DXA },
              children: [new Paragraph({ alignment: AlignmentType.END, children: [arialRun('Division subtotal', { bold: true })] })],
            }),
            bodyCell(moneyUsd(sub), c4, { bold: true, shadingFill: BRAND.light }),
          ],
        })
      );
    }
    estRows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 4,
            width: { size: c0 + c1 + c2 + c3, type: WidthType.DXA },
            shading: { fill: BRAND.mid, type: ShadingType.CLEAR },
            children: [new Paragraph({ alignment: AlignmentType.END, children: [arialRun('GRAND TOTAL', { bold: true, color: BRAND.white })] })],
          }),
          new TableCell({
            width: { size: c4, type: WidthType.DXA },
            shading: { fill: BRAND.mid, type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [arialRun(moneyUsd(total), { bold: true, color: BRAND.white })] })],
          }),
        ],
      })
    );
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [c0, c1, c2, c3, c4],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows: estRows,
      })
    );
  }

  // --- 6. Cost Summary ---
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [arialRun('Cost Summary', { bold: true, color: BRAND.primary })] }));
  {
    const wA = Math.floor(CONTENT_W * 0.62);
    const wB = CONTENT_W - wA;
    const rows = [new TableRow({ children: [headerCell('Cost Category', wA), headerCell('Amount', wB)] })];
    for (const r of costSummaryRows.slice(0, 30)) {
      rows.push(
        new TableRow({
          children: [bodyCell(String(r.category || ''), wA), bodyCell(moneyUsd(r.amount), wB, { bold: true })],
        })
      );
    }
    rows.push(
      new TableRow({
        children: [
          bodyCell('Grand Total', wA, { bold: true, shadingFill: BRAND.light }),
          bodyCell(moneyUsd(total), wB, { bold: true, shadingFill: BRAND.light }),
        ],
      })
    );
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [wA, wB],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows,
      })
    );
  }

  // --- 7. Terms & Conditions ---
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [arialRun('Terms & Conditions', { bold: true, color: BRAND.primary })],
    })
  );
  {
    const wA = Math.floor(CONTENT_W * 0.28);
    const wB = CONTENT_W - wA;
    const terms = [
      ['Validity', 'This proposal is valid for 30 days from the proposal date unless otherwise noted.'],
      ['Payment', 'Net 30; progress billing aligned with completed milestones unless otherwise agreed in writing.'],
      ['Change Orders', 'Changes in scope, schedule, or price require written approval prior to execution.'],
      ['Warranty', 'Five (5) years workmanship on installed systems and applicable roofing/electrical work, subject to manufacturer terms.'],
      ['Exclusions', 'Permits/fees unless listed; hazardous materials abatement; latent conditions; owner-furnished delays; utility service upgrades beyond noted scope.'],
      ['Insurance', `${COMPANY.name} maintains statutory workers compensation and commercial general liability as required for California licensed contractors.`],
    ];
    const rows = [new TableRow({ children: [headerCell('Topic', wA), headerCell('Terms', wB)] })];
    for (const [a, b] of terms) {
      rows.push(new TableRow({ children: [bodyCell(a, wA, { bold: true }), bodyCell(b, wB)] }));
    }
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [wA, wB],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows,
      })
    );
  }

  // --- 8. Construction Timeline ---
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [arialRun('Construction Timeline', { bold: true, color: BRAND.primary })],
    })
  );
  {
    const w0 = 1600;
    const w1 = CONTENT_W - w0 - 900 - 900;
    const w2 = 900;
    const w3 = 900;
    const rows = [
      new TableRow({
        children: [
          headerCell('Phase', w0),
          headerCell('Task / Activity', w1),
          headerCell('Start Wk', w2),
          headerCell('End Wk', w3),
        ],
      }),
    ];
    for (const ph of phases.slice(0, 25)) {
      const taskStr = (Array.isArray(ph.tasks) ? ph.tasks : []).join('; ');
      rows.push(
        new TableRow({
          children: [
            bodyCell(String(ph.phase || ''), w0),
            bodyCell(taskStr, w1),
            bodyCell(String(ph.startWeek ?? ''), w2),
            bodyCell(String(ph.endWeek ?? ''), w3),
          ],
        })
      );
    }
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [w0, w1, w2, w3],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows,
      })
    );
  }
  children.push(
    p([
      arialRun(
        'Note: Dates and week numbers are planning placeholders. A detailed CPM schedule will be submitted after award and notice to proceed.',
        { italics: true, size: 20 }
      ),
    ])
  );

  // --- 9. Proposal Acceptance ---
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [arialRun('Proposal Acceptance', { bold: true, color: BRAND.primary })],
    })
  );
  {
    const half = Math.floor(CONTENT_W / 2);
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [half, half],
        layout: TableLayoutType.FIXED,
        borders: tableThinBorders(),
        rows: [
          new TableRow({
            children: [
              bodyCell(`${COMPANY.name}\n\nSignature: ___________________________\n\nName: ${COMPANY.ceo}\n\nTitle: CEO`, half),
              bodyCell(
                `${projectDetails.ownerOrGC || 'Client / GC'}\n\nSignature: ___________________________\n\nName: ___________________________\n\nTitle: ___________________________`,
                half
              ),
            ],
          }),
        ],
      })
    );
  }

  const doc = new Document({
    creator: COMPANY.name,
    title: `${safeProjectName} — Proposal`,
    description: 'Commercial construction proposal',
    numbering,
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 22 },
          paragraph: { spacing: { after: 120 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H },
            margin: {
              top: MARGIN_1IN,
              right: MARGIN_1IN,
              bottom: MARGIN_1IN,
              left: MARGIN_1IN,
            },
          },
        },
        footers: { default: footer },
        children,
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(docxPath, buf);
  return { docxPath, projectDir, safeProjectName };
}

async function generateHumanTasksPDF(folderPath, projectName, tasks) {
  const children = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: `${projectName} — Human Task Instructions`,
          bold: true,
          size: 32,
          font: 'Arial',
        }),
      ],
      spacing: { after: 320 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: 'The following tasks require manual completion by a team member.',
          size: 22,
          font: 'Arial',
          italics: true,
        }),
      ],
      spacing: { after: 400 },
    }),
  ];

  tasks.forEach((task, i) => {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [
          new TextRun({
            text: `Step ${i + 1}: ${task.task || 'Task'}`,
            bold: true,
            size: 26,
            font: 'Arial',
          }),
        ],
        spacing: { before: 320, after: 120 },
      })
    );
    if (task.item) children.push(new Paragraph({ children: [new TextRun({ text: `• Item: ${task.item}`, size: 22, font: 'Arial' })] }));
    if (task.destination) children.push(new Paragraph({ children: [new TextRun({ text: `• Destination: ${task.destination}`, size: 22, font: 'Arial' })] }));
    if (task.contact) children.push(new Paragraph({ children: [new TextRun({ text: `• Contact: ${task.contact}`, size: 22, font: 'Arial' })] }));
    if (task.deadline) children.push(new Paragraph({ children: [new TextRun({ text: `• Deadline: ${task.deadline}`, size: 22, font: 'Arial' })] }));
    children.push(new Paragraph({ spacing: { after: 200 } }));
  });

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children,
      },
    ],
  });

  const safeProjectName = String(projectName || 'project')
    .replace(/[^a-zA-Z0-9-_ ]/g, '')
    .replace(/\s+/g, '-');
  const outPath = path.join(folderPath, `${safeProjectName}-Instructions.docx`);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  console.log(`[smart-docs] Human tasks instructions saved: ${outPath}`);
}

async function fillHDGBidForm(filePath, projectDetails, proposalContent, _anthropicClient) {
  const XLSX = require('xlsx');
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const today = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
  const setCell = (r, c, val, type = 's') => {
    const addr = XLSX.utils.encode_cell({ r, c });
    ws[addr] = type === 'n' ? { t: 'n', v: val } : { t: 's', v: String(val) };
  };

  setCell(7, 1, 'Transform Energy');
  setCell(8, 1, today);

  const lineItems = proposalContent.lineItems || [];
  const scopeMap = {};
  for (const div of lineItems) {
    for (const item of div.items || []) {
      const desc = String(item.description || '').toLowerCase().trim();
      if (!desc) continue;
      const tc = Number(item.totalCost);
      if (Number.isFinite(tc)) scopeMap[desc] = tc;
    }
  }

  const transformScopes = [
    'electrical',
    'lighting',
    'led',
    'conduit',
    'site lighting',
    'roofing',
    'tpo',
    'metal roof',
    'insulation',
    'solar',
    'pv',
    'battery',
  ];

  for (let r = 13; r <= 77; r++) {
    const row = data[r];
    if (!row) continue;
    const scopeLabel = (String(row[0] || '') + ' ' + String(row[1] || '')).toLowerCase().trim();
    if (!scopeLabel) continue;

    const doesScope = transformScopes.some((s) => scopeLabel.includes(s));

    if (doesScope) {
      setCell(r, 2, 'YES');
      let cost = 0;
      const words = scopeLabel.split(/\s+/).filter((w) => w.length > 3);
      for (const [desc, c] of Object.entries(scopeMap)) {
        if (words.some((w) => desc.includes(w))) {
          cost = c;
          break;
        }
      }
      if (cost > 0) setCell(r, 5, cost, 'n');
    } else if (scopeLabel.length > 2) {
      setCell(r, 2, 'NO');
    }
  }

  let maxR = 77;
  let maxC = 5;
  const curRef = ws['!ref'];
  if (curRef) {
    const rng = XLSX.utils.decode_range(curRef);
    maxR = Math.max(maxR, rng.e.r);
    maxC = Math.max(maxC, rng.e.c);
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });

  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);
  const outPath = path.join(path.dirname(filePath), `${base}_Transform Energy_filled${ext}`);
  XLSX.writeFile(wb, outPath);
  console.log('[smart-docs] Filled HDG bid form:', outPath);
  return outPath;
}

async function handleSmartDocuments(folderPath, projectDetails, proposalContent, anthropicClient) {
  const XLSX = require('xlsx');
  const files = fs
    .readdirSync(folderPath, { withFileTypes: true })
    .filter((f) => !f.isDirectory())
    .map((f) => path.join(folderPath, f.name));

  const knownData = {
    projectName: projectDetails.projectName,
    projectAddress: projectDetails.projectAddress,
    city: projectDetails.city,
    state: projectDetails.state,
    ownerOrGC: projectDetails.ownerOrGC,
    contactName: projectDetails.contactName,
    contactEmail: projectDetails.contactEmail,
    totalContractPrice: proposalContent.totalContractPrice,
    startDate: proposalContent.timeline?.startDate,
    completionDate: proposalContent.timeline?.completionDate,
    durationWeeks: proposalContent.timeline?.durationWeeks,
    companyName: 'Transform Energy',
    companyPhone: '209.606.0191',
    companyEmail: 'Todd.filbrun@transformenergy.com',
    companyLicense: 'Lic. #1063970',
    ceoName: 'Todd Filbrun',
  };

  const humanTasks = [];

  // A) Auto-fill Excel files
  for (const filePath of files) {
    const ext = path.extname(filePath).toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) continue;

    try {
      const wb = XLSX.readFile(filePath);
      let modified = false;

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z100');

        for (let R = range.s.r; R <= range.e.r; R++) {
          for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = ws[cellAddr];
            if (!cell || cell.t !== 's') continue;
            const val = String(cell.v || '').toLowerCase().trim();

            const fieldMap = {
              'contractor name': knownData.companyName,
              subcontractor: knownData.companyName,
              contractor: knownData.companyName,
              license: knownData.companyLicense,
              'license number': knownData.companyLicense,
              'contact name': knownData.ceoName,
              contact: knownData.ceoName,
              phone: knownData.companyPhone,
              email: knownData.companyEmail,
              'start date': knownData.startDate,
              'completion date': knownData.completionDate,
              'project name': knownData.projectName,
              'project address': knownData.projectAddress,
              'total price': knownData.totalContractPrice?.toString(),
              'bid amount': knownData.totalContractPrice?.toString(),
            };

            for (const [pattern, fillValue] of Object.entries(fieldMap)) {
              if (val.includes(pattern) && fillValue) {
                const nextAddr = XLSX.utils.encode_cell({ r: R, c: C + 1 });
                const nextCell = ws[nextAddr];
                if (!nextCell || !nextCell.v) {
                  ws[nextAddr] = { t: 's', v: fillValue };
                  modified = true;
                }
              }
            }
          }
        }
      }

      if (modified) {
        const baseName = path.basename(filePath, ext);
        const outPath = path.join(folderPath, `${baseName}_filled${ext}`);
        XLSX.writeFile(wb, outPath);
        console.log(`[smart-docs] Filled Excel: ${outPath}`);
      }
    } catch (err) {
      console.error(`[smart-docs] Excel error ${filePath}:`, err.message);
    }
  }

  // HDG / breakdown bid forms (pattern-specific fill)
  for (const filePath of files) {
    const name = path.basename(filePath).toUpperCase();
    if (
      path.extname(filePath).toLowerCase() === '.xlsx' &&
      (name.includes('HDG') || name.includes('REQUIRED') || name.includes('BREAKDOWN'))
    ) {
      try {
        await fillHDGBidForm(filePath, projectDetails, proposalContent, anthropicClient);
      } catch (err) {
        console.error(`[smart-docs] HDG form error ${filePath}:`, err.message);
      }
    }
  }

  // B) Use Claude to identify human tasks from all docs
  try {
    const { readAllFilesInFolder } = require('./file-ingestor');
    const extractedText = await readAllFilesInFolder(folderPath);

    const taskResponse = await anthropicClient.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1500,
      system: `You are analyzing project documents for a construction subcontractor (Transform Energy).
Identify any tasks that require a human to physically complete — things like:
mailing documents to a location, delivering items, signing and returning forms, attending a pre-bid meeting, obtaining a bond or insurance certificate, submitting a physical form by mail.
Return ONLY valid JSON array:
[{ "task": "description", "item": "what needs to be sent/done", "destination": "where it goes", "contact": "contact name and info if available", "deadline": "deadline if mentioned" }]
Return empty array [] if no human physical tasks found.`,
      messages: [
        {
          role: 'user',
          content: `Project: ${projectDetails.projectName}\n\n${extractedText.slice(0, 40000)}`,
        },
      ],
    });

    const raw = taskResponse.content.find((b) => b.type === 'text')?.text || '[]';
    const clean = raw.replace(/```json|```/g, '').trim();
    const tasks = JSON.parse(clean);
    if (Array.isArray(tasks)) humanTasks.push(...tasks);
  } catch (err) {
    console.error('[smart-docs] Task extraction error:', err.message);
  }

  // C) Generate human task instructions PDF
  if (humanTasks.length > 0) {
    await generateHumanTasksPDF(folderPath, projectDetails.projectName, humanTasks);
  }
}

function convertToPdf(docxPath, projectDir) {
  const candidates = [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/bin/libreoffice',
    '/usr/local/bin/libreoffice',
    'soffice',
  ];
  const base = path.basename(docxPath, '.docx');
  for (const bin of candidates) {
    try {
      const r = spawnSync(
        bin,
        ['--headless', '--convert-to', 'pdf', '--outdir', projectDir, docxPath],
        { encoding: 'utf-8', maxBuffer: 20 * 1024 * 1024 }
      );
      if (r.error) continue;
      if (r.status !== 0 && r.status !== null) continue;
      const pdfPath = path.join(projectDir, `${base}.pdf`);
      if (fs.existsSync(pdfPath)) return pdfPath;
    } catch {
      /* try next */
    }
  }
  return null;
}

function slackFitEmoji(rec) {
  const u = String(rec || '').toUpperCase();
  if (u === 'BID') return ':large_green_circle:';
  if (u === 'PASS') return ':red_circle:';
  return ':large_yellow_circle:';
}

async function postProposalToSlack(
  slackClient,
  channel,
  threadTs,
  projectDetails,
  fitResult,
  proposalContent,
  docxPath,
  pdfPath
) {
  const total = Number(proposalContent.totalContractPrice) || 0;
  const tl = proposalContent.timeline || {};
  const lines = [
    `${slackFitEmoji(fitResult.recommendation)} *Transform Energy proposal ready*`,
    `*Project:* ${projectDetails.projectName || 'Unknown'}`,
    `*Fit score:* ${fitResult.score}/100 (${fitResult.recommendation})`,
    `*Matched services:* ${(fitResult.matchedServices || []).join(', ') || '—'}`,
    `*Total contract price (planning):* ${moneyUsd(total)}`,
    `*Timeline:* ${tl.startDate || 'TBD'} → ${tl.completionDate || 'TBD'} (${tl.durationWeeks != null ? `${tl.durationWeeks} wks` : 'duration TBD'})`,
    '',
    `*Why:* ${String(fitResult.reasoning || '').slice(0, 1500)}`,
    '',
    `*Word:* \`${docxPath}\``,
    pdfPath ? `*PDF:* \`${pdfPath}\`` : '*PDF:* LibreOffice not available on server — open the .docx locally and export PDF if needed.',
    '',
    'Reply `approve` in this thread to finalize (acknowledgment placeholder for downstream workflow).',
  ];
  await slackClient.chat.postMessage({
    channel,
    thread_ts: threadTs || undefined,
    text: lines.join('\n'),
  });
}

async function postSlackText(slackClient, channel, threadTs, text) {
  await slackClient.chat.postMessage({
    channel,
    thread_ts: threadTs || undefined,
    text,
  });
}

async function handleProposalPipeline(userText, { slackClient, channel, threadTs, anthropicClient }) {
  console.log('[proposal-pipeline] handleProposalPipeline called with: ' + String(userText || '').slice(0, 80));
  const key = threadKey(channel, threadTs);
  let effectiveText = String(userText || '').trim();
  console.log('[proposal-pipeline] effectiveText first 120:', effectiveText.slice(0, 120));
  let fileIngestJustCompleted = false;
  let proposalOutputDir = getProposalsOutputDir();
  let useProjectFolderForDocx = false;

  try {
    const projectQaName = detectProjectFileQuestion(effectiveText);
    if (projectQaName) {
      const folderPath = findProjectFolder(projectQaName);
      if (folderPath) {
        await postSlackText(
          slackClient,
          channel,
          threadTs,
          `:mag: Reading project documents for *${projectQaName}* and answering your question...`
        );
        const docsText = await readAllFilesInFolder(folderPath);
        if (!docsText || !docsText.trim()) {
          await postSlackText(
            slackClient,
            channel,
            threadTs,
            `:warning: I found the project folder for *${projectQaName}*, but I could not extract readable text from its supported files.`
          );
          return 'PROJECT_QA_EMPTY';
        }
        const qaMsg = await anthropicClient.messages.create({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1200,
          system:
            "You are a construction document analyst. Answer the user's question based only on the provided project documents. Be specific — cite which document contains the answer.",
          messages: [
            {
              role: 'user',
              content: `Project folder: ${folderPath}\n\nUser question:\n${effectiveText}\n\nProject documents text:\n${docsText}`,
            },
          ],
        });
        const qaAnswer = extractAssistantText(qaMsg) || 'I could not find enough information in the provided documents to answer that question.';
        await postSlackText(slackClient, channel, threadTs, qaAnswer.slice(0, 39000));
        return 'PROJECT_QA';
      }
    }

    const pendingFiles = pendingFileIngestByThread.get(key);
    if (pendingFiles) {
      if (Date.now() - pendingFiles.createdAtMs > PENDING_TTL_MS) {
        pendingFileIngestByThread.delete(key);
      } else if (looksLikeConfirmation(effectiveText)) {
        const folderPath = pendingFiles.folderPath;
        const blob = await readAllFilesInFolder(folderPath);
        effectiveText = `${pendingFiles.baseUserText}\n\n--- Ingested project files ---\n\n${blob}`;
        pendingFileIngestByThread.delete(key);
        fileIngestJustCompleted = true;
        proposalOutputDir = folderPath;
        useProjectFolderForDocx = true;
      } else if (detectFileReference(effectiveText)) {
        const fname = detectFileReference(effectiveText);
        let fp = null;
        for (const f of pendingFiles.files) {
          if (path.basename(f).toLowerCase().includes(fname.toLowerCase())) {
            fp = f;
            break;
          }
        }
        if (!fp) fp = findFileByName(fname);
        if (fp) {
          const blob = await readFile(fp);
          effectiveText = `${pendingFiles.baseUserText}\n\n--- Ingested single file ---\n\n${blob}`;
        } else {
          await postSlackText(
            slackClient,
            channel,
            threadTs,
            `:x: Could not find a file matching *${fname}*. Reply *yes* for all listed files, or \`use file <name>\` with a partial filename.`
          );
          return 'ERROR';
        }
        pendingFileIngestByThread.delete(key);
        fileIngestJustCompleted = true;
        proposalOutputDir = pendingFiles.folderPath;
        useProjectFolderForDocx = true;
      } else {
        await postSlackText(
          slackClient,
          channel,
          threadTs,
          `:hourglass_flowing_sand: Reply *yes* to ingest every listed file, or \`use file <filename>\` for one file.`
        );
        return 'AWAITING_FILE_CHOICE';
      }
    }

    const alreadyHasIngested =
      effectiveText.includes('--- Ingested project files ---') ||
      effectiveText.includes('--- Ingested single file ---');

    console.log('[pp] fileIngestJustCompleted:', fileIngestJustCompleted);
    console.log('[pp] alreadyHasIngested:', alreadyHasIngested);
    console.log('[pp] looksLikeProposalRequest:', looksLikeProposalRequest(effectiveText));
    const projName = detectProjectNameReference(effectiveText);
    console.log('[pp] detectProjectNameReference:', projName);
    console.log('[pp] findProjectFolder:', projName ? findProjectFolder(projName) : null);
    console.log(
      '[pp] listProjectFiles count:',
      (() => {
        const f = projName ? findProjectFolder(projName) : null;
        return f ? listProjectFiles(f).length : 'no folder';
      })()
    );

    if (
      !fileIngestJustCompleted &&
      !alreadyHasIngested &&
      looksLikeProposalRequest(effectiveText)
    ) {
      if (projName) {
        const folder = findProjectFolder(projName);
        if (folder) {
          const files = listProjectFiles(folder);
          if (files.length > 0) {
            pendingFileIngestByThread.set(key, {
              folderPath: folder,
              files,
              baseUserText: effectiveText,
              createdAtMs: Date.now(),
            });
            await postSlackText(slackClient, channel, threadTs, formatFileListForSlack(folder, files));
            return 'AWAITING_FILE_CHOICE';
          }
        }
      }
    }

    const pending = pendingProposalByThread.get(key);
    if (pending && looksLikeConfirmation(effectiveText) && !fileIngestJustCompleted) {
      effectiveText = pending.sourceText;
      pendingProposalByThread.delete(key);
    } else if (
      !fileIngestJustCompleted &&
      !looksLikeExplicitProposalRequest(effectiveText) &&
      looksLikePastedOpportunity(effectiveText)
    ) {
      pendingProposalByThread.set(key, { sourceText: effectiveText, createdAtMs: Date.now() });
      await postSlackText(
        slackClient,
        channel,
        threadTs,
        `:clipboard: I can draft a full Transform Energy proposal from this opportunity.\n\nWant me to generate it? Reply *yes* or *approve* in this thread.`
      );
      return 'AWAITING_CONFIRMATION';
    }

    await postSlackText(
      slackClient,
      channel,
      threadTs,
      ':gear: Analyzing project and generating proposal... ~30–60 seconds.'
    );

    const projectDetails = await extractProjectDetails(effectiveText, anthropicClient);
    const fitResult = await scoreProjectFit(projectDetails, anthropicClient);

    if (fitResult.recommendation === 'PASS') {
      await postSlackText(
        slackClient,
        channel,
        threadTs,
        `:octagonal_sign: *Skipping proposal generation*\nFit score ${fitResult.score}/100 (PASS).\n${fitResult.reasoning}`
      );
      return 'PASS';
    }

    const proposalContent = await generateProposalContent(projectDetails, fitResult, anthropicClient);
    const { docxPath, projectDir, safeProjectName } = await buildDocx(
      projectDetails,
      proposalContent,
      proposalOutputDir,
      { useOutputDirAsProjectDir: useProjectFolderForDocx }
    );

    // Run smart document handling if we have a project folder
    if (proposalOutputDir && useProjectFolderForDocx) {
      try {
        await handleSmartDocuments(proposalOutputDir, projectDetails, proposalContent, anthropicClient);
      } catch (err) {
        console.error('[smart-docs] Error:', err.message);
      }
    }

    const pdfPath = convertToPdf(docxPath, projectDir);
    await postProposalToSlack(
      slackClient,
      channel,
      threadTs,
      projectDetails,
      fitResult,
      proposalContent,
      docxPath,
      pdfPath
    );
    return 'OK';
  } catch (err) {
    console.error('proposal-pipeline error:', err);
    await postSlackText(
      slackClient,
      channel,
      threadTs,
      `:x: Proposal pipeline failed: ${String(err?.message || err).slice(0, 2500)}`
    );
    pendingProposalByThread.delete(key);
    pendingFileIngestByThread.delete(key);
    return 'ERROR';
  }
}

module.exports = {
  shouldHandleAsProposalPipeline,
  looksLikeProposalRequest,
  looksLikePastedOpportunity,
  handleProposalPipeline,
};
