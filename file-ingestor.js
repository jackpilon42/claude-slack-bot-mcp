/**
 * file-ingestor.js
 * Transform Energy — Project File Ingestion
 *
 * Reads PDF, Excel/CSV, and Word (.doc/.docx) files from a project folder
 * and extracts all text/data into a single string for Claude to process.
 *
 * Supported formats:
 *   .pdf     — pdf-parse
 *   .xlsx    — xlsx
 *   .xls     — xlsx
 *   .csv     — xlsx
 *   .docx    — mammoth
 *   .doc     — mammoth (best effort)
 *   .txt     — native fs
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const PROJECT_BASE_DIR =
  process.env.TRANSFORM_PROJECTS_DIR ||
  path.join(require('os').homedir(), 'Downloads', 'transform-projects');

// ─── Directory helpers ────────────────────────────────────────────────────────

/**
 * Returns the full path to a project folder by fuzzy-matching the project name.
 * e.g. "Shelter Cove" matches ~/Downloads/transform-projects/Shelter Cove/
 */
function findProjectFolder(projectName) {
  if (!fs.existsSync(PROJECT_BASE_DIR)) return null;

  const entries = fs.readdirSync(PROJECT_BASE_DIR, { withFileTypes: true });
  const folders = entries.filter((e) => e.isDirectory()).map((e) => e.name);

  const lower = projectName.toLowerCase().trim();

  // Exact match first
  const exact = folders.find((f) => f.toLowerCase() === lower);
  if (exact) return path.join(PROJECT_BASE_DIR, exact);

  // Partial match
  const partial = folders.find(
    (f) => f.toLowerCase().includes(lower) || lower.includes(f.toLowerCase())
  );
  if (partial) return path.join(PROJECT_BASE_DIR, partial);

  return null;
}

/**
 * Lists all supported files under a project folder (recursive into subfolders).
 * Skips .zip archives and Office temp lock files (~$…).
 */
const SUPPORTED_EXTENSIONS = ['.pdf', '.xlsx', '.xls', '.csv', '.docx', '.doc', '.txt'];

function listProjectFiles(folderPath) {
  if (!fs.existsSync(folderPath)) return [];
  const out = [];

  function walk(dir) {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
        continue;
      }
      const name = entry.name;
      if (name.startsWith('~$')) continue;
      const ext = path.extname(name).toLowerCase();
      if (ext === '.zip') continue;
      if (SUPPORTED_EXTENSIONS.includes(ext)) out.push(fullPath);
    }
  }

  walk(folderPath);
  return out.sort();
}

/**
 * Finds a single file by name (fuzzy) within the project base dir and subfolders.
 */
function findFileByName(fileName) {
  const lower = fileName.toLowerCase();

  function search(dir) {
    if (!fs.existsSync(dir)) return null;
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        const found = search(fullPath);
        if (found) return found;
      } else if (entry.name.toLowerCase().includes(lower)) {
        return fullPath;
      }
    }
    return null;
  }

  return search(PROJECT_BASE_DIR);
}

// ─── File readers ─────────────────────────────────────────────────────────────

async function readPdf(filePath) {
  const pdfParse = require('pdf-parse');
  const buffer = fs.readFileSync(filePath);
  const data = await pdfParse(buffer);
  return `[PDF: ${path.basename(filePath)}]\n${data.text}`;
}

async function readExcel(filePath) {
  const XLSX = require('xlsx');
  const workbook = XLSX.readFile(filePath);
  const parts = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
    if (csv.trim()) {
      parts.push(`[Sheet: ${sheetName}]\n${csv}`);
    }
  }
  return `[Excel: ${path.basename(filePath)}]\n${parts.join('\n\n')}`;
}

async function readCsv(filePath) {
  const XLSX = require('xlsx');
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
  return `[CSV: ${path.basename(filePath)}]\n${csv}`;
}

async function readDocx(filePath) {
  const mammoth = require('mammoth');
  const result = await mammoth.extractRawText({ path: filePath });
  return `[Word Doc: ${path.basename(filePath)}]\n${result.value}`;
}

async function readTxt(filePath) {
  return `[Text: ${path.basename(filePath)}]\n${fs.readFileSync(filePath, 'utf8')}`;
}

/**
 * Reads a single file and returns extracted text.
 */
async function readFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  try {
    switch (ext) {
      case '.pdf':  return await readPdf(filePath);
      case '.xlsx':
      case '.xls':  return await readExcel(filePath);
      case '.csv':  return await readCsv(filePath);
      case '.docx':
      case '.doc':  return await readDocx(filePath);
      case '.txt':  return await readTxt(filePath);
      default:      return null;
    }
  } catch (err) {
    return `[Error reading ${path.basename(filePath)}: ${err.message}]`;
  }
}

/**
 * Reads all supported files in a folder and returns combined extracted text.
 * Truncates to ~80,000 chars to stay within Claude's context window.
 */
async function readAllFilesInFolder(folderPath) {
  const files = listProjectFiles(folderPath);
  const parts = [];
  for (const f of files) {
    const text = await readFile(f);
    if (text) parts.push(text);
  }
  const combined = parts.join('\n\n' + '─'.repeat(60) + '\n\n');
  return combined.slice(0, 80000); // ~80k chars is safe for Claude context
}

// ─── Slack helpers ────────────────────────────────────────────────────────────

/**
 * Formats a file list for a Slack confirmation message.
 */
function formatFileListForSlack(folderPath, files) {
  const names = files.map((f) => `• \`${path.basename(f)}\``).join('\n');
  return (
    `:file_folder: *Found ${files.length} file${files.length !== 1 ? 's' : ''} in:*\n` +
    `\`${folderPath}\`\n\n${names}\n\n` +
    `Reply *yes* to use all files, or name a specific file to use just that one.`
  );
}

// ─── Intent detection helpers ─────────────────────────────────────────────────

/**
 * Detects "use file <filename>" pattern.
 * Returns the filename string or null.
 */
function detectFileReference(text) {
  const match = text.match(/\buse\s+file\s+["']?([^"'\n]+?)["']?\s*$/i);
  return match ? match[1].trim() : null;
}

/**
 * Detects "generate proposal for <project name>" pattern.
 * Uses the first line only (Slack often appends newlines). Trims trailing quotes / punctuation.
 */
function detectProjectNameReference(text) {
  const first = String(text || '')
    .trim()
    .split(/\r?\n/)[0]
    .trim();
  const match = first.match(
    /\b(?:generate|create|make|write|build|draft)\s+(?:a\s+)?(?:proposal|bid|quote|estimate)\s+(?:for|on)\s+["']?([^"'\n]+?)["']?(?:[.,;:!?]+)?\s*$/i
  );
  if (!match) return null;
  return match[1]
    .trim()
    .replace(/^["']|["']$/g, '')
    .replace(/[.,;:!?]+$/g, '')
    .trim() || null;
}

// ─── Exports ──────────────────────────────────────────────────────────────────

module.exports = {
  PROJECT_BASE_DIR,
  findProjectFolder,
  listProjectFiles,
  findFileByName,
  readFile,
  readAllFilesInFolder,
  formatFileListForSlack,
  detectFileReference,
  detectProjectNameReference,
  SUPPORTED_EXTENSIONS,
};
