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

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');
const { promisify } = require('util');

const execFileAsync = promisify(execFile);

const PROJECT_BASE_DIR =
  process.env.TRANSFORM_PROJECTS_DIR ||
  path.join(require('os').homedir(), 'Downloads', 'transform-projects');

// ─── Directory helpers ────────────────────────────────────────────────────────

/** Lowercase, strip spaces and punctuation — for similarity / retail normalization. */
function normalizeComparable(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

/** Common retail spellings after `normalizeComparable` (hyphens/spaces already removed). */
function normalizeRetailSpelling(comp) {
  if (!comp) return '';
  let s = comp;
  s = s.replace(/penney/g, 'penny');
  s = s.replace(/walmarts/g, 'walmart');
  return s;
}

/**
 * Extra spellings for fuzzy match: trailing …ey vs …y (Penney/Penny style) and reverse.
 */
function spellingVariantsForComparable(comp) {
  const base = normalizeRetailSpelling(comp);
  const out = new Set([base]);
  if (base.length >= 5 && base.endsWith('ey')) {
    out.add(base.slice(0, -2) + 'y');
  }
  if (base.length >= 5 && base.endsWith('y') && !base.endsWith('ey')) {
    out.add(base.slice(0, -1) + 'ey');
  }
  return [...out];
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = new Array(n + 1);
  let cur = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1;
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    const t = prev;
    prev = cur;
    cur = t;
  }
  return prev[n];
}

/** Best similarity in [0,1] across spelling variants; uses normalized Levenshtein (1 - dist/maxLen). */
function bestNameSimilarity(queryName, folderName) {
  const qNorm = normalizeComparable(queryName);
  const fNorm = normalizeComparable(folderName);
  const qVars = spellingVariantsForComparable(qNorm);
  const fVars = spellingVariantsForComparable(fNorm);
  let best = 0;
  for (const q of qVars) {
    for (const f of fVars) {
      const maxLen = Math.max(q.length, f.length);
      if (maxLen < 4) continue;
      const sim = 1 - levenshtein(q, f) / maxLen;
      if (sim > best) best = sim;
    }
  }
  return best;
}

/**
 * Returns the full path to a project folder by fuzzy-matching the project name.
 * e.g. "Shelter Cove" matches ~/Downloads/transform-projects/Shelter Cove/
 */
function findProjectFolder(projectName) {
  if (!projectName || !String(projectName).trim()) return null;
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

  // Fuzzy: retail spelling + …ey/…y variants + ≥80% character similarity (normalized Levenshtein)
  const MIN_SIM = 0.8;
  let bestFolder = null;
  let bestScore = 0;
  for (const f of folders) {
    const score = bestNameSimilarity(projectName, f);
    if (score >= MIN_SIM && score > bestScore) {
      bestScore = score;
      bestFolder = f;
    }
  }
  if (bestFolder) return path.join(PROJECT_BASE_DIR, bestFolder);

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
  const basename = path.basename(filePath);
  let bytes = 0;
  try {
    bytes = fs.statSync(filePath).size;
  } catch {
    bytes = 0;
  }

  if (bytes > 55 * 1024 * 1024) {
    const tmpHuge = path.join(
      os.tmpdir(),
      `pdftxt-huge-${process.pid}-${Date.now()}-${basename.replace(/[^a-zA-Z0-9._-]/g, '_')}.txt`
    );
    try {
      await execFileAsync(
        'pdftotext',
        ['-layout', '-nopgbrk', '-f', '1', '-l', '30', filePath, tmpHuge],
        { timeout: 300000, maxBuffer: 24 * 1024 * 1024 }
      );
      if (fs.existsSync(tmpHuge)) {
        const head = fs.readFileSync(tmpHuge, 'utf8');
        try {
          fs.unlinkSync(tmpHuge);
        } catch {
          /* ignore */
        }
        const note =
          '[Large PDF — first ~30 pages as text (title blocks usually here). Remainder may be vector/scanned; use page images when provided.]';
        return `[PDF: ${basename}]\n${note}\n\n${head.trim()}`;
      }
    } catch {
      try {
        if (fs.existsSync(tmpHuge)) fs.unlinkSync(tmpHuge);
      } catch {
        /* ignore */
      }
    }
  }

  let text = '';
  try {
    const { PDFParse } = require('pdf-parse');
    const parser = new PDFParse({ verbosity: 0, url: filePath });
    const data = await parser.getText();
    text = String(data?.text || '');
  } catch {
    text = '';
  }

  const compactLen = text.replace(/\s+/g, ' ').trim().length;
  if (compactLen < 120) {
    const tmpOut = path.join(os.tmpdir(), `pdftxt-${process.pid}-${Date.now()}-${basename.replace(/[^a-zA-Z0-9._-]/g, '_')}.txt`);
    try {
      await execFileAsync('pdftotext', ['-layout', '-nopgbrk', filePath, tmpOut], {
        timeout: 180000,
        maxBuffer: 50 * 1024 * 1024,
      });
      if (fs.existsSync(tmpOut)) {
        const fallback = fs.readFileSync(tmpOut, 'utf8');
        if (fallback.replace(/\s+/g, ' ').trim().length > compactLen) {
          text = fallback;
        }
      }
    } catch {
      /* keep pdf-parse text */
    } finally {
      try {
        if (fs.existsSync(tmpOut)) fs.unlinkSync(tmpOut);
      } catch {
        /* ignore */
      }
    }
  }

  const body =
    (text || '').trim() ||
    '(minimal or no text layer — vector/scanned title blocks may only appear in page images.)';
  return `[PDF: ${basename}]\n${body}`;
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

/** Lower = read / raster earlier in contact-style jobs (title blocks, plansets). */
function contactFilePriority(filePath) {
  const n = path.basename(filePath).toLowerCase();
  const ext = path.extname(n);
  if (ext !== '.pdf') return 40;
  if (
    /planset|plan.?set|combined.?set|full.?set|master.?plan|drawing.?index|plan.?book|set.?of.?draw/.test(n)
  ) {
    return 0;
  }
  if (
    /geotech|spec|addendum|bid|itb|prequal|transmittal|00[-_]|report|proposal|contract|instruction|submittal|cover|directory|form|notice|pre-bid|rfp|rfq/.test(
      n
    )
  ) {
    return 1;
  }
  if (/plan|draw|sheet|arch|struct|civil|mep|survey|landscape|code/.test(n)) return 2;
  return 10;
}

function sortFilesForContactRead(files) {
  return [...files].sort((a, b) => {
    const pa = contactFilePriority(a);
    const pb = contactFilePriority(b);
    if (pa !== pb) return pa - pb;
    try {
      const sa = fs.statSync(a).size;
      const sb = fs.statSync(b).size;
      if (path.extname(a).toLowerCase() === '.pdf' && path.extname(b).toLowerCase() === '.pdf') {
        return sb - sa;
      }
      return 0;
    } catch {
      return 0;
    }
  });
}

/**
 * Reads all supported files in a folder and returns combined extracted text.
 * Truncates to maxChars (default ~80k). Optional contactReadOrder boosts plansets/specs before the slice cuts off huge sets.
 */
async function readAllFilesInFolder(folderPath, maxChars = 80000, readOpts = {}) {
  let files = listProjectFiles(folderPath);
  if (readOpts.contactReadOrder) {
    files = sortFilesForContactRead(files);
  }
  const parts = [];
  for (const f of files) {
    const text = await readFile(f);
    if (text) parts.push(text);
  }
  const combined = parts.join('\n\n' + '─'.repeat(60) + '\n\n');
  return combined.slice(0, maxChars);
}

/**
 * Rasterize PDF page range to JPEG (requires `pdftoppm`, e.g. poppler).
 * @param maxPages page count starting at opts.firstPage (default 1)
 * @param opts.firstPage 1-based first page; opts.timeoutMs override; opts.skipDpiClamp
 * Returns { data, media_type, filename, sourcePdf, pageLabel }[] for Claude vision.
 */
async function rasterizePdfPages(filePath, maxPages = 3, dpi = 100, opts = {}) {
  const firstPage = Math.max(1, parseInt(String(opts.firstPage ?? 1), 10) || 1);
  const n = Math.max(1, parseInt(String(maxPages), 10) || 1);
  const lastPage = firstPage + n - 1;

  let bytes = 0;
  try {
    bytes = fs.statSync(filePath).size;
  } catch {
    bytes = 0;
  }
  let timeoutMs = Number(opts.timeoutMs);
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    if (bytes > 80 * 1024 * 1024) timeoutMs = 420000;
    else if (bytes > 40 * 1024 * 1024) timeoutMs = 300000;
    else if (bytes > 15 * 1024 * 1024) timeoutMs = 180000;
    else timeoutMs = 120000;
  }
  let dpiUse = dpi;
  if (!opts.skipDpiClamp && bytes > 45 * 1024 * 1024) {
    dpiUse = Math.min(dpiUse, 72);
  }

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdfpages-'));
  const prefix = path.join(tmpDir, 'page');
  const baseArgs = ['-jpeg', '-r', String(dpiUse), '-f', String(firstPage), '-l', String(lastPage)];
  try {
    try {
      await execFileAsync(
        'pdftoppm',
        [...baseArgs, '-scale-to', '2048', filePath, prefix],
        { timeout: timeoutMs, maxBuffer: 64 * 1024 * 1024 }
      );
    } catch (e1) {
      await execFileAsync('pdftoppm', [...baseArgs, filePath, prefix], {
        timeout: timeoutMs,
        maxBuffer: 64 * 1024 * 1024,
      });
    }
    let files = fs.readdirSync(tmpDir).filter((f) => f.endsWith('.jpg'));
    files.sort((a, b) => {
      const na = parseInt(String(a).replace(/\D/g, ''), 10) || 0;
      const nb = parseInt(String(b).replace(/\D/g, ''), 10) || 0;
      return na - nb;
    });
    files = files.map((f) => path.join(tmpDir, f));
    const baseName = path.basename(filePath);
    const images = files.map((f) => {
      const m = path.basename(f).match(/-(\d+)\.jpg$/i);
      const pageLabel = m ? m[1] : String(firstPage);
      return {
        data: fs.readFileSync(f).toString('base64'),
        media_type: 'image/jpeg',
        filename: baseName,
        sourcePdf: baseName,
        pageLabel,
      };
    });
    for (const f of files) {
      try {
        fs.unlinkSync(f);
      } catch {
        /* ignore */
      }
    }
    fs.rmdirSync(tmpDir);
    return images;
  } catch (err) {
    console.error('[pdf-rasterize]', path.basename(filePath), err.message);
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      /* ignore */
    }
    return [];
  }
}

/**
 * Text from `readPdf` plus rasterized page images for one PDF.
 */
async function readPdfWithVision(filePath, options = {}) {
  const { maxPages = 3, dpi = 100 } = options;
  const [text, images] = await Promise.all([
    readPdf(filePath),
    rasterizePdfPages(filePath, maxPages, dpi, { firstPage: 1 }),
  ]);
  return { text, images };
}

/**
 * Full-folder text via `readAllFilesInFolder` plus JPEGs from PDFs (up to caps).
 * @param options.contactCoverPages — for contact extraction: page 1 of many PDFs first (title blocks), then optional page 2 pass
 */
async function readAllFilesWithVision(folderPath, options = {}) {
  const {
    maxImagesPerPdf = 3,
    maxTotalImages = 12,
    imageDpi = 100,
    contactCoverPages = false,
    maxTextChars = 80000,
  } = options;
  const text = await readAllFilesInFolder(folderPath, maxTextChars, {
    contactReadOrder: contactCoverPages,
  });
  const allImages = [];
  let pdfs = listProjectFiles(folderPath).filter((fp) => fp.toLowerCase().endsWith('.pdf'));

  if (contactCoverPages) {
    pdfs = sortFilesForContactRead(pdfs);
  } else {
    pdfs.sort((a, b) => {
      const pa = contactFilePriority(a);
      const pb = contactFilePriority(b);
      if (pa !== pb) return pa - pb;
      try {
        return fs.statSync(a).size - fs.statSync(b).size;
      } catch {
        return 0;
      }
    });
  }

  if (contactCoverPages) {
    for (const pdfPath of pdfs) {
      if (allImages.length >= maxTotalImages) break;
      let bytes = 0;
      try {
        bytes = fs.statSync(pdfPath).size;
      } catch {
        continue;
      }
      const timeoutMs = bytes > 80 * 1024 * 1024 ? 420000 : bytes > 40 * 1024 * 1024 ? 300000 : bytes > 15 * 1024 * 1024 ? 180000 : 120000;
      let dpiUse = imageDpi;
      if (bytes > 45 * 1024 * 1024) dpiUse = Math.min(dpiUse, 72);
      const pages = await rasterizePdfPages(pdfPath, 1, dpiUse, { firstPage: 1, timeoutMs });
      allImages.push(...pages);
    }
    if (maxImagesPerPdf >= 2) {
      for (const pdfPath of pdfs) {
        if (allImages.length >= maxTotalImages) break;
        let bytes = 0;
        try {
          bytes = fs.statSync(pdfPath).size;
        } catch {
          continue;
        }
        const timeoutMs = bytes > 80 * 1024 * 1024 ? 420000 : bytes > 40 * 1024 * 1024 ? 300000 : 180000;
        let dpiUse = imageDpi;
        if (bytes > 45 * 1024 * 1024) dpiUse = Math.min(dpiUse, 72);
        const pages = await rasterizePdfPages(pdfPath, 1, dpiUse, { firstPage: 2, timeoutMs });
        allImages.push(...pages);
      }
    }
  } else {
    for (const pdfPath of pdfs) {
      if (allImages.length >= maxTotalImages) break;
      const remaining = maxTotalImages - allImages.length;
      const nPages = Math.min(maxImagesPerPdf, remaining);
      if (nPages <= 0) break;
      const pages = await rasterizePdfPages(pdfPath, nPages, imageDpi, { firstPage: 1 });
      allImages.push(...pages);
    }
  }
  return { text, images: allImages };
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

function looksLikeQuestionText(text) {
  const t = String(text || '').trim();
  if (!t) return false;
  if (/[?]\s*$/.test(t)) return true;
  return /\b(which|what|where|when|how|find|does|is)\b/i.test(t);
}

/**
 * Detects project-file question intent by matching a known project folder name
 * in a question-like user message. Returns matched project folder name or null.
 */
function detectProjectFileQuestion(text) {
  if (!looksLikeQuestionText(text)) return null;
  if (!fs.existsSync(PROJECT_BASE_DIR)) return null;
  const entries = fs.readdirSync(PROJECT_BASE_DIR, { withFileTypes: true });
  const folders = entries.filter((e) => e.isDirectory()).map((e) => e.name);
  if (folders.length === 0) return null;

  const lower = String(text || '').toLowerCase();
  const normText = normalizeComparable(text);

  let best = null;
  for (const folder of folders) {
    const folderLower = folder.toLowerCase();
    const folderNorm = normalizeComparable(folder);
    const exactish = lower.includes(folderLower);
    const normContains = folderNorm.length >= 4 && normText.includes(folderNorm);
    if (exactish || normContains) {
      if (!best || folder.length > best.length) best = folder;
    }
  }
  return best;
}

function detectDocumentTask(text) {
  const lower = String(text || '').toLowerCase();
  const taskKeywords = /\b(make|create|generate|build|extract|find|gather|pull|list|compile)\b/i.test(text);
  const docKeywords = /\b(pdf|contact|info|sheet|summary|list|directory)\b/i.test(lower);
  const fileRef = /transform-projects|downloads folder|files in|folder|downloaded/i.test(lower);
  return taskKeywords && docKeywords && fileRef;
}

// ─── Exports ──────────────────────────────────────────────────────────────────

module.exports = {
  PROJECT_BASE_DIR,
  findProjectFolder,
  listProjectFiles,
  findFileByName,
  readFile,
  readAllFilesInFolder,
  rasterizePdfPages,
  readPdfWithVision,
  readAllFilesWithVision,
  formatFileListForSlack,
  detectFileReference,
  detectProjectNameReference,
  detectProjectFileQuestion,
  detectDocumentTask,
  SUPPORTED_EXTENSIONS,
};
