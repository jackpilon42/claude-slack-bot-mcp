#!/usr/bin/env node

/**
 * pdf-extractor.js (Enhanced)
 * Standalone PDF Text Extraction Tool for macOS
 *
 * Requires: Homebrew, Poppler (brew install poppler)
 *
 * Usage:
 *   node pdf-extractor.js extract <pdf-file>
 *   node pdf-extractor.js inspect <pdf-file>
 *   node pdf-extractor.js pages <pdf-file> <page-range>
 *   node pdf-extractor.js images <pdf-file>
 *   node pdf-extractor.js raster <pdf-file> <page>
 *   node pdf-extractor.js batch extract [--dry-run] -- <paths...>
 *   node pdf-extractor.js batch inspect [--dry-run] -- <paths...>
 *
 * Examples:
 *   node pdf-extractor.js extract bid.pdf
 *   node pdf-extractor.js batch extract -- *.pdf
 *   node pdf-extractor.js batch extract -- /path/to/pdfs/
 *   node pdf-extractor.js batch extract --dry-run -- *.pdf
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const { globSync } = require('glob');

// ─── Configuration ───────────────────────────────────────────────────────────

const CONFIG = {
  tempDir: path.join(os.tmpdir(), 'pdf-extractor'),
  outputDir: process.cwd(),
  pdfTool: 'pdftotext',
  infoPdfInfo: 'pdfinfo',
  imageTool: 'pdfimages',
  rasterTool: 'pdftoppm',
  dpi: 150,
};

// Ensure temp directory exists
if (!fs.existsSync(CONFIG.tempDir)) {
  fs.mkdirSync(CONFIG.tempDir, { recursive: true });
}

// ─── Batch Output Structure ──────────────────────────────────────────────────

/**
 * Create timestamped batch run directory
 * Returns: { runDir, timestamp }
 */
function createBatchRunDir() {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const runDir = path.join(CONFIG.outputDir, `pdf-extract-${timestamp}`);

  if (!fs.existsSync(runDir)) {
    fs.mkdirSync(runDir, { recursive: true });
  }

  return { runDir, timestamp };
}

/**
 * Create per-PDF subdirectory
 */
function createPdfOutputDir(runDir, pdfFileName) {
  const safeName = path.basename(pdfFileName, '.pdf').replace(/[^a-zA-Z0-9_-]/g, '_');
  const pdfDir = path.join(runDir, safeName);

  if (!fs.existsSync(pdfDir)) {
    fs.mkdirSync(pdfDir, { recursive: true });
  }

  return pdfDir;
}

/**
 * Initialize manifest for batch run
 */
function createManifest(runDir, timestamp) {
  return {
    timestamp,
    runDir,
    startTime: new Date().toISOString(),
    endTime: null,
    totalFiles: 0,
    successful: 0,
    failed: 0,
    files: {},
    errors: [],
  };
}

/**
 * Write manifest to disk
 */
function writeManifest(manifest) {
  const manifestPath = path.join(manifest.runDir, 'manifest.json');
  manifest.endTime = new Date().toISOString();
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
  return manifestPath;
}

/**
 * Write error report
 */
function writeErrorReport(manifest) {
  if (manifest.errors.length === 0) return null;

  const reportPath = path.join(manifest.runDir, 'errors.json');
  const report = {
    timestamp: new Date().toISOString(),
    totalErrors: manifest.errors.length,
    errors: manifest.errors,
  };

  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
  return reportPath;
}

// ─── Helper Functions ────────────────────────────────────────────────────────

/**
 * Check if a command exists in PATH
 */
function commandExists(cmd) {
  try {
    execSync(`which ${cmd}`, { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

/**
 * Verify all required tools are installed
 */
function checkDependencies() {
  const required = [CONFIG.pdfTool, CONFIG.infoPdfInfo, CONFIG.imageTool, CONFIG.rasterTool];
  const missing = required.filter((cmd) => !commandExists(cmd));

  if (missing.length > 0) {
    console.error('❌ Missing required tools:');
    missing.forEach((cmd) => console.error(`   - ${cmd}`));
    console.error('\n📦 Install Poppler with:');
    console.error('   brew install poppler');
    process.exit(1);
  }
}

/**
 * Expand file patterns and directories
 * Handles wildcards, relative paths, and directory recursion
 */
function expandFilePaths(patterns) {
  const files = new Set();

  patterns.forEach((pattern) => {
    const resolved = path.resolve(pattern);
    const stats = fs.existsSync(resolved) ? fs.statSync(resolved) : null;

    if (stats && stats.isDirectory()) {
      const findPdfs = (dir) => {
        const contents = fs.readdirSync(dir);
        contents.forEach((file) => {
          const filePath = path.join(dir, file);
          const fileStats = fs.statSync(filePath);
          if (fileStats.isDirectory()) {
            findPdfs(filePath);
          } else if (filePath.toLowerCase().endsWith('.pdf')) {
            files.add(filePath);
          }
        });
      };
      findPdfs(resolved);
    } else if (pattern.includes('*') || pattern.includes('?')) {
      const matches = globSync(pattern, { absolute: true, nodir: true });
      matches.forEach((f) => {
        if (f.toLowerCase().endsWith('.pdf')) {
          files.add(f);
        }
      });
    } else if (fs.existsSync(resolved) && resolved.toLowerCase().endsWith('.pdf')) {
      files.add(resolved);
    }
  });

  return Array.from(files).sort();
}

/**
 * Get PDF file information
 */
function getPdfInfo(pdfPath) {
  try {
    const info = execSync(`${CONFIG.infoPdfInfo} "${pdfPath}"`, { encoding: 'utf-8' });
    const lines = info.split('\n');

    const result = {
      path: pdfPath,
      raw: info,
      parsed: {},
    };

    lines.forEach((line) => {
      const [key, ...valueParts] = line.split(/:\s+/);
      if (key && valueParts.length > 0) {
        result.parsed[key.toLowerCase()] = valueParts.join(': ').trim();
      }
    });

    return result;
  } catch (error) {
    throw new Error(`Failed to read PDF info: ${error.message}`);
  }
}

/**
 * Extract text from entire PDF
 */
function extractFullText(pdfPath, outputPath) {
  try {
    execSync(`${CONFIG.pdfTool} -layout "${pdfPath}" "${outputPath}"`, {
      stdio: 'pipe',
    });

    const stats = fs.statSync(outputPath);
    return {
      success: true,
      outputPath,
      sizeKb: (stats.size / 1024).toFixed(2),
    };
  } catch (error) {
    throw new Error(`Extraction failed: ${error.message}`);
  }
}

/**
 * Extract text from specific page range
 */
function extractPageRange(pdfPath, pageRange, outputPath) {
  const [startStr, endStr] = pageRange.split('-');
  const start = parseInt(startStr, 10);
  const end = parseInt(endStr, 10) || start;

  if (isNaN(start) || isNaN(end) || start < 1 || end < start) {
    throw new Error('Invalid page range. Use format: "1-10" or "5"');
  }

  try {
    execSync(`${CONFIG.pdfTool} -f ${start} -l ${end} -layout "${pdfPath}" "${outputPath}"`, {
      stdio: 'pipe',
    });

    const stats = fs.statSync(outputPath);
    return {
      success: true,
      pages: `${start}-${end}`,
      outputPath,
      sizeKb: (stats.size / 1024).toFixed(2),
    };
  } catch (error) {
    throw new Error(`Page range extraction failed: ${error.message}`);
  }
}

/**
 * Extract images from PDF
 */
function extractImages(pdfPath, outputDir) {
  const outputPrefix = path.join(outputDir, 'images');

  try {
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    execSync(`${CONFIG.imageTool} -png "${pdfPath}" "${outputPrefix}"`, {
      stdio: 'pipe',
    });

    const files = fs.readdirSync(outputDir);
    const images = files.filter((f) => f.startsWith('images') && f.endsWith('.png'));

    return {
      success: true,
      count: images.length,
      images,
      outputDir,
    };
  } catch (error) {
    throw new Error(`Image extraction failed: ${error.message}`);
  }
}

/**
 * Rasterize a PDF page for visual inspection
 */
function rasterizePage(pdfPath, pageNum, outputDir) {
  const baseName = path.basename(pdfPath, '.pdf');
  const outputPath = path.join(outputDir, baseName);

  try {
    execSync(
      `${CONFIG.rasterTool} -jpeg -r ${CONFIG.dpi} -f ${pageNum} -l ${pageNum} "${pdfPath}" "${outputPath}"`,
      { stdio: 'pipe' }
    );

    const files = fs.readdirSync(outputDir).filter((f) => f.startsWith(baseName));
    if (files.length > 0) {
      const actualFile = path.join(outputDir, files[0]);
      const stats = fs.statSync(actualFile);
      return {
        success: true,
        page: pageNum,
        outputFile: actualFile,
        sizeKb: (stats.size / 1024).toFixed(2),
      };
    }
    throw new Error('Raster output file not found after pdftoppm');
  } catch (error) {
    throw new Error(`Rasterization failed: ${error.message}`);
  }
}

/**
 * Validate PDF file exists and is readable
 */
function validatePdf(pdfPath) {
  if (!fs.existsSync(pdfPath)) {
    throw new Error(`File not found: ${pdfPath}`);
  }

  if (!fs.statSync(pdfPath).isFile()) {
    throw new Error(`Not a file: ${pdfPath}`);
  }

  if (!pdfPath.toLowerCase().endsWith('.pdf')) {
    console.warn(`⚠️  Warning: File doesn't have .pdf extension`);
  }
}

/**
 * Display help message
 */
function showHelp() {
  console.log(`
PDF Text Extraction Tool (Enhanced)
===================================

Usage:
  node pdf-extractor.js <command> <file> [options]
  node pdf-extractor.js batch <subcommand> [options] -- <paths...>

Single File Commands:
  extract <file>           Extract all text (outputs to .txt)
  inspect <file>           Show PDF info and diagnostics
  pages <file> <range>     Extract specific pages (e.g., "1-10" or "5")
  images <file>            Extract embedded images
  raster <file> <page>     Rasterize a page to JPEG for visual inspection

Batch Commands:
  batch extract [--dry-run] -- <paths...>
                           Batch extract all PDFs
  batch inspect [--dry-run] -- <paths...>
                           Batch inspect all PDFs

Paths can be:
  - File globs: *.pdf, *.PDF
  - Directories: /path/to/pdfs/ (recursive)
  - Individual files: /path/to/bid.pdf

Options:
  --dry-run                Show files that would be processed without processing

Examples:
  # Single file
  node pdf-extractor.js extract bid.pdf
  node pdf-extractor.js inspect proposal.pdf
  node pdf-extractor.js pages spec.pdf 1-10

  # Batch processing
  node pdf-extractor.js batch extract -- *.pdf
  node pdf-extractor.js batch extract -- /path/to/pdfs/
  node pdf-extractor.js batch extract --dry-run -- *.pdf

Output:
  - Single files: .txt saved to current directory
  - Batch runs: Timestamped directory with per-PDF folders, manifest.json, errors.json

Setup:
  If Poppler tools are not found, install with:
    brew install poppler
`);
}

// ─── Batch Processing ───────────────────────────────────────────────────────

/**
 * Process batch of PDFs
 */
function processBatch(subcommand, filePaths, options = {}) {
  const { dryRun = false } = options;

  if (filePaths.length === 0) {
    console.error('❌ No PDF files found');
    process.exit(1);
  }

  if (dryRun) {
    console.log(`📋 Dry run: Would process ${filePaths.length} file(s):\n`);
    filePaths.forEach((f, i) => {
      console.log(`   ${i + 1}. ${f}`);
    });
    console.log(`\nRemove --dry-run to actually process files.`);
    return;
  }

  const { runDir, timestamp } = createBatchRunDir();
  const manifest = createManifest(runDir, timestamp);
  manifest.totalFiles = filePaths.length;

  console.log(`📁 Batch run directory: ${runDir}\n`);

  filePaths.forEach((pdfPath, index) => {
    const fileNum = index + 1;
    const pdfName = path.basename(pdfPath);
    const manifestKey = path.resolve(pdfPath);

    process.stdout.write(`[${fileNum}/${filePaths.length}] ${pdfName} ... `);

    try {
      validatePdf(pdfPath);
      const pdfDir = createPdfOutputDir(runDir, pdfName);

      let result = {};

      switch (subcommand) {
        case 'extract': {
          const outputPath = path.join(pdfDir, `${path.basename(pdfPath, '.pdf')}.txt`);
          result = extractFullText(pdfPath, outputPath);
          break;
        }

        case 'inspect': {
          result = getPdfInfo(pdfPath);
          const infoPath = path.join(pdfDir, 'info.json');
          fs.writeFileSync(infoPath, JSON.stringify(result.parsed, null, 2));
          result.infoFile = infoPath;
          break;
        }

        default:
          throw new Error(`Unknown batch subcommand: ${subcommand}`);
      }

      manifest.files[manifestKey] = {
        status: 'success',
        basename: pdfName,
        pdfDir,
        result,
        processedAt: new Date().toISOString(),
      };

      manifest.successful++;
      console.log('✅');
    } catch (error) {
      const errorEntry = {
        file: pdfName,
        path: pdfPath,
        error: error.message,
        timestamp: new Date().toISOString(),
      };

      manifest.files[manifestKey] = {
        status: 'error',
        basename: pdfName,
        error: error.message,
        processedAt: new Date().toISOString(),
      };

      manifest.errors.push(errorEntry);
      manifest.failed++;

      console.log(`❌ ${error.message}`);
    }
  });

  const manifestPath = writeManifest(manifest);
  const errorReportPath = writeErrorReport(manifest);

  console.log(`\n${'═'.repeat(60)}`);
  console.log(`📊 Batch Summary:`);
  console.log(`   Total: ${manifest.totalFiles} | Success: ${manifest.successful} | Failed: ${manifest.failed}`);
  console.log(`   📄 Manifest: ${manifestPath}`);
  if (errorReportPath) {
    console.log(`   ⚠️  Errors: ${errorReportPath}`);
  }
  console.log(`   📁 Output: ${runDir}`);
}

// ─── Main CLI ────────────────────────────────────────────────────────────────

function main() {
  const args = process.argv.slice(2);

  if (args.length === 0 || args[0] === '--help' || args[0] === '-h') {
    showHelp();
    return;
  }

  checkDependencies();

  const command = args[0];

  if (command === 'batch') {
    if (args.length < 2) {
      console.error('❌ Batch requires a subcommand (extract, inspect)\n');
      showHelp();
      process.exit(1);
    }

    const subcommand = args[1];
    const doubleHyphenIndex = args.indexOf('--');

    if (doubleHyphenIndex === -1) {
      console.error('❌ Batch commands require -- separator before file paths\n');
      console.error('Example: batch extract -- *.pdf');
      process.exit(1);
    }

    const options = {};
    const optionArgs = args.slice(2, doubleHyphenIndex);
    const pathArgs = args.slice(doubleHyphenIndex + 1);

    if (optionArgs.includes('--dry-run')) {
      options.dryRun = true;
    }

    if (pathArgs.length === 0) {
      console.error('❌ No file paths specified after --\n');
      showHelp();
      process.exit(1);
    }

    const expandedPaths = expandFilePaths(pathArgs);
    processBatch(subcommand, expandedPaths, options);
    return;
  }

  const pdfFile = args[1];

  if (!pdfFile) {
    console.error(`❌ No PDF file specified\n`);
    showHelp();
    process.exit(1);
  }

  const pdfPath = path.resolve(pdfFile);

  try {
    validatePdf(pdfPath);
    const baseName = path.basename(pdfPath, '.pdf');

    switch (command) {
      case 'extract': {
        console.log(`\n📖 Extracting text from ${path.basename(pdfPath)}...\n`);
        const outputPath = path.join(CONFIG.outputDir, `${baseName}.txt`);
        const result = extractFullText(pdfPath, outputPath);
        console.log(`✅ Text extracted successfully`);
        console.log(`   Output: ${outputPath}`);
        console.log(`   Size: ${result.sizeKb} KB\n`);
        break;
      }

      case 'inspect': {
        const result = getPdfInfo(pdfPath);
        console.log('\n📄 PDF Information:\n');
        console.log(result.raw);
        break;
      }

      case 'pages': {
        const pageRange = args[2];
        if (!pageRange) {
          console.error('❌ Page range required (e.g., "1-10")');
          process.exit(1);
        }
        console.log(`\n📖 Extracting pages ${pageRange}...\n`);
        const outputPath = path.join(CONFIG.outputDir, `${baseName}_pages_${pageRange}.txt`);
        const result = extractPageRange(pdfPath, pageRange, outputPath);
        console.log(`✅ Pages extracted`);
        console.log(`   Output: ${outputPath}`);
        console.log(`   Size: ${result.sizeKb} KB\n`);
        break;
      }

      case 'images': {
        console.log(`\n🖼️  Extracting images...\n`);
        const result = extractImages(pdfPath, CONFIG.outputDir);
        console.log(`✅ Extracted ${result.count} image(s)\n`);
        break;
      }

      case 'raster': {
        const pageNum = args[2];
        if (!pageNum) {
          console.error('❌ Page number required (e.g., "1")');
          process.exit(1);
        }
        console.log(`\n🖼️  Rasterizing page ${pageNum}...\n`);
        const result = rasterizePage(pdfPath, pageNum, CONFIG.tempDir);
        console.log(`✅ Page rasterized`);
        console.log(`   Output: ${result.outputFile}`);
        console.log(`   Size: ${result.sizeKb} KB\n`);
        break;
      }

      default: {
        console.error(`❌ Unknown command: ${command}\n`);
        showHelp();
        process.exit(1);
      }
    }
  } catch (error) {
    console.error(`\n❌ Error: ${error.message}\n`);
    process.exit(1);
  }
}

main();
