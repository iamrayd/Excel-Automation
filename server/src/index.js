import express from 'express';
import cors from 'cors';
import multer from 'multer';
import ExcelJS from 'exceljs';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

dotenv.config();

// Bypass certificate verification issues caused by corporate firewalls/proxies
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
if (process.env.NODE_TLS_REJECT_UNAUTHORIZED === '0') {
  console.warn('WARNING: NODE_TLS_REJECT_UNAUTHORIZED is set to "0". TLS certificate validation is disabled. Do not use this configuration in production.');
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

// Set up upload directories
const uploadDir = path.join(__dirname, '../uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + '-' + file.originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB file size limit to prevent Denial of Service
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xlsm' || ext === '.xls') {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xlsx, .xlsm, .xls) are allowed.'));
    }
  }
});

// Helper function to extract structured data from excel sheet
function getRowValues(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  if (!row || !row.hasValues) return null;
  const values = [];
  const colCount = worksheet.columnCount || 10;
  for (let col = 1; col <= colCount; col++) {
    const cell = row.getCell(col);
    let val = cell.value;
    if (val && typeof val === 'object') {
      if (val.result !== undefined) {
        val = val.result;
      } else if (val.richText) {
        val = val.richText.map(t => t.text).join('');
      } else if (val.text) {
        val = val.text;
      } else if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      } else {
        val = JSON.stringify(val);
      }
    }
    values.push(val === null || val === undefined ? '' : String(val));
  }
  return values;
}

// Parses workbook using ExcelJS
async function parseExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheets = [];
  workbook.eachSheet((worksheet) => {
    const sheetInfo = {
      name: worksheet.name,
      headers: [],
      samples: [],
    };

    const headers = getRowValues(worksheet, 1);
    if (headers) {
      sheetInfo.headers = headers;
    }

    for (let r = 2; r <= 4; r++) {
      const rowVals = getRowValues(worksheet, r);
      if (rowVals && rowVals.some(v => v !== '')) {
        sheetInfo.samples.push(rowVals);
      }
    }

    sheets.push(sheetInfo);
  });

  return sheets;
}

function extractJSON(text) {
  if (!text) return null;
  const match = text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/) || [null, text];
  const rawJson = match[1] ? match[1].trim() : text.trim();
  try {
    return JSON.parse(rawJson);
  } catch (e) {
    console.error("Failed to parse JSON:", rawJson);
    throw new Error("Failed to parse JSON mapping returned by model.");
  }
}

async function executeTransformation(fromFilePath, toFilePath, outputFilePath, sheetMapping) {
  const fromWorkbook = new ExcelJS.Workbook();
  await fromWorkbook.xlsx.readFile(fromFilePath);

  const toWorkbook = new ExcelJS.Workbook();
  await toWorkbook.xlsx.readFile(toFilePath);

  const fromSheetName = sheetMapping.fromSheet || fromWorkbook.worksheets[0].name;
  const toSheetName = sheetMapping.toSheet || toWorkbook.worksheets[0].name;

  const fromSheet = fromWorkbook.getWorksheet(fromSheetName);
  const toSheet = toWorkbook.getWorksheet(toSheetName);

  if (!fromSheet || !toSheet) {
    throw new Error(`Sheet not found: "${fromSheetName}" in FROM file or "${toSheetName}" in TO file.`);
  }

  const fromHeaders = getRowValues(fromSheet, 1);
  const toHeaders = getRowValues(toSheet, 1);

  if (!fromHeaders || !toHeaders) {
    throw new Error("Could not read header rows from the sheets.");
  }

  const fromHeaderMap = {};
  fromHeaders.forEach((h, idx) => {
    if (h) fromHeaderMap[h.trim().toLowerCase()] = idx + 1;
  });

  const toHeaderMap = {};
  toHeaders.forEach((h, idx) => {
    if (h) toHeaderMap[h.trim().toLowerCase()] = idx + 1;
  });

  const targetGroups = {};
  sheetMapping.mapping.forEach(m => {
    const toColNameClean = m.to.trim().toLowerCase();
    const fromColNameClean = m.from.trim().toLowerCase();
    const targetColIdx = toHeaderMap[toColNameClean];
    const srcColIdx = fromHeaderMap[fromColNameClean];

    if (targetColIdx && srcColIdx) {
      if (!targetGroups[targetColIdx]) {
        targetGroups[targetColIdx] = [];
      }
      if (!targetGroups[targetColIdx].includes(srcColIdx)) {
        targetGroups[targetColIdx].push(srcColIdx);
      }
    }
  });

  let nextRowTo = 2;
  const lastRowToNum = toSheet.lastRow ? toSheet.lastRow.number : 1;
  
  let sheetIsEmpty = true;
  for (let r = 1; r <= lastRowToNum; r++) {
    const row = toSheet.getRow(r);
    if (row && row.hasValues) {
      sheetIsEmpty = false;
      break;
    }
  }

  if (!sheetIsEmpty && toSheet.lastRow) {
    nextRowTo = toSheet.lastRow.number + 1;
  }

  // Find ID column index in target sheet mapping to avoid duplicate rows
  let targetIdColIdx = null;
  sheetMapping.mapping.forEach(m => {
    const toClean = m.to.trim().toLowerCase();
    if (toClean === 'customer id' || toClean === 'id' || toClean === 'client id' || toClean === 'identifier') {
      targetIdColIdx = toHeaderMap[toClean];
    }
  });

  const seenIds = new Set();
  if (targetIdColIdx) {
    // Populate seenIds from existing data in target sheet (rows 2 to nextRowTo - 1)
    for (let r = 2; r < nextRowTo; r++) {
      const row = toSheet.getRow(r);
      if (row && row.hasValues) {
        const cell = row.getCell(targetIdColIdx);
        const val = cell.value !== null && cell.value !== undefined ? String(cell.value).trim() : '';
        if (val !== '') {
          seenIds.add(val);
        }
      }
    }
  }

  const lastRowFromNum = fromSheet.lastRow ? fromSheet.lastRow.number : 1;

  for (let r = 2; r <= lastRowFromNum; r++) {
    const fromRow = fromSheet.getRow(r);
    if (!fromRow || !fromRow.hasValues) continue;

    // Check for duplicate ID values before copying
    if (targetIdColIdx) {
      const srcColIndices = targetGroups[targetIdColIdx];
      if (srcColIndices) {
        const idVal = srcColIndices.map(srcIdx => {
          const cell = fromRow.getCell(srcIdx);
          let val = cell.value;
          if (val && typeof val === 'object') {
            if (val.result !== undefined) val = val.result;
            else if (val.richText) val = val.richText.map(t => t.text).join('');
            else if (val.text) val = val.text;
            else if (val instanceof Date) val = val.toISOString().split('T')[0];
            else val = JSON.stringify(val);
          }
          return val !== null && val !== undefined ? String(val).trim() : '';
        }).filter(v => v !== '').join(' ');

        if (idVal !== '') {
          if (seenIds.has(idVal)) {
            console.log(`Skipping duplicate row ${r} with ID: ${idVal}`);
            continue; // Skip copying this row!
          }
          seenIds.add(idVal);
        }
      }
    }

    const toRow = toSheet.getRow(nextRowTo);
    let rowHasData = false;

    Object.keys(targetGroups).forEach(targetColIdxStr => {
      const targetColIdx = parseInt(targetColIdxStr, 10);
      const srcColIndices = targetGroups[targetColIdx];

      const values = srcColIndices.map(srcIdx => {
        const cell = fromRow.getCell(srcIdx);
        let val = cell.value;
        if (val && typeof val === 'object') {
          if (val.result !== undefined) val = val.result;
          else if (val.richText) val = val.richText.map(t => t.text).join('');
          else if (val.text) val = val.text;
          else if (val instanceof Date) val = val.toISOString().split('T')[0];
          else val = JSON.stringify(val);
        }
        return val !== null && val !== undefined ? String(val).trim() : '';
      }).filter(v => v !== '');

      if (values.length > 0) {
        toRow.getCell(targetColIdx).value = values.join(' ');
        rowHasData = true;
      }
    });

    if (rowHasData) {
      toRow.commit();
      nextRowTo++;
    }
  }

  const outDir = path.dirname(outputFilePath);
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true });
  }

  await toWorkbook.xlsx.writeFile(outputFilePath);
}

function constructAuditPrompt(fromStructure, toStructure) {
  const formatStructure = (structure) => {
    return structure.map(sheet => {
      let sheetStr = `Sheet Name: "${sheet.name}"\n`;
      sheetStr += `Headers: [${sheet.headers.map(h => `"${h}"`).join(', ')}]\n`;
      sheetStr += `Sample Rows:\n`;
      if (sheet.samples.length === 0) {
        sheetStr += `  (No data rows found)\n`;
      } else {
        sheet.samples.forEach((sample, idx) => {
          sheetStr += `  Row ${idx + 2}: [${sample.map(s => `"${s}"`).join(', ')}]\n`;
        });
      }
      return sheetStr;
    }).join('\n');
  };

  return `You are an expert data quality analyst.
Your task is to audit the output of a spreadsheet copy-paste transformation.

Here is the sample of the raw FROM spreadsheet:
${formatStructure(fromStructure)}

Here is the sample of the generated TO spreadsheet:
${formatStructure(toStructure)}

Analyze if the columns from the FROM sheet were correctly matched and copied to the TO sheet. Check for:
1. Mismatched column mapping (e.g. email data in a name column).
2. Missing or blank columns in the output that exist in the source sheet.
3. Mismatched values or data formats.

Output ONLY a JSON object containing your audit result. Do not include markdown code fences (other than optionally \`\`\`json), explanations, or notes outside the JSON block.

JSON Schema:
{
  "status": "Success" | "Warning" | "Error",
  "summary": "Short 1-2 sentence overall summary.",
  "details": [
    "Detail bullet point 1",
    "Detail bullet point 2"
  ]
}
`;
}

// Parse headers and auto-match route
app.post('/api/parse', upload.fields([
  { name: 'from', maxCount: 1 },
  { name: 'to', maxCount: 1 }
]), async (req, res) => {
  try {
    if (!req.files || !req.files['from'] || !req.files['to']) {
      return res.status(400).json({ error: 'Both FROM and TO files are required.' });
    }

    const fromFile = req.files['from'][0];
    const toFile = req.files['to'][0];

    const fromExt = path.extname(fromFile.originalname).toLowerCase();
    const toExt = path.extname(toFile.originalname).toLowerCase();

    if (fromExt === '.xls' || toExt === '.xls') {
      return res.status(400).json({
        error: 'Legacy .xls format is not supported for structure parsing. Please re-save your files as .xlsx or .xlsm before uploading.'
      });
    }

    console.log(`Parsing FROM file: ${fromFile.path}`);
    const fromStructure = await parseExcelFile(fromFile.path);

    console.log(`Parsing TO file: ${toFile.path}`);
    const toStructure = await parseExcelFile(toFile.path);

    const fromHeaders = fromStructure[0]?.headers || [];
    const toHeaders = toStructure[0]?.headers || [];

    // Auto-map headers case-insensitively
    const mappings = [];
    toHeaders.forEach(toH => {
      if (!toH) return;
      const toClean = toH.trim().toLowerCase();
      const match = fromHeaders.find(fromH => fromH && fromH.trim().toLowerCase() === toClean);
      if (match) {
        mappings.push({ from: match, to: toH });
      }
    });

    return res.json({
      fromFileId: fromFile.filename,
      toFileId: toFile.filename,
      fromHeaders,
      toHeaders,
      mappings
    });

  } catch (error) {
    console.error('Error parsing files:', error);
    return res.status(500).json({ error: error.message || 'Internal server error' });
  }
});

// Transfer operation route using custom mappings
app.post('/api/transfer', async (req, res) => {
  try {
    const { fromFileId, toFileId, mappings } = req.body;

    if (!fromFileId || !toFileId || !Array.isArray(mappings)) {
      return res.status(400).json({ error: 'fromFileId, toFileId, and mappings array are required.' });
    }

    // Safety checks against directory traversal / arbitrary file read
    const isSafeFilename = (filename) => {
      if (typeof filename !== 'string') return false;
      const base = path.basename(filename);
      return base === filename && !filename.includes('..') && filename !== '.' && filename !== '';
    };

    if (!isSafeFilename(fromFileId) || !isSafeFilename(toFileId)) {
      return res.status(400).json({ error: 'Security Exception: Invalid file ID format.' });
    }

    // Robust validation of mapping values to prevent server crashes
    const isValidMapping = mappings.every(m => 
      m && 
      typeof m === 'object' && 
      typeof m.from === 'string' && 
      typeof m.to === 'string' &&
      m.from.trim() !== '' &&
      m.to.trim() !== ''
    );
    if (!isValidMapping) {
      return res.status(400).json({ error: 'Invalid mappings format. Each mapping must contain non-empty "from" and "to" strings.' });
    }

    const fromFilePath = path.join(uploadDir, fromFileId);
    const toFilePath = path.join(uploadDir, toFileId);

    if (!fs.existsSync(fromFilePath) || !fs.existsSync(toFilePath)) {
      return res.status(400).json({ error: 'Uploaded source or target file not found on the server. Please upload them again.' });
    }

    console.log(`Resolving file headers for transformation...`);
    const fromStructure = await parseExcelFile(fromFilePath);
    const toStructure = await parseExcelFile(toFilePath);

    const fromSheetName = fromStructure[0]?.name || 'Sheet1';
    const toSheetName = toStructure[0]?.name || 'Sheet1';

    const sheetMapping = {
      fromSheet: fromSheetName,
      toSheet: toSheetName,
      mapping: mappings
    };

    const outputFilename = `${Date.now()}-${path.basename(toFilePath).substring(path.basename(toFilePath).indexOf('-') + 1)}`;
    const outputFilePath = path.join(__dirname, '../outputs', outputFilename);

    console.log('Applying direct custom column transformation...');
    await executeTransformation(fromFilePath, toFilePath, outputFilePath, sheetMapping);

    // AI Audit step with graceful fallback
    let auditReport = {
      status: "Warning",
      summary: "AI Audit skipped.",
      details: ["API key is missing or OpenRouter was rate-limited. The transformation succeeded, but the AI audit was bypassed."]
    };

    const apiKey = process.env.OPENROUTER_API_KEY;
    if (apiKey) {
      try {
        console.log("Parsing generated output file for AI audit...");
        const outputStructure = await parseExcelFile(outputFilePath);
        
        const auditPrompt = constructAuditPrompt(fromStructure, outputStructure);
        const modelName = process.env.OPENROUTER_MODEL || 'deepseek/deepseek-chat-v3-0324:free';
        console.log(`Calling OpenRouter for AI audit using model: ${modelName}`);

        const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`,
            'HTTP-Referer': 'http://localhost:3001',
            'X-Title': 'Sheetshift',
          },
          body: JSON.stringify({
            model: modelName,
            messages: [
              {
                role: 'user',
                content: auditPrompt
              }
            ],
            max_tokens: 1000
          })
        });

        if (response.ok) {
          const data = await response.json();
          const rawContent = data?.choices?.[0]?.message?.content || '';
          const parsed = extractJSON(rawContent);
          if (parsed && parsed.status && parsed.summary) {
            auditReport = parsed;
          }
        } else {
          console.error(`OpenRouter audit request failed with status: ${response.status}`);
        }
      } catch (err) {
        console.error('Failed to run AI audit:', err);
      }
    }

    return res.json({
      message: 'Successfully transferred columns and copied rows',
      downloadUrl: `/api/download/${outputFilename}`,
      audit: auditReport
    });

  } catch (error) {
    console.error('Error executing custom transfer:', error);
    return res.status(500).json({ error: error.message || 'Internal server error' });
  }
});

// Download endpoint
app.get('/api/download/:filename', (req, res) => {
  const filename = req.params.filename;
  
  // Safe filename check against directory traversal
  const isSafeDownloadName = (name) => {
    if (typeof name !== 'string') return false;
    const base = path.basename(name);
    return base === name && !name.includes('..') && name !== '.' && name !== '';
  };
  
  if (!isSafeDownloadName(filename)) {
    return res.status(400).json({ error: 'Security Exception: Invalid download name.' });
  }

  const filePath = path.join(__dirname, '../outputs', filename);
  if (fs.existsSync(filePath)) {
    res.download(filePath, filename);
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

// Custom error handling middleware to capture multer/unhandled exceptions and format cleanly as JSON
app.use((err, req, res, next) => {
  console.error('Unhandled application error:', err);
  const status = err.status || 500;
  res.status(status).json({ error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`Sheetshift backend listening on port ${PORT}`);
});
