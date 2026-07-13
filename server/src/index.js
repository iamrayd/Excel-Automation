import express from 'express';
import cors from 'cors';
import multer from 'multer';
import ExcelJS from 'exceljs';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import os from 'os';
import sql from 'mssql';

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

// Set up upload and output directories
const uploadDir = process.env.VERCEL ? path.join(os.tmpdir(), 'uploads') : path.join(__dirname, '../uploads');
const outputDir = process.env.VERCEL ? path.join(os.tmpdir(), 'outputs') : path.join(__dirname, '../outputs');

if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
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

// Helper function to extract structured data from excel sheet row
function getRowValues(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  if (!row || !row.hasValues) return [];
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

// Generate Mock SAP Data for testing
function generateMockSAPData(table, offset, limit) {
  const list = [];
  const recordsToGenerate = Math.min(limit, 450000 - offset); // Cap mockup at 450k records
  
  for (let i = 0; i < recordsToGenerate; i++) {
    const idx = offset + i + 1;
    if (table === 'OCRD') {
      list.push({
        CardCode: `C${String(idx).padStart(5, '0')}`,
        CardName: `Acme Partner Corp ${idx}`,
        CardType: idx % 4 === 0 ? 'S' : 'C',
        Balance: (Math.random() * 5000).toFixed(2),
        E_Mail: `billing.partner${idx}@acme-corp.com`,
        Phone1: `555-${String(idx % 10000).padStart(4, '0')}`,
        CntctPrsn: `John Doe ${idx}`
      });
    } else if (table === 'OITM') {
      list.push({
        ItemCode: `P${String(idx).padStart(5, '0')}`,
        ItemName: `HP LaserJet Printer Scanner ${idx}`,
        OnHand: Math.floor(Math.random() * 100),
        Price: (Math.random() * 200).toFixed(2),
        BarCode: `718029${idx}`
      });
    } else {
      list.push({
        DocEntry: idx,
        DocNum: 1000 + idx,
        CardCode: `C${String(idx % 100).padStart(5, '0')}`,
        DocTotal: (Math.random() * 1500).toFixed(2),
        DocDate: new Date(Date.now() - (idx * 24 * 3600 * 1000)).toISOString().split('T')[0]
      });
    }
  }
  return list;
}

// SQL Server connection Configuration pool Builder
function buildSqlConfig(body) {
  return {
    user: body.username || process.env.DB_USER,
    password: body.password || process.env.DB_PASSWORD,
    server: body.server || process.env.DB_SERVER || 'localhost',
    database: body.companyDb || process.env.DB_DATABASE,
    port: parseInt(body.port || process.env.DB_PORT || '1433', 10),
    options: {
      encrypt: false,
      trustServerCertificate: true
    },
    connectionTimeout: 8000,
    requestTimeout: 15000
  };
}

// 1. CONNECTION ENDPOINT
app.post('/api/sap/connect', async (req, res) => {
  try {
    const { server, companyDb, username, password } = req.body;

    if (process.env.SAP_MOCK === 'true' || (!server && !companyDb)) {
      console.log('SAP Mock Connection: SUCCESS');
      return res.json({ status: 'connected', db: 'SBO_DemoDB (MOCK)' });
    }

    if (!server || !companyDb || !username) {
      return res.status(400).json({ error: 'Server URL, Company DB, and Username are required.' });
    }

    const config = buildSqlConfig(req.body);
    console.log(`Connecting directly to MS SQL Server: ${config.server}:${config.port}/${config.database}`);

    const pool = await sql.connect(config);
    await pool.request().query('SELECT 1 as heartbeat');
    
    return res.json({ status: 'connected', db: config.database });
  } catch (error) {
    console.error('SQL Connection Error:', error);
    return res.status(500).json({ error: `Connection failed: ${error.message}` });
  }
});

// 2. DISCOVER TABLES ENDPOINT
app.get('/api/sap/tables', (req, res) => {
  const { search } = req.query;

  const allTables = [
    { code: 'OCRD', name: 'Business Partners', count: 1248, fields: 7, desc: 'Customer and vendor account records, billing addresses, and terms.' },
    { code: 'OITM', name: 'Item Master', count: 10450, fields: 5, desc: 'Product database, SKUs, inventory status, and warehouse locations.' },
    { code: 'OINV', name: 'Sales Invoices', count: 45600, fields: 5, desc: 'Historical sales invoice transactions, totals, tax codes, and customer codes.' }
  ];

  if (search) {
    const filtered = allTables.filter(t => 
      t.code.toLowerCase().includes(search.toLowerCase()) || 
      t.name.toLowerCase().includes(search.toLowerCase())
    );
    return res.json(filtered);
  }

  return res.json(allTables);
});

// 3. SCHEMA COLUMNS DISCOVERY ENDPOINT
app.get('/api/sap/schema', async (req, res) => {
  try {
    const { table, server, companyDb } = req.query;

    if (!table) {
      return res.status(400).json({ error: 'Table code is required.' });
    }

    if (process.env.SAP_MOCK === 'true' || (!server && !companyDb)) {
      if (table === 'OCRD') return res.json(['CardCode', 'CardName', 'CardType', 'Balance', 'E_Mail', 'Phone1', 'CntctPrsn']);
      if (table === 'OITM') return res.json(['ItemCode', 'ItemName', 'OnHand', 'Price', 'BarCode']);
      return res.json(['DocEntry', 'DocNum', 'CardCode', 'DocTotal', 'DocDate']);
    }

    const config = buildSqlConfig(req.query);
    const pool = await sql.connect(config);
    
    const result = await pool.request()
      .input('tableName', sql.VarChar, table)
      .query(`
        SELECT COLUMN_NAME 
        FROM INFORMATION_SCHEMA.COLUMNS 
        WHERE TABLE_NAME = @tableName
      `);

    const columns = result.recordset.map(row => row.COLUMN_NAME);
    if (columns.length === 0) {
      // Fallback in case metadata tables differ in some SAP configurations
      if (table === 'OCRD') return res.json(['CardCode', 'CardName', 'CardType', 'Balance', 'E_Mail', 'Phone1', 'CntctPrsn']);
      if (table === 'OITM') return res.json(['ItemCode', 'ItemName', 'OnHand', 'Price', 'BarCode']);
      return res.json(['DocEntry', 'DocNum', 'CardCode', 'DocTotal', 'DocDate']);
    }

    return res.json(columns);
  } catch (error) {
    console.error('Schema Parsing Error:', error);
    return res.status(500).json({ error: error.message });
  }
});

// 4. STREAMING TRANSFER DATA ENDPOINT
app.post('/api/sap/transfer', upload.single('template'), async (req, res) => {
  let pool = null;
  try {
    const { table, server, companyDb, mappings, deduplicate } = req.body;
    const templateFile = req.file;

    if (!templateFile) return res.status(400).json({ error: 'Excel template file is required.' });
    
    const parsedMappings = JSON.parse(mappings || '[]');
    const activeMappings = parsedMappings.filter(m => m.from && m.to);

    if (activeMappings.length === 0) {
      return res.status(400).json({ error: 'Please configure at least one column mapping.' });
    }

    const isMock = process.env.SAP_MOCK === 'true' || (!server && !companyDb);
    const templatePath = templateFile.path;
    const outputFilename = `export-${Date.now()}-${templateFile.originalname}`;
    const outputFilePath = path.join(outputDir, outputFilename);

    console.log(`Executing SQL transfer (Mock=${isMock}) -> writing to: ${outputFilePath}`);

    // Load original template spreadsheet to capture headings and fonts
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.readFile(templatePath);
    const targetSheet = templateWorkbook.worksheets[0];
    const targetHeaders = getRowValues(targetSheet, 1);
    
    const toHeaderMap = {};
    targetHeaders.forEach((h, idx) => {
      if (h) toHeaderMap[h.trim().toLowerCase()] = idx + 1;
    });

    // Populate deduplication IDs from the template
    const seenIds = new Set();
    let targetIdColIdx = null;

    if (deduplicate === 'true') {
      activeMappings.forEach(m => {
        const toClean = m.to.trim().toLowerCase();
        if (toClean === 'customer id' || toClean === 'id' || toClean === 'client id' || toClean === 'identifier' || toClean === 'item code' || toClean === 'product id') {
          targetIdColIdx = toHeaderMap[toClean];
        }
      });

      if (targetIdColIdx) {
        targetSheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1) {
            const val = String(row.getCell(targetIdColIdx).value || '').trim();
            if (val) seenIds.add(val);
          }
        });
      }
    }

    // Initialize the ExcelJS Streaming Writer
    const options = {
      filename: outputFilePath,
      useStyles: true,
      useSharedStrings: true
    };
    const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter(options);
    const writeSheet = workbookWriter.addWorksheet(targetSheet.name);

    // Write original rows and headers to the stream
    targetSheet.eachRow((row, rowNum) => {
      const rowValues = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowValues.push(cell.value);
      });
      writeSheet.addRow(rowValues).commit();
    });

    // Establish DB Connection
    if (!isMock) {
      const config = buildSqlConfig(req.body);
      pool = await sql.connect(config);
    }

    let hasMoreData = true;
    let offset = 0;
    const limit = 10000;
    let totalImported = 0;
    let totalDuplicates = 0;

    const startTime = Date.now();

    // Whitelist query columns
    const selectColumns = activeMappings.map(m => `[${m.from}]`).join(', ');

    while (hasMoreData) {
      let records = [];

      if (isMock) {
        records = generateMockSAPData(table, offset, limit);
        hasMoreData = records.length === limit && offset < 50000; // Cap mockup at 50k for demo
      } else {
        const result = await pool.request()
          .input('limit', sql.Int, limit)
          .input('offset', sql.Int, offset)
          .query(`
            SELECT ${selectColumns} 
            FROM [${table}]
            ORDER BY (SELECT NULL)
            OFFSET @offset ROWS FETCH NEXT @limit ROWS ONLY
          `);
        records = result.recordset;
        hasMoreData = records.length === limit;
      }

      records.forEach(record => {
        // ID Deduplication Check
        if (targetIdColIdx) {
          const fromMapping = activeMappings.find(m => m.to.trim().toLowerCase() === Object.keys(toHeaderMap).find(k => toHeaderMap[k] === targetIdColIdx));
          if (fromMapping) {
            const recordId = String(record[fromMapping.from] || '').trim();
            if (recordId && seenIds.has(recordId)) {
              totalDuplicates++;
              return; // Skip duplicate row
            }
            if (recordId) seenIds.add(recordId);
          }
        }

        // Build values array
        const excelRowValues = new Array(targetHeaders.length + 1).fill('');
        activeMappings.forEach(m => {
          const toIdx = toHeaderMap[m.to.trim().toLowerCase()];
          if (toIdx) {
            excelRowValues[toIdx] = record[m.from] !== undefined ? String(record[m.from]) : '';
          }
        });

        excelRowValues.shift(); // remove index 0
        writeSheet.addRow(excelRowValues).commit();
        totalImported++;
      });

      offset += limit;
    }

    // Flush and commit file write
    await workbookWriter.commit();

    const stats = fs.statSync(outputFilePath);

    return res.json({
      message: 'Successfully transferred SAP data to template',
      downloadUrl: `/api/download/${outputFilename}`,
      summary: {
        totalImported,
        totalDuplicates,
        timeElapsed: ((Date.now() - startTime) / 1000).toFixed(1),
        fileSize: (stats.size / (1024 * 1024)).toFixed(1)
      }
    });

  } catch (error) {
    console.error('Transfer Error:', error);
    return res.status(500).json({ error: error.message || 'Transformation failed' });
  } finally {
    if (pool) {
      try {
        await pool.close();
      } catch (err) {
        console.error('Failed to close SQL connection pool:', err);
      }
    }
  }
});

// 5. SECURE DOWNLOAD ENDPOINT
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

  const filePath = path.join(outputDir, filename);
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
