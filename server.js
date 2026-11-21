// server.js - Integrated Backend & REST API Server for BDO Bani Portal
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const whatsappService = require('./whatsapp_service');

const PORT = process.env.PORT || 3000;
const DB_PATH = path.join(__dirname, 'data', 'database.json');
const DB_JS_PATH = path.join(__dirname, 'data', 'database.js');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function readDatabase() {
  if (fs.existsSync(DB_PATH)) {
    return JSON.parse(fs.readFileSync(DB_PATH, 'utf8'));
  }
  return {};
}

function writeDatabase(db) {
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf8');
  // Also regenerate database.js for standalone offline opening
  const jsContent = `/**
 * BDO BANI DISTRICT CAPEX & MLA CDF MASTER DATA
 * Rural Development Department, Block Bani, District Kathua, J&K
 */
window.BDO_INITIAL_DATA = ${JSON.stringify(db, null, 2)};
`;
  fs.writeFileSync(DB_JS_PATH, jsContent, 'utf8');
}

// Fetch live Kathua Google Sheet (Read-Only)
function fetchKathuaGoogleSheet() {
  return new Promise((resolve, reject) => {
    const url = 'https://docs.google.com/spreadsheets/d/1RaKzdz9wSCQaQdtjxIsPcZO2e4Pw9Od2USq39KQWeZg/gviz/tq?tqx=out:csv&gid=2114934525';
    https.get(url, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    }).on('error', reject);
  });
}

function parseCSVLine(line) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      result.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  result.push(cur);
  return result;
}

const server = http.createServer(async (req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const [reqPath, queryStr] = req.url.split('?');

  // API ROUTE 1: GET /api/data
  if (reqPath === '/api/data' && req.method === 'GET') {
    try {
      const db = readDatabase();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data: db }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // API ROUTE 2: POST /api/save
  if (reqPath === '/api/save' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body);
        writeDatabase(payload);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Database saved successfully!' }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: e.message }));
      }
    });
    return;
  }

  // API ROUTE 3: GET /api/kathua-sync (Analyze & Map BDO Bani row in Google Sheet)
  if (reqPath === '/api/kathua-sync' && req.method === 'GET') {
    try {
      const rawCSV = await fetchKathuaGoogleSheet();
      const lines = rawCSV.split('\n').filter(l => l.trim().length > 0);
      const headers = parseCSVLine(lines[0]);

      let baniRowIndex = -1;
      let baniRowData = null;

      for (let i = 1; i < lines.length; i++) {
        const cols = parseCSVLine(lines[i]);
        if (cols.some(c => c.toLowerCase().includes('bdo-bani') || c.toLowerCase().includes('bani'))) {
          baniRowIndex = i + 1; // 1-based spreadsheet row
          baniRowData = cols;
          break;
        }
      }

      // Calculate figures from current local database
      const db = readDatabase();
      const priNew = db.priNew || [];
      const ddc = db.ddcWorks || [];
      const bdc = db.bdcWorks || [];
      const apdp = db.apdpWorks || [];
      const st = db.stWorks || [];

      // Capex New Works Allotted Count & Cost
      const totalAllottedWorks = priNew.length + ddc.length + bdc.length + apdp.length; // 78 + 17 + 1 + 1 = 97
      const totalAllottedCost = (
        priNew.reduce((s, w) => s + (w.allotmentCost || 0), 0) +
        ddc.reduce((s, w) => s + (w.allotmentCost || 0), 0) +
        bdc.reduce((s, w) => s + (w.allotmentCost || 0), 0) +
        apdp.reduce((s, w) => s + (w.allotmentCost || 0), 0)
      ); // 161.085 + 44.18 + 14.165 + 3.82 = 223.25 Lakhs!

      const priCompleted = priNew.filter(w => (w.physicalStatus || '').toLowerCase().includes('comp')).length;
      const ddcCompleted = ddc.filter(w => (w.physicalStatus || '').toLowerCase().includes('comp')).length;
      const totalCompletedWorks = priCompleted + ddcCompleted;

      const responsePayload = {
        success: true,
        googleSheet: {
          sheetName: "DAILY MONITORING FORMAT — DISTRICT CAPEX BUDGET 2026-27",
          baniRowNumber: baniRowIndex,
          headers: headers,
          currentRowData: baniRowData
        },
        localExcelCalculated: {
          department: "BDO-Bani",
          component: "DDC/BDC/PRI/ABDP/APDP/Additionality",
          allottedWorksCount: totalAllottedWorks,
          completedWorksCount: totalCompletedWorks,
          allottedAmountLakh: totalAllottedCost.toFixed(2),
          remarks: `${priCompleted} works under PRI and ${ddcCompleted} works under DDC physically Completed`
        }
      };

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(responsePayload, null, 2));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // API ROUTE 4: GET /api/whatsapp/status & /api/whatsapp/qr
  if (reqPath === '/api/whatsapp/status' || reqPath === '/api/whatsapp/qr') {
    try {
      const status = await whatsappService.getStatus();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, ...status }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // API ROUTE 5: POST /api/whatsapp/send-otp
  if (reqPath === '/api/whatsapp/send-otp' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body || '{}');
        const result = await whatsappService.sendOtp(payload.phone, payload.otp, payload.user);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: e.message }));
      }
    });
    return;
  }

  // API ROUTE 6: GET /api/daily-bani-sheet (Live Fetch Daily Bani Google Sheet)
  if (reqPath === '/api/daily-bani-sheet' && req.method === 'GET') {
    try {
      const gUrl = 'https://docs.google.com/spreadsheets/d/10vQXS5D36tGuvaBNEpe4xYBRs6RKUqv6BHq_eboosNQ/export?format=xlsx';
      const XLSX = require('xlsx');

      function fetchXLSXBuffer(targetUrl) {
        return new Promise((resolve, reject) => {
          https.get(targetUrl, res => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
              return fetchXLSXBuffer(res.headers.location).then(resolve).catch(reject);
            }
            const chunks = [];
            res.on('data', c => chunks.push(c));
            res.on('end', () => resolve(Buffer.concat(chunks)));
          }).on('error', reject);
        });
      }

      const buf = await fetchXLSXBuffer(gUrl);
      const wb = XLSX.read(buf, { type: 'buffer' });
      const finRows = XLSX.utils.sheet_to_json(wb.Sheets['Financial-2026-27'], { header: 1 });
      const physRows = XLSX.utils.sheet_to_json(wb.Sheets['Physical-2026-27'], { header: 1 });

      res.writeHead(200, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      });
      res.end(JSON.stringify({ success: true, finRows, physRows }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // API ROUTE 6: POST /api/whatsapp/disconnect
  if (reqPath === '/api/whatsapp/disconnect' && req.method === 'POST') {
    try {
      const result = await whatsappService.disconnect();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // API ROUTE 7: POST /api/activity-logs (Permanent Server-Side Audit Log)
  if (reqPath === '/api/activity-logs' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        const db = readDatabase();
        if (!Array.isArray(db.activityLogs)) db.activityLogs = [];
        if (payload.log) {
          db.activityLogs.unshift(payload.log);
          if (db.activityLogs.length > 2500) db.activityLogs.pop(); // keep last 2500 logs on disk
          writeDatabase(db);
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, count: db.activityLogs.length }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: e.message }));
      }
    });
    return;
  }

  // API ROUTE 8: GET /api/activity-logs
  if (reqPath === '/api/activity-logs' && req.method === 'GET') {
    try {
      const db = readDatabase();
      const logs = Array.isArray(db.activityLogs) ? db.activityLogs : [];
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, logs: logs }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: e.message }));
    }
    return;
  }

  // STATIC FILE HANDLER
  let filePath = path.join(__dirname, reqPath === '/' ? 'index.html' : reqPath);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, content) => {
    if (err) {
      if (err.code === 'ENOENT') {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('404 Not Found');
      } else {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end(`Server Error: ${err.code}`);
      }
    } else {
      res.writeHead(200, { 
        'Content-Type': contentType,
        'Cache-Control': 'no-cache'
      });
      res.end(content);
    }
  });
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    /* console.log(`[Server] Port ${PORT} already active, using existing server process.`); */
  } else {
    console.error('[Server] Error:', err);
  }
});

server.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  /* console.log(`=======================================================`); */
  /* console.log(`  🚀 BDO BANI INTEGRATED WORKS PORTAL & BACKEND ACTIVE`); */
  /* console.log(`  🌐 Server running at: ${url}`); */
  /* console.log(`  📡 REST API available at: ${url}/api/data & ${url}/api/kathua-sync`); */
  /* console.log(`=======================================================`); */
});
