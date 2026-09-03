const express    = require('express');
const path       = require('path');
const fs         = require('fs');
const initSqlJs  = require('sql.js');
const JSZip      = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  convertInchesToTwip, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;
// Vercel's filesystem is read-only except /tmp
const DB_FILE = process.env.VERCEL
  ? '/tmp/od_records.db'
  : path.join(__dirname, 'od_records.db');

// ─── sql.js setup ──────────────────────────────────────────────────────────────
let db;

async function initDB() {
  const SQL = await initSqlJs({
    // In Vercel serverless, local .wasm lookup fails — load from CDN instead
    locateFile: file => `https://sql.js.org/dist/${file}`
  });

  if (fs.existsSync(DB_FILE)) {
    const fileBuffer = fs.readFileSync(DB_FILE);
    db = new SQL.Database(fileBuffer);
  } else {
    db = new SQL.Database();
  }

  db.run(`
    CREATE TABLE IF NOT EXISTS od_entries (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      student_name TEXT NOT NULL,
      vtu_id       TEXT NOT NULL,
      department   TEXT NOT NULL,
      od_date      TEXT NOT NULL,
      reason       TEXT NOT NULL,
      status       TEXT NOT NULL DEFAULT 'Pending',
      notes        TEXT DEFAULT '',
      created_at   TEXT DEFAULT (datetime('now')),
      updated_at   TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS settings (
      key   TEXT PRIMARY KEY,
      value TEXT
    );
    INSERT OR IGNORE INTO settings (key, value) VALUES ('hod_name', 'HOD Name');
    INSERT OR IGNORE INTO settings (key, value) VALUES ('hod_dept', 'Department');
  `);

  saveDB();
}

function saveDB() {
  const data = db.export();
  fs.writeFileSync(DB_FILE, Buffer.from(data));
}

// Helper: run a SELECT and return array of row objects
function queryAll(sql, params = []) {
  const stmt = db.prepare(sql);
  stmt.bind(params);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
}

function queryOne(sql, params = []) {
  const rows = queryAll(sql, params);
  return rows[0] || null;
}

// ─── Express setup ─────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Settings API ──────────────────────────────────────────────────────────────
app.get('/api/settings', (req, res) => {
  const rows = queryAll('SELECT key, value FROM settings');
  const s = {};
  rows.forEach(r => s[r.key] = r.value);
  res.json(s);
});

app.post('/api/settings', (req, res) => {
  const { key, value } = req.body;
  if (!key) return res.status(400).json({ error: 'key required' });
  db.run('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)', [key, value]);
  saveDB();
  res.json({ success: true });
});

// ─── OD Entries API ────────────────────────────────────────────────────────────
app.get('/api/entries', (req, res) => {
  const { date, status, dept, search } = req.query;
  let sql = 'SELECT * FROM od_entries WHERE 1=1';
  const params = [];

  if (date)   { sql += ' AND od_date = ?';                  params.push(date); }
  if (status) { sql += ' AND status = ?';                   params.push(status); }
  if (dept)   { sql += ' AND department LIKE ?';            params.push('%' + dept + '%'); }
  if (search) {
    sql += ' AND (student_name LIKE ? OR vtu_id LIKE ?)';
    params.push('%' + search + '%', '%' + search + '%');
  }
  sql += ' ORDER BY od_date DESC, id DESC';

  const rows = queryAll(sql, params);
  res.json(rows);
});

app.post('/api/entries', (req, res) => {
  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;
  if (!student_name || !vtu_id || !department || !od_date || !reason)
    return res.status(400).json({ error: 'Missing required fields' });

  db.run(
    `INSERT INTO od_entries (student_name, vtu_id, department, od_date, reason, status, notes)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [student_name, vtu_id, department, od_date, reason, status || 'Pending', notes || '']
  );
  saveDB();
  const row = queryOne('SELECT last_insert_rowid() as id');
  res.json({ id: row ? row.id : null, success: true });
});

app.put('/api/entries/:id', (req, res) => {
  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;
  const id = parseInt(req.params.id);
  db.run(
    `UPDATE od_entries
     SET student_name=?, vtu_id=?, department=?, od_date=?, reason=?, status=?, notes=?,
         updated_at=datetime('now')
     WHERE id=?`,
    [student_name, vtu_id, department, od_date, reason, status, notes || '', id]
  );
  saveDB();
  res.json({ success: true });
});

app.delete('/api/entries/:id', (req, res) => {
  db.run('DELETE FROM od_entries WHERE id = ?', [parseInt(req.params.id)]);
  saveDB();
  res.json({ success: true });
});

app.get('/api/stats', (req, res) => {
  const total     = queryOne('SELECT COUNT(*) as c FROM od_entries').c || 0;
  const pending   = queryOne("SELECT COUNT(*) as c FROM od_entries WHERE status='Pending'").c || 0;
  const processed = queryOne("SELECT COUNT(*) as c FROM od_entries WHERE status='Processed'").c || 0;
  const today     = new Date().toISOString().split('T')[0];
  const todayCount= queryOne('SELECT COUNT(*) as c FROM od_entries WHERE od_date=?', [today]).c || 0;
  res.json({ total, pending, processed, todayCount });
});

// ─── DOCX Builder ──────────────────────────────────────────────────────────────
function tnr(text, opts = {}) {
  return new TextRun({ text, font: 'Times New Roman', size: 24, ...opts });
}

function para(children, alignment = AlignmentType.LEFT, afterSpacing = 160) {
  return new Paragraph({ alignment, spacing: { after: afterSpacing }, children });
}

function blank() {
  return new Paragraph({ children: [tnr('')], spacing: { after: 120 } });
}

function fmtDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });
}

// ─── Combined DOCX: one letter for multiple students on one date ───────────────
function buildCombinedDocx(entries, letterDate) {
  if (!entries || !entries.length) throw new Error('No entries');

  const odDate        = entries[0].od_date;
  const odDateStr     = fmtDate(odDate);
  const letterDateStr = letterDate
    ? fmtDate(letterDate)
    : new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });

  // Collect unique reasons
  const reasons = [...new Set(entries.map(e => e.reason))];
  const reasonText = reasons.join('; ');

  // Collect departments for "From" section
  const depts = [...new Set(entries.map(e => e.department))];
  const deptText = depts.join(' / ');

  // ── Table border helper ──
  const thinBorder = {
    top:    { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    left:   { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    right:  { style: BorderStyle.SINGLE, size: 1, color: '000000' }
  };

  // ── Header row ──
  const headers = ['S.No', 'Student Name', 'VTU ID', 'Department', 'Reason for OD'];
  const colWidths = [600, 2200, 1800, 2000, 3400]; // twips

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      new TableCell({
        width:   { size: colWidths[i], type: WidthType.DXA },
        borders: thinBorder,
        shading: { fill: 'D3D3D3', type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [tnr(h, { bold: true })],
          spacing: { before: 60, after: 60 }
        })]
      })
    )
  });

  // ── Data rows ──
  const dataRows = entries.map((e, idx) =>
    new TableRow({
      children: [
        new TableCell({
          width: { size: colWidths[0], type: WidthType.DXA },
          borders: thinBorder,
          children: [new Paragraph({ children: [tnr(String(idx + 1))], spacing: { before: 60, after: 60 } })]
        }),
        new TableCell({
          width: { size: colWidths[1], type: WidthType.DXA },
          borders: thinBorder,
          children: [new Paragraph({ children: [tnr(e.student_name)], spacing: { before: 60, after: 60 } })]
        }),
        new TableCell({
          width: { size: colWidths[2], type: WidthType.DXA },
          borders: thinBorder,
          children: [new Paragraph({ children: [tnr(e.vtu_id)], spacing: { before: 60, after: 60 } })]
        }),
        new TableCell({
          width: { size: colWidths[3], type: WidthType.DXA },
          borders: thinBorder,
          children: [new Paragraph({ children: [tnr(e.department)], spacing: { before: 60, after: 60 } })]
        }),
        new TableCell({
          width: { size: colWidths[4], type: WidthType.DXA },
          borders: thinBorder,
          children: [new Paragraph({ children: [tnr(e.reason)], spacing: { before: 60, after: 60 } })]
        })
      ]
    })
  );

  const studentTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows]
  });

  const bodyIntro =
    `The following ${entries.length} student${entries.length > 1 ? 's' : ''} of ` +
    `${deptText}, Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology, ` +
    `were engaged in ${reasonText} on ${odDateStr} and are hereby requested to be granted On Duty (OD) for the same.`;

  const bodyClose =
    `We kindly request you to consider this application and issue OD letters for the ` +
    `aforementioned students for ${odDateStr}.`;

  return new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top:    convertInchesToTwip(1),
            right:  convertInchesToTwip(1),
            bottom: convertInchesToTwip(1),
            left:   convertInchesToTwip(1.25)
          }
        }
      },
      children: [
        // FROM
        para([tnr('From,', { bold: true })]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        // DATE
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(),
        // TO
        para([tnr('To,', { bold: true })]),
        para([tnr('The Dean')]),
        para([tnr('School of Computing')]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        // SUBJECT
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) for ${odDateStr} — ${entries.length} Student${entries.length > 1 ? 's' : ''}`, { bold: true })
        ]),
        blank(),
        // SALUTATION
        para([tnr('Respected Sir/Madam,')]),
        blank(),
        // BODY
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 200 },
          children: [tnr(bodyIntro)]
        }),
        blank(),
        // TABLE
        studentTable,
        blank(),
        // CLOSE
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 200 },
          children: [tnr(bodyClose)]
        }),
        blank(),
        para([tnr('Thanking you,')]),
        blank(),
        para([tnr('Yours sincerely,')]),
        blank(),
        para([tnr('(Authorised Signatory)')]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')])
      ]
    }]
  });
}

function buildDocx(entry, letterDate) {
  const odDateStr     = fmtDate(entry.od_date);
  const letterDateStr = letterDate
    ? fmtDate(letterDate)
    : new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });

  const bodyText =
    `I, ${entry.student_name} (VTU ID: ${entry.vtu_id}), a student of ${entry.department}, ` +
    `humbly request you to kindly grant me an On Duty (OD) for ${odDateStr}, ` +
    `as I was engaged in ${entry.reason}.`;

  const closingText =
    `I kindly request you to consider my application and issue an OD letter for the ` +
    `aforementioned date. I assure you that I have fulfilled all the necessary requirements for the same.`;

  return new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top:    convertInchesToTwip(1),
            right:  convertInchesToTwip(1),
            bottom: convertInchesToTwip(1),
            left:   convertInchesToTwip(1.25)
          }
        }
      },
      children: [
        // FROM
        para([tnr('From,', { bold: true })]),
        para([tnr(entry.student_name)]),
        para([tnr(entry.department)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        // DATE
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(),
        // TO
        para([tnr('To,', { bold: true })]),
        para([tnr('The Dean')]),
        para([tnr('School of Computing')]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        // SUBJECT
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) Letter for ${odDateStr}`, { bold: true })
        ]),
        blank(),
        // SALUTATION
        para([tnr('Respected Sir/Madam,')]),
        blank(),
        // BODY paragraph 1
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 200 },
          children: [tnr(bodyText)]
        }),
        // BODY paragraph 2
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 200 },
          children: [tnr(closingText)]
        }),
        blank(),
        // CLOSING
        para([tnr('Thanking you,')]),
        blank(),
        para([tnr('Yours sincerely,')]),
        para([tnr(entry.student_name)]),
        para([tnr(entry.vtu_id)]),
        para([tnr(entry.department)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')])
      ]
    }]
  });
}

// ─── Download routes ───────────────────────────────────────────────────────────
app.post('/api/download', async (req, res) => {
  const { id, letterDate } = req.body;
  const entry = queryOne('SELECT * FROM od_entries WHERE id = ?', [id]);
  if (!entry) return res.status(404).json({ error: 'Entry not found' });

  const doc    = buildDocx(entry, letterDate);
  const buffer = await Packer.toBuffer(doc);
  const fname  = `OD_${entry.student_name.replace(/\s+/g, '_')}_${entry.od_date}.docx`;
  res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.send(buffer);
});

app.post('/api/download-bulk', async (req, res) => {
  const { ids, letterDate } = req.body;
  if (!ids || !ids.length) return res.status(400).json({ error: 'No entries selected' });

  const zip = new JSZip();
  for (const id of ids) {
    const entry = queryOne('SELECT * FROM od_entries WHERE id = ?', [id]);
    if (!entry) continue;
    const doc    = buildDocx(entry, letterDate);
    const buffer = await Packer.toBuffer(doc);
    const fname  = `OD_${entry.student_name.replace(/\s+/g, '_')}_${entry.od_date}.docx`;
    zip.file(fname, buffer);
  }

  const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
  res.setHeader('Content-Disposition', 'attachment; filename="OD_Letters.zip"');
  res.setHeader('Content-Type', 'application/zip');
  res.send(zipBuf);
});

// Combined letter for multiple students (same or mixed dates) — one DOCX with a table
app.post('/api/download-combined', async (req, res) => {
  const { ids, letterDate } = req.body;
  if (!ids || !ids.length) return res.status(400).json({ error: 'No entries selected' });

  const entries = ids
    .map(id => queryOne('SELECT * FROM od_entries WHERE id = ?', [id]))
    .filter(Boolean);

  if (!entries.length) return res.status(404).json({ error: 'No valid entries found' });

  // Group by date: if all same date → one file; if multiple dates → one file per date in ZIP
  const byDate = {};
  entries.forEach(e => {
    if (!byDate[e.od_date]) byDate[e.od_date] = [];
    byDate[e.od_date].push(e);
  });

  const dates = Object.keys(byDate).sort();

  if (dates.length === 1) {
    // Single date → single combined DOCX
    const doc    = buildCombinedDocx(byDate[dates[0]], letterDate);
    const buffer = await Packer.toBuffer(doc);
    const fname  = `OD_Combined_${dates[0]}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } else {
    // Multiple dates → one combined DOCX per date, all in a ZIP
    const zip = new JSZip();
    for (const date of dates) {
      const doc    = buildCombinedDocx(byDate[date], letterDate);
      const buffer = await Packer.toBuffer(doc);
      zip.file(`OD_Combined_${date}.docx`, buffer);
    }
    const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
    res.setHeader('Content-Disposition', 'attachment; filename="OD_Combined_Letters.zip"');
    res.setHeader('Content-Type', 'application/zip');
    res.send(zipBuf);
  }
});

// ─── Start ─────────────────────────────────────────────────────────────────────
initDB().then(() => {
  if (!process.env.VERCEL) {
    // Local dev: start HTTP server
    app.listen(PORT, () => {
      console.log(`\n✅  OD Letter Processing System`);
      console.log(`🌐  http://localhost:${PORT}\n`);
    });
  }
}).catch(err => {
  console.error('Failed to initialise database:', err);
  if (!process.env.VERCEL) process.exit(1);
});

// Export for Vercel serverless runtime
module.exports = app;
