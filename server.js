const express    = require('express');
const path       = require('path');
const JSZip      = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  convertInchesToTwip, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── In-Memory Data Store (replaces sql.js / SQLite) ──────────────────────────
// NOTE: Data resets on cold start (Vercel serverless limitation).
// For persistent storage, connect a cloud DB like Supabase or PlanetScale.
let nextId = 1;
let odEntries = [];
let settings  = { hod_name: 'HOD Name', hod_dept: 'Department' };

function now() {
  return new Date().toISOString().replace('T', ' ').substring(0, 19);
}

// ─── Express setup ─────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Settings API ──────────────────────────────────────────────────────────────
app.get('/api/settings', (req, res) => {
  res.json(settings);
});

app.post('/api/settings', (req, res) => {
  const { key, value } = req.body;
  if (!key) return res.status(400).json({ error: 'key required' });
  settings[key] = value;
  res.json({ success: true });
});

// ─── OD Entries API ────────────────────────────────────────────────────────────
app.get('/api/entries', (req, res) => {
  const { date, status, dept, search } = req.query;
  let rows = [...odEntries];

  if (date)   rows = rows.filter(r => r.od_date === date);
  if (status) rows = rows.filter(r => r.status === status);
  if (dept)   rows = rows.filter(r => r.department.toLowerCase().includes(dept.toLowerCase()));
  if (search) {
    const s = search.toLowerCase();
    rows = rows.filter(r =>
      r.student_name.toLowerCase().includes(s) ||
      r.vtu_id.toLowerCase().includes(s)
    );
  }

  // Sort: od_date DESC, id DESC
  rows.sort((a, b) => {
    if (b.od_date !== a.od_date) return b.od_date.localeCompare(a.od_date);
    return b.id - a.id;
  });

  res.json(rows);
});

app.post('/api/entries', (req, res) => {
  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;
  if (!student_name || !vtu_id || !department || !od_date || !reason)
    return res.status(400).json({ error: 'Missing required fields' });

  const entry = {
    id: nextId++,
    student_name,
    vtu_id,
    department,
    od_date,
    reason,
    status: status || 'Pending',
    notes: notes || '',
    created_at: now(),
    updated_at: now()
  };
  odEntries.push(entry);
  res.json({ id: entry.id, success: true });
});

app.put('/api/entries/:id', (req, res) => {
  const id  = parseInt(req.params.id);
  const idx = odEntries.findIndex(e => e.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Entry not found' });

  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;
  odEntries[idx] = {
    ...odEntries[idx],
    student_name, vtu_id, department, od_date, reason, status,
    notes: notes || '',
    updated_at: now()
  };
  res.json({ success: true });
});

app.delete('/api/entries/:id', (req, res) => {
  const id  = parseInt(req.params.id);
  const idx = odEntries.findIndex(e => e.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Entry not found' });
  odEntries.splice(idx, 1);
  res.json({ success: true });
});

app.get('/api/stats', (req, res) => {
  const total      = odEntries.length;
  const pending    = odEntries.filter(e => e.status === 'Pending').length;
  const processed  = odEntries.filter(e => e.status === 'Processed').length;
  const today      = new Date().toISOString().split('T')[0];
  const todayCount = odEntries.filter(e => e.od_date === today).length;
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

  const reasons    = [...new Set(entries.map(e => e.reason))];
  const reasonText = reasons.join('; ');
  const depts      = [...new Set(entries.map(e => e.department))];
  const deptText   = depts.join(' / ');

  const thinBorder = {
    top:    { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    left:   { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    right:  { style: BorderStyle.SINGLE, size: 1, color: '000000' }
  };

  const headers   = ['S.No', 'Student Name', 'VTU ID', 'Department', 'Reason for OD'];
  const colWidths = [600, 2200, 1800, 2000, 3400];

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

  const dataRows = entries.map((e, idx) =>
    new TableRow({
      children: [
        new TableCell({ width: { size: colWidths[0], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(String(idx + 1))], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[1], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.student_name)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[2], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.vtu_id)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[3], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.department)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[4], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.reason)], spacing: { before: 60, after: 60 } })] })
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
        para([tnr('From,', { bold: true })]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(),
        para([tnr('To,', { bold: true })]),
        para([tnr('The Dean')]),
        para([tnr('School of Computing')]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) for ${odDateStr} — ${entries.length} Student${entries.length > 1 ? 's' : ''}`, { bold: true })
        ]),
        blank(),
        para([tnr('Respected Sir/Madam,')]),
        blank(),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, children: [tnr(bodyIntro)] }),
        blank(),
        studentTable,
        blank(),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, children: [tnr(bodyClose)] }),
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
        para([tnr('From,', { bold: true })]),
        para([tnr(entry.student_name)]),
        para([tnr(entry.department)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(),
        para([tnr('To,', { bold: true })]),
        para([tnr('The Dean')]),
        para([tnr('School of Computing')]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        blank(),
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) Letter for ${odDateStr}`, { bold: true })
        ]),
        blank(),
        para([tnr('Respected Sir/Madam,')]),
        blank(),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, children: [tnr(bodyText)] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, children: [tnr(closingText)] }),
        blank(),
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
  const entry = odEntries.find(e => e.id === parseInt(id));
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
    const entry = odEntries.find(e => e.id === parseInt(id));
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

app.post('/api/download-combined', async (req, res) => {
  const { ids, letterDate } = req.body;
  if (!ids || !ids.length) return res.status(400).json({ error: 'No entries selected' });

  const entries = ids
    .map(id => odEntries.find(e => e.id === parseInt(id)))
    .filter(Boolean);

  if (!entries.length) return res.status(404).json({ error: 'No valid entries found' });

  const byDate = {};
  entries.forEach(e => {
    if (!byDate[e.od_date]) byDate[e.od_date] = [];
    byDate[e.od_date].push(e);
  });

  const dates = Object.keys(byDate).sort();

  if (dates.length === 1) {
    const doc    = buildCombinedDocx(byDate[dates[0]], letterDate);
    const buffer = await Packer.toBuffer(doc);
    const fname  = `OD_Combined_${dates[0]}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } else {
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
if (!process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`\n✅  OD Letter Processing System`);
    console.log(`🌐  http://localhost:${PORT}\n`);
  });
}

// Export for Vercel serverless runtime
module.exports = app;
