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
// ─── DOCX Builder ──────────────────────────────────────────────────────────────
function tnr(text, opts = {}) {
  return new TextRun({ text: text || '', font: 'Times New Roman', size: 24, ...opts });
}

function para(children, alignment = AlignmentType.LEFT, afterSpacing = 140) {
  return new Paragraph({ alignment, spacing: { after: afterSpacing, line: 276 }, children });
}

function blank(spacing = 100) {
  return new Paragraph({ children: [tnr('')], spacing: { after: spacing } });
}

function fmtDate(dateStr) {
  if (!dateStr) return '';
  // If already formatted or text range
  if (typeof dateStr === 'string' && (dateStr.includes(' ') || dateStr.includes('-') === false)) {
    return dateStr;
  }
  try {
    const d = new Date(dateStr + (dateStr.length === 10 ? 'T00:00:00' : ''));
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });
  } catch (e) {
    return dateStr;
  }
}

// ─── Combined DOCX: Professional Letter for multiple students ─────────────────
function buildCombinedDocx(entries, letterDate, meta = {}) {
  if (!entries || !entries.length) throw new Error('No entries');

  const odDatesList = [...new Set(entries.map(e => e.od_date))];
  const defaultOdDateStr = odDatesList.map(d => fmtDate(d)).join(', ');
  const odDatesStr = meta.odDates ? meta.odDates : defaultOdDateStr;

  const letterDateStr = letterDate
    ? fmtDate(letterDate)
    : new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });

  const depts = [...new Set(entries.map(e => e.department))];
  const deptText = meta.fromDept || depts.join(' / ') || 'School of Computing';

  const fromName = meta.fromName || 'Head of Department / Faculty Coordinator';
  const designation = meta.fromDesignation || 'Authorised Signatory';
  const eventName = meta.eventName || 'the Scheduled Academic / Co-Curricular Event';
  const eventDates = meta.eventDates ? meta.eventDates : odDatesStr;
  const commonReason = meta.reason || [...new Set(entries.map(e => e.reason))].join('; ');
  const toAuthority = meta.toAuthority || 'The Dean';
  const toSchool = meta.toSchool || 'School of Computing';

  const thinBorder = {
    top:    { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    left:   { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    right:  { style: BorderStyle.SINGLE, size: 1, color: '333333' }
  };

  const headers = ['S.No', 'Student Name', 'VTU ID', 'Department', 'Reason / Activity', 'OD Date(s)'];
  const colWidths = [600, 2400, 1800, 2000, 2600, 1600];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      new TableCell({
        width: { size: colWidths[i], type: WidthType.DXA },
        borders: thinBorder,
        shading: { fill: 'EAEAEA', type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [tnr(h, { bold: true })],
          spacing: { before: 80, after: 80 }
        })]
      })
    )
  });

  const dataRows = entries.map((e, idx) =>
    new TableRow({
      children: [
        new TableCell({ width: { size: colWidths[0], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(String(idx + 1))], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[1], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.student_name, { bold: true })], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[2], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.vtu_id)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[3], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.department)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[4], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.reason || commonReason)], spacing: { before: 60, after: 60 } })] }),
        new TableCell({ width: { size: colWidths[5], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(fmtDate(e.od_date))], spacing: { before: 60, after: 60 } })] })
      ]
    })
  );

  const studentTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows]
  });

  const bodyIntro =
    `With reference to the subject cited above, we bring to your kind notice that the following ` +
    `${entries.length} student${entries.length > 1 ? 's' : ''} of ${deptText}, ` +
    `Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology, ` +
    `participated in "${eventName}" scheduled from ${eventDates}. ` +
    `The primary purpose and reason for On Duty (OD) is "${commonReason}".`;

  const bodyDetails =
    `The student${entries.length > 1 ? 's were' : ' was'} actively on duty during the period: ${odDatesStr}. ` +
    `Due to their official representation and active participation in the event, ` +
    `they could not attend the regularly scheduled academic classes and laboratory sessions during this duration.`;

  const bodyClose =
    `In view of the above, we kindly request your good office to grant On Duty (OD) attendance ` +
    `for the aforementioned student${entries.length > 1 ? 's' : ''} for ${odDatesStr} and enable the appropriate attendance recording in the academic registry.`;

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
        para([tnr(fromName, { bold: true })]),
        para([tnr(designation)]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        para([tnr('Avadi, Chennai - 600 062, Tamil Nadu')]),
        blank(80),
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(80),
        para([tnr('To,', { bold: true })]),
        para([tnr(toAuthority)]),
        para([tnr(toSchool)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        para([tnr('Avadi, Chennai - 600 062, Tamil Nadu')]),
        blank(100),
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) Permission — "${eventName}" — ${odDatesStr} — Reg.`, { bold: true })
        ]),
        blank(80),
        para([tnr('Respected Sir/Madam,')]),
        blank(80),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160, line: 276 }, children: [tnr(bodyIntro)] }),
        blank(60),
        studentTable,
        blank(80),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 }, children: [tnr(bodyDetails)] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 180, line: 276 }, children: [tnr(bodyClose)] }),
        blank(80),
        para([tnr('Thanking you,')]),
        blank(120),
        para([tnr('Yours faithfully,')]),
        blank(200),
        para([tnr(fromName, { bold: true })]),
        para([tnr(designation)]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')])
      ]
    }]
  });
}

// ─── Single Student DOCX ───────────────────────────────────────────────────────
function buildDocx(entry, letterDate, meta = {}) {
  const odDateStr = meta.odDates ? meta.odDates : fmtDate(entry.od_date);
  const letterDateStr = letterDate
    ? fmtDate(letterDate)
    : new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });

  const fromName = meta.fromName || entry.student_name;
  const designation = meta.fromDesignation || `Student (VTU ID: ${entry.vtu_id})`;
  const deptText = meta.fromDept || entry.department;
  const eventName = meta.eventName || 'the Scheduled Academic / Co-Curricular Event';
  const eventDates = meta.eventDates ? meta.eventDates : odDateStr;
  const reasonText = meta.reason || entry.reason;
  const toAuthority = meta.toAuthority || 'The Dean';
  const toSchool = meta.toSchool || 'School of Computing';

  const thinBorder = {
    top:    { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    left:   { style: BorderStyle.SINGLE, size: 1, color: '333333' },
    right:  { style: BorderStyle.SINGLE, size: 1, color: '333333' }
  };

  const studentTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ['Student Name', 'VTU ID', 'Department', 'OD Date(s)'].map(h =>
          new TableCell({
            borders: thinBorder,
            shading: { fill: 'EAEAEA', type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [tnr(h, { bold: true })], spacing: { before: 60, after: 60 } })]
          })
        )
      }),
      new TableRow({
        children: [
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.student_name, { bold: true })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.vtu_id)], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.department)], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(odDateStr)], spacing: { before: 60, after: 60 } })] })
        ]
      })
    ]
  });

  const bodyIntro =
    `With reference to the subject cited above, I wish to submit that I, ${entry.student_name} ` +
    `(VTU ID: ${entry.vtu_id}), a bona fide student of ${entry.department}, ` +
    `participated in "${eventName}" scheduled from ${eventDates}. ` +
    `The reason and purpose for On Duty (OD) was "${reasonText}".`;

  const bodyDetails =
    `I was officially engaged in duty for this event during: ${odDateStr}. ` +
    `On account of this official representation, I could not attend regular instructional lectures and practical laboratory sessions on the said date(s).`;

  const closingText =
    `I therefore kindly request your good office to grant On Duty (OD) attendance for ${odDateStr} ` +
    `and facilitate regularisation of my attendance. I have completed all prerequisites and documentation for the same.`;

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
        para([tnr(fromName, { bold: true })]),
        para([tnr(designation)]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        para([tnr('Avadi, Chennai - 600 062, Tamil Nadu')]),
        blank(80),
        para([tnr('Date: ', { bold: true }), tnr(letterDateStr)]),
        blank(80),
        para([tnr('To,', { bold: true })]),
        para([tnr(toAuthority)]),
        para([tnr(toSchool)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')]),
        para([tnr('Avadi, Chennai - 600 062, Tamil Nadu')]),
        blank(100),
        para([
          tnr('Subject: ', { bold: true }),
          tnr(`Request for On Duty (OD) Letter — "${eventName}" — ${odDateStr} — Reg.`, { bold: true })
        ]),
        blank(80),
        para([tnr('Respected Sir/Madam,')]),
        blank(80),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160, line: 276 }, children: [tnr(bodyIntro)] }),
        blank(60),
        studentTable,
        blank(80),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 }, children: [tnr(bodyDetails)] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 180, line: 276 }, children: [tnr(closingText)] }),
        blank(80),
        para([tnr('Thanking you,')]),
        blank(120),
        para([tnr('Yours faithfully,')]),
        blank(200),
        para([tnr(fromName, { bold: true })]),
        para([tnr(designation)]),
        para([tnr(deptText)]),
        para([tnr('Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology')])
      ]
    }]
  });
}

// ─── Download routes ───────────────────────────────────────────────────────────
app.post('/api/download', async (req, res) => {
  const { id, letterDate, meta } = req.body;
  const entry = odEntries.find(e => e.id === parseInt(id));
  if (!entry) return res.status(404).json({ error: 'Entry not found' });

  const doc    = buildDocx(entry, letterDate, meta || {});
  const buffer = await Packer.toBuffer(doc);
  const fname  = `OD_${entry.student_name.replace(/\s+/g, '_')}_${entry.od_date}.docx`;
  res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.send(buffer);
});

app.post('/api/download-bulk', async (req, res) => {
  const { ids, letterDate, meta } = req.body;
  if (!ids || !ids.length) return res.status(400).json({ error: 'No entries selected' });

  const zip = new JSZip();
  for (const id of ids) {
    const entry = odEntries.find(e => e.id === parseInt(id));
    if (!entry) continue;
    const doc    = buildDocx(entry, letterDate, meta || {});
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
  const { ids, letterDate, meta } = req.body;
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
    const doc    = buildCombinedDocx(byDate[dates[0]], letterDate, meta || {});
    const buffer = await Packer.toBuffer(doc);
    const fname  = `OD_Combined_${dates[0]}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } else {
    const zip = new JSZip();
    for (const date of dates) {
      const doc    = buildCombinedDocx(byDate[date], letterDate, meta || {});
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
