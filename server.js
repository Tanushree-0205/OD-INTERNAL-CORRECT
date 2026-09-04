require('dotenv').config();
const express    = require('express');
const path       = require('path');
const JSZip      = require('jszip');
const { createClient } = require('@supabase/supabase-js');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  convertInchesToTwip, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── Supabase Cloud Database Client ────────────────────────────────────────────
const SUPABASE_URL = process.env.SUPABASE_URL || 'https://crjomxucwjlpggvniwnr.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY || process.env.SUPABASE_ANON_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY || 'sb_publishable_rFk_LoeVlHcWkmngwdrFBw_wJ2I6_k_';

let supabase = null;
if (SUPABASE_URL && SUPABASE_KEY) {
  try {
    supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
    console.log('✅ Connected to Supabase Cloud Database:', SUPABASE_URL);
  } catch (e) {
    console.warn('⚠️ Supabase initialization warning:', e.message);
  }
} else {
  console.log('ℹ️ Running in local memory store mode. (Add SUPABASE_URL & SUPABASE_KEY to connect cloud database)');
}

// ─── In-Memory Fallback Store ──────────────────────────────────────────────────
let nextId = 1;
let odEntries = [];
let settings  = { hod_name: 'Dr. M. Senthil Kumar', hod_desig: 'Professor & Head of the Department', hod_dept: 'Computer Science & Engineering' };

function now() {
  return new Date().toISOString().replace('T', ' ').substring(0, 19);
}

// ─── Express setup ─────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Database Health & Status API ──────────────────────────────────────────────
app.get('/api/db/health', async (req, res) => {
  if (!supabase) {
    return res.json({
      connected: false,
      provider: 'Local Memory Storage',
      supabaseUrl: null,
      message: 'SUPABASE_URL and SUPABASE_KEY environment variables are not configured.',
      totalCount: odEntries.length
    });
  }

  try {
    const { data, count, error } = await supabase
      .from('od_entries')
      .select('*', { count: 'exact', head: true });

    if (error) throw error;
    res.json({
      connected: true,
      provider: 'Supabase PostgreSQL Cloud Database',
      supabaseUrl: SUPABASE_URL,
      message: 'Active cloud database connection verified',
      totalCount: count || 0
    });
  } catch (err) {
    res.json({
      connected: false,
      provider: 'Supabase PostgreSQL Cloud Database (Error)',
      supabaseUrl: SUPABASE_URL,
      message: err.message || 'Error communicating with Supabase',
      totalCount: odEntries.length
    });
  }
});

// ─── Settings API ──────────────────────────────────────────────────────────────
app.get('/api/settings', async (req, res) => {
  if (supabase) {
    try {
      const { data, error } = await supabase.from('settings').select('*');
      if (!error && data && data.length > 0) {
        data.forEach(item => {
          settings[item.key] = item.value;
        });
      }
    } catch (e) {
      console.error('Supabase settings fetch error:', e.message);
    }
  }
  res.json(settings);
});

app.post('/api/settings', async (req, res) => {
  const { key, value } = req.body;
  if (!key) return res.status(400).json({ error: 'key required' });
  settings[key] = value;

  if (supabase) {
    try {
      await supabase.from('settings').upsert({ key, value, updated_at: now() });
    } catch (e) {
      console.error('Supabase settings update error:', e.message);
    }
  }

  res.json({ success: true });
});

// ─── OD Entries API ────────────────────────────────────────────────────────────
app.get('/api/entries', async (req, res) => {
  const { date, status, dept, search } = req.query;

  if (supabase) {
    try {
      let query = supabase.from('od_entries').select('*');
      if (date) query = query.eq('od_date', date);
      if (status) query = query.eq('status', status);
      if (dept) query = query.ilike('department', `%${dept}%`);

      const { data, error } = await query.order('od_date', { ascending: false }).order('id', { ascending: false });
      if (!error && data) {
        let rows = data;
        if (search) {
          const s = search.toLowerCase();
          rows = rows.filter(r =>
            (r.student_name && r.student_name.toLowerCase().includes(s)) ||
            (r.vtu_id && r.vtu_id.toLowerCase().includes(s))
          );
        }
        odEntries = rows; // Sync local cache
        return res.json(rows);
      }
    } catch (e) {
      console.error('Supabase fetch entries error:', e.message);
    }
  }

  // Local fallback
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

  rows.sort((a, b) => {
    if (b.od_date !== a.od_date) return b.od_date.localeCompare(a.od_date);
    return b.id - a.id;
  });

  res.json(rows);
});

app.post('/api/entries', async (req, res) => {
  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;
  if (!student_name || !vtu_id || !department || !od_date || !reason)
    return res.status(400).json({ error: 'Missing required fields' });

  const entryData = {
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

  let insertedId = nextId++;

  if (supabase) {
    try {
      const { data, error } = await supabase.from('od_entries').insert([entryData]).select();
      if (!error && data && data.length > 0) {
        insertedId = data[0].id;
        odEntries.push(data[0]);
        return res.json({ id: insertedId, success: true, cloud: true });
      }
    } catch (e) {
      console.error('Supabase insert error:', e.message);
    }
  }

  const localEntry = { id: insertedId, ...entryData };
  odEntries.push(localEntry);
  res.json({ id: insertedId, success: true, cloud: false });
});

app.post('/api/entries/bulk', async (req, res) => {
  const entriesList = Array.isArray(req.body) ? req.body : req.body.entries;
  if (!Array.isArray(entriesList) || entriesList.length === 0) {
    return res.status(400).json({ error: 'Invalid or empty entries array' });
  }

  const validEntries = [];
  for (const item of entriesList) {
    const { student_name, vtu_id, department, od_date, reason, status, notes } = item;
    if (!student_name || !vtu_id || !department || !od_date) continue;
    validEntries.push({
      student_name,
      vtu_id,
      department,
      od_date,
      reason: reason || 'On Duty Academic Activity',
      status: status || 'Pending',
      notes: notes || '',
      created_at: now(),
      updated_at: now()
    });
  }

  if (supabase) {
    try {
      const { data, error } = await supabase.from('od_entries').insert(validEntries).select();
      if (!error && data) {
        data.forEach(e => odEntries.push(e));
        return res.json({ success: true, count: data.length, entries: data, cloud: true });
      }
    } catch (e) {
      console.error('Supabase bulk insert error:', e.message);
    }
  }

  const created = [];
  for (const item of validEntries) {
    const entry = { id: nextId++, ...item };
    odEntries.push(entry);
    created.push(entry);
  }

  res.json({ success: true, count: created.length, entries: created, cloud: false });
});

app.put('/api/entries/:id', async (req, res) => {
  const id  = parseInt(req.params.id);
  const { student_name, vtu_id, department, od_date, reason, status, notes } = req.body;

  if (supabase) {
    try {
      const { error } = await supabase
        .from('od_entries')
        .update({ student_name, vtu_id, department, od_date, reason, status, notes: notes || '', updated_at: now() })
        .eq('id', id);

      if (!error) {
        const idx = odEntries.findIndex(e => e.id === id);
        if (idx !== -1) {
          odEntries[idx] = { ...odEntries[idx], student_name, vtu_id, department, od_date, reason, status, notes, updated_at: now() };
        }
        return res.json({ success: true, cloud: true });
      }
    } catch (e) {
      console.error('Supabase update error:', e.message);
    }
  }

  const idx = odEntries.findIndex(e => e.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Entry not found' });

  odEntries[idx] = {
    ...odEntries[idx],
    student_name, vtu_id, department, od_date, reason, status,
    notes: notes || '',
    updated_at: now()
  };
  res.json({ success: true, cloud: false });
});

app.delete('/api/entries/:id', async (req, res) => {
  const id  = parseInt(req.params.id);

  if (supabase) {
    try {
      const { error } = await supabase.from('od_entries').delete().eq('id', id);
      if (!error) {
        const idx = odEntries.findIndex(e => e.id === id);
        if (idx !== -1) odEntries.splice(idx, 1);
        return res.json({ success: true, cloud: true });
      }
    } catch (e) {
      console.error('Supabase delete error:', e.message);
    }
  }

  const idx = odEntries.findIndex(e => e.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Entry not found' });
  odEntries.splice(idx, 1);
  res.json({ success: true });
});

app.get('/api/stats', async (req, res) => {
  let rows = odEntries;
  if (supabase) {
    try {
      const { data } = await supabase.from('od_entries').select('*');
      if (data) rows = data;
    } catch (e) {}
  }

  const total      = rows.length;
  const pending    = rows.filter(e => e.status === 'Pending').length;
  const processed  = rows.filter(e => e.status === 'Processed').length;
  const today      = new Date().toISOString().split('T')[0];
  const todayCount = rows.filter(e => e.od_date === today).length;
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

function formatDMY(dateStr) {
  if (!dateStr) return '';
  if (typeof dateStr === 'string' && dateStr.includes('/')) return dateStr;
  const str = String(dateStr).trim();
  const parts = str.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`; // DD/MM/YYYY
  }
  return str;
}

function makeLetterhead() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [
        tnr('VEL TECH RANGARAJAN DR. SAGUNTHALA R&D INSTITUTE OF SCIENCE AND TECHNOLOGY', {
          bold: true,
          size: 26, // 13pt
          color: '0F172A'
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [
        tnr('(Deemed to be University Estd. u/s 3 of UGC Act, 1956)', {
          italics: true,
          size: 19, // 9.5pt
          color: '475569'
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 8, color: '1D4ED8' }
      },
      children: [
        tnr('School of Computing  ·  Avadi, Chennai - 600 062, Tamil Nadu, India', {
          size: 21, // 10.5pt
          color: '1E293B',
          bold: true
        })
      ]
    }),
    blank(100)
  ];
}

// ─── Combined DOCX: Multi-Date Grouped Letter with Student Rosters ────────────
function buildCombinedDocx(entries, letterDate, meta = {}) {
  if (!entries || !entries.length) throw new Error('No entries');

  // Group entries by date
  const byDate = {};
  entries.forEach(e => {
    const d = e.od_date || 'Undated';
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(e);
  });
  const sortedDates = Object.keys(byDate).sort();

  const formattedDateList = sortedDates.map(d => formatDMY(d));
  const defaultOdDateStr = formattedDateList.join(', ');
  const odDatesStr = meta.odDates ? meta.odDates : defaultOdDateStr;

  const letterDateStr = letterDate
    ? formatDMY(letterDate)
    : new Date().toLocaleDateString('en-GB'); // DD/MM/YYYY

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

  const headers = ['S.No', 'Student Name', 'VTU ID', 'Department'];
  const colWidths = [800, 3200, 2400, 3600];

  // Build multi-date sections
  const multiDateSections = [];
  for (const d of sortedDates) {
    const dateEntries = byDate[d];

    // Date Header: Date: DD/MM/YYYY
    multiDateSections.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 140, after: 80 },
        children: [
          tnr('Date: ', { bold: true, size: 24, color: '0F172A' }),
          tnr(formatDMY(d), { bold: true, size: 24, color: '1D4ED8' })
        ]
      })
    );

    // Table Header Row
    const headerRow = new TableRow({
      tableHeader: true,
      children: headers.map((h, i) =>
        new TableCell({
          width: { size: colWidths[i], type: WidthType.DXA },
          borders: thinBorder,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({
            children: [tnr(h, { bold: true, size: 22 })],
            spacing: { before: 80, after: 80 }
          })]
        })
      )
    });

    // Table Data Rows
    const dataRows = dateEntries.map((e, idx) =>
      new TableRow({
        children: [
          new TableCell({ width: { size: colWidths[0], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(String(idx + 1), { size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ width: { size: colWidths[1], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.student_name, { bold: true, size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ width: { size: colWidths[2], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.vtu_id, { size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ width: { size: colWidths[3], type: WidthType.DXA }, borders: thinBorder, children: [new Paragraph({ children: [tnr(e.department, { size: 22 })], spacing: { before: 60, after: 60 } })] })
        ]
      })
    );

    multiDateSections.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [headerRow, ...dataRows]
      })
    );

    multiDateSections.push(blank(80));
  }

  const bodyIntro =
    `With reference to the subject cited above, we bring to your kind notice that the following ` +
    `${entries.length} student${entries.length > 1 ? 's' : ''} of ${deptText}, ` +
    `Vel Tech Rangarajan Dr. Sagunthala R&D Institute of Science and Technology, ` +
    `participated in "${eventName}" scheduled from ${eventDates}. ` +
    `The primary purpose and reason for On Duty (OD) is "${commonReason}".`;

  const bodyDetails =
    `The student${entries.length > 1 ? 's were' : ' was'} actively on duty during the period: ${odDatesStr}. ` +
    `Due to their official representation and active participation in the event, ` +
    `they could not attend regular instructional classes and practical laboratory sessions during this duration.`;

  const bodyClose =
    `In view of the above, we kindly request your good office to grant On Duty (OD) attendance ` +
    `for the aforementioned student${entries.length > 1 ? 's' : ''} for the respective date(s) indicated below and enable regularisation of attendance in the academic registry.`;

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
        ...makeLetterhead(),
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
        ...multiDateSections,
        blank(60),
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
  const odDateStr = meta.odDates ? meta.odDates : formatDMY(entry.od_date);
  const letterDateStr = letterDate
    ? formatDMY(letterDate)
    : new Date().toLocaleDateString('en-GB');

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
        children: ['Student Name', 'VTU ID', 'Department', 'OD Date'].map(h =>
          new TableCell({
            borders: thinBorder,
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [tnr(h, { bold: true, size: 22 })], spacing: { before: 60, after: 60 } })]
          })
        )
      }),
      new TableRow({
        children: [
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.student_name, { bold: true, size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.vtu_id, { size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(entry.department, { size: 22 })], spacing: { before: 60, after: 60 } })] }),
          new TableCell({ borders: thinBorder, children: [new Paragraph({ children: [tnr(odDateStr, { size: 22 })], spacing: { before: 60, after: 60 } })] })
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
        ...makeLetterhead(),
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
  const fname  = `OD_${entry.student_name.replace(/\s+/g, '_')}_${formatDMY(entry.od_date).replace(/\//g, '-')}.docx`;
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
    const fname  = `OD_${entry.student_name.replace(/\s+/g, '_')}_${formatDMY(entry.od_date).replace(/\//g, '-')}.docx`;
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

  const doc    = buildCombinedDocx(entries, letterDate, meta || {});
  const buffer = await Packer.toBuffer(doc);
  const fname  = 'OD_Official_Combined_Letter.docx';
  res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.send(buffer);
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
