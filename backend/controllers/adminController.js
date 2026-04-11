const crypto = require('crypto');
const XLSX = require('xlsx');
const { db, auth } = require('../firebase');

function normalizeKeyPart(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

function buildImportKey(row) {
  const base = [row.registerNumber, row.name, row.department, row.batch, row.programName].filter(Boolean).join('|');
  return normalizeKeyPart(base) || crypto.createHash('md5').update(JSON.stringify(row)).digest('hex').slice(0, 16);
}

function inferProgramType(programName = '') {
  const v = programName.toLowerCase();
  if (v.includes('internship')) return 'Internship';
  if (v.includes('cert')) return 'Certification';
  if (v.includes('workshop')) return 'Workshop';
  return 'Training';
}

function mapSheetStatus(status = '') {
  const v = String(status).trim().toLowerCase();
  if (!v) return 'Registered';
  if (v === 'completed' || v === 'complete') return 'Completed';
  if (v === 'doing' || v === 'in progress' || v === 'ongoing') return 'Doing';
  if (v === 'registered' || v === 'enrolled') return 'Registered';
  return String(status).trim();
}

function parseSheetRows(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('The uploaded workbook is empty.');
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });
  if (rows.length < 2) throw new Error('The uploaded sheet does not contain any student rows.');
  const headers = rows[0].map(h => String(h || '').trim());
  return rows.slice(1).map((row, index) => {
    const v = headers.reduce((acc, h, i) => { acc[h || `__col_${i}`] = String(row[i] || '').trim(); return acc; }, {});
    return {
      rowNumber: index + 2,
      name: v.Name || v.name || v.Student || '',
      registerNumber: v['Register Number'] || v['Register No'] || v['Reg No'] || v['Reg Number'] || v['__col_1'] || '',
      department: v.Department || v.department || '',
      batch: v.Batch || v.batch || v.Year || '',
      programName: v.Program || v['Program Name'] || '',
      domain: v.Domain || v.domain || '',
      status: mapSheetStatus(v.Status || v.status || ''),
    };
  }).filter(r => r.name || r.department || r.programName);
}

exports.getPendingStudents = async (req, res) => {
  try {
    const snap = await db.collection('students').where('status', '==', 'pending').get();
    res.json(snap.docs.map(d => ({ id: d.id, ...d.data() })));
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.approveStudent = async (req, res) => {
  try {
    const { uid } = req.params;
    await auth.updateUser(uid, { disabled: false });
    await db.collection('students').doc(uid).update({ status: 'approved', approvedAt: new Date().toISOString() });
    res.json({ message: 'Student approved.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.rejectStudent = async (req, res) => {
  try {
    const { uid } = req.params;
    const { reason } = req.body;
    await auth.updateUser(uid, { disabled: true });
    await db.collection('students').doc(uid).update({ status: 'rejected', rejectionReason: reason || 'Not specified', rejectedAt: new Date().toISOString() });
    res.json({ message: 'Student rejected.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.getAllStudents = async (req, res) => {
  try {
    const { dept, year } = req.query;
    const snap = await db.collection('students').where('status', '==', 'approved').get();
    let students = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    if (dept) students = students.filter(s => s.department === dept);
    if (year) students = students.filter(s => s.year === year);
    res.json(students);
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.getDashboardStats = async (req, res) => {
  try {
    const [studSnap, docsSnap, partsSnap, pendingSnap] = await Promise.all([
      db.collection('students').where('status', '==', 'approved').get(),
      db.collection('documents').get(),
      db.collection('participation').get(),
      db.collection('students').where('status', '==', 'pending').get(),
    ]);
    const docs = docsSnap.docs.map(d => d.data());
    const parts = partsSnap.docs.map(d => d.data());
    res.json({
      totalStudents: studSnap.size,
      pendingApprovals: pendingSnap.size,
      enrolled: parts.length,
      completed: parts.filter(p => p.status === 'Completed').length,
      verifiedDocs: docs.filter(d => d.status === 'Verified').length,
      pendingDocs: docs.filter(d => d.status === 'Under Review').length,
      rejectedDocs: docs.filter(d => d.status === 'Rejected').length,
    });
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.getReportData = async (req, res) => {
  try {
    const { dept, year, program, status, domain, type } = req.query;
    const ci = (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim();

    const [partsSnap, docsSnap] = await Promise.all([
      db.collection('participation').get(),
      db.collection('documents').get(),
    ]);

    let parts = partsSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    let docs = docsSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    if (dept) parts = parts.filter(p => ci(p.department, dept));
    if (year) parts = parts.filter(p => ci(p.year, year));
    if (program) parts = parts.filter(p => ci(p.programName, program));
    if (domain) parts = parts.filter(p => ci(p.domain, domain));
    if (status) parts = parts.filter(p => ci(p.status, status));

    if (type === 'documents') {
      return res.json(docs.map(d => ({
        'Student Name': d.studentName, 'Register No': d.registerNumber,
        Department: d.department, 'Program Name': d.programName,
        'Doc Type': d.docType, Status: d.status,
        'Admin Remark': d.adminRemark || '', 'Submitted At': d.submittedAt,
        'Drive Link': d.driveLink,
      })));
    }

    res.json(parts.map(p => ({
      'Student Name': p.name,
      'Reg ID': p.regId || p.registerNumber || '',
      Department: p.department,
      Batch: p.year,
      Domain: p.domain || '',
      'Program Name': p.programName,
      'Program Type': p.programType || '',
      Status: p.status,
      'Enroll Date': p.enrollDate || '',
    })));
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.importStudentSheet = async (req, res) => {
  try {
    if (!req.file?.buffer) return res.status(400).json({ message: 'Please upload an Excel file.' });

    const rows = parseSheetRows(req.file.buffer);
    if (!rows.length) return res.status(400).json({ message: 'No usable rows were found in the uploaded sheet.' });

    const importedAt = new Date().toISOString();
    const results = { imported: 0, skipped: 0, errors: [] };

    for (const row of rows) {
      if (!row.name || !row.department || !row.batch || !row.programName) {
        results.skipped += 1;
        results.errors.push(`Row ${row.rowNumber}: missing required fields.`);
        continue;
      }

      const importKey = buildImportKey(row);
      const studentId = `sheet-${normalizeKeyPart(row.registerNumber) || normalizeKeyPart(`${row.name}-${row.department}-${row.batch}`)}`;
      const programId = `sheet-${normalizeKeyPart(row.programName)}`;
      const participationId = `${studentId}__${programId}`;

      const participationPayload = {
        studentId, name: row.name, registerNumber: row.registerNumber,
        department: row.department, year: row.batch, domain: row.domain,
        programId, programName: row.programName, programType: inferProgramType(row.programName),
        status: row.status, regId: row.registerNumber,
        enrollDate: importedAt, submittedOn: importedAt,
        importKey, importSource: req.file.originalname,
      };

      const studentDoc = await db.collection('students').doc(studentId).get();
      const sd = studentDoc.exists ? studentDoc.data() : {};
      await db.collection('students').doc(studentId).set({
        uid: studentId, email: sd.email || '', name: row.name,
        registerNumber: row.registerNumber, department: row.department,
        year: row.batch, domain: row.domain,
        section: sd.section || '', phone: sd.phone || '', college: sd.college || '',
        gender: sd.gender || '', cgpa: sd.cgpa || '', arrears: sd.arrears || '',
        courses: sd.courses || '0', certs: sd.certs || '0',
        accommodation: sd.accommodation || '', native: sd.native || '',
        bus: sd.bus || '', plan: sd.plan || '',
        role: 'student', status: 'approved', importKey,
        importSource: req.file.originalname,
        createdAt: sd.createdAt || importedAt, importedAt, approvedAt: importedAt,
      }, { merge: true });

      const programDoc = await db.collection('programs').doc(programId).get();
      const pd = programDoc.exists ? programDoc.data() : {};
      await db.collection('programs').doc(programId).set({
        name: row.programName,
        type: pd.type || inferProgramType(row.programName),
        desc: pd.desc || 'Imported from Excel sheet',
        duration: pd.duration || '', eligibility: pd.eligibility || 'All Students',
        importSource: req.file.originalname, importedAt,
      }, { merge: true });

      await db.collection('participation').doc(participationId).set(participationPayload, { merge: true });
      await db.collection('participations').doc(participationId).set(participationPayload, { merge: true });

      results.imported += 1;
    }

    res.json({ message: 'Excel sheet imported successfully.', ...results });
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
};
