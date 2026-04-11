const crypto = require('crypto');
const XLSX = require('xlsx');
const { db, auth } = require('../firebase');

function normalizeKeyPart(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function buildImportKey(row) {
  const base = [
    row.registerNumber,
    row.name,
    row.department,
    row.batch,
    row.programName,
  ].filter(Boolean).join('|');

  return normalizeKeyPart(base) || crypto.createHash('md5').update(JSON.stringify(row)).digest('hex').slice(0, 16);
}

function inferProgramType(programName = '') {
  const value = programName.toLowerCase();
  if (value.includes('internship')) return 'Internship';
  if (value.includes('cert')) return 'Certification';
  if (value.includes('workshop')) return 'Workshop';
  return 'Training';
}

function mapSheetStatus(status = '') {
  const value = String(status).trim().toLowerCase();
  if (!value) return 'Registered';
  if (value === 'completed' || value === 'complete') return 'Completed';
  if (value === 'doing' || value === 'in progress' || value === 'ongoing') return 'Doing';
  if (value === 'registered' || value === 'enrolled') return 'Registered';
  return String(status).trim();
}

function parseSheetRows(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error('The uploaded workbook is empty.');
  }

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
    header: 1,
    defval: '',
    raw: false,
  });

  if (rows.length < 2) {
    throw new Error('The uploaded sheet does not contain any student rows.');
  }

  const headers = rows[0].map(h => String(h || '').trim());

  return rows
    .slice(1)
    .map((row, index) => {
      const values = headers.reduce((acc, header, colIndex) => {
        acc[header || `__col_${colIndex}`] = String(row[colIndex] || '').trim();
        return acc;
      }, {});

      return {
        rowNumber: index + 2,
        name: values.Name || values.name || values.Student || '',
        registerNumber: values['Register Number'] || values['Register No'] || values['Reg No'] || values['Reg Number'] || values['__col_1'] || '',
        department: values.Department || values.department || '',
        batch: values.Batch || values.batch || values.Year || '',
        programName: values.Program || values['Program Name'] || '',
        domain: values.Domain || values.domain || '',
        status: mapSheetStatus(values.Status || values.status || ''),
      };
    })
    .filter(row => row.name || row.department || row.programName);
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
    let q = db.collection('students').where('status', '==', 'approved');
    const snap = await q.get();
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
    const { dept, year, program, status, type } = req.query;

    const [studSnap, partsSnap, docsSnap] = await Promise.all([
      db.collection('students').where('status', '==', 'approved').get(),
      db.collection('participation').get(),
      db.collection('documents').get(),
    ]);

    let students = studSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    let parts = partsSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    let docs = docsSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    if (dept) { students = students.filter(s => s.department === dept); parts = parts.filter(p => p.department === dept); }
    if (year) { students = students.filter(s => s.year === year); parts = parts.filter(p => p.year === year); }
    if (program) parts = parts.filter(p => p.programName === program);
    if (req.query.domain) parts = parts.filter(p => p.domain === req.query.domain);
    if (status) parts = parts.filter(p => p.status === status);

    if (type === 'documents') {
      return res.json(docs.map(d => ({
        'Student Name': d.studentName, 'Register No': d.registerNumber,
        Department: d.department, 'Program Name': d.programName,
        'Doc Type': d.docType, Status: d.status,
        'Admin Remark': d.adminRemark || '', 'Submitted At': d.submittedAt,
        'Drive Link': d.driveLink,
      })));
    }

    // Default: participation report
    res.json(parts.map(p => ({
      'Student Name': p.name, Department: p.department, Batch: p.year,
      'Program Name': p.programName, 'Program Type': p.programType || '',
      'Reg ID': p.regId || '', 'Enroll Date': p.enrollDate || '',
      Status: p.status, 'Submitted On': p.submittedOn,
    })));
  } catch (e) { res.status(500).json({ message: e.message }); }
};

exports.importStudentSheet = async (req, res) => {
  try {
    if (!req.file?.buffer) {
      return res.status(400).json({ message: 'Please upload an Excel file.' });
    }

    const rows = parseSheetRows(req.file.buffer);

    if (!rows.length) {
      return res.status(400).json({ message: 'No usable rows were found in the uploaded sheet.' });
    }

    const importedAt = new Date().toISOString();
    const results = {
      imported: 0,
      skipped: 0,
      errors: [],
    };

    for (const row of rows) {
      if (!row.name || !row.department || !row.batch || !row.programName) {
        results.skipped += 1;
        results.errors.push(`Row ${row.rowNumber}: missing required fields (name, department, batch, or program).`);
        continue;
      }

      const importKey = buildImportKey(row);
      const studentId = `sheet-${normalizeKeyPart(row.registerNumber) || normalizeKeyPart(`${row.name}-${row.department}-${row.batch}`)}`;
      const programId = `sheet-${normalizeKeyPart(row.programName)}`;
      const participationId = `${studentId}__${programId}`;

      const studentRef = db.collection('students').doc(studentId);
      const programRef = db.collection('programs').doc(programId);
      const participationPayload = {
        studentId,
        name: row.name,
        registerNumber: row.registerNumber,
        department: row.department,
        year: row.batch,
        domain: row.domain,
        programId,
        programName: row.programName,
        programType: inferProgramType(row.programName),
        status: row.status,
        regId: row.registerNumber,
        enrollDate: importedAt,
        submittedOn: importedAt,
        importKey,
        importSource: req.file.originalname,
      };

      const studentDoc = await studentRef.get();
      await studentRef.set({
        uid: studentId,
        email: studentDoc.exists ? (studentDoc.data().email || '') : '',
        name: row.name,
        registerNumber: row.registerNumber,
        department: row.department,
        year: row.batch,
        section: studentDoc.exists ? (studentDoc.data().section || '') : '',
        phone: studentDoc.exists ? (studentDoc.data().phone || '') : '',
        college: studentDoc.exists ? (studentDoc.data().college || '') : '',
        gender: studentDoc.exists ? (studentDoc.data().gender || '') : '',
        domain: row.domain,
        cgpa: studentDoc.exists ? (studentDoc.data().cgpa || '') : '',
        arrears: studentDoc.exists ? (studentDoc.data().arrears || '') : '',
        courses: studentDoc.exists ? (studentDoc.data().courses || '0') : '0',
        certs: studentDoc.exists ? (studentDoc.data().certs || '0') : '0',
        accommodation: studentDoc.exists ? (studentDoc.data().accommodation || '') : '',
        native: studentDoc.exists ? (studentDoc.data().native || '') : '',
        bus: studentDoc.exists ? (studentDoc.data().bus || '') : '',
        plan: studentDoc.exists ? (studentDoc.data().plan || '') : '',
        role: 'student',
        status: 'approved',
        importKey,
        importSource: req.file.originalname,
        createdAt: studentDoc.exists ? (studentDoc.data().createdAt || importedAt) : importedAt,
        importedAt,
        approvedAt: importedAt,
      }, { merge: true });

      const programDoc = await programRef.get();
      await programRef.set({
        name: row.programName,
        type: programDoc.exists ? (programDoc.data().type || inferProgramType(row.programName)) : inferProgramType(row.programName),
        desc: programDoc.exists ? (programDoc.data().desc || 'Imported from Excel sheet') : 'Imported from Excel sheet',
        duration: programDoc.exists ? (programDoc.data().duration || '') : '',
        eligibility: programDoc.exists ? (programDoc.data().eligibility || 'All Students') : 'All Students',
        importSource: req.file.originalname,
        importedAt,
      }, { merge: true });

      await db.collection('participation').doc(participationId).set(participationPayload, { merge: true });
      await db.collection('participations').doc(participationId).set(participationPayload, { merge: true });

      results.imported += 1;
    }

    res.json({
      message: 'Excel sheet imported successfully.',
      ...results,
    });
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
};

