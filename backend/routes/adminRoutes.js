const express = require('express');
const router = express.Router();
const multer = require('multer');
const adminController = require('../controllers/adminController');
const { verifyToken, requireAdmin } = require('../middleware/auth');
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
});

router.use(verifyToken, requireAdmin);

router.get('/students/pending', adminController.getPendingStudents);
router.get('/students', adminController.getAllStudents);
router.put('/students/:uid/approve', adminController.approveStudent);
router.put('/students/:uid/reject', adminController.rejectStudent);
router.get('/dashboard', adminController.getDashboardStats);
router.get('/report', adminController.getReportData);
router.get('/debug-participation', async (req, res) => {
  const { db } = require('../firebase');
  const snap = await db.collection('participation').get();
  const sample = snap.docs.slice(0, 3).map(d => d.data());
  res.json({ total: snap.size, sample });
});
router.post('/import-sheet', upload.single('sheet'), adminController.importStudentSheet);

module.exports = router;
