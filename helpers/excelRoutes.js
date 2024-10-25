const router = require('express').Router();
const excelController = require('../controllers/excelController');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });

router.post('/upload', upload.single('file'), excelController.uploadExcel);
router.get('/sheet/:sheetName', excelController.getSheetData); // Add this line

module.exports = router;
