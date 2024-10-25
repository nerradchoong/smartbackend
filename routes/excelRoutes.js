const express = require('express');
const router = express.Router();
const excelController = require('../controllers/excelController');

// Define route for interpreting formulas
router.post('/interpretFormulas', excelController.interpretFormulas);

module.exports = router;
