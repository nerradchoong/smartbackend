const excelService = require('../services/excelService');

exports.uploadExcel = async (req, res) => {
    if (!req.file) {
        console.error("No file received");
        return res.status(400).send("No file uploaded.");
    }
    try {
        const buffer = req.file.buffer;
        const { sheets } = excelService.processExcel(buffer);
        req.session.workbookBuffer = buffer; // Save the buffer to session for later use
        res.json({ sheets }); // Initially return only sheets
    } catch (error) {
        console.error("Error processing Excel file:", error);
        res.status(500).send("Failed to process file");
    }
};

exports.getSheetData = async (req, res) => {
    const sheetName = req.params.sheetName;
    if (!req.session.workbookBuffer) {
        return res.status(400).send("Session expired or workbook not uploaded.");
    }
    try {
        const { items, formulas } = excelService.processSheet(req.session.workbookBuffer, sheetName);
        res.json({ items, formulas });
    } catch (error) {
        console.error("Error processing sheet:", error);
        res.status(500).send("Failed to process sheet");
    }
};
