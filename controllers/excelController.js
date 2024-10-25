const excelService = require('../services/excelService');

exports.interpretFormulas = (req, res) => {
    const { formulas, headers } = req.body;
    try {
        const interpretedFormulas = excelService.interpretFormulas(formulas, headers);
        res.json({ formulas: interpretedFormulas });
    } catch (error) {
        console.error("Error interpreting formulas:", error);
        res.status(500).send("Failed to interpret formulas");
    }
};
