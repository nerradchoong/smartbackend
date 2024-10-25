const XLSX = require('xlsx');

exports.processExcel = (buffer) => {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheets = workbook.SheetNames.map(name => name.trim().toLowerCase());  // Normalize and store sheet names
    return { sheets, buffer };  // Return normalized sheet names and buffer for later use
};

exports.processSheet = (buffer, requestedSheetName) => {
    const normalizedRequestedName = requestedSheetName.trim().toLowerCase();
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const normalizedSheetNames = workbook.SheetNames.map(name => name.trim().toLowerCase());

    const sheetIndex = normalizedSheetNames.indexOf(normalizedRequestedName);
    const sheetName = sheetIndex !== -1 ? workbook.SheetNames[sheetIndex] : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
        console.error(`No worksheet named ${sheetName} found.`);
        return { items: [], formulas: [] };
    }

    const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0] || [];
    const items = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: false });
    const formulas = interpretFormulas(worksheet, headers);
    return { items, formulas, sheetUsed: sheetName };
};

function interpretFormulas(worksheet, headers) {
    const formulaCells = XLSX.utils.sheet_to_formulae(worksheet);
    let formulaInterpretations = {};
    Object.keys(formulaCells).forEach(cell => {
        const formulaStr = formulaCells[cell];
        const { c } = XLSX.utils.decode_cell(cell);
        const headerName = headers[c] || `Column ${c + 1}`;
        formulaInterpretations[headerName] = interpretSingleFormula(formulaStr, headers);
    });
    return Object.entries(formulaInterpretations).map(([header, formula]) => ({ header, formula }));
}

function interpretSingleFormula(formulaStr, headers) {
    const match = formulaStr.match(/(SUM|AVERAGE)\(([^)]+)\)/);
    if (match) {
        const operation = match[1];
        const range = match[2];
        const rangeParts = range.split(':');
        if (rangeParts.length === 2) {
            const startCol = XLSX.utils.decode_col(rangeParts[0].match(/[A-Z]+/)[0]);
            const endCol = XLSX.utils.decode_col(rangeParts[1].match(/[A-Z]+/)[0]);
            const colsInvolved = headers.slice(startCol, endCol + 1).join(' + ');
            return `${operation}(${colsInvolved})`;
        }
    }
    return formulaStr; // Default return if no pattern match
}
