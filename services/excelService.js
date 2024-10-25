const XLSX = require('xlsx');

exports.interpretFormulas = (formulasRaw, headers) => {
    let formulaInterpretations = {};
    let cachedFormulas = {};  // Cache to store interpretations of formulas

    console.log("Headers received for interpretation:", headers);

    Object.keys(formulasRaw).forEach(cell => {
        const formulaStr = formulasRaw[cell];
        const colIndex = XLSX.utils.decode_cell(cell).c;
        const headerName = headers[colIndex];

        console.log(`Processing cell ${cell} with formula ${formulaStr}. Mapped header: ${headerName}`);

        if (headerName === undefined || headerName === '') {
            console.warn(`Warning: Header for column ${colIndex} is undefined or empty. Skipping this formula.`);
            return;
        }

        if (cachedFormulas[formulaStr]) {
            formulaInterpretations[headerName] = cachedFormulas[formulaStr];
        } else {
            const match = formulaStr.match(/(SUM|AVERAGE)\(([^)]+)\)/);
            if (match) {
                const operation = match[1];
                const range = match[2];
                const rangeParts = range.split(':');

                if (rangeParts.length === 2) {
                    const startCol = XLSX.utils.decode_col(rangeParts[0].match(/[A-Z]+/)[0]);
                    const endCol = XLSX.utils.decode_col(rangeParts[1].match(/[A-Z]+/)[0]);
                    const colsInvolved = headers.slice(startCol, endCol + 1).join(' + ');

                    const interpretation = `${operation}(${colsInvolved})`;
                    formulaInterpretations[headerName] = interpretation;
                    cachedFormulas[formulaStr] = interpretation;
                }
            }
        }
    });

    return Object.entries(formulaInterpretations).map(([key, value]) => ({ header: key, formula: value }));
};
