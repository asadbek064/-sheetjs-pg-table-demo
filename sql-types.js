const format = require("pg-format");
const XLSX = require('xlsx');

// Utility func for number handling
const isNumeric = value => !isNaN(String(value).replace(/[,$\s]/g,'')) && !isNaN(parseFloat(String(value).replace(/[,$\s]/g,'')));
const cleanNumericValue = value => typeof value === 'string' ? value.replace(/[,$\s]/g, '') : value;

// Determines PostgreSQL data type based on column values
function deduceType(values) {
    if (!values || values.length === 0) return 'text';

    const isDateCell = cell => cell &&  (cell.t === 'd' | (cell.t === 'n'));
    const isBooleanCell = cell => cell?.t === 'b';
    const isValidNumber = cell => cell && (cell.t === 'n' || isNumeric(cell.v));
    const needsPrecision = num => {
        const str = num.toString();
        return str.includes('e') ||
               (str.includes('.') && str.split('.')[1].length > 6) ||
               Math.abs(num) > 1e15;
    };

    // Type detection priority: dates > booleans > numbers > text
    if (values.some(isDateCell)) return 'date';
    if (values.some(isBooleanCell) && values.every(cell => !cell || isBooleanCell(cell))) return 'boolean';

    const numberValues = values
        .filter(isValidNumber)
        .map(cell => parseFloat(cleanNumericValue(cell.v)));

    if (numberValues.length && values.every(cell => !cell || isValidNumber(cell))) {
        return numberValues.some(needsPrecision) ? 'numeric' : 'double precision';
    }
    return 'text';
}

// Converts Sheetjs cell value to PostgreSQL compatible format
function parseValue(cell, type) {
    if (!cell || cell.v == null || cell.v === '') return null;
    
    switch (type) {
        case 'date':
            if (cell.t === 'd') return cell.v.toISOString().split('T')[0];
            if (cell.t === 'n') return new Date((cell.v - 25569) * 86400 * 1000).toISOString().split('T')[0];
            return null;
        case 'double precision':
            if (cell.t === 'n') return cell.v;
            if (isNumeric(cell.v)) return parseFloat(cleanNumericValue(cell.v));
            return null;
        case 'boolean':
            return cell.t === 'b' ? cell.v : null;
        default:
            return String(cell.v);
    }
}

async function sheet_to_pg_table(client, worksheet, tableName) {
    if (!worksheet['!ref']) return;
    
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    
    // Extract headers from first row, clean names for PostgreSQL
    const headers = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const cell = worksheet[cellAddress];
        const headerValue = cell ? String(cell.v).replace(/[^a-zA-Z0-9_]/g, '_') : `column_${col + 1}`;
        headers.push(headerValue.toLowerCase());
    }

    // Group cell values by column for type deduction
    const columnValues = headers.map(() => []);
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            columnValues[col].push(cell);
        }
    }

    // Deduc PostgreSQL type for each column
    const types = {};
    headers.forEach((header, idx) => {
        types[header] = deduceType(columnValues[idx]);
    });

    await client.query(format('DROP TABLE IF EXISTS %I', tableName));
    
    const createTableSQL = format(
        'CREATE TABLE %I (%s)',
        tableName,
        headers.map(header => format('%I %s', header, types[header])).join(', ')
    );
    await client.query(createTableSQL);

    // Insert data row by row
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const values = headers.map((header, col) => {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            return parseValue(cell, types[header]);
        });

        const insertSQL = format(
            'INSERT INTO %I (%s) VALUES (%s)',
            tableName,
            headers.map(h => format('%I', h)).join(', '),
            values.map(() => '%L').join(', ')
        );
        await client.query(format(insertSQL, ...values));
    }
}

module.exports = { sheet_to_pg_table };