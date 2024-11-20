const XLSX = require('xlsx');
const { Client } = require('pg');
const { sheet_to_pg_table } = require('./sql-types');
const path = require('path');

async function readExcelAndTest(filename, tableName) {
    console.log(`\nTesting ${filename}...`);
    
    // Read Excel file
    const workbook = XLSX.readFile(path.join('test_files', filename),  { dense: true } );
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to array of objects
    const data = XLSX.utils.sheet_to_json(worksheet); // keep number formatting
    console.log('Parsed Excel data:', data);
    
    // Connect to PostgreSQL
    const client = new Client({
        host: 'localhost',
        database: 'SheetJSPG',
        user: 'postgres',
        password: '7509'
    });
    
    try {
        await client.connect();
        
        // Import data
        await sheet_to_pg_table(client, workbook, tableName);
        
        // Verify table structure
        const structure = await client.query(`
            SELECT column_name, data_type 
            FROM information_schema.columns 
            WHERE table_name = $1
            ORDER BY ordinal_position;
        `, [tableName]);
        console.log('\nTable structure:', structure.rows);
        
        // Verify data
        const results = await client.query(`SELECT * FROM ${tableName}`);
        console.log('\nImported data from DB:', results.rows);
        
    } catch (error) {
        console.error(`Error testing ${filename}:`, error);
        throw error;
    } finally {
        await client.end();
    }
}

async function runAllTests() {
    try {
        // Test each Excel file
        await readExcelAndTest('number_formats.xlsx', 'test_number_formats');
        await readExcelAndTest('date_formats.xlsx', 'test_dates');
        await readExcelAndTest('special_values.xlsx', 'test_special_values');
        await readExcelAndTest('precision.xlsx', 'test_precision');
        await readExcelAndTest('string_formats.xlsx', 'test_string_formats');
        await readExcelAndTest('boolean_formats.xlsx', 'test_boolean_formats');
        
        console.log('\nAll tests completed successfully');
    } catch (error) {
        console.error('\nTests failed:', error);
        process.exit(1);
    }
}

runAllTests().catch(console.error);