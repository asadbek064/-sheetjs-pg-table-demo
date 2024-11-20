# Sheetjs to PostgreSQL Creating a Table Demo

A Node.js utility that intelligently converts Sheetjs `worksheet` to PostgreSQL tables while preserving appropriate data types.

> This demo project serves as a refernce implementation for SheetJS + PostgreSQL integration. For more details, vist the [SheetJS Documentation](https://docs.sheetjs.com/docs/demos/data/postgresql/#creating-a-table).

### Features
* Automatic data type detection from Excel columns
* Support various data formats:
    * Numbers (integer and floating-point)
    * Dates
    * Booleans
    * Text
* Handles special number formats (scientific notations, high precision)
* Clean column names for PostgreSQL compatibility

### Prerequisites
* Node.js
* PostgreSQL (16)
* Python 3.x

### Installation
1. Install Python dependencies:

```bash
pip install -r requirements.txt
```

2. Install Node.js dependencies:
```
npm install i
```

### Setup 
1. Generate test_files:
```bash
python3 gen_test_files.py
```
2. Configure PostgreSQL connection in `test.js`
```javascript
const client = new Client({
    host: 'localhost',
    database: 'SheetJSPG',
    user: 'postgres',
    password: '7509'
});
```

### Run
```bash
node test.js
```

### Test Files
The test suite includes various Excel files testing different data scenarios:

* `number_formats.xlsx`: Various numeric formats
* `date_formats.xlsx`: Date handling
* `special_values.xlsx`: Edge cases
* `precision.xlsx`: High-precision numbers
* `string_formats.xlsx`: Text handling
* `boolean_formats.xlsx`: Boolean values

### Type Mapping
* Excel dates → PostgreSQL `date`
* Booleans → PostgreSQL `boolean`
* High-precision numbers → PostgreSQL `numeric`
* Standard numbers → PostgreSQL `double precision`
* Text/other → PostgreSQL `text`