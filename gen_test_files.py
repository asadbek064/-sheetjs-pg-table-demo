import pandas as pd
from datetime import datetime
import numpy as np
import os

def create_test_directory():
    """Create a directory for test files if it doesn't exist"""
    if not os.path.exists('test_files'):
        os.makedirs('test_files')
        
def generate_number_formats_test():
    """Test Case 1: Common spreadsheet number formats"""
    df = pd.DataFrame({
        'id': range(1, 7),
        'value': [
            1234.56,              # Plain number
            '1,234.56',           # Thousands separator
            1234.5600,            # Fixed decimal places
            0.1234,               # Will be formatted as percentage
            -1234.56,             # Will be formatted as parentheses
            -1230                 # Will be formatted as scientific
        ]
    })
    
    # Create Excel writer with xlsxwriter engine
    writer = pd.ExcelWriter('test_files/number_formats.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # Get workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Add formats
    percent_format = workbook.add_format({'num_format': '0.00%'})
    accounting_format = workbook.add_format({'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})
    scientific_format = workbook.add_format({'num_format': '0.00E+00'})
    
    # Apply formats to specific cells
    worksheet.set_column('B:B', 15)  # Set column width
    worksheet.write('B5', 0.1234, percent_format)  # Percentage
    worksheet.write('B6', -1234.56, accounting_format)  # Parentheses
    worksheet.write('B7', -1230, scientific_format)  # Scientific
    
    writer.close()

def generate_date_formats_test():
    """Test Case 3: Date and timestamp formats"""
    df = pd.DataFrame({
        'id': range(1, 5),
        'date': [
            datetime(2024, 1, 1),                    # ISO format
            datetime(2024, 1, 1),                    # US format
            datetime(2024, 1, 1),                    # Excel format
            45292,                                   # Excel serial date
        ]
    })
    
    writer = pd.ExcelWriter('test_files/date_formats.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Add different date and timestamp formats
    iso_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    us_format = workbook.add_format({'num_format': 'm/d/yyyy'})
    excel_format = workbook.add_format({'num_format': 'dd-mmm-yyyy'})
    
    # New timestamp formats
    datetime_24h_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
    datetime_12h_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss AM/PM'})
    datetime_ms_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss.000'})
    iso_timestamp_format = workbook.add_format({'num_format': 'yyyy-mm-ddThh:mm:ss'})
    
    # Set column width to accommodate timestamps
    worksheet.set_column('B:B', 25)
    
    # Apply formats
    worksheet.write('B2', datetime(2024, 1, 1), iso_format)
    worksheet.write('B3', datetime(2024, 1, 1), us_format)
    worksheet.write('B4', datetime(2024, 1, 1), excel_format)
    worksheet.write('B5', 45292, excel_format)

    
    writer.close()
    

def generate_special_values_test():
    """Test Case 4: Empty and special values"""
    df = pd.DataFrame({
        'id': range(1, 6),
        'value': [
            np.nan,               # NULL
            '',                   # Empty string
            '#N/A',               # Excel error
            '#DIV/0!',            # Excel error
            '-'                   # Common placeholder
        ]
    })
    
    writer = pd.ExcelWriter('test_files/special_values.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()

def generate_precision_test():
    """Test Case 5: Number precision"""
    df = pd.DataFrame({
        'id': range(1, 8),
        'value': [
            1.234567890123456,              # High precision decimal
            12345678901234567890,           # Large integer (> 15 digits)
            -0.00000000123456,              # Small decimal
            9.99999e20,                     # Scientific notation large
            -1.23456e-10,                   # Scientific notation small
            123456789.123456789,            # Mixed large number with decimals
            1234567890123456.789            # Edge case for precision
        ]
    })
    
    writer = pd.ExcelWriter('test_files/precision.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Add formats for different number types
    precision_format = workbook.add_format({'num_format': '0.000000000000000'})
    scientific_format = workbook.add_format({'num_format': '0.000000E+00'})
    
    worksheet.write('B3', 1.234567890123456, precision_format)
    worksheet.write('B4', 9.99999e20, scientific_format)
    
   # Apply formats
    worksheet.set_column('B:B', 20)
    for row in range(1, 8):
        if row in [4, 5]:  # Scientific notation
            worksheet.write(row, 1, df['value'][row-1], scientific_format)
        else:
            worksheet.write(row, 1, df['value'][row-1], precision_format)
    
    writer.close()
    
def generate_string_formats_test():
    """Test Case 2: String formats and special characters"""
    df = pd.DataFrame({
        'id': range(1, 9),
        'value': [
            'Simple text',                    # Plain text
            'Text with spaces   ',            # Trailing spaces
            '   Text with spaces',            # Leading spaces
            'Text with\nnewline',             # Newline character
            'Text with "quotes"',             # Quoted text
            'Text with special chars: @#$%',  # Special characters
            'Very long text ' * 10,           # Long text
            'Super long text ' * 100           # Super long text
        ]
    })
    
    writer = pd.ExcelWriter('test_files/string_formats.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Set column width to show long text
    worksheet.set_column('B:B', 50)
    
    writer.close()
    
def generate_boolean_formats_test():
    """Test Case: Boolean formats in Excel"""
    df = pd.DataFrame({
        'id': range(1, 5),
        'value': [
            True,                   # Simple True
            False,                  # Simple False
            'TRUE',                 # String TRUE
            'FALSE'                 # String FALSE
        ]
    })
    
    writer = pd.ExcelWriter('test_files/boolean_formats.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Add boolean formats
    bool_format = workbook.add_format({'num_format': 'BOOLEAN'})
    custom_true_false = workbook.add_format({'num_format': '"True";;"False"'})
    yes_no_format = workbook.add_format({'num_format': '"YES";;NO'})
    
    # Set column width
    worksheet.set_column('B:B', 15)
    
    # Apply different boolean formats
    worksheet.write('B2', True, bool_format)          # Standard TRUE
    worksheet.write('B3', False, bool_format)         # Standard FALSE
    worksheet.write('B4', True, yes_no_format)        # Yes/No format
    worksheet.write('B5', False, yes_no_format)       # Yes/No format
    
    writer.close()


def main():
    """Geneate all test Excel files"""
    create_test_directory()

    print("Generating test Excel files...")
    generate_number_formats_test()
    generate_date_formats_test()
    generate_special_values_test()
    generate_precision_test()
    generate_string_formats_test()
    generate_boolean_formats_test()
    print("Test files generated in 'test_files' directory")

if __name__ == "__main__":
    main()