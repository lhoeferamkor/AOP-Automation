import xlsxwriter

# Create a new Excel file and add a worksheet for the data.
workbook = xlsxwriter.Workbook('test_pivot_mcve.xlsx')
data_worksheet = workbook.add_worksheet('Data')

# Write header row
data_worksheet.write_row('A1', ['Product', 'Region', 'Sales'])

# Write some data for the pivot table.
data = [
    ['Apples',  'East',   9000],
    ['Pears',   'East',   5000],
    ['Bananas', 'East',   6000],
    ['Oranges', 'East',   8000],
    ['Apples',  'West',   3000],
    ['Pears',   'West',   4000],
    ['Bananas', 'West',   7000],
    ['Oranges', 'West',   5000],
]

row_num = 1
for item, region, sales in data:
    data_worksheet.write(row_num, 0, item)
    data_worksheet.write(row_num, 1, region)
    data_worksheet.write(row_num, 2, sales)
    row_num += 1

# Add a worksheet for the pivot table.
pivot_worksheet = workbook.add_worksheet('PivotTableSheet')

# Define the properties for the pivot table.
pivot_table_options = {
    'data': '=Data!A1:C9',  # Range of the source data, including headers
    'rows': [
        {'field': 'Product'}
    ],
    'columns': [
        {'field': 'Region'}
    ],
    'values': [
        {'field': 'Sales', 'function': 'sum', 'name': 'Total Sales'}
    ],
    # 'name': 'MySalesPivot', # Optional name for the pivot table in Excel
    # 'style': 10,             # Optional Excel built-in style
}

# Add the pivot table to the 'PivotTableSheet' worksheet.
# 'A3' is the top-left cell where the pivot table will be inserted.
try:
    pivot_worksheet.add_pivot_table('A3', pivot_table_options)
    print(f"Attempting to add pivot table. Using xlsxwriter version: {xlsxwriter.__version__}")
except AttributeError as e:
    print(f"AttributeError encountered: {e}")
    print(f"This strongly suggests an issue with your xlsxwriter version ({xlsxwriter.__version__}) or installation.")
    print("Please ensure you have upgraded xlsxwriter (pip install --upgrade xlsxwriter).")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

# Close the workbook to save the file.
workbook.close()

if 'AttributeError' not in locals() or not isinstance(e, AttributeError): # Check if AttributeError was raised
    print("Excel file 'test_pivot_mcve.xlsx' with pivot table should have been created successfully.")