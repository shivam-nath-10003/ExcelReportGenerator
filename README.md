# Excel Data Processing and Reporting

This Python script processes Excel data, extracts relevant columns, performs data manipulation, and generates formatted Excel reports with highlighting and visualizations.

## Dependencies
- pandas
- numpy
- xlsxwriter


## Usage
1. Clone or download the repository.
2. Install dependencies using the above command.
3. Replace the input Excel file path (`input_excel_file`) and output Excel file path (`output_excel_file`) in the script.
4. Run the script `process_excel.py`.
5. View the generated Excel file with processed data, highlighting, and visualizations.

## Functionality
- The script reads an input Excel file and extracts relevant columns for processing.
- It creates a new Excel file with formatted sheets and highlighted cells.
- Cells with multiple quadrants are highlighted in yellow, cells with a single quadrant in pink, and cells with 'No' in the 'MSME status' column in red.
- For each of the top 15 sellers, a new sheet is created with date-wise business data, and bar graphs are added for visual representation.

## Customization
- Adjust the input and output file paths according to your specific requirements.
- Modify the script to handle different column names or data formats in the input Excel file if necessary.

## Example Usage
```python
input_excel_file = 'input.xlsx'
output_excel_file = 'output.xlsx'
process_excel(input_excel_file, output_excel_file)

