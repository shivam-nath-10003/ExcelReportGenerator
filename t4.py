import pandas as pd
import numpy as np
import xlsxwriter

# Function to process the Excel file
def process_excel(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Extract the relevant columns
    entry_column = df.iloc[:, 15]  # 16th column (0-based index)
    associated_data_column = df.iloc[:, 16]  # 17th column (0-based index)
    msme_status_column = df.iloc[:, 17]  # 18th column (0-based index)
    quadrant_column = df.iloc[:, 24]  # 25th column (0-based index)
    date_column = df.iloc[:, 1]  # 2nd column (0-based index)
    business_column = df.iloc[:, 28]  # 29th column (0-based index)

    # Create a DataFrame for the first sheet
    result_df = pd.DataFrame({
        'Entry': entry_column,
        'Associated Data': associated_data_column,
        'MSME status': msme_status_column
    })

    # Count occurrences and create a new column for it
    result_df['Occurrence Count'] = result_df.groupby('Entry')['Entry'].transform('count')

    # Remove duplicate rows to keep only unique entries
    result_df = result_df.drop_duplicates(subset='Entry')

    # Sort the DataFrame by 'Occurrence Count' in descending order
    result_df = result_df.sort_values(by='Occurrence Count', ascending=False)

    # Add a new column for quadrants and highlight cells with more than one quadrant in yellow
    quadrant_dict = {}  # Dictionary to store unique quadrants for each seller
    for index, row in df.iterrows():
        entry = row['GEM Seller ID']
        quadrant = row['Quadrant']
        if entry not in quadrant_dict:
            quadrant_dict[entry] = set()  # Use a set to store unique quadrants
        if not pd.isna(quadrant):  # Skip NaN values
            # Check for the presence of 'Q1', 'Q2', 'Q3', 'Q4' within the cell
            for q in ['Q1', 'Q2', 'Q3', 'Q4']:
                if q in str(quadrant):
                    quadrant_dict[entry].add(q)

    # Function to highlight cells with more than one quadrant in yellow
    def highlight_multiple_quadrants_yellow(s):
        entry = s['Entry']
        if entry in quadrant_dict:
            is_multiple = len(quadrant_dict[entry]) > 1
            return ['background-color: yellow' if is_multiple else '' for _ in s]
        else:
            return ['' for _ in s]

    # Function to highlight cells with only one quadrant in pink
    def highlight_single_quadrant_pink(s):
        entry = s['Entry']
        if entry in quadrant_dict:
            is_single = len(quadrant_dict[entry]) == 1
            return ['background-color: pink' if is_single else '' for _ in s]
        else:
            return ['' for _ in s]

    # Function to highlight cells with 'No' in the 'MSME status' column in red
    def highlight_msme_no_red(s):
        return ['background-color: red' if str(s['MSME status']).lower() == 'no' else '' for _ in s]

    # Add a new column for quadrants
    result_df['Quadrants'] = result_df.apply(lambda row: ', '.join(quadrant_dict.get(row['Entry'], [])), axis=1)

    # Highlight cells with more than one quadrant in yellow
    result_df_styled = result_df.style.apply(highlight_multiple_quadrants_yellow, axis=1, subset=['Entry'])

    # Highlight cells with only one quadrant in pink
    result_df_styled = result_df_styled.apply(highlight_single_quadrant_pink, axis=1, subset=['Entry'])

    # Highlight cells with 'No' in the 'MSME status' column in red
    result_df_styled = result_df_styled.apply(highlight_msme_no_red, axis=1, subset=['Associated Data', 'MSME status'])

    # Write the first sheet to a new Excel file with highlighting
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        result_df_styled.to_excel(writer, sheet_name='All Sellers', index=False)

        # Create a new sheet for each of the top 15 sellers
        for seller_id in result_df['Entry'].head(15):
            # Filter the data for the current seller
            seller_data = df[df['GEM Seller ID'] == seller_id]

            # Create a DataFrame for the date-wise business
            seller_sheet_df = pd.DataFrame({
                'Date': seller_data['Date'],  # Replace with the correct column name for dates
                'Business Done': pd.to_numeric(seller_data['Total Order Value'], errors='coerce').fillna(0)
            })

            # Create a new sheet for the current seller
            sheet_name = seller_id[:31]  # Use the first 31 characters of the seller ID as the sheet name
            seller_sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Create a bar graph for the current seller
            chart_sheet = writer.sheets[sheet_name]
            chart = writer.book.add_chart({'type': 'bar'})
            chart.add_series({
                'categories': f"='{sheet_name}'!$A$2:$A${len(seller_sheet_df) + 1}",
                'values': f"='{sheet_name}'!$B$2:$B${len(seller_sheet_df) + 1}",
                'name': 'Business Done',
                'fill': {'color': 'yellow'}  # Adjust the color based on quadrant if needed
            })
            chart_sheet.insert_chart('D2', chart)

# Example usage
input_excel_file = 'i.xlsx'  # Replace with your input file
output_excel_file = 'o008.xlsx'  # Replace with your desired output file
process_excel(input_excel_file, output_excel_file)
