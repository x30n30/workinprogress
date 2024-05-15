import pandas as pd
import numpy as np


def identify_and_extract_tables(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Dictionary to hold the extracted tables from each sheet
    extracted_tables = {}

    # Process each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)

        # Drop completely empty rows and columns to minimize the dataset for processing
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)

        # Initialize variables to track potential tables
        table_start = None
        tables = []

        # Iterate over the DataFrame rows to find blocks of non-empty cells
        for i, row in df.iterrows():
            if not row.isnull().all():  # Check if the row is not completely empty
                if table_start is None:
                    table_start = i  # Mark the start of a potential table
            else:
                if table_start is not None:
                    # Identify the end of the table and append the slice to the list
                    tables.append(df.loc[table_start:i - 1])
                    table_start = None  # Reset the table start

        # Check if the last rows are part of a table
        if table_start is not None:
            tables.append(df.loc[table_start:])

        # Store the tables in the dictionary with the sheet name as the key
        extracted_tables[sheet_name] = tables

    # Output the tables (could also save to new Excel files or process further)
    for sheet, tables in extracted_tables.items():
        print(f"Sheet: {sheet}")
        for i, table in enumerate(tables):
            print(f"Table {i + 1} in {sheet}")
            print(table)
            print("\n")  # Add space between tables for readability


# Usage example
file_path = 'C:\#Code\workinprogress\Input\PoP_Visma_2403.xlsx'
identify_and_extract_tables(file_path)
