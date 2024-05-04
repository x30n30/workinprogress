import pandas as pd

def load_excel_data(filepath, header_row_index):
    """
    Loads data from an Excel file, starting from a specific header row.

    Parameters:
        filepath (str): Path to the Excel file.
        header_row_index (int): Index of the row containing column headers (0-indexed).

    Returns:
        pd.DataFrame: DataFrame containing the loaded data.
    """
    try:
        df = pd.read_excel(filepath, header=header_row_index)
        return df
    except Exception as e:
        print(f"Failed to load data: {e}")
        return None

# Specify the path to your Excel file and the row index where the headers are found
filepath = r'C:\#Code\workinprogress\data.xlsx'  # Adjust this path to the actual location of your Excel file
header_row_index = 9  # Adjust this index based on the actual position in your file

# Load the data
dataframe = load_excel_data(filepath, header_row_index)

# Print the DataFrame to verify contents
if dataframe is not None:
    if not dataframe.empty:
        print(dataframe)
    else:
        print("The DataFrame is empty.")
else:
    print("No data loaded.")
