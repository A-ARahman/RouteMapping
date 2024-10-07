import pandas as pd

def check_string_data_types(file_path):
    # Read the Excel file
    df = pd.read_excel('datanode2.xlsx')
    
    # Check for string data types in each row and show the location
    string_columns = df.select_dtypes(include=['object']).columns.tolist()
    
    if string_columns:
        for col in string_columns:
            for idx, value in df[col].iterrows():
                if isinstance(value, str):
                    print(f"String data found at row {idx}, column '{col}': {value}")
    else:
        print("No columns with string data types found.")

if __name__ == "__main__":
    file_path = 'path_to_your_excel_file.xlsx'  # Replace with your actual file path
    check_string_data_types(file_path)