import pandas as pd
import os

def read_employee_data(file_path): 
    try:
        # Read the Excel file using the provided file path
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Check for missing values
        if df.isnull().any().any():
            print("Warning: Missing values detected in the Excel file.")
        
        print("Employee data loaded successfully:")
        print(df)
        
        return df
    
    except FileNotFoundError:
        print(f"Error: The file was not found at the specified path: {file_path}")
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")

# Example usage
if __name__ == "__main__":
    file_path =  r"C:\Users\uncommonStudent\Desktop\python\payslips\employee.xlsx"
    employee_df = read_employee_data(file_path)

    if employee_df is not None:
        for index, row in employee_df.iterrows():
            print(f"\nEmployee {index + 1}:")
            print(f"ID: {row['Employee ID']}")
            print(f"Name: {row['Name']}")
            print(f"Email: {row['Email']}")
            print(f"Basic Salary: {row['Basic Salary']}")
            print(f"Allowances: {row['Allowances']}")
            print(f"Deductions: {row['Deductions']}")