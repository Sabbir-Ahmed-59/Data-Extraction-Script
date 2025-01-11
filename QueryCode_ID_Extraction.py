import re
import pandas as pd

# Function to extract IDs from the text pattern
def extract_ids_from_file(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    # Regular expressions to match the patterns
    balance_type_pattern = r'<#elseif\s+item\.BalanceTypeID\s*==\s*(\d+)#>'
    fu_type_pattern = r'<#elseif\s+item2\.FUTypeID\s*==\s*(\d+)#>'

    # Extract IDs using regex
    balance_type_ids = re.findall(balance_type_pattern, content)
    fu_type_ids = re.findall(fu_type_pattern, content)

    # Combine the results into a list of dictionaries for structured output
    extracted_data = [{'Type': 'BalanceTypeID', 'ID': id} for id in balance_type_ids] + \
                     [{'Type': 'FUTypeID', 'ID': id} for id in fu_type_ids]

    return extracted_data

# Function to save extracted IDs to an Excel file
def save_to_excel(extracted_data, excel_file_path):
    df = pd.DataFrame(extracted_data, columns=['Type', 'ID'])
    df.to_excel(excel_file_path, index=False)

# Main execution
if __name__ == "__main__":
    text_file_path = r'E:\RPA\Notification Template Simplification\6. Query Code Template\QueryTemplate.txt'  # Replace with your text file path
    excel_file_path = r'E:\RPA\Notification Template Simplification\6. Query Code Template\Extracted_IDs.xlsx'  # Replace with your desired output Excel file path

    # Extract IDs
    extracted_data = extract_ids_from_file(text_file_path)

    # Save to Excel
    save_to_excel(extracted_data, excel_file_path)

    print(f"IDs extracted and saved to {excel_file_path}")
