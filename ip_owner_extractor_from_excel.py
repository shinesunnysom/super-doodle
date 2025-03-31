# Written in Spyder IDE
# Learning purposes, give where credit is due...
# Extraction of IPs and Owners from muddle information of text, just simply take an list of information and print an excel sheet with IPs and Owners

import pandas as pd
import re

def extract_ip_owner_from_excel(file_path, column_name):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Check if the specified column exists
    if column_name not in df.columns:
        print(f"Error: Column '{column_name}' not found in the Excel file")
        return None
    
    # Regular expression pattern for IP and owner
    ip_pattern = r'Client Address: (\d+\.\d+\.\d+\.\d+)\s*\((.*?)\)'

    # List to store results
    results = []

    # Process each cell in the specified column
    for text in df[column_name]:
        if pd.isna(text):  # Skip empty cells
            continue
            
        # Find all matches in the current cell
        matches = re.finditer(ip_pattern, str(text))
        
        # Extract IP and owner for each match
        for match in matches:
            ip = match.group(1)
            owner = match.group(2)
            results.append({'IP': ip, 'Owner': owner})
    
    # Convert results to DataFrame for better Spyder integration
    result_df = pd.DataFrame(results)
    
    # Print results (visible in Spyder's IPython console)
    if not result_df.empty:
        print("Extracted IP and Owner information:")
        for index, row in result_df.iterrows():
            print(f"{row['IP']}, {row['Owner']}")
    
    # Save to Excel
    output_file = 'ip_owner_output.xlsx'
    result_df.to_excel(output_file, index=False)
    print(f"\nResults saved to '{output_file}'")
    
    # Return DataFrame for use in Spyder's Variable Explorer
    return result_df

# Usage with specific values
try:
    # Specific file path and column name
    file_path = 'C:\\Users\\JohnDoe\\Documents\\network_logs.xlsx'
    column_name = 'Details' #whatever
    
    # Run the function and store result
    extracted_data = extract_ip_owner_from_excel(file_path, column_name)
    
except Exception as e:
    print(f"An error occurred: {str(e)}")
