import os
import pandas as pd
import requests
import re

# Define the input Excel file and output folder
input_file = 'PDF_links.xlsx'  # Replace with your Excel file name
output_folder = 'PDF_all'

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read the Excel file
df = pd.read_excel(input_file)

# Iterate through the rows of the dataframe
for index, row in df.iterrows():
    quarter = '2025-01'
    pdf_url = row[quarter]       # actual column name for links

    cur_num = row['Number']
    # Convert numbers to strings with leading zeros
    cur_num = str(cur_num).zfill(3)

    fund_name = row['FUND']
    fund_name = fund_name.replace(" ", "_")

    if isinstance(pdf_url, str) and not pdf_url.isspace():

        # Extract the language code surrounded by underscores
        match = re.search(r'_([a-z]{2})_', pdf_url)
        if match:
            lang = match.group(1)
            #print(f"Extracted language code: {lang}")
        else:
            lang = ""

        try:
            # Download the PDF
            response = requests.get(pdf_url)
            response.raise_for_status()  # Raise an error for bad HTTP status codes
            
            # Save the PDF with the fund name
            file_path = os.path.join(output_folder, f"{cur_num}.{fund_name}.{lang}.{quarter}.pdf")
            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            print(f"Downloaded: {fund_name}")
        except Exception as e:
            print(f"Failed to download {fund_name} from {pdf_url}: {e}")
    else:
        print(f"Failed to download {fund_name}: NO LINK FOUND")

