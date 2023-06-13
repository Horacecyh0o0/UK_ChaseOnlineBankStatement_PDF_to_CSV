import pdfplumber
import os
import re
import pandas as pd
from datetime import datetime, timedelta
import shutil

#Funcitons

def extract_desired_text(file_path):
    with pdfplumber.open(file_path) as pdf:
        desired_text = ""
        start_flag = False
        first_line_found = False
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                if not start_flag and "Date Transaction details Amount Balance" in line:
                    start_flag = True
                    first_line_found = True
                    continue  # Skip adding the first line
                if "Some useful information" in line:
                    start_flag = False
                    break
                if start_flag and re.match(r"^\d", line) and len(line) > 5:
                    desired_text += line + '\n'
    return desired_text

def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


#main

active_folder_path = 'D:/Chase Bank Statement/Active'
active_files = os.listdir(active_folder_path)
active_file_name = active_files[0]
active_file_path = os.path.join(active_folder_path, active_file_name)

pdf_desired_text = extract_desired_text(active_file_path)

lines = pdf_desired_text.split('\n')
lines = [line for line in lines if line]
dates = []
transaction_details = []
amounts = []
balances = []

# Iterate over the lines and extract the information
for line in lines:
    elements = line.split()
    if elements[3].startswith('-'):
        continue  # Skip the current iteration and move to the next line
    
    date_str = " ".join(elements[:3])
    date = datetime.strptime(date_str, '%d %b %Y').date()
    transaction_detail = " ".join(elements[3:-2])

    amountstr = elements[-2]
    amountstr = amountstr.replace('£', '').replace(',', '')
    if isfloat(amountstr):
        amount = float(amountstr)
    else:
        transaction_detail = " ".join(elements[3:-1])
        amount = None

    
    balancestr = elements[-1].lstrip('£')
    balancestr = balancestr.replace(',', '')
    if isfloat(balancestr):
        balance = float(balancestr)
    else:
        balance = None
    
    

   # Append the data to the respective lists
    dates.append(date)
    transaction_details.append(transaction_detail)
    amounts.append(amount)
    balances.append(balance)

# Create a DataFrame from the extracted data
data = {
    'Date': dates,
    'Transaction details': transaction_details,
    'Amount': amounts,
    'Balance': balances
}
df = pd.DataFrame(data)

# Export the DataFrame to an Excel file
today = dates[0]
year = today.year
month = today.month
DesiredExcelName = f"{year} {month:02d}"
output_path = r"D:\Chase Bank Statement\Output"

output_file = fr"{output_path}\{DesiredExcelName}.xlsx"
df.to_excel(output_file, index=False)


destination_folder = 'D:\Chase Bank Statement\Prcessed Raw PDF'
shutil.move(active_file_path, destination_folder)
