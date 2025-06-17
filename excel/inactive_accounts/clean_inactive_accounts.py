import os
import pandas as pd
import xlwings as xw
import openpyxl

dirname = os.path.dirname(__file__)
path = os.path.join(dirname, 'files')

print(f'Reading files in folder: {path} . . . ', end='')
files = os.listdir(path)
csv_files = []
for file in files:
    if (file.endswith('.csv')):
        csv_files.append(file)
    else:
        excel_file = file
excel_file_path = os.path.join(path, excel_file)
print('Completed')

print('Cleaning process started')
openpyxl.Workbook().save(os.path.join(path, 'clean_data.xlsx')) # Creates a new excel
# ========================================================================================================================================================================
for i in range(0, len(csv_files)):
    csv_file_path = os.path.join(path, csv_files[i])
    site = csv_files[i].split('_')[1].rstrip('.csv')
    print(f'\tCleaning Data: {site} . . . ')

    try:
        df = pd.read_csv(csv_file_path) #data frame
    except (UnicodeDecodeError):
        df = pd.read_csv(csv_file_path, encoding='unicode_escape') #data frame

    site_codes = {'Mississauga':6012,'Richmond':6070,'Manitoba':6090}

    if (site in ['Mississauga', 'Richmond', 'Manitoba']):
        df['Pharmacy #'] = site_codes[site]

    # Converts data in column 'Deceased Date' and 'Discharge Date' from type object to type datetime64[ns]
    # Clean whitespace and invisible characters
    df['Deceased Date'] = df['Deceased Date'].astype(str).str.strip().str.replace('\u00A0', '', regex=True)
    df['Discharge Date'] = df['Discharge Date'].astype(str).str.strip().str.replace('\u00A0', '', regex=True)

    # Convert to datetime
    df['Deceased Date'] = pd.to_datetime(df['Deceased Date'], dayfirst=True, format='mixed', errors='coerce')
    df['Discharge Date'] = pd.to_datetime(df['Discharge Date'], dayfirst=True, format='mixed', errors='coerce')

    # Creates a new column
    # 'Deceased' if the row 'Deceased Date' has a date, otherwise 'Discharge'
    df['Column1'] = df.apply(lambda x: 'Deceased' if pd.notna(x['Deceased Date']) else 'Discharged', axis=1) # x['Deceased Date'] is the value in each row.

    # Deceased Date is updated to replace N/A with a date in 'Discharge Date' column
    df['Deceased Date'] = df.apply(lambda x: x['Deceased Date'] if pd.notna(x['Deceased Date']) else x['Discharge Date'], axis=1)

    # Filter out the necessary columns
    cols = ['Pharmacy #',
            'Home', 
            'Home Address', 
            'Patient', 
            'Deceased Date', 
            'AR Account',
            'Column1']

    clean_data_df = df.reindex(columns = cols)

    clean_data_df.rename(columns={'Deceased Date': 'Deceased/ Discharged Date'}, inplace=True)

    #save cleaned data
    with pd.ExcelWriter(os.path.join(path, 'clean_data.xlsx'), mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        clean_data_df.to_excel(writer, sheet_name=site, index=False)

    # Append Clean DataFrame to a excel sheet
    book = xw.Book(excel_file_path)
    ws = book.sheets[site]

    last_row = ws.range('A1').end('down').row # gets the last row in the sheet
    ws.range(f'A{last_row + 1}').value = clean_data_df.values.tolist() # insert new data in the sheet

    book.save()
    book.close()

    print(f'\tCompleted')
print('Cleaning Process completed')