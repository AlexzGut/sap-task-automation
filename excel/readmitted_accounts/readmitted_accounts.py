import pandas as pd
import os


def main():
    # Get the current directory of the application
    current_dir = os.path.dirname(__file__)

    # Get the path to the files directory
    path = os.path.join(current_dir, 'files')

    # Check if the path exists
    if not os.path.exists(path):
        print(f"Path does not exist: {path}")
        return
    
    # Get the file containing admitted accounts
    new_discharged_file = '1 Jan 2017 - 31 Dec 2025 Inactive Patient Listing Report_Toronto.csv'
    new_discharged_accounts = os.path.join(path, 'new_inactive_accounts', new_discharged_file)

    # Get the file containing inactive accounts
    old_inactive_accounts_file = 'Updated_ALL INACTIVE Account List.xlsx'
    old_inactive_accounts = os.path.join(path, 'old_inactive_accounts',old_inactive_accounts_file)

    # Read the files
    df_new_discharged = pd.read_csv(new_discharged_accounts)
    df_old_inactive = pd.read_excel(old_inactive_accounts, sheet_name='Toronto')

    # Filter the dataframes to only include rows with a valid 'AR Account'
    df_new_discharged.dropna(subset=['AR Account'], inplace=True)
    df_old_inactive_ar = df_old_inactive.dropna(subset=['AR Account'])

    # Filter out rows where 'Column1' is 'Deceased' 
    df_old_discharged = df_old_inactive_ar[df_old_inactive_ar['Column1'] == 'Discharged']
    
    # Keep accounts that are in the old inactive accounts but not in the new inactive accounts 
    df_readmitted = df_old_discharged[~df_old_discharged['AR Account'].isin(df_new_discharged['AR Account'])]
    
    # Select the necessary columns
    cols = ['Pharmacy #',
            'AR Account']

    df_readmitted = df_readmitted.reindex(columns = cols)
    df_readmitted.rename(columns={'Pharmacy #': 'Site'}, inplace=True)
    df_readmitted.to_csv(os.path.join(path, 'readmitted_accounts.csv'), index=False)

    # Remove readmitted accounts from the old inactive accounts
    df_readmitted_rm = df_old_inactive[~df_old_inactive['AR Account'].isin(df_readmitted['AR Account'])]
    df_readmitted_rm.to_csv(os.path.join(path, 'readmitted_accounts_removed.csv'), index=False)


if '__main__' == __name__:
    main()