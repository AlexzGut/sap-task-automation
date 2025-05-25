import pandas as pd
import os
import sys
import openpyxl


def get_files(folder_name : str) -> list[str]:
    """ Get the files in the given folder name """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(__file__)

    path = os.path.join(current_dir, 'files', folder_name)
    files = os.listdir(path)

    return [os.path.join(path, file) for file in files] 


def get_site_from_file(file_name : str):
    """ Get the site from the file name """
    return file_name.split('_')[-1].rstrip('.csv')


def filter_accounts_by_exclusion(df: pd.DataFrame, exclusion_df: pd.DataFrame, key: str = 'AR Account') -> pd.DataFrame:
    """ Return rows from df where the key column is NOT present in exclusion_df[key]."""
    return df[~df[key].isin(exclusion_df[key])]


def main():
    """ Main function to get the readmitted accounts """
    new_inactive_accounts = get_files('new_inactive_accounts')
    old_inactive_accounts = get_files('old_inactive_accounts')[0]
    admitted_flag_accounts = get_files('admitted_flag_accounts')

    for i in range(len(new_inactive_accounts)):
        site = get_site_from_file(new_inactive_accounts[i])

        # Read the files
        df_old = pd.read_excel(old_inactive_accounts, sheet_name=site)
        df_new = pd.read_csv(new_inactive_accounts[i])
        df_adflag = pd.read_csv(admitted_flag_accounts[i])

         # Filter the dataframes to only include rows with a valid 'AR Account'
        df_old_clean = df_old.dropna(subset=['AR Account'])
        df_new.dropna(subset=['AR Account'], inplace = True)
        df_adflag.dropna(subset=['AR Account'], inplace=True)

        # Filter out rows where in the old inactive accounts 'Column1' is 'Deceased' 
        df_old_discharged_accounts  = df_old_clean[df_old_clean['Column1'] == 'Discharged']

        # Get the readmitted accounts
        df_readmitted = filter_accounts_by_exclusion(df_old_discharged_accounts , df_new) # accounts that were discharged but are now active, however, some discharged/deceased accounts are in the readmitted list
        # This step, uses the readmitted accounts, and a report of admitted accounts with the discharged/deceased flag on to remove those accounts from the readmitted list.
        # This step is necessary because Kroll's report of Discharged/Deceased accounts is based on the Discharged/Deceased date.
        df_readmitted = filter_accounts_by_exclusion(df_readmitted, df_adflag)

        # Select the necessary columns
        cols = ['Pharmacy #',
                'AR Account']

        df_readmitted = df_readmitted.reindex(columns = cols)
        df_readmitted.rename(columns={'Pharmacy #': 'Site'}, inplace=True)
        df_readmitted.to_csv(os.path.join(os.path.join(os.path.dirname(__file__), 'files'), f'readmitted_accounts_{site}.csv'), index=False)

        # Remove readmitted accounts from the old inactive accounts
        df_readmitted_rm = filter_accounts_by_exclusion(df_old, df_readmitted)
        df_readmitted_rm.to_csv(os.path.join(os.path.join(os.path.dirname(__file__), 'files'), f'readmitted_accounts_removed_{site}.csv'), index=False)

        # Save the readmitted accounts to an excel file
        with pd.ExcelWriter(os.path.join(os.path.join(os.path.dirname(__file__), 'files'), 'readmitted.xlsx'), mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_readmitted.to_excel(writer, sheet_name=site, index=False)


if '__main__' == __name__:
    main()




# I need to pull a report for admitted accounts where the Discharged/Deceased flag is on
# Then, do exactly the same cleaning for the old inactive accounts,
# and then compare the two files to see which accounts are in the (admitted + flag) accounts but not in the new inactive accounts