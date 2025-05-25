import datetime
import pandas as pd
import os
import sys
import openpyxl
import xlwings as xw


def main():
    inactive_accounts = get_files('inactive_accounts')[0]
    discharged_flag_accounts = get_files('admitted_flag_accounts')
    discharged_accounts = get_files('discharged_accounts')

    if (len(discharged_flag_accounts) != len(discharged_accounts)):
        raise ValueError('The number of admitted_flag_accounts and discharged_accounts files do not match')
    
    for i in range(8, len(discharged_flag_accounts)):
        site = get_site_from_file(discharged_accounts[i])
        print(f'Processing {site}...')
        discharged_accounts_df = process_discharged_pipeline(discharged_accounts[i])
        flagged_accounts_df = process_discharged_pipeline(discharged_flag_accounts[i])
        inactive_accounts_df = process_inactive_pipeline(inactive_accounts, site)

        # Get readmitted accounts
        readmitted_candidates_df = filter_accounts_by_exclusion(inactive_accounts_df, discharged_accounts_df)
        readmitted_accounts_df = filter_accounts_by_exclusion(readmitted_candidates_df, flagged_accounts_df)

        # Select the necessary columns and save to excel
        process_readmitted_pipeline(readmitted_accounts_df, site)

        # Remove readmitted accounts from the inactive accounts
        inactive_accounts_df = load_excel_file({'file_name' : inactive_accounts, 'sheet_name' : site})['df']
        inactive_accounts_updated_df = filter_accounts_by_exclusion(inactive_accounts_df, readmitted_accounts_df)

        # Save the updated inactive accounts to the excel file
        book = xw.Book(inactive_accounts)
        sheet = book.sheets[site]

        # Get the last row and column of the sheet
        # This is necessary to clear the existing data in the sheet
        last_row = sheet.range('A1').end('down').row
        last_col = sheet.range('A1').end('right').column

        # Clear the existing data in the sheet
        sheet.range('A2', (last_row, last_col)).clear_contents()

        # Write the updated inactive accounts to the sheet
        sheet.range('A2').value = inactive_accounts_updated_df.values.tolist()
        book.save()
        book.close()

        print(f'Finished processing {site}...')


# === Pipeline classes ===
class PipelineStage:
    def __init__(self, func):
        self.func = func

    def __call__(self, context):
        """ Call the function with the context """
        return self.func(context)


class Pipeline:
    def __init__(self, stages: list[PipelineStage]):
        self.stages = stages

    def run(self, context):
        """ Run the pipeline """
        for stage in self.stages:
            context = stage(context)
        return context
    

# === Pipelines ===
def process_discharged_pipeline(file_name: str) -> pd.DataFrame:
    """ Process a file and return a dataframe """
    context = Pipeline([
        PipelineStage(load_csv_file),
        PipelineStage(dropna_ar_account),
    ]).run({'file_name' : file_name})
    return context['df']


def process_inactive_pipeline(file_name: str, site: str) -> pd.DataFrame:
    """ Process a file and return a dataframe """
    context = Pipeline([
        PipelineStage(load_excel_file),
        PipelineStage(dropna_ar_account),
        PipelineStage(filter_discharge_accounts),
    ]).run({'file_name' : file_name, 'sheet_name' : site})
    return context['df']


def process_readmitted_pipeline(df: pd.DataFrame, site : str) -> pd.DataFrame:
    """ Process a file and return a dataframe """
    context = Pipeline([
        PipelineStage(select_columns),
        PipelineStage(save_to_excel)
    ]).run({'df' : df, 'columns' : ['Pharmacy #', 'AR Account'], 'file_path': os.path.join(os.path.dirname(__file__)), 'sheet_name': site})
    return context['df']


# === Pipeline stages ===
def select_columns(context : dict) -> dict:
    """ Select the columns from the dataframe """
    context['df'] = context['df'].reindex(columns=context['columns'])
    return context


def save_to_excel(context : dict) -> dict:
    """ Save the dataframe to an excel file """
    current_date_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.join(context['file_path'], f'readmitted_accounts_{current_date_time}.xlsx')
    openpyxl.Workbook().save(file_name) # Ensure the workbook is created with the current date and time
    with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        context['df'].to_excel(writer, sheet_name=context['sheet_name'], index=False)
    return context


def dropna_ar_account(context : dict) -> dict:
    """ Drop rows with NaN in the 'AR Account' column """
    context['df'] = context['df'].dropna(subset=['AR Account'])
    return context


def filter_discharge_accounts(context : dict) -> dict:
    """ Get the accounts that are discharged """
    context['df'] = context['df'][context['df']['Column1'] == 'Discharged']
    return context


def load_csv_file(context : dict) -> dict:
    """ Load a csv file into a dataframe """
    encodings = ['utf-8', 'unicode_escape']
    for encoding in encodings:
        try:
            context['df'] = pd.read_csv(context['file_name'], encoding=encoding)
            return context
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f'Error loading file {context['file_name']}: {e}')
            context['df'] = pd.DataFrame()
            return context
        

def load_excel_file(context : dict) -> dict:
    """ Load an excel file into a dataframe """
    try:
        context['df'] =  pd.read_excel(context['file_name'], sheet_name=context['sheet_name'])
    except UnicodeDecodeError as e:
        print(f'Error loading file {context['file_name']}: {e}')
        context['df'] = pd.DataFrame()
    return context


# === Helper functions ===
def filter_accounts_by_exclusion(df: pd.DataFrame, exclusion_df: pd.DataFrame, key: str = 'AR Account') -> pd.DataFrame:
    """ Return rows from df where the key column is NOT present in exclusion_df[key]."""
    return df[~df[key].isin(exclusion_df[key])]


def get_site_from_file(file_name : str) -> str:
    """ Get the site from the file name """
    return file_name.split('_')[-1].rstrip('.csv')


def get_files(folder_name : str) -> list[str]:
    """ Get the files in the given folder name """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(__file__)

    path = os.path.join(current_dir, 'files', folder_name)
    files = os.listdir(path)

    return [os.path.join(path, file) for file in files] 


if '__main__' == __name__:
    main()