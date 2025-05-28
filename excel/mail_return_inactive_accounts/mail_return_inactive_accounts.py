import pandas as pd
import os

def main():
    dir = os.path.dirname(__file__)
    EXCEL_PATH = os.path.join(dir, '2025 - Mail Return.xlsx')

    df = process_inactive_accounts(EXCEL_PATH)
    print(df)
    print(df.info())


class PipelineStage:
    def __init__(self, func):
        self.func = func
    
    def __call__(self, context):
        return self.func(context)

class Pipeline:
    def __init__(self, stages: list[PipelineStage]):
        self.stages = stages

    def run(self, context):
        for stage in self.stages:
            context = stage(context)
        return context
    

# === Pipelines ===
def process_inactive_accounts(file_path) -> pd.DataFrame:
    context = Pipeline([
        PipelineStage(load_excel_file),
        PipelineStage(filter_inactive_accounts),
        PipelineStage(dropna_for_column)
    ]).run({'file_path': file_path, 'sheet_name' : 'MailReturn', 'col_dropna' : 'Active/Inactive'})
    return context['df']

# === Pipelines Stages ===
def load_excel_file(context : dict) -> dict:
    """ Load an excel file into a dataframe """
    try:
        context['df'] =  pd.read_excel(context['file_path'], sheet_name=context['sheet_name'], skiprows=2)
    except UnicodeDecodeError as e:
        print(f'Error loading file {context['file_path']}: {e}')
        context['df'] = pd.DataFrame()
    return context


def filter_inactive_accounts(context : dict) -> dict:
    df = context['df']
    context['df'] =df[df['Active/Inactive'] == 'Inactive']
    return context

def dropna_for_column(context: dict) -> dict:
    context['df'] = context['df'].dropna(subset = [context['col_dropna']])
    return context


if __name__ == '__main__':
    main()
