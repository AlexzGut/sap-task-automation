import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def generate_medical_expense_report(file_path : str, customer_name : str):
    df = clean_data_pipeline(file_path, customer_name)

    pdf_filename = f'Medical Expense Report {customer_name}.pdf'
    with PdfPages(pdf_filename) as pdf:

        # Add DataFrame as a table
        rows_per_page = 20  # Set how many rows per pages
        num_pages = (len(df) + rows_per_page - 1) // rows_per_page

        for page in range(num_pages):
            start = page * rows_per_page
            end = min(start + rows_per_page, len(df))
            df_chunk = df.iloc[start:end]

            fig_table, ax_table = plt.subplots(figsize=(11.69, 8.27))  # Adjust height for rows
            ax_table.axis('tight')
            ax_table.axis('off')
            col_widths = [0.09] * df.shape[1] # Adjust as needed

            table = ax_table.table(
                cellText=df_chunk.values,
                colLabels=df.columns,
                loc='center',
                colWidths=col_widths
            )
            pdf.savefig(fig_table)
            plt.close(fig_table) # Close the figure to free memory


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
    
# === Pipeline ===
def clean_data_pipeline(file_path: str, customer_name : str) -> pd.DataFrame:
    """ Process a file and return a dataframe """
    context = Pipeline([
        PipelineStage(load_excel_file),
        PipelineStage(filter_by_customer_name),
        PipelineStage(select_columns)
    ]).run({
        'file_path' : file_path,
          'customer_name' : customer_name,
            'columns' : [
                'Rx #',
                'Rx Date',
                'Transaction #', 
                'Rx Name', 
                'Quantity', 
                'Drug # Din', 
                'Doctor First Name',
                'Doctor last Name',
                'Total Price', 
                'Total Fee', 
                'Total Plan Paid', 
                'Total 3rd Party Pays', 
                'Patient Pays',
                'Adjudication Date',
            ]
        })
    return context['df']


# === Pipeline stages ===
def load_excel_file(context : dict) -> dict:
    """ Load an excel file into a dataframe """
    try:
        context['df'] =  pd.read_excel(context['file_path'])
        if 'Rx Date' in context['df'].columns:
            context['df']['Rx Date'] = pd.to_datetime(context['df']['Rx Date'], errors='coerce')
            context['df'] = context['df'][context['df']['Rx Date'].notnull()]
            # Remove time from datetime Date
            context['df']['Rx Date'] = context['df']['Rx Date'].dt.date 
            context['df']['Adjudication Date'] = context['df']['Adjudication Date'].dt.date 
    except UnicodeDecodeError as e:
        print(f'Error loading file {context['file_path']}: {e}')
        context['df'] = pd.DataFrame()
    return context

def select_columns(context : dict) -> dict:
    """ Select the columns from the dataframe """
    context['df'] = context['df'].reindex(columns=context['columns'])
    return context

def filter_by_customer_name(context : dict) -> dict:
    """ Filter dataframe rows by customer name """
    context['df'] = context['df'][context['df']['Patient First Name'] == context['customer_name']]
    return context