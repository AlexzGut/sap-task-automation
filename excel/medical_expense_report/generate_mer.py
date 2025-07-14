import pandas as pd
import pdfkit
import os
import random
from datetime import datetime, timedelta


def generate_medical_expense_report(file_path : str, customer_name : str):
    basedir = os.path.dirname(__file__)
    
    # df = clean_data_pipeline(file_path, customer_name)

    # =================== Testing ===========================
    # Random name generators
    first_names = ['John', 'Jane', 'Alice', 'Robert', 'Emily', 'Michael', 'Laura', 'David', 'Sarah', 'Chris']
    last_names = ['Smith', 'Johnson', 'Brown', 'Taylor', 'Anderson', 'Lee', 'White', 'Martin', 'Clark', 'Lewis']

    # Generate 50 rows of sample data
    n = 50
    df = pd.DataFrame({
        'RX #': [f"RX{1000+i}" for i in range(n)],
        'Transaction #': [f"T{5000+i}" for i in range(n)],
        'RX Date': [datetime.today() - timedelta(days=random.randint(0, 365)) for _ in range(n)],
        'Quantiry': [random.randint(10, 100) for _ in range(n)],
        'Drug # Din': [f"DIN{random.randint(100000, 999999)}" for _ in range(n)],
        'Doctor First Name': [random.choice(first_names) for _ in range(n)],
        'Doctor Last Name': [random.choice(last_names) for _ in range(n)],
        'Total Price': [round(random.uniform(10.0, 300.0), 2) for _ in range(n)],
        'Total Fee': [round(random.uniform(5.0, 50.0), 2) for _ in range(n)],
        'Total Plan Paid': [round(random.uniform(0.0, 200.0), 2) for _ in range(n)],
        'Total 3rd Party Pays': [round(random.uniform(0.0, 100.0), 2) for _ in range(n)],
        'Patient Pays': [round(random.uniform(0.0, 100.0), 2) for _ in range(n)],
        'Adjudication Date': [datetime.today() - timedelta(days=random.randint(0, 365)) for _ in range(n)],
    })

    #==================================================================

    pdf_filename = f'Medical Expense Report {customer_name}.pdf'
   
    # Load HTML template
    with open(os.path.join(basedir, 'resources', 'html_template','medical_expense_report.html'), 'r') as file:
        html_template = file.read()

    # Convert DataFrame to HTML table
    # Extract column headers separately
    header_html = '<tr>' + ''.join(f'<th><span>{col}</span></th>' for col in df.columns) + '</tr>'

    #Extract data without headers from dataframe
    html_rows = df.to_html(index=False, header=False)

    # Insert headers manually into a table
    mer_html = f'''
    <table class="dataframe">
    <thead>
    {header_html}
    </thead>
    {''.join(html_rows.splitlines()[1:-1])}
    </table>
    '''

    # Replace <data> tag with table HTML
    updated_html = html_template.replace('<data>', mer_html)

    # Replace <img alt="medisystem_logo"> tag with MediSystem Logo
    medi_logo_tag = f'<img src="{os.path.join(basedir, 'resources', 'img', 'mediSystem_logo.png')}" alt="mediSystem_logo">'
    updated_html = updated_html.replace('<img alt="medisystem_logo">', medi_logo_tag)

    # Save updated HTML
    # with open(os.path.join(basedir, 'temp.html'), 'w') as file:
    #     file.write(updated_html)

    path_wkhtmltopdf = os.path.join(basedir, 'resources', 'wkhtmltopdf', 'wkhtmltopdf.exe')  # Adjust for your system
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    options = {
        'page-size': 'A4',
        'orientation': 'landscape',
        'encoding' : 'utf-8',
        'enable-local-file-access': None,
        'margin-top': '0.5in',
        'margin-right': '0.2in',
        'margin-bottom': '1.0in',
        'margin-left': '0.2in',
    }

    pdfkit.from_string(updated_html, pdf_filename, configuration=config, options=options)
    print('Completed')


generate_medical_expense_report('', 'John')
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