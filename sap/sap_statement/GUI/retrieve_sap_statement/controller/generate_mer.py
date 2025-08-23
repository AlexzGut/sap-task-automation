import pandas as pd
import pdfkit
from datetime import date
import os


def generate_medical_expense_report(file_path : str, customer_name : str, lower_date : date=None, upper_date : date=None):
    basedir = os.path.dirname(__file__)
    pdf_filename = f'Medical Expense Report {customer_name}.pdf'
    
    context = clean_data_pipeline(file_path, customer_name, lower_date, upper_date)
    df = context['df']

    # Load HTML template
    with open(os.path.join(basedir, 'resources', 'html_template','medical_expense_report.html'), 'r') as file:
        html_template = file.read()

    # Set Facility Name and Resident Name
    report_html = html_template.replace('resident name', context['patient name'])
    report_html = report_html.replace('resident home', context['facility name'])

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
    <tr id="totals">
        <th>Total</th>
        <td></td><td></td><td></td><td></td><td></td><td></td>
        <td>{context['totals']['Total Price']}</td>
        <td>{context['totals']['Total Fee']}</td>
        <td>{context['totals']['Total Plan Paid']}</td>
        <td>{context['totals']['Total 3rd Party Pays']}</td>
        <td>{context['totals']['Patient Pays']}</td>
    </tr>
    </table>
    '''

    # Replace <data> tag with table HTML
    report_html = report_html.replace('<data>', mer_html)

    # Replace <img alt="medisystem_logo"> tag with MediSystem Logo
    medi_logo_tag = f'<img src="{os.path.join(basedir, 'resources', 'img', 'mediSystem_logo.png')}" alt="mediSystem_logo">'
    report_html = report_html.replace('<img alt="medisystem_logo">', medi_logo_tag)

    # Save updated HTML
    with open(os.path.join(basedir, 'temp.html'), 'w') as file:
        file.write(report_html)

    path_wkhtmltopdf = os.path.join(basedir, 'resources', 'wkhtmltopdf', 'wkhtmltopdf.exe') 
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    options = {
        'page-size': 'A4',
        'orientation': 'landscape',
        'encoding' : 'UTF-8',
        'enable-local-file-access': None,
        'margin-top': '0.5in',
        'margin-right': '0.2in',
        'margin-bottom': '1.0in',
        'margin-left': '0.2in',
        'zoom': 1.0,
        'dpi': 300,
    }

    user = os.getlogin()
    pdfkit.from_string(report_html,  os.path.join('C:\\', 'Users', user, 'Downloads', pdf_filename), configuration=config, options=options)
    

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
def clean_data_pipeline(file_path: str, customer_name : str, lower_date : date=None, upper_date : date=None) -> pd.DataFrame:
    """ Process a file and return a dataframe """
    context = Pipeline([
        PipelineStage(load_excel_file),
        PipelineStage(filter_by_customer_name),
        PipelineStage(extract_patient_data),
        PipelineStage(filter_by_date),
        PipelineStage(filter_by_patient_pays),
        PipelineStage(concatanate_doctor_name),
        PipelineStage(select_columns),
        PipelineStage(to_str),
        PipelineStage(get_totals)
    ]).run({
        'file_path' : file_path,
        'customer_name' : customer_name,
        'columns' : [
            'Rx Date',
            'Transaction #', 
            'Doctor Name',
            'Quantity', 
            'Rx #',
            'Drug # Din', 
            'Rx Name', 
            'Total Price', 
            'Total Fee', 
            'Total Plan Paid', 
            'Total 3rd Party Pays', 
            'Patient Pays',
            ],
        'lower_date' : lower_date,
        'upper_date' : upper_date,
        })
    return context


# === Pipeline stages ===
def load_excel_file(context : dict) -> dict:
    """ Load an excel file into a dataframe """
    try:
        context['df'] =  pd.read_excel(context['file_path'])
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
    context['df'] = context['df'][context['df']['Patient Last Name'] == context['customer_name']]
    return context

def extract_patient_data(context : dict) -> dict:
    context['patient name'] = f"{context['df']['Patient First Name'].values[0]} {context['df']['Patient Last Name'].values[0]}"
    context['facility name'] =  context['df']['Facility Name'].values[0]
    return context

def filter_by_date(context : dict) -> dict:
    """ Filter dataframe rows by date (From, To)"""
    if context['lower_date'] and context['upper_date']:
        context['df'] = context['df'][(context['df']['Rx Date'] >= pd.to_datetime(context['lower_date'])) & (context['df']['Rx Date'] <= pd.to_datetime(context['upper_date']))]
    return context

def filter_by_patient_pays(context : dict) -> dict:
    # context['df'] = context['df'][(context['df']['Patient Pays'] != 0)]
    context['df'] = context['df'][(context['df']['Patient Pays'] == 0)]
    return context

def to_str(context : dict) -> dict:
    context['df']['Rx #'] = context['df']['Rx #'].astype(str)
    context['df']['Transaction #'] = context['df']['Transaction #'].astype(str)
    context['df']['Drug # Din'] = context['df']['Drug # Din'].astype(str)
    return context

def get_totals(context : dict) -> dict:
    context['totals'] = context['df'].sum(numeric_only=True)
    context['totals'] = context['totals'].round(decimals=2)
    return context

def concatanate_doctor_name(context : dict) -> dict:
    context['df']['Doctor Name'] = context['df']['Doctor First Name'].astype(str) + ' ' + context['df']['Doctor last Name'].astype(str)
    return context