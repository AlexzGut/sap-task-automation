import pandas as pd
import pdfkit
import os


def generate_medical_expense_report(file_path : str, customer_name : str):
    basedir = os.path.dirname(__file__)
    pdf_filename = f'Medical Expense Report {customer_name}.pdf'
    
    context = clean_data_pipeline(file_path, customer_name)
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

    path_wkhtmltopdf = os.path.join(basedir, 'resources', 'wkhtmltopdf', 'wkhtmltopdf.exe')  # Adjust for your system
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

    pdfkit.from_string(report_html, pdf_filename, configuration=config, options=options)
    print('Completed')


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
        PipelineStage(extract_patient_data),
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
    context['df'] = context['df'][context['df']['Patient First Name'] == context['customer_name']]
    return context

def extract_patient_data(context : dict) -> dict:
    context['patient name'] = f"{context['df'].at[0, 'Patient First Name']} {context['df'].at[0, 'Patient Last Name']}"
    context['facility name'] =  context['df'].at[0, 'Facility Name']
    return context