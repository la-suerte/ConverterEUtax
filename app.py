import os
import pandas as pd
from flask import Flask, request, render_template_string, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import tempfile
import uuid
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

# Configuration
UPLOAD_FOLDER = '/tmp'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_required_sections(excel_data):
    """Validate that all required sections are present in the Excel file"""
    required_sections = [
        'General Information',
        'Country-by-Country Overview',
        'Subsidiaries and Activities',
        'Omitted Information'
    ]
   
    missing_sections = []
    available_sheets = list(excel_data.keys()) if isinstance(excel_data, dict) else []
   
    for section in required_sections:
        # Check if section exists in sheet names (case insensitive)
        if not any(section.lower() in sheet.lower() for sheet in available_sheets):
            missing_sections.append(section)
   
    return missing_sections

def validate_general_info(df):
    """Validate required fields in General Information section"""
    required_fields = [
        'Ultimate Parent Name',
        'Country of Registered Office',
        'Financial Year Start Date',
        'Financial Year End Date',
        'Reporting Currency',
        'OECD Instructions Used'
    ]
   
    missing_fields = []
   
    if df.empty:
        return required_fields
   
    # Check if required fields exist in the dataframe
    for field in required_fields:
        if not any(field.lower() in str(col).lower() for col in df.columns):
            # Also check in the first column values
            if not any(field.lower() in str(val).lower() for val in df.iloc[:, 0].values if pd.notna(val)):
                missing_fields.append(field)
   
    return missing_fields

def validate_country_data(df):
    """Validate required fields in Country-by-Country section"""
    required_fields = [
        'Tax Jurisdiction',
        'Country Code',
        'Revenues',
        'Profit (Loss) Before Tax',
        'Income Tax Paid',
        'Income Tax Accrued',
        'Accumulated Earnings',
        'Number of Employees'
    ]
   
    missing_fields = []
   
    if df.empty:
        return required_fields
   
    for field in required_fields:
        if not any(field.lower() in str(col).lower() for col in df.columns):
            missing_fields.append(field)
   
    return missing_fields

def extract_general_info(df):
    """Extract general information from the dataframe"""
    info = {}
   
    # Try to extract from key-value pairs in first two columns
    if len(df.columns) >= 2:
        for _, row in df.iterrows():
            key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
           
            if "ultimate parent" in key.lower():
                info['ultimate_parent'] = value
            elif "country of registered office" in key.lower():
                info['country_office'] = value
            elif "financial year start" in key.lower():
                info['fy_start'] = value
            elif "financial year end" in key.lower():
                info['fy_end'] = value
            elif "reporting currency" in key.lower():
                info['currency'] = value
            elif "oecd" in key.lower():
                info['oecd_instructions'] = value.lower() in ['yes', 'true', '1']
   
    return info

def format_date(date_str):
    """Format date to YYYY-MM-DD"""
    try:
        # Try different date formats
        for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y']:
            try:
                date_obj = datetime.strptime(str(date_str), fmt)
                return date_obj.strftime('%Y-%m-%d')
            except ValueError:
                continue
        return str(date_str)  # Return as-is if parsing fails
    except:
        return str(date_str)

def generate_xhtml(excel_data):
    """Generate XHTML with iXBRL markup from Excel data"""
   
    # Extract data from different sheets
    general_info_df = None
    country_data_df = None
    subsidiaries_df = None
    omitted_info_df = None
   
    # Find the appropriate sheets
    for sheet_name, df in excel_data.items():
        if 'general' in sheet_name.lower():
            general_info_df = df
        elif 'country' in sheet_name.lower() or 'overview' in sheet_name.lower():
            country_data_df = df
        elif 'subsid' in sheet_name.lower() or 'activities' in sheet_name.lower():
            subsidiaries_df = df
        elif 'omit' in sheet_name.lower():
            omitted_info_df = df
   
    # Extract general information
    general_info = extract_general_info(general_info_df) if general_info_df is not None else {}
   
    # Generate unique entity ID
    entity_id = f"entity_{uuid.uuid4().hex[:8]}"
   
    xhtml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"
      xmlns:ix="http://www.xbrl.org/2013/inlineXBRL"
      xmlns:ixt="http://www.xbrl.org/inlineXBRL/transformation/2020-02-12"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
      xmlns:xbrli="http://www.xbrl.org/2003/instance"
      xmlns:cbcr="http://xbrl.ifrs.org/taxonomy/2024-03-14/ifrs-cbcr">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Country-by-Country Report - {general_info.get('ultimate_parent', 'Company')}</title>
    <ix:header>
        <ix:references>
            <ix:relationship fromDocument="http://xbrl.ifrs.org/taxonomy/2024-03-14/ifrs-cbcr" />
        </ix:references>
        <ix:resources>
            <xbrli:context id="duration">
                <xbrli:entity>
                    <xbrli:identifier scheme="http://www.company-registry.eu">{entity_id}</xbrli:identifier>
                </xbrli:entity>
                <xbrli:period>
                    <xbrli:startDate>{format_date(general_info.get('fy_start', '2025-01-01'))}</xbrli:startDate>
                    <xbrli:endDate>{format_date(general_info.get('fy_end', '2025-12-31'))}</xbrli:endDate>
                </xbrli:period>
            </xbrli:context>
            <xbrli:unit id="currency">
                <xbrli:measure>{general_info.get('currency', 'EUR')}</xbrli:measure>
            </xbrli:unit>
            <xbrli:unit id="pure">
                <xbrli:measure>xbrli:pure</xbrli:measure>
            </xbrli:unit>
        </ix:resources>
    </ix:header>
</head>
<body>
    <h1>Country-by-Country Report</h1>
   
    <h2>Section 1: General Information</h2>
    <table border="1">
        <tr>
            <td>Name of Ultimate Parent Undertaking:</td>
            <td><ix:nonNumeric name="cbcr:NameOfUltimateParentOfGroupOfStandaloneCompany" contextRef="duration">{general_info.get('ultimate_parent', 'N/A')}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Country of Registered Office:</td>
            <td><ix:nonNumeric name="cbcr:CountryOfRegisteredOfficeOfUltimateParentUndertaking" contextRef="duration">{general_info.get('country_office', 'N/A')}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Financial Year Start Date:</td>
            <td><ix:nonNumeric name="cbcr:DateOfStartOfFinancialYear" contextRef="duration">{format_date(general_info.get('fy_start', '2025-01-01'))}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Financial Year End Date:</td>
            <td><ix:nonNumeric name="cbcr:DateOfEndOfFinancialYear" contextRef="duration">{format_date(general_info.get('fy_end', '2025-12-31'))}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Reporting Currency:</td>
            <td><ix:nonNumeric name="cbcr:ReportingCurrency" contextRef="duration">{general_info.get('currency', 'EUR')}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>OECD Instructions Used:</td>
            <td><ix:nonNumeric name="cbcr:ApplicationOfOptionToReportInAccordanceWithTaxationReportingInstructions" contextRef="duration">{'Yes' if general_info.get('oecd_instructions', False) else 'No'}</ix:nonNumeric></td>
        </tr>
    </table>
   
    <h2>Section 2: Overview of Information on a Country-by-Country Basis</h2>
    <table border="1">
        <thead>
            <tr>
                <th>Tax Jurisdiction</th>
                <th>Country Code</th>
                <th>Revenues</th>
                <th>Profit (Loss) Before Tax</th>
                <th>Income Tax Paid</th>
                <th>Income Tax Accrued</th>
                <th>Accumulated Earnings</th>
                <th>Number of Employees</th>
            </tr>
        </thead>
        <tbody>'''
   
    # Add country data rows
    if country_data_df is not None and not country_data_df.empty:
        for _, row in country_data_df.iterrows():
            if pd.notna(row.iloc[0]):  # Skip empty rows
                xhtml_content += f'''
            <tr>
                <td><ix:nonNumeric name="cbcr:TaxJurisdiction" contextRef="duration">{row.iloc[0] if pd.notna(row.iloc[0]) else 'N/A'}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:CountryCodeOfMemberStateOrTaxJurisdiction" contextRef="duration">{row.iloc[1] if pd.notna(row.iloc[1]) else 'N/A'}</ix:nonNumeric></td>
                <td><ix:nonFraction name="cbcr:Revenues" contextRef="duration" unitRef="currency" decimals="0">{int(float(row.iloc[2])) if pd.notna(row.iloc[2]) and str(row.iloc[2]).replace('.','').replace('-','').isdigit() else 0}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:ProfitLossBeforeTax" contextRef="duration" unitRef="currency" decimals="0">{int(float(row.iloc[3])) if pd.notna(row.iloc[3]) and str(row.iloc[3]).replace('.','').replace('-','').isdigit() else 0}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:IncomeTaxPaidOnCashBasis" contextRef="duration" unitRef="currency" decimals="0">{int(float(row.iloc[4])) if pd.notna(row.iloc[4]) and str(row.iloc[4]).replace('.','').replace('-','').isdigit() else 0}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:IncomeTaxAccrued" contextRef="duration" unitRef="currency" decimals="0">{int(float(row.iloc[5])) if pd.notna(row.iloc[5]) and str(row.iloc[5]).replace('.','').replace('-','').isdigit() else 0}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:AccumulatedEarnings" contextRef="duration" unitRef="currency" decimals="0">{int(float(row.iloc[6])) if pd.notna(row.iloc[6]) and str(row.iloc[6]).replace('.','').replace('-','').isdigit() else 0}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:NumberOfEmployees" contextRef="duration" unitRef="pure" decimals="0">{int(float(row.iloc[7])) if pd.notna(row.iloc[7]) and str(row.iloc[7]).replace('.','').isdigit() else 0}</ix:nonFraction></td>
            </tr>'''
   
    xhtml_content += '''
        </tbody>
    </table>
   
    <h2>Section 3: List of Subsidiaries and Activities</h2>
    <table border="1">
        <thead>
            <tr>
                <th>Tax Jurisdiction</th>
                <th>Country Code</th>
                <th>Subsidiary Name</th>
                <th>Nature of Activities</th>
            </tr>
        </thead>
        <tbody>'''
   
    # Add subsidiary data
    if subsidiaries_df is not None and not subsidiaries_df.empty:
        for _, row in subsidiaries_df.iterrows():
            if pd.notna(row.iloc[0]):
                xhtml_content += f'''
            <tr>
                <td><ix:nonNumeric name="cbcr:TaxJurisdiction" contextRef="duration">{row.iloc[0] if pd.notna(row.iloc[0]) else 'N/A'}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:CountryCodeOfMemberStateOrTaxJurisdiction" contextRef="duration">{row.iloc[1] if pd.notna(row.iloc[1]) else 'N/A'}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:DisclosureOfNamesOfSubsidiaryUndertakingsConsolidatedInFinancialStatementsOfUltimateParentUndertakingExplanatory" contextRef="duration">{row.iloc[2] if pd.notna(row.iloc[2]) else 'N/A'}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:DescriptionOfNatureOfActivitiesOfSubsidiaryUndertakingsInMemberStateOrTaxJurisdictionExplanatory" contextRef="duration">{row.iloc[3] if pd.notna(row.iloc[3]) else 'N/A'}</ix:nonNumeric></td>
            </tr>'''
   
    xhtml_content += '''
        </tbody>
    </table>
   
    <h2>Section 4: Omitted Information</h2>
    <div>
        <p><strong>Information Omitted:</strong></p>
        <ix:nonNumeric name="cbcr:DisclosureOfTypeOfInformationOmittedExplanatory" contextRef="duration">'''
   
    if omitted_info_df is not None and not omitted_info_df.empty:
        omitted_text = str(omitted_info_df.iloc[0, 0]) if pd.notna(omitted_info_df.iloc[0, 0]) else "No information omitted"
    else:
        omitted_text = "No information omitted"
   
    xhtml_content += f'''{omitted_text}</ix:nonNumeric>
    </div>
   
    <h2>Section 5: Explanations for Material Discrepancies</h2>
    <div>
        <ix:nonNumeric name="cbcr:ExplanationOfAnyMaterialDiscrepanciesBetweenIncomeTaxPaidAndAccruedExplanatory" contextRef="duration">No material discrepancies identified</ix:nonNumeric>
    </div>
   
    <hr/>
    <p><em>This report was generated in compliance with Commission Implementing Regulation (EU) 2024/2952.</em></p>
</body>
</html>'''
   
    return xhtml_content

# HTML template for the upload form
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EU CbCR Converter</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary-color: #007bff; /* A nice blue */
            --primary-dark: #0056b3;
            --text-color: #333;
            --light-grey: #f8f9fa;
            --medium-grey: #e9ecef;
            --border-color: #dee2e6;
            --shadow-color: rgba(0, 0, 0, 0.08);
            --success-bg: #d4edda;
            --success-text: #155724;
            --error-bg: #f8d7da;
            --error-text: #721c24;
        }

        body {
            font-family: 'Inter', sans-serif;
            line-height: 1.6;
            color: var(--text-color);
            margin: 0;
            padding: 40px 20px;
            background-color: var(--light-grey);
            display: flex;
            justify-content: center;
            align-items: flex-start; /* Align to top for longer content */
            min-height: 100vh;
            box-sizing: border-box;
        }

        .container {
            background-color: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 10px 30px var(--shadow-color);
            max-width: 900px; /* Slightly wider for better content flow */
            width: 100%;
            box-sizing: border-box;
        }

        h1 {
            font-size: 2.5em;
            color: var(--primary-dark);
            text-align: center;
            margin-bottom: 15px;
            font-weight: 700;
        }

        h2 {
            font-size: 1.8em;
            color: var(--text-color);
            margin-top: 30px;
            margin-bottom: 15px;
            border-bottom: 1px solid var(--medium-grey);
            padding-bottom: 8px;
            font-weight: 600;
        }

        h3 {
            font-size: 1.4em;
            margin-top: 25px;
            margin-bottom: 10px;
            color: var(--text-color);
            font-weight: 600;
        }

        p {
            margin-bottom: 15px;
            font-weight: 400;
        }

        .subtitle {
            text-align: center;
            color: #6c757d;
            margin-top: -10px;
            margin-bottom: 30px;
            font-size: 1.1em;
        }

        .info-box, .requirements {
            background-color: var(--light-grey);
            border-left: 5px solid var(--primary-color);
            padding: 20px 25px;
            margin-bottom: 25px;
            border-radius: 8px;
            font-size: 0.95em;
            color: var(--text-color);
            box-shadow: 0 2px 8px rgba(0,0,0,0.05); /* Subtle shadow for depth */
        }

        .requirements {
            border-left: 5px solid #ffc107; /* Yellow for requirements */
            background-color: #fffde7; /* Lighter yellow background */
        }

        .upload-area {
            border: 2px dashed var(--primary-color);
            border-radius: 10px;
            padding: 50px;
            text-align: center;
            margin-bottom: 30px;
            transition: background-color 0.3s ease;
            cursor: pointer; /* Indicate it's clickable */
        }
        .upload-area:hover {
            background-color: #f0f8ff; /* Lighter blue on hover */
        }

        .btn {
            background-color: var(--primary-color);
            color: white;
            padding: 14px 28px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1.1em;
            font-weight: 500;
            transition: background-color 0.3s ease, transform 0.2s ease;
            display: inline-block; /* For centering */
        }
        .btn:hover {
            background-color: var(--primary-dark);
            transform: translateY(-2px); /* Slight lift on hover */
        }

        .file-input-label {
            display: inline-block;
            padding: 10px 20px;
            background-color: #6c757d;
            color: white;
            border-radius: 5px;
            cursor: pointer;
            font-weight: 500;
            transition: background-color 0.3s ease;
        }
        .file-input-label:hover {
            background-color: #5a6268;
        }
        input[type="file"] {
            display: none; /* Hide default file input */
        }
        #file-name {
            margin-top: 15px;
            font-style: italic;
            color: #555;
            font-size: 0.9em;
        }

        .message {
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            font-weight: 500;
            font-size: 0.95em;
        }
        .error {
            background-color: var(--error-bg);
            color: var(--error-text);
            border: 1px solid #f5c6cb;
        }
        .success {
            background-color: var(--success-bg);
            color: var(--success-text);
            border: 1px solid #c3e6cb;
        }

        .footer-info {
            margin-top: 40px;
            padding-top: 25px;
            border-top: 1px solid var(--border-color);
            font-size: 0.9em;
            color: #6c757d;
        }
        .footer-info ul {
            list-style: none;
            padding: 0;
            margin-top: 10px;
        }
        .footer-info li {
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>EU CbCR Converter</h1>
        <p class="subtitle">Streamlining Country-by-Country Reporting for EU Tax Compliance</p>
       
        <div class="info-box">
            <p><strong>Purpose:</strong> This tool simplifies compliance with the new EU public Country-by-Country Reporting (CbCR) directive (Commission Implementing Regulation (EU) 2024/2952). It converts your Excel-based income tax data into the required XHTML format with Inline XBRL (iXBRL) markups.</p>
        </div>
       
        <div class="requirements">
            <h3>Excel File Structure Requirements:</h3>
            <ul>
                <li><strong>Sheet 1: "General Information"</strong> &ndash; Details about your ultimate parent undertaking, financial year, and reporting currency.</li>
                <li><strong>Sheet 2: "Country-by-Country Overview"</strong> &ndash; Essential tax data for each tax jurisdiction.</li>
                <li><strong>Sheet 3: "Subsidiaries and Activities"</strong> &ndash; A comprehensive list of your subsidiaries and their primary business activities.</li>
                <li><strong>Sheet 4: "Omitted Information"</strong> &ndash; A clear disclosure of any information that has been intentionally omitted from the report.</li>
            </ul>
            <p><strong>Important:</strong> All four sections are mandatory. Missing sections will result in an error.</p>
        </div>
       
        {% with messages = get_flashed_messages() %}
            {% if messages %}
                {% for message in messages %}
                    <div class="message error">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}
       
        <form method="post" enctype="multipart/form-data">
            <div class="upload-area" onclick="document.getElementById('file').click()">
                <h3>Upload Your Excel File</h3>
                <p>Click here or drag & drop your .xlsx or .xls file.</p>
                <input type="file" name="file" id="file" accept=".xlsx,.xls" required>
                <p id="file-name"></p>
                <div style="margin-top: 20px;">
                    <button type="submit" class="btn">Convert to XHTML</button>
                </div>
            </div>
        </form>
       
        <div class="footer-info">
            <p><strong>Regulatory Compliance Details:</strong></p>
            <ul>
                <li>Output adheres to Commission Implementing Regulation (EU) 2024/2952.</li>
                <li>Utilizes the Inline XBRL (iXBRL) 1.1 specification for digital reporting.</li>
                <li>Applicable for financial years commencing on or after 1 January 2025.</li>
                <li>Mandatory for multinational undertakings with consolidated revenues exceeding EUR 750 million.</li>
            </ul>
        </div>
    </div>
   
    <script>
        document.getElementById('file').addEventListener('change', function(e) {
            const fileName = e.target.files[0]?.name;
            document.getElementById('file-name').textContent = fileName ? `Selected file: ${fileName}` : '';
        });

        // Prevent default drag and drop behavior
        const uploadArea = document.querySelector('.upload-area');
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = '#e0efff'; // Highlight on drag over
        });
        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = 'var(--light-grey)'; // Reset on drag leave
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = 'var(--light-grey)'; // Reset on drop
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                document.getElementById('file').files = files;
                document.getElementById('file-name').textContent = `Selected file: ${files[0].name}`;
            }
        });
    </script>
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file selected')
            return redirect(request.url)
       
        file = request.files['file']
        if file.filename == '':
            flash('No file selected')
            return redirect(request.url)
       
        if file and allowed_file(file.filename):
            try:
                # Read Excel file
                excel_data = pd.read_excel(file, sheet_name=None)
               
                # Validate required sections
                missing_sections = validate_required_sections(excel_data)
                if missing_sections:
                    flash(f'Missing required sections: {", ".join(missing_sections)}. Please ensure your Excel file contains sheets for: General Information, Country-by-Country Overview, Subsidiaries and Activities, and Omitted Information.')
                    return redirect(request.url)
               
                # Additional validation for required fields
                errors = []
               
                # Validate General Information
                general_sheet = None
                for sheet_name, df in excel_data.items():
                    if 'general' in sheet_name.lower():
                        general_sheet = df
                        break
               
                if general_sheet is not None:
                    missing_general = validate_general_info(general_sheet)
                    if missing_general:
                        errors.append(f'Missing fields in General Information: {", ".join(missing_general)}')
               
                # Validate Country-by-Country data
                country_sheet = None
                for sheet_name, df in excel_data.items():
                    if 'country' in sheet_name.lower() or 'overview' in sheet_name.lower():
                        country_sheet = df
                        break
               
                if country_sheet is not None:
                    missing_country = validate_country_data(country_sheet)
                    if missing_country:
                        errors.append(f'Missing fields in Country-by-Country Overview: {", ".join(missing_country)}')
               
                if errors:
                    for error in errors:
                        flash(error)
                    return redirect(request.url)
               
                # Generate XHTML
                xhtml_content = generate_xhtml(excel_data)
               
                # Create temporary file for download
                temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.xhtml', delete=False, encoding='utf-8')
                temp_file.write(xhtml_content)
                temp_file.close()
               
                # Send file for download
                return send_file(
                    temp_file.name,
                    as_attachment=True,
                    download_name=f'country_by_country_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xhtml',
                    mimetype='application/xhtml+xml'
                )
               
            except Exception as e:
                flash(f'Error processing file: {str(e)}')
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)')
            return redirect(request.url)
   
    return render_template_string(HTML_TEMPLATE)

# Remove the following lines:
# def handler(request):
#     return app(request.environ, lambda *args: None)

# if __name__ == '__main__':
#     app.run(debug=True)
