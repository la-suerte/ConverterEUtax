import os
import pandas as pd
from flask import Flask, request, render_template_string, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import tempfile
import uuid
from datetime import datetime
import re
from xml.sax.saxutils import escape # <-- ADDED: Import for XML escaping

app = Flask(__name__)
app.secret_key = 'your-secret-key-here' # IMPORTANT: Change this to a strong, unique secret key for production

# Configuration
UPLOAD_FOLDER = '/tmp'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename):
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_required_sections(excel_data):
    """
    Validates that all required sections (sheet names) are present in the Excel file.
    Sheet names are checked case-insensitively.
    """
    required_sections = [
        'General Information',
        'Country-by-Country Overview',
        'Subsidiaries and Activities',
        'Omitted Information'
    ]
    
    missing_sections = []
    available_sheets = list(excel_data.keys()) if isinstance(excel_data, dict) else []
    
    for section in required_sections:
        if not any(section.lower() in sheet.lower() for sheet in available_sheets):
            missing_sections.append(section)
    
    return missing_sections

def validate_general_info(df):
    """
    Validates required fields in the 'General Information' section.
    Checks for field names in the first row (headers) or first column (key-value pairs).
    """
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
    
    for field in required_fields:
        if not any(field.lower() in str(col).lower() for col in df.columns):
            if not any(field.lower() in str(val).lower() for val in df.iloc[:, 0].values if pd.notna(val)):
                missing_fields.append(field)
    
    return missing_fields

def validate_country_data(df):
    """
    Validates required fields (column headers) in the 'Country-by-Country Overview' section.
    """
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
    """
    Extracts general information from the 'General Information' DataFrame.
    Assumes data is in key-value pairs in the first two columns.
    """
    info = {}
    print("\n--- DEBUG: Extracting General Info ---") # DEBUG
    if df is not None and len(df.columns) >= 2:
        for i, row in df.iterrows():
            key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # DEBUG: Print raw key-value pair from General Info sheet
            print(f"  Raw GI Row {i}: Key='{key}', Value='{value}'")

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
    print("--- DEBUG: Extracted General Info Dict ---") # DEBUG
    print(info) # DEBUG
    return info

def format_date(date_str):
    """
    Formats a date string into 'YYYY-MM-DD'.
    Tries multiple common date formats for parsing.
    """
    try:
        date_formats = ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d']
        
        if isinstance(date_str, pd.Timestamp):
            return date_str.strftime('%Y-%m-%d')
        
        s_date_str = str(date_str)

        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(s_date_str, fmt)
                return date_obj.strftime('%Y-%m-%d')
            except ValueError:
                continue
        return s_date_str
    except Exception:
        return str(date_str)

def generate_xhtml(excel_data):
    """
    Generates XHTML content with iXBRL markup from the parsed Excel data.
    All string data inserted into the XHTML is XML-escaped to prevent parsing errors.
    """
    print("\n--- DEBUG: Starting generate_xhtml ---") # DEBUG

    general_info_df = None
    country_data_df = None
    subsidiaries_df = None
    omitted_info_df = None
    
    for sheet_name, df in excel_data.items():
        if 'general' in sheet_name.lower():
            general_info_df = df
        elif 'country' in sheet_name.lower() or 'overview' in sheet_name.lower():
            country_data_df = df
        elif 'subsid' in sheet_name.lower() or 'activities' in sheet_name.lower():
            subsidiaries_df = df
        elif 'omit' in sheet_name.lower():
            omitted_info_df = df
    
    general_info = extract_general_info(general_info_df) if general_info_df is not None else {}
    
    entity_id = f"entity_{uuid.uuid4().hex[:8]}"
    
    # DEBUG: Print and escape general info items one by one
    raw_parent_name = str(general_info.get('ultimate_parent', 'N/A'))
    escaped_parent_name = escape(raw_parent_name)
    print(f"  DEBUG GI: Raw Parent Name='{raw_parent_name}', Escaped='{escaped_parent_name}'")

    raw_country_office = str(general_info.get('country_office', 'N/A'))
    escaped_country_office = escape(raw_country_office)
    print(f"  DEBUG GI: Raw Country Office='{raw_country_office}', Escaped='{escaped_country_office}'")

    raw_fy_start = str(format_date(general_info.get('fy_start', '2025-01-01')))
    escaped_fy_start = escape(raw_fy_start) # Dates typically don't need escaping, but for consistency
    print(f"  DEBUG GI: Raw FY Start='{raw_fy_start}', Escaped='{escaped_fy_start}'")

    raw_fy_end = str(format_date(general_info.get('fy_end', '2025-12-31')))
    escaped_fy_end = escape(raw_fy_end)
    print(f"  DEBUG GI: Raw FY End='{raw_fy_end}', Escaped='{escaped_fy_end}'")
    
    raw_currency = str(general_info.get('currency', 'EUR'))
    escaped_currency = escape(raw_currency)
    print(f"  DEBUG GI: Raw Currency='{raw_currency}', Escaped='{escaped_currency}'")

    raw_oecd_instructions = 'Yes' if general_info.get('oecd_instructions', False) else 'No'
    escaped_oecd_instructions = escape(raw_oecd_instructions) # 'Yes'/'No' don't need escaping but for consistency
    print(f"  DEBUG GI: Raw OECD='{raw_oecd_instructions}', Escaped='{escaped_oecd_instructions}'")

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
    <title>Country-by-Country Report - {escaped_parent_name}</title>
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
                    <xbrli:startDate>{escaped_fy_start}</xbrli:startDate>
                    <xbrli:endDate>{escaped_fy_end}</xbrli:endDate>
                </xbrli:period>
            </xbrli:context>
            <xbrli:unit id="currency">
            <xbrli:measure>iso4217:{escaped_currency}</xbrli:measure>
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
            <td><ix:nonNumeric name="cbcr:NameOfUltimateParentOfGroupOfStandaloneCompany" contextRef="duration">{escaped_parent_name}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Country of Registered Office:</td>
            <td><ix:nonNumeric name="cbcr:CountryOfRegisteredOfficeOfUltimateParentUndertaking" contextRef="duration">{escaped_country_office}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Financial Year Start Date:</td>
            <td><ix:nonNumeric name="cbcr:DateOfStartOfFinancialYear" contextRef="duration" format="ixt:date-day-month-year">{escaped_fy_start}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Financial Year End Date:</td>
            <td><ix:nonNumeric name="cbcr:DateOfEndOfFinancialYear" contextRef="duration" format="ixt:date-day-month-year">{escaped_fy_end}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>Reporting Currency:</td>
            <td><ix:nonNumeric name="cbcr:ReportingCurrency" contextRef="duration">{escaped_currency}</ix:nonNumeric></td>
        </tr>
        <tr>
            <td>OECD Instructions Used:</td>
            <td><ix:nonNumeric name="cbcr:ApplicationOfOptionToReportInAccordanceWithTaxationReportingInstructions" contextRef="duration">{escaped_oecd_instructions}</ix:nonNumeric></td>
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
    
    if country_data_df is not None and not country_data_df.empty:
        print("\n  --- DEBUG: Processing Country Data ---") # DEBUG
        for i, row in country_data_df.iterrows():
            if pd.notna(row.iloc[0]): 
                print(f"    Country Data Row {i}:") # DEBUG
                raw_jurisdiction = str(row.iloc[0]) if pd.notna(row.iloc[0]) else 'N/A'
                jurisdiction = escape(raw_jurisdiction)
                print(f"      Raw Jurisdiction='{raw_jurisdiction}', Escaped='{jurisdiction}'") # DEBUG

                raw_country_code = str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A'
                country_code = escape(raw_country_code)
                print(f"      Raw Country Code='{raw_country_code}', Escaped='{country_code}'") # DEBUG
                
                revenues = int(float(row.iloc[2])) if pd.notna(row.iloc[2]) and str(row.iloc[2]).replace('.','',1).replace('-','',1).isdigit() else 0 # Adjusted isdigit check slightly for floats
                profit_loss = int(float(row.iloc[3])) if pd.notna(row.iloc[3]) and str(row.iloc[3]).replace('.','',1).replace('-','',1).isdigit() else 0
                tax_paid = int(float(row.iloc[4])) if pd.notna(row.iloc[4]) and str(row.iloc[4]).replace('.','',1).replace('-','',1).isdigit() else 0
                tax_accrued = int(float(row.iloc[5])) if pd.notna(row.iloc[5]) and str(row.iloc[5]).replace('.','',1).replace('-','',1).isdigit() else 0
                accum_earnings = int(float(row.iloc[6])) if pd.notna(row.iloc[6]) and str(row.iloc[6]).replace('.','',1).replace('-','',1).isdigit() else 0
                num_employees = int(float(row.iloc[7])) if pd.notna(row.iloc[7]) and str(row.iloc[7]).replace('.','',1).isdigit() else 0 # Employees are usually whole numbers

                xhtml_content += f'''
            <tr>
                <td><ix:nonNumeric name="cbcr:TaxJurisdiction" contextRef="duration">{jurisdiction}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:CountryCodeOfMemberStateOrTaxJurisdiction" contextRef="duration">{country_code}</ix:nonNumeric></td>
                <td><ix:nonFraction name="cbcr:Revenues" contextRef="duration" unitRef="currency" decimals="0" scale="0">{revenues}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:ProfitLossBeforeTax" contextRef="duration" unitRef="currency" decimals="0" scale="0">{profit_loss}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:IncomeTaxPaidOnCashBasis" contextRef="duration" unitRef="currency" decimals="0" scale="0">{tax_paid}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:IncomeTaxAccrued" contextRef="duration" unitRef="currency" decimals="0" scale="0">{tax_accrued}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:AccumulatedEarnings" contextRef="duration" unitRef="currency" decimals="0" scale="0">{accum_earnings}</ix:nonFraction></td>
                <td><ix:nonFraction name="cbcr:NumberOfEmployees" contextRef="duration" unitRef="pure" decimals="0" scale="0">{num_employees}</ix:nonFraction></td>
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
    
    if subsidiaries_df is not None and not subsidiaries_df.empty:
        print("\n  --- DEBUG: Processing Subsidiaries Data ---") # DEBUG
        for i, row in subsidiaries_df.iterrows():
            if pd.notna(row.iloc[0]):
                print(f"    Subsidiary Data Row {i}:") # DEBUG

                raw_sub_jurisdiction = str(row.iloc[0]) if pd.notna(row.iloc[0]) else 'N/A'
                sub_jurisdiction = escape(raw_sub_jurisdiction)
                print(f"      Raw Sub Jurisdiction='{raw_sub_jurisdiction}', Escaped='{sub_jurisdiction}'") # DEBUG

                raw_sub_country_code = str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A'
                sub_country_code = escape(raw_sub_country_code)
                print(f"      Raw Sub Country Code='{raw_sub_country_code}', Escaped='{sub_country_code}'") # DEBUG
                
                raw_subsidiary_name = str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A'
                subsidiary_name = escape(raw_subsidiary_name)
                print(f"      Raw Subsidiary Name='{raw_subsidiary_name}', Escaped='{subsidiary_name}'") # DEBUG

                raw_nature_of_activities = str(row.iloc[3]) if pd.notna(row.iloc[3]) else 'N/A'
                nature_of_activities = escape(raw_nature_of_activities)
                print(f"      Raw Nature of Activities='{raw_nature_of_activities}', Escaped='{nature_of_activities}'") # DEBUG

                xhtml_content += f'''
            <tr>
                <td><ix:nonNumeric name="cbcr:TaxJurisdiction" contextRef="duration">{sub_jurisdiction}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:CountryCodeOfMemberStateOrTaxJurisdiction" contextRef="duration">{sub_country_code}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:DisclosureOfNamesOfSubsidiaryUndertakingsConsolidatedInFinancialStatementsOfUltimateParentUndertakingExplanatory" contextRef="duration">{subsidiary_name}</ix:nonNumeric></td>
                <td><ix:nonNumeric name="cbcr:DescriptionOfNatureOfActivitiesOfSubsidiaryUndertakingsInMemberStateOrTaxJurisdictionExplanatory" contextRef="duration">{nature_of_activities}</ix:nonNumeric></td>
            </tr>'''
    
    xhtml_content += '''
        </tbody>
    </table>
    
    <h2>Section 4: Omitted Information</h2>
    <div>
        <p><strong>Information Omitted:</strong></p>
        <ix:nonNumeric name="cbcr:DisclosureOfTypeOfInformationOmittedExplanatory" contextRef="duration">'''
    
    if omitted_info_df is not None and not omitted_info_df.empty:
        raw_omitted_text = str(omitted_info_df.iloc[0, 0]) if pd.notna(omitted_info_df.iloc[0, 0]) else "No information omitted"
    else:
        raw_omitted_text = "No information omitted"
    
    omitted_text = escape(raw_omitted_text)
    print(f"\n  --- DEBUG: Omitted Info ---") # DEBUG
    print(f"    Raw Omitted Text='{raw_omitted_text}', Escaped='{omitted_text}'") # DEBUG
    
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
    
    print("\n--- DEBUG: Finished generate_xhtml ---") # DEBUG
    return xhtml_content

# HTML template for the upload form (remains the same)
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

        const uploadArea = document.querySelector('.upload-area');
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = '#e0efff'; 
        });
        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = 'transparent'; // Reset, or to initial if different from white
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.backgroundColor = 'transparent'; 
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
                excel_data = pd.read_excel(file, sheet_name=None)
                
                missing_sections = validate_required_sections(excel_data)
                if missing_sections:
                    flash(f'Missing required sections: {", ".join(missing_sections)}. Please ensure your Excel file contains sheets for: General Information, Country-by-Country Overview, Subsidiaries and Activities, and Omitted Information.')
                    return redirect(request.url)
                
                errors = []
                
                general_sheet = None
                for sheet_name, df in excel_data.items():
                    if 'general' in sheet_name.lower():
                        general_sheet = df
                        break
                
                if general_sheet is not None:
                    missing_general = validate_general_info(general_sheet)
                    if missing_general:
                        errors.append(f'Missing fields in General Information: {", ".join(missing_general)}')
                
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
                
                xhtml_content = generate_xhtml(excel_data)
                
                # DEBUG: Print the full XHTML content to console
                print("\n----------- BEGIN GENERATED XHTML CONTENT (DEBUG) -----------")
                print(xhtml_content)
                print("------------ END GENERATED XHTML CONTENT (DEBUG) ------------\n")

                # DEBUG: Optionally save to a file for easier inspection if console output is too much
                # with open("debug_output.xhtml", "w", encoding="utf-8") as f_debug:
                #     f_debug.write(xhtml_content)
                # print("DEBUG: XHTML content also saved to debug_output.xhtml")
                # End DEBUG section

                temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.xhtml', delete=False, encoding='utf-8')
                temp_file.write(xhtml_content)
                temp_file.close()

                return send_file(
                    temp_file.name,
                    as_attachment=True,
                    download_name=f'country_by_country_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xhtml',
                    mimetype='application/xhtml+xml'
                )
                
            except Exception as e:
                flash(f'Error processing file: {str(e)}')
                # DEBUG: Print full traceback for server-side debugging
                import traceback
                print("--- ERROR TRACEBACK ---")
                traceback.print_exc()
                print("-----------------------")
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)')
            return redirect(request.url)
    
    return render_template_string(HTML_TEMPLATE)

# For local development, if you are not using Vercel's `vercel dev`
if __name__ == '__main__':
    app.run(debug=True) # Make sure debug is True to see print statements and auto-reload
