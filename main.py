from flask import Flask, render_template, request, send_file
from datetime import datetime
import zipfile
import os
import shutil
import tempfile
import re
from docx import Document

app = Flask(__name__)

AGREEMENT_TEMPLATE = "template.docx"
IRR_TEMPLATE = "irrtemplate.docx"
"


def ordinal(n):
    if 11 <= n % 100 <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"


def format_date_ordinal(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{ordinal(dt.day)} {dt.strftime('%B %Y')}"


def merge_runs_in_paragraph(paragraph):
    if len(paragraph.runs) == 0:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = full_text
        else:
            run.text = ""


def replace_placeholders(text, replacements):
    def replace_match(match):
        key = match.group(1).strip()
        if key in replacements:
            return str(replacements[key])
        return match.group(0)
    return re.sub(r'\{\{([^}]+)\}\}', replace_match, text)


def replace_in_document(doc, replacements):
    for paragraph in doc.paragraphs:
        merge_runs_in_paragraph(paragraph)
        if len(paragraph.runs) > 0:
            paragraph.runs[0].text = replace_placeholders(paragraph.runs[0].text, replacements)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    merge_runs_in_paragraph(paragraph)
                    if len(paragraph.runs) > 0:
                        paragraph.runs[0].text = replace_placeholders(paragraph.runs[0].text, replacements)
    
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            merge_runs_in_paragraph(paragraph)
            if len(paragraph.runs) > 0:
                paragraph.runs[0].text = replace_placeholders(paragraph.runs[0].text, replacements)
        
        for paragraph in section.footer.paragraphs:
            merge_runs_in_paragraph(paragraph)
            if len(paragraph.runs) > 0:
                paragraph.runs[0].text = replace_placeholders(paragraph.runs[0].text, replacements)


def generate_document(template_path, replacements, output_name):
    doc = Document(template_path)
    replace_in_document(doc, replacements)
    temp_dir = tempfile.mkdtemp()
    output_path = os.path.join(temp_dir, output_name)
    doc.save(output_path)
    return output_path


def calculate_weekly_rent(monthly_rent):
    try:
        monthly = float(monthly_rent.replace("£", "").replace(",", "").strip())
        weekly = (monthly * 12) / 52
        return f"{weekly:.2f}"
    except:
        return "0.00"


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        mobile = request.form.get('mobile', '').strip()
        kin_name = request.form.get('kin_name', '').strip()
        kin_phone = request.form.get('kin_phone', '').strip()
        kin_email = request.form.get('kin_email', '').strip()
        employer = request.form.get('employer', '').strip()
        agreement_date = request.form.get('agreement_date', '').strip()
        start_date = request.form.get('start_date', '').strip()
        end_date = request.form.get('end_date', '').strip()
        rent = request.form.get('rent', '').strip()
        deposit = request.form.get('deposit', '').strip()
        room = request.form.get('room', '').strip()
        property_addr = request.form.get('property', '123 Daren Avenue SE1 4FR').strip()
        term_months = request.form.get('term_months', '12').strip()
        utilities = request.form.get('utilities', '250').strip()
        ref = request.form.get('ref', '').strip()
        
        if not all([name, start_date, rent, deposit, room]):
            return "All required fields must be filled", 400
        
        surname = name.split()[-1] if name else "Tenant"
        
        if agreement_date:
            formatted_agreement_date = format_date_ordinal(agreement_date)
        else:
            formatted_agreement_date = format_date_ordinal(datetime.now().strftime("%Y-%m-%d"))
        
        formatted_start = format_date_ordinal(start_date)
        formatted_end = format_date_ordinal(end_date) if end_date else "To be agreed"
        
        try:
            rent_clean = rent.replace("£", "").replace(",", "").strip()
            rent_formatted = f"{float(rent_clean):,.2f}"
        except:
            rent_formatted = rent
        
        try:
            dep_clean = deposit.replace("£", "").replace(",", "").strip()
            dep_formatted = f"{float(dep_clean):,.2f}"
        except:
            dep_formatted = deposit
        
        weekly_rent = calculate_weekly_rent(rent_formatted)
        
        try:
            util_clean = utilities.replace("£", "").replace(",", "").strip()
            utilities_formatted = f"{float(util_clean):,.2f}"
        except:
            utilities_formatted = utilities
        
        if not ref:
            ref = f"{surname.upper()}.{property_addr.split()[0].upper()}"
        
        kin_full = f"{kin_name}"
        if kin_phone:
            kin_full += f" - {kin_phone}"
        if kin_email:
            kin_full += f" - {kin_email}"
        
        # Generate Agreement
        agreement_replacements = {
            "DATE": formatted_agreement_date,
            "NAME": name,
            "RENT": rent_formatted,
            "START": formatted_start,
            "DEPOSIT": dep_formatted,
            "ROOM": room.upper(),
            "PROPERTY": property_addr,
            "ADDRESS": property_addr,
            "END": formatted_end,
            "REF": ref,
        }
        
        if not os.path.exists(AGREEMENT_TEMPLATE):
            return f"Agreement template not found: {os.path.abspath(AGREEMENT_TEMPLATE)}", 500
        
        agreement_path = generate_document(
            AGREEMENT_TEMPLATE, 
            agreement_replacements, 
            f"Agreement{surname}.docx"
        )
        
        # Generate IRR
        irr_replacements = {
            "DATE": formatted_agreement_date,
            "NAME": name,
            "EMAIL": email,
            "MOBILE": mobile,
            "KIN": kin_full,
            "EMPLOYER": employer,
            "ADDRESS": property_addr,
            "ROOM": room.upper(),
            "PW": weekly_rent,
            "PCM": rent_formatted,
            "MONTHS": term_months,
            "START": formatted_start,
            "END": formatted_end,
        }
        
        if not os.path.exists(IRR_TEMPLATE):
            return f"IRR template not found: {os.path.abspath(IRR_TEMPLATE)}", 500
        
        irr_path = generate_document(
            IRR_TEMPLATE, 
            irr_replacements, 
            f"IRR{surname}.docx"
        )
        
        # ZIP BOTH FILES TOGETHER
        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, f"KPI_{surname}_Documents.zip")
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(agreement_path, f"Agreement{surname}.docx")
            zipf.write(irr_path, f"IRR{surname}.docx")
        
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f"KPI_{surname}_Documents.zip",
            mimetype='application/zip'
        )
        
    except Exception as e:
        import traceback
        return f"<pre>Error: {str(e)}\n\n{traceback.format_exc()}</pre>", 500


if __name__ == '__main__':
    app.run(debug=True)
