from flask import Flask, render_template, request, send_file
from datetime import datetime
import zipfile
import os
import shutil
import tempfile
from docx import Document

app = Flask(__name__)

TEMPLATE_PATH = "template.docx"


def ordinal(n):
    """Add ordinal suffix to day number: 1 -> 1st, 2 -> 2nd, 3 -> 3rd, 4 -> 4th"""
    if 11 <= n % 100 <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"


def format_date_ordinal(date_str):
    """Convert YYYY-MM-DD to '23rd April 2026' format"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{ordinal(dt.day)} {dt.strftime('%B %Y')}"


def merge_runs_in_paragraph(paragraph):
    """Merge all runs in a paragraph into a single run, preserving text."""
    if len(paragraph.runs) == 0:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = full_text
        else:
            run.text = ""


def replace_in_document(doc, replacements):
    """
    Replace all {{PLACEHOLDER}} in document by:
    1. Merging runs in each paragraph (fixes Word's split-run issue)
    2. Doing simple text replacement
    """
    for paragraph in doc.paragraphs:
        merge_runs_in_paragraph(paragraph)
        if len(paragraph.runs) > 0:
            for key, value in replacements.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.runs[0].text:
                    paragraph.runs[0].text = paragraph.runs[0].text.replace(placeholder, str(value))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    merge_runs_in_paragraph(paragraph)
                    if len(paragraph.runs) > 0:
                        for key, value in replacements.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in paragraph.runs[0].text:
                                paragraph.runs[0].text = paragraph.runs[0].text.replace(placeholder, str(value))
    
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            merge_runs_in_paragraph(paragraph)
            if len(paragraph.runs) > 0:
                for key, value in replacements.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in paragraph.runs[0].text:
                        paragraph.runs[0].text = paragraph.runs[0].text.replace(placeholder, str(value))
        
        for paragraph in section.footer.paragraphs:
            merge_runs_in_paragraph(paragraph)
            if len(paragraph.runs) > 0:
                for key, value in replacements.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in paragraph.runs[0].text:
                        paragraph.runs[0].text = paragraph.runs[0].text.replace(placeholder, str(value))


def generate_contract(template_path, replacements):
    """Generate contract by loading template, replacing placeholders, saving output."""
    doc = Document(template_path)
    replace_in_document(doc, replacements)
    temp_dir = tempfile.mkdtemp()
    output_path = os.path.join(temp_dir, "output.docx")
    doc.save(output_path)
    return output_path


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        name = request.form.get('name', '').strip()
        agreement_date = request.form.get('agreement_date', '').strip()
        start_date = request.form.get('start_date', '').strip()
        end_date = request.form.get('end_date', '').strip()
        rent = request.form.get('rent', '').strip()
        deposit = request.form.get('deposit', '').strip()
        room = request.form.get('room', '').strip()
        property_addr = request.form.get('property', '123 Daren Avenue SE1 4FR').strip()
        ref = request.form.get('ref', '').strip()
        
        if not all([name, start_date, rent, deposit, room]):
            return "All fields are required", 400
        
        # Format dates WITH ordinal suffixes (1st, 2nd, 3rd, 4th)
        if agreement_date:
            formatted_agreement_date = format_date_ordinal(agreement_date)
        else:
            formatted_agreement_date = format_date_ordinal(datetime.now().strftime("%Y-%m-%d"))
        
        formatted_start = format_date_ordinal(start_date)
        formatted_end = format_date_ordinal(end_date) if end_date else "To be agreed"
        
        # Format currency
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
        
        # Auto-generate reference if blank
        if not ref:
            ref = f"KPI-{name.replace(' ', '-').upper()}-{datetime.now().strftime('%Y%m%d')}"
        
        # ALL placeholders
        replacements = {
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
        
        if not os.path.exists(TEMPLATE_PATH):
            return f"Template not found: {os.path.abspath(TEMPLATE_PATH)}", 500
        
        output_path = generate_contract(TEMPLATE_PATH, replacements)
        
        safe_name = name.replace(" ", "_").replace("/", "_")
        download_name = f"Licence_Agreement_{safe_name}_{datetime.now().strftime('%Y%m%d')}.docx"
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        import traceback
        return f"<pre>Error: {str(e)}\n\n{traceback.format_exc()}</pre>", 500


if __name__ == '__main__':
    app.run(debug=True)
