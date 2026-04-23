from flask import Flask, render_template, request, send_file
from datetime import datetime
import zipfile
import os
import shutil
import tempfile
from docx import Document

app = Flask(__name__)

TEMPLATE_PATH = "template.docx"


def merge_runs_in_paragraph(paragraph):
    """Merge all runs in a paragraph into a single run, preserving text."""
    if len(paragraph.runs) == 0:
        return

    full_text = "".join(run.text for run in paragraph.runs)

    # Clear all runs except first
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
    # Process all paragraphs in main document
    for paragraph in doc.paragraphs:
        merge_runs_in_paragraph(paragraph)
        if len(paragraph.runs) > 0:
            for key, value in replacements.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.runs[0].text:
                    paragraph.runs[0].text = paragraph.runs[0].text.replace(placeholder, str(value))

    # Process all tables
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

    # Process headers and footers
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

    # Save to temp file
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

        # Format agreement date
        if agreement_date:
            try:
                dt = datetime.strptime(agreement_date, "%d/%m/%Y")
                formatted_agreement_date = dt.strftime("%d %B %Y")
            except ValueError:
                formatted_agreement_date = agreement_date
        else:
            formatted_agreement_date = datetime.now().strftime("%d %B %Y")

        # Format start date
        try:
            dt = datetime.strptime(start_date, "%d/%m/%Y")
            formatted_start = dt.strftime("%d %B %Y")
        except ValueError:
            formatted_start = start_date

        # Format end date
        try:
            dt_end = datetime.strptime(end_date, "%d/%m/%Y")
            formatted_end = dt_end.strftime("%d %B %Y")
        except ValueError:
            formatted_end = end_date if end_date else "To be agreed"

        # Format currency
        try:
            rent_clean = rent.replace("Â£", "").replace(",", "").strip()
            rent_formatted = f"{float(rent_clean):,.2f}"
        except:
            rent_formatted = rent

        try:
            dep_clean = deposit.replace("Â£", "").replace(",", "").strip()
            dep_formatted = f"{float(dep_clean):,.2f}"
        except:
            dep_formatted = deposit

        # Auto-generate reference if blank
        if not ref:
            ref = f"KPI-{name.replace(' ', '-').upper()}-{datetime.now().strftime('%Y%m%d')}"

        # ALL placeholders - will replace EVERY instance in the document
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
