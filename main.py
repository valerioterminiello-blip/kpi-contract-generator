from flask import Flask, render_template, request, send_file
from datetime import datetime
import zipfile
import os
import shutil
import tempfile

app = Flask(__name__)

TEMPLATE_PATH = "template.docx"

def replace_all_placeholders_in_xml(xml_path, replacements):
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    for key, value in replacements.items():
        placeholder = f"{{{{{key}}}}}"
        content = content.replace(placeholder, str(value))
    
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(content)


def generate_contract(template_path, replacements):
    temp_dir = tempfile.mkdtemp(prefix="docx_")
    
    try:
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_files = ["word/document.xml"]
        word_dir = os.path.join(temp_dir, "word")
        if os.path.exists(word_dir):
            for f in os.listdir(word_dir):
                if f.endswith(".xml"):
                    xml_files.append(f"word/{f}")
        
        for xml_file in xml_files:
            file_path = os.path.join(temp_dir, xml_file)
            if os.path.exists(file_path):
                replace_all_placeholders_in_xml(file_path, replacements)
        
        output_path = os.path.join(temp_dir, "output.docx")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file == "output.docx":
                        continue
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)
        
        return output_path
        
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise e


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
        
        # Format agreement date (the date the form is filled out / contract is dated)
        if agreement_date:
            try:
                dt = datetime.strptime(agreement_date, "%d/%m/%Y")
                formatted_agreement_date = dt.strftime("%d %B %Y")
            except ValueError:
                formatted_agreement_date = agreement_date
        else:
            # Default to today if not provided
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
            "DATE": formatted_agreement_date,   # User-selected date (or today)
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
