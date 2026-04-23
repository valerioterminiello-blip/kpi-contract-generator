from flask import Flask, render_template, request, send_file
from datetime import datetime
import zipfile
import os
import shutil
import tempfile

app = Flask(__name__)

TEMPLATE_PATH = "licence_template.docx"

def generate_contract(template_path, replacements):
    temp_dir = tempfile.mkdtemp(prefix="docx_")
    
    try:
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_files = ["word/document.xml"]
        word_dir = os.path.join(temp_dir, "word")
        if os.path.exists(word_dir):
            for f in os.listdir(word_dir):
                if f.startswith(("header", "footer")) and f.endswith(".xml"):
                    xml_files.append(f"word/{f}")
        
        for xml_file in xml_files:
            file_path = os.path.join(temp_dir, xml_file)
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                for key, value in replacements.items():
                    placeholder = "{{" + key + "}}"
                    content = content.replace(placeholder, str(value))
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
        
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
        start_date = request.form.get('start_date', '').strip()
        rent = request.form.get('rent', '').strip()
        deposit = request.form.get('deposit', '').strip()
        room = request.form.get('room', '').strip()
        property_addr = request.form.get('property', '123 Daren Avenue SE1 4FR').strip()
        
        if not all([name, start_date, rent, deposit, room]):
            return "All fields are required", 400
        
        try:
            dt = datetime.strptime(start_date, "%d/%m/%Y")
            formatted_date = dt.strftime("%d %B %Y")
        except ValueError:
            formatted_date = start_date
        
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
        
        replacements = {
            "DATE": datetime.now().strftime("%d %B %Y"),
            "NAME": name,
            "RENT": rent_formatted,
            "START": formatted_date,
            "DEPOSIT": dep_formatted,
            "ROOM": room.upper(),
            "PROPERTY": property_addr,
        }
        
        if not os.path.exists(TEMPLATE_PATH):
            return f"Template not found: {TEMPLATE_PATH}", 500
        
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
        return f"Error generating document: {str(e)}", 500


if __name__ == '__main__':
    app.run(debug=True)
