from flask import Flask, request, send_file, render_template_string
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

HTML_FORM = """
<h2>KPI Licence Generator</h2>
<form method="post">
Name: <input name="name"><br><br>
Room: <input name="room"><br><br>
Property Address: <input name="address"><br><br>
Start Date: <input name="start"><br><br>
End Date: <input name="end"><br><br>
Monthly Fee: <input name="rent"><br><br>
Deposit: <input name="deposit"><br><br>
Payment Reference: <input name="ref"><br><br>
<button type="submit">Generate Contract</button>
</form>
"""

def replace_text(doc, data):
    for p in doc.paragraphs:
        for key, val in data.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        today = datetime.today().strftime("%d %B %Y")

        data = {
            "{{NAME}}": request.form["name"],
            "{{ROOM}}": request.form["room"],
            "{{ADDRESS}}": request.form["address"],
            "{{START}}": request.form["start"],
            "{{END}}": request.form["end"],
            "{{RENT}}": request.form["rent"],
            "{{DEPOSIT}}": request.form["deposit"],
            "{{REF}}": request.form["ref"],
            "{{TODAY}}": today
        }

        doc = Document("template.docx")
        replace_text(doc, data)

        filename = "contract.docx"
        doc.save(filename)

        return send_file(filename, as_attachment=True)

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
