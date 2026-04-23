from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

TEMPLATE_PATH = "template.docx"
OUTPUT_PATH = "output.docx"


def replace_placeholders(doc, data):
    """
    Replaces placeholders while preserving formatting as much as possible.
    Uses run-level replacement (prevents loss of bold/size).
    """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for key, value in data.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)

    # tables support (important for contracts)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        for key, value in data.items():
                            if key in run.text:
                                run.text = run.text.replace(key, value)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        name = request.form.get("name")
        room = request.form.get("room")
        rent = request.form.get("rent")
        deposit = request.form.get("deposit")
        address = request.form.get("address")
        reference = request.form.get("reference")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        today = datetime.today().strftime("%d %B %Y")

        doc = Document(TEMPLATE_PATH)

        data = {
            "{{NAME}}": name,
            "{{ROOM}}": room,
            "{{RENT}}": rent,
            "{{DEPOSIT}}": deposit,
            "{{ADDRESS}}": address,
            "{{REFERENCE}}": reference,
            "{{START_DATE}}": start_date,
            "{{END_DATE}}": end_date,
            "{{TODAY}}": today
        }

        replace_placeholders(doc, data)

        doc.save(OUTPUT_PATH)

        return send_file(OUTPUT_PATH, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
