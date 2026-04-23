from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import re
import os


class LicenceAgreementGenerator:
    def __init__(self):
        self.licensor = "Kensington Properties & Investments Ltd"
        self.today = datetime.now().strftime("%d %B %Y")
        
    def validate_date(self, date_str):
        """Strict DD/MM/YYYY validation with leap year support."""
        pattern = r"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}$"
        if not re.match(pattern, date_str):
            return None
        try:
            return datetime.strptime(date_str, "%d/%m/%Y").strftime("%d %B %Y")
        except ValueError:
            return None
    
    def validate_currency(self, amount_str):
        """Accepts '500', '500.00', '£500' — returns clean numeric string."""
        cleaned = re.sub(r"[£,]", "", amount_str.strip())
        try:
            val = float(cleaned)
            if val <= 0:
                return None
            return f"{val:,.2f}"
        except ValueError:
            return None
    
    def validate_name(self, name):
        """Ensures printable name, strips excessive whitespace."""
        cleaned = " ".join(name.strip().split())
        if len(cleaned) < 2 or len(cleaned) > 100:
            return None
        if not re.match(r"^[A-Za-z\s\-\.']+$", cleaned):
            return None
        return cleaned
    
    def validate_room(self, room):
        """Normalizes room identifiers."""
        cleaned = room.strip().upper()
        if not cleaned:
            return None
        # Handle "Room 5", "5", "A-12" etc.
        cleaned = re.sub(r"^ROOM\s*", "", cleaned)
        return cleaned
    
    def get_safe_filename(self, name):
        """Prevents overwrite, handles special characters."""
        base = re.sub(r'[\\/*?:"<>|]', "", name)
        base = base.replace(" ", "_")
        filename = f"Licence_Agreement_{base}.docx"
        counter = 1
        while os.path.exists(filename):
            filename = f"Licence_Agreement_{base}_{counter}.docx"
            counter += 1
        return filename
    
    def create_document(self, name, start_date_raw, rent_raw, deposit_raw, room_raw):
        # --- VALIDATION ---
        name = self.validate_name(name)
        if not name:
            raise ValueError("Invalid name. Use 2-100 letters, spaces, hyphens, or apostrophes only.")
        
        start_date = self.validate_date(start_date_raw)
        if not start_date:
            raise ValueError("Invalid date. Use DD/MM/YYYY format (e.g. 01/05/2026).")
        
        rent = self.validate_currency(rent_raw)
        if not rent:
            raise ValueError("Invalid rent. Enter a positive number (e.g. 850 or £850).")
        
        deposit = self.validate_currency(deposit_raw)
        if not deposit:
            raise ValueError("Invalid deposit. Enter a positive number.")
        
        room = self.validate_room(room_raw)
        if not room:
            raise ValueError("Invalid room number.")
        
        # --- DOCUMENT BUILD ---
        doc = Document()
        
        # Page setup for professional print
        sections = doc.sections[0]
        sections.top_margin = Inches(1)
        sections.bottom_margin = Inches(1)
        sections.left_margin = Inches(1.25)
        sections.right_margin = Inches(1.25)
        
        # Header
        header = doc.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header.add_run("LICENCE AGREEMENT")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        doc.add_paragraph()  # Spacer
        
        # Meta info
        meta = doc.add_paragraph()
        meta.add_run(f"Date of Agreement: ").bold = True
        meta.add_run(self.today)
        
        doc.add_paragraph()
        
        # Parties
        p = doc.add_paragraph()
        p.add_run("Parties:\n").bold = True
        p.add_run(f"1. {self.licensor} (\"Licensor\")\n")
        p.add_run(f"2. {name} (\"Licensee\")")
        
        doc.add_paragraph()
        
        # Recitals
        recital = doc.add_paragraph()
        recital.add_run("Recitals:\n").bold = True
        recital.add_run(
            f"The Licensor agrees to grant the Licensee a licence to occupy "
            f"Room {room} subject to the terms and conditions set out below."
        )
        
        doc.add_paragraph()
        
        # Terms
        terms = [
            ("1. Premises", f"The Licensee is granted a non-exclusive licence to occupy Room {room} within the property managed by the Licensor."),
            ("2. Commencement", f"This agreement shall commence on {start_date} and shall continue until terminated in accordance with Clause 6."),
            ("3. Licence Fee", f"The Licensee shall pay a monthly licence fee of £{rent}, payable in advance on or before the 1st day of each calendar month."),
            ("4. Deposit", f"A security deposit of £{deposit} shall be paid by the Licensee prior to taking occupation. This deposit shall be held in accordance with the applicable tenancy deposit protection regulations and returned within 14 days of the agreement's termination, less any deductions for damages or outstanding charges."),
            ("5. Licensee Obligations", "The Licensee agrees to: (a) comply with all house rules and reasonable instructions issued by the Licensor; (b) refrain from causing nuisance, noise, or disturbance to other occupants; (c) maintain the Room and shared spaces in a clean and tidy condition; (d) not sub-let, assign, or share occupation without prior written consent."),
            ("6. Termination", "Either party may terminate this agreement by providing not less than 28 days' written notice to the other party. The Licensor may terminate immediately in the event of a material breach by the Licensee of any term of this agreement."),
            ("7. Governing Law", "This agreement shall be governed by and construed in accordance with the laws of England and Wales.")
        ]
        
        for title, content in terms:
            p = doc.add_paragraph()
            p.add_run(f"{title}\n").bold = True
            p.add_run(content)
            p.paragraph_format.space_after = Pt(12)
        
        doc.add_paragraph()
        
        # Signature block
        sig_heading = doc.add_paragraph()
        sig_heading.add_run("Executed as a deed").bold = True
        sig_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()
        
        # Two-column signature layout
        table = doc.add_table(rows=3, cols=2)
        table.autofit = False
        table.allow_autofit = False
        table.columns[0].width = Inches(3)
        table.columns[1].width = Inches(3)
        
        # Licensee signature
        table.cell(0, 0).text = "Signed by the Licensee:"
        table.cell(1, 0).text = "_" * 40
        table.cell(2, 0).text = f"Name: {name}\nDate: _______________"
        
        # Licensor signature
        table.cell(0, 1).text = "Signed for and on behalf of the Licensor:"
        table.cell(1, 1).text = "_" * 40
        table.cell(2, 1).text = f"Name: _______________\nDate: _______________"
        
        # Style the table
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(6)
        
        # Save
        filename = self.get_safe_filename(name)
        doc.save(filename)
        return filename


def main():
    generator = LicenceAgreementGenerator()
    
    print("=" * 50)
    print("LICENCE AGREEMENT GENERATOR")
    print("=" * 50)
    
    try:
        name = input("\nEnter full name: ")
        start_date = input("Enter start date (DD/MM/YYYY): ")
        rent = input("Enter monthly rent (£): ")
        deposit = input("Enter deposit (£): ")
        room = input("Enter room number: ")
        
        print("\nGenerating document...")
        filename = generator.create_document(name, start_date, rent, deposit, room)
        print(f"\n✓ Success: '{filename}' created.")
        print(f"  Agreement date: {generator.today}")
        
    except ValueError as e:
        print(f"\n✗ Validation Error: {e}")
    except Exception as e:
        print(f"\n✗ Unexpected Error: {e}")


if __name__ == "__main__":
    main()
