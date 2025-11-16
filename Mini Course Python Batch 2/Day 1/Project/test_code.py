# ============================================
#  Auto Invoice Generator
#  Excel → Word (docxtpl) with Date & Rupiah Format
# ============================================

import openpyxl
import docxtpl
from datetime import datetime
import os


# ===========================
#  Helper Functions
# ===========================

def format_date(value):
    """Format datetime menjadi 'YYYY-MM-DD'. Selain datetime tidak diubah."""
    if isinstance(value, datetime):
        return value.date().strftime("%Y-%m-%d")
    return value if value not in (None, "") else ""


def format_currency(value):
    """Format angka menjadi Rupiah: Rp 1.234.567"""
    if value in (None, ""):
        return ""
    try:
        value = float(value)
        formatted = "Rp {:,}".format(int(value)).replace(",", ".")
        return formatted
    except (ValueError, TypeError):
        return value


# ===========================
#  Main Process
# ===========================

def generate_invoices(
    excel_path,
    template_path,
    output_dir="output_invoice"
):
    """Generate invoice Word files from Excel + docxtpl template."""

    # Pastikan folder output tersedia
    os.makedirs(output_dir, exist_ok=True)

    # Load Excel
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = workbook.active
    rows = list(sheet.values)

    # Load template docx
    template = docxtpl.DocxTemplate(template_path)

    # Loop setiap data (skip header)
    for row in rows[1:]:
        clean = [cell if cell is not None else "" for cell in row]

        # Mapping data ke template
        context = {
            "DATE": format_date(clean[0]),
            "INVOICE": clean[1],
            "NAME": clean[2],
            "ALAMAT": clean[3],
            "POS": clean[4],
            "NOTE": clean[5],

            # Produk 1
            "Q_SATU": clean[6],
            "PRODUCT_SATU": clean[7],
            "PRICE_SATU": format_currency(clean[8]),
            "TOTAL_SATU": format_currency(clean[9]),

            # Produk 2
            "Q_DUA": clean[10],
            "PRODUCT_DUA": clean[11],
            "PRICE_DUA": format_currency(clean[12]),
            "TOTAL_DUA": format_currency(clean[13]),

            # Ringkasan biaya
            "SUBTOTAL": format_currency(clean[14]),
            "TAX": format_currency(clean[15]),
            "HANDLING": format_currency(clean[16]),
            "TOTAL": format_currency(clean[17]),
        }

        # Render template
        template.render(context)

        # Nama file output
        invoice_number = str(clean[1]).replace("/", "-")
        output_file = os.path.join(output_dir, f"Invoice-{invoice_number}.docx")

        # Simpan
        template.save(output_file)
        print(f"[DONE] Saved → {output_file}")


# ===========================
#  Execute
# ===========================

if __name__ == "__main__":

    excel_file = r"D:\Mini Course Python Batch 2\Project\data.xlsx"
    template_file = r"D:\Mini Course Python Batch 2\Project\Business invoice Basic.docx"

    generate_invoices(excel_file, template_file)
