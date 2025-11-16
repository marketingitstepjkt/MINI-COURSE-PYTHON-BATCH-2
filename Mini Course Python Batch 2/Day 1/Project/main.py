# Import library yang dibutuhkan
import openpyxl
import docxtpl

# Cek file excel
file_excel = "D:\Mini Course Python Batch 2\Project\data.xlsx"
load = openpyxl.load_workbook(file_excel, data_only=True)

# Cek sheet active
sheet = load.active
print(sheet)

# Cek file dokumen
file_doc = docxtpl.DocxTemplate("D:\Mini Course Python Batch 2\Project\Business invoice Basic.docx")
print(file_doc)

# Ambil value dalam file excel
get_value = list(sheet.values)
print(get_value)

# Looping : mix and match data excel dan data dokumen
for value in get_value[1:]:

    # Exception untuk none
    clean_value = [clean if clean is not None else"" for clean in value]

    file_doc.render({
        "DATE": clean_value[0],
        "INVOICE": clean_value[1],
        "NAME": clean_value[2],
        "ALAMAT": clean_value[3],
        "POS": clean_value[4],
        "NOTE": clean_value[5],
        "Q_SATU": clean_value[6],
        "PRODUCT_SATU": clean_value[7],
        "PRICE_SATU": clean_value[8],
        "TOTAL_SATU": clean_value[9],
        "Q_DUA": clean_value[10],
        "PRODUCT_DUA": clean_value[11],
        "PRICE_DUA": clean_value[12],
        "TOTAL_DUA": clean_value[13],
        "SUBTOTAL": clean_value[14],
        "TAX": clean_value[15],
        "HANDLING": clean_value[16],
        "TOTAL": clean_value[17]
    })

    # Simpan file ke dalam komputer
    file_doc.name = f"Ini adalah invoice untuk ke - {value[1]}.docx"
    file_doc.save(file_doc.name)
