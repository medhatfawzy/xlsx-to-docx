from docx import Document
from docx.shared import Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_TABLE_DIRECTION

from load_excel_file import load_excel_file

data = load_excel_file("record.xlsx")


# Setting up some properties of the page
document = Document()
section = document.sections[0]
section.page_width = Inches(8.27) #A4 width
section.page_height = Inches(11.69) #A4 height
section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.2)

keys = [
    "اسم الصنف",
    "رقم الصفحة",
    "اسم الدفتر",
    "الكمية",
    "الوحدة",
    "إحداثي التخزين"
]


for row in data:
    heading = document.add_heading("كارت الصنف")
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = document.add_table(len(keys), 2, style="Light Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.table_direction = WD_TABLE_DIRECTION.RTL
    table.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.columns[0].width = Emu(1100000)
    table.columns[1].width = Emu(3000000)

    for i, j in enumerate(keys):
        table.cell(i, 0).text = j
        table.cell(i, 1).text = str(row[i])

document.save("items_tags.docx")