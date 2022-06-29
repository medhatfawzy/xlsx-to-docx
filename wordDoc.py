from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_TABLE_DIRECTION
import pandas as pd
from coordinate import coordination

data = pd.read_excel("record.xlsx")
code_columns = ["المخزن", "بلوك", "راك", "رف"]

data["coordinate"] = data[code_columns].apply(coordination, axis=1)
data.drop(code_columns, inplace=True, axis=1)

document = Document()
section = document.sections[0]
section.page_width = Inches(8.27)
section.page_height = Inches(11.69)
section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.2)

keys = [
    "اسم الصنف",
    "رقم الصفحة",
    "اسم الدفتر",
    "الكمية",
    "الوحدة",
    "إحداثي التخزين"
]


for row in data.values:
    document.add_heading("كارت الصنف")
    table = document.add_table(len(keys), 2, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    table.table_direction = WD_TABLE_DIRECTION.RTL
    table.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, j in enumerate(keys):
        table_cell_right = table.cell(i, 0)
        table_cell_left = table.cell(i, 1)
        table_cell_right.text = j
        table_cell_left.text = str(row[i])

document.save("test.docx")