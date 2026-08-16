from docx import Document
from office_utils import *

TABLE_STYLE_DEFAULT = "Table Grid"          # 黑色细边框、无底色
CN_FONT = "宋体"

doc = Document()

for style_name in ["Normal", "Title", "Heading 1", "Heading 2",
                   "List Bullet", "List Number"]:
    set_style_font(doc.styles[style_name], CN_FONT)

table = doc.add_table(3,5)
table.style = TABLE_STYLE_DEFAULT
table.cell(0,0).merge(table.cell(2,0))
table.cell(0,0).text = '检验记录'
table.cell(0,0).paragraphs[0].runs[0].bold = True
doc.add_paragraph('\n'*2)

doc.save('../output/检验记录模板.docx')