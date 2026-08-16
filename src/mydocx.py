import os.path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

from demo_docx import fill_table
# if os.path.exists(r'D:\Project\python\pyOffice\src\output'):
#     print("-----------------")
from office_utils import *


CN_FONT = "宋体"   # 中文字体, 可换成 "宋体"/"等线" 等(需本机装有该字体)
TABLE_STYLE_DEFAULT = "Table Grid"          # 黑色细边框、无底色
TABLE_STYLE_BLUE = "Light Grid Accent 1"    # 蓝色系底纹样式

doc =Document()
for style_name in ["Normal", "Title", "Heading 1", "Heading 2",
                   "List Bullet", "List Number"]:
    set_style_font(doc.styles[style_name], CN_FONT)

doc.add_heading("大标题", level=0)
doc.add_heading("一级标题", level=1)

p = doc.add_paragraph('这是一个段落，默认样式')
set_run_font(p.runs[0],CN_FONT)
p0 = p.insert_paragraph_before('插入到段落前')


p2 = doc.add_paragraph()
r1 = p2.add_run("加粗 ")
r1.bold = True                           # 加粗 (取消用 r1.bold = None)
r2 = p2.add_run("斜体 ")
r2.italic = True                         # 斜体
r3 = p2.add_run("红色字 ")
r3.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)   # 颜色: RGB 三元组
r4 = p2.add_run("16号字")
r4.font.size = Pt(16)                    # 字号: Pt(磅), Word 里的"号"要换算

for run in p2.runs:
    print(run.text)
    set_run_font(run, CN_FONT)

# 4. 对齐方式: 常用 LEFT / CENTER / RIGHT / JUSTIFY(两端对齐)
pCenter = doc.add_paragraph("居中显示的段落")
pCenter.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 7. 表格
doc.add_heading("二、表格", level=1)
headers = ["姓名", "部门", "工资"]
rows = [
    ["张三", "研发部", "15000"],
    ["李四", "市场部", "12000"],
    ["王五", "人事部", "10000"],
]
table = doc.add_table(rows=1, cols=3)
table.style = TABLE_STYLE_DEFAULT
fill_table(table,headers,rows)


doc.save('../output/demo.docx')

# import os
# print(__file__)
# # 当前脚本完整路径
# script_path = os.path.abspath(__file__)
# # 脚本所在目录
# script_dir = os.path.dirname(script_path)
#
# print(script_path)
# print(script_dir)
