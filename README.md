# pyOffice — docx/xlsx 读写练习与工具封装

练习项目，目标：熟悉 python-docx / openpyxl 的读写操作，逐步把确定会用到的
操作封装成 `office_utils` 包，最终合并到射频组件测试项目（测试结果 → word/excel 报告）。

## 环境与运行

- 虚拟环境：`.venv`（Python 3.12.7）
- 依赖：`python-docx`、`openpyxl`（见 `requirements.txt`，换机器可 `pip install -r requirements.txt` 一键恢复）

```bash
# Git Bash
.venv/Scripts/python.exe demo_docx.py    # docx 读写演示
.venv/Scripts/python.exe demo_xlsx.py    # xlsx 读写演示
```

## 目录结构

```
pyOffice/
├── demo_docx.py          # docx 读写演示（详细注释，学习材料）
├── demo_xlsx.py          # xlsx 读写演示（详细注释）
├── office_utils/         # 封装包（成长中）
│   ├── __init__.py       # 统一导出所有公开 API
│   └── bookmark.py       # docx 书签操作（创建/查找/填充）
├── output/               # 脚本生成的文件
└── requirements.txt
```

## office_utils API

| 函数 | 签名 | 用途 |
|---|---|---|
| `set_run_font` | `(run, name)` | 给单个 run 设中西文字体（补 eastAsia 槽位） |
| `set_style_font` | `(style, name)` | 给样式设中西文字体，所有用该样式的文字生效 |
| `add_bookmark` | `(run, name, bookmark_id)` | 给 run 打书签（书签只包住该 run） |
| `find_bookmark` | `(doc, name) → (start, end)` | 按名查找单个书签，找不到返回 `(None, None)` |
| `find_bookmarks` | `(doc) → {name: (start, end)}` | 一次遍历返回全部书签 |
| `fill_bookmark` | `(doc, name, text, font=None, *, clear=True)` | 单个书签填充 |
| `fill_bookmarks` | `(doc, data, font=None, *, clear=True)` | 批量填充，`data` 为 `{书签名: 文字}` |

要点：

- **`clear` 参数**：默认 `True` 清除书签内旧内容（覆盖语义）；`clear=False` 追加语义，
  旧内容保留、新文字插在书签起点后。
- **批量性能**：`fill_bookmarks` 只遍历一次 XML 树。实测 100 次循环查找 36.8ms vs
  单次批量索引 0.3ms，书签多时必须用批量版。
- **错误提示**：批量填充时缺失的书签会一次性全部列出（含现有书签清单），方便对照模板排查。
- **典型用法**：

```python
from docx import Document
from office_utils import fill_bookmarks

doc = Document("template.docx")            # Word 里排版好模板并插好书签
fill_bookmarks(doc, {
    "reporter": "王五",
    "report_date": "2026-08-14",
    "result": "合格",
}, font="微软雅黑")
doc.save("output.docx")
```

## 架构决策（已确定）

1. **按功能域分模块，不按读/写分**：`bookmark.py` 是书签从创建到填充的完整闭环。
   拆成 `read.py`/`write.py` 会把一个功能割到两个文件。
2. **`__init__.py` 统一导出**：调用方永远 `from office_utils import xxx`，不感知内部
   文件结构。内部移动/重命名文件只改 `__init__.py`，不破坏调用方。
3. **无状态纯函数，不封装类**：状态在 `Document`/`Workbook` 对象上（python-docx /
   openpyxl 的职责），工具函数无状态就不需要类。类的触发点（出现再上）：
   - 配置传参重复到烦（先模块常量，再配置对象）；
   - 高层生命周期管理（`ReportFiller`: 打开→填充→另存）——类适合出现在高层 API，
     不适合包在底层工具函数外；
   - 多种输出格式需要统一接口（多态）。
4. **分层**：低层函数做单一小事（现在的 7 个）；将来加面向场景的高层函数
   （如 `fill_report(template_path, data, output_path)`），测试代码只碰高层。
5. **拆分标准**：单模块超 ~300 行、或混入不相关功能域时拆。当前已知信号：
   `set_run_font`/`set_style_font` 与书签无关，引入 xlsx 功能时应挪到 `docx_font.py`。

### 演进路线

```
阶段一（当前）     office_utils/bookmark.py 单模块
阶段二（加 xlsx）  扁平加文件: docx_table.py / docx_font.py / xlsx_writer.py
阶段三（模块多）   按格式分子包: office_utils/docx/... + office_utils/xlsx/...
```

不跳级，避免过度设计。

## 关键知识点 / 踩坑记录

### docx

- **docx 本质**：zip + XML（`word/document.xml` 正文、`styles.xml` 样式、`theme1.xml` 主题）。
  python-docx 没有的功能可直接操作 XML。
- **中文字体（最大坑）**：默认模板主题里中文字体为空 → 打开显示"方框叉"。
  `run.font.name` 只设 ascii/hAnsi 槽位，中文必须额外设 `w:eastAsia`
  （`set_run_font` 已封装）。字体四槽位：ascii / hAnsi / eastAsia / cs。
- **对象模型**：`Document → paragraphs → runs`；格式加在 run 上不在段落上。
  `doc.paragraphs` 不含表格内/页眉页脚的段落。
- **书签 XML 结构**：`<w:bookmarkStart w:id w:name/> ... <w:bookmarkEnd w:id/>` 配对。
  书签要包住"占位符 run"而不是整段（否则替换时误删前缀文字）；书签元素不能插到
  段落属性 `pPr` 之前（非法 XML 顺序）。
- **表格样式**：内置 100 种。`Table Grid` = 黑色细边框无底色（最"默认"）；
  带 `Accent N` 的才有彩色（1蓝 2红 3绿 4紫 5青 6橙）。速查表在 `demo_docx.py` 顶部注释。
- **合并单元格**：读取时被合并的格子返回重复内容。

### xlsx

- **公式**：`load_workbook(path, data_only=True)` 读公式计算值，但值是 Excel 打开
  保存时的缓存——文件从未被 Excel 打开过则读到 `None`（openpyxl 不计算公式）。
- **三种读法**：`iter_rows()` 遍历 / `values_only=True` 直接取值 /
  `ws["B2"]`、`ws.cell(row, col)` 定位。
- **写入**：`append()` 按行追加、单元格赋值、公式字符串 `=SUM(...)`；
  样式 `Font`/`PatternFill`/`Border`/`Alignment` 加在单元格上；
  `freeze_panes` 冻结首行；`BarChart`+`Reference` 生成图表。

### 通用（lxml XML 操作，bookmark.py 内部用的）

- `qn("w:id")`：把带命名空间前缀的标签转成完整 URI 形式。
- `OxmlElement("w:r")`：手工创建 docx 内部 XML 节点。
- `el.addprevious()/addnext()`：相对某节点前后插入。
- `el.getnext()/getparent()/remove()`：兄弟遍历与删除。
- 修改 XML 树后，之前拿到的元素引用可能失效——批量修改时注意重新查找。

## 迁移到测试项目的计划

- **确定会用**：书签填充 `fill_bookmarks`（word 报告模板）。
- **待验证后封装**：docx 表格写入（测试结果表）、xlsx 数据追加（测试记录）。
- **模板与代码解耦**：模板文件放独立目录；docx 书签名 / xlsx 单元格坐标就是
  "代码 ↔ 模板编辑者"之间的接口约定，改模板不改代码。
- 数据与格式解耦：测试逻辑产出统一 dict，输出层只认 dict。
