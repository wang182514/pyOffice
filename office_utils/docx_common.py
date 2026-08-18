# -*- coding: utf-8 -*-
"""文档级通用操作（一次设置, 整篇生效）"""
from docx.oxml.ns import qn

from .docx_font import set_style_font


# 默认要设置中文字体的样式列表 —— 大多数文档用到这些就够了
DEFAULT_STYLE_NAMES = [
    "Normal", "Title", "Heading 1", "Heading 2", "Heading 3",
    "Heading 4", "Heading 5", "Heading 6",
    "List Bullet", "List Number", "List Paragraph",
    "Caption", "Quote", "Intense Quote",
]


def set_doc_fonts(doc, cn_font, en_font=None, style_names=None):
    """一键设置文档常用样式的中英文字体

    解决实际问题: python-docx 默认模板的主题里中文字体为空,
    打开文档时中文全显示成"方框叉"。
    在样式层面设一次, 后面所有用这些样式的段落自动继承。

    Args:
        doc:         python-docx 的 Document 对象
        cn_font:     中文字体, 如 "微软雅黑" / "宋体"
        en_font:     西文字体, 缺省与 cn_font 相同
        style_names: 要设置字体的样式名列表, 缺省 DEFAULT_STYLE_NAMES
                     (模板里有自定义样式时再额外传)

    Example:
        from docx import Document
        from office_utils import set_doc_fonts

        doc = Document("template.docx")
        set_doc_fonts(doc, "微软雅黑")
    """
    if en_font is None:
        en_font = cn_font
    if style_names is None:
        style_names = DEFAULT_STYLE_NAMES

    for name in style_names:
        if name in doc.styles:             # 模板没这个样式就跳过, 不报错
            style = doc.styles[name]
            style.font.name = en_font      # 西文
            style.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
