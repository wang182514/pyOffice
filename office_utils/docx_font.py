# -*- coding: utf-8 -*-
"""字体设置工具 —— 从原 bookmark.py 拆分出来

字体不属于书签功能域, 单独一个模块, 后续 cell/paragraph 格式化也会用到。
"""
from docx.oxml.ns import qn


def set_run_font(run, name):
    """给单个 run 设置字体 —— 西文和中文要分别设置

    Word 字体四槽位: w:ascii / w:hAnsi / w:eastAsia / w:cs
    run.font.name 只设 ascii + hAnsi, 中文必须额外设 eastAsia
    (否则打开文档中文显示"方框叉")。

    Args:
        run:  python-docx 的 Run 对象
        name: 字体名, 如 "微软雅黑" / "宋体" / "等线"
    """
    run.font.name = name                                  # 西文 (ascii/hAnsi)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)   # 中文 (eastAsia)


def set_style_font(style, name):
    """给样式设置字体 —— 一次设置, 所有使用该样式的文字全部生效

    直接格式(加在单个 run 上)只影响那一段文字;
    样式是命名的格式预设, 改了 Word 样式面板里的定义,
    影响所有使用该样式的段落。

    Args:
        style: python-docx 的 Style 对象 (如 doc.styles["Normal"])
        name:  字体名
    """
    style.font.name = name
    style.element.rPr.rFonts.set(qn("w:eastAsia"), name)
