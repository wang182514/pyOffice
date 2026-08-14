# -*- coding: utf-8 -*-
"""office_utils —— Office 文档(docx/xlsx)读写工具包

当前模块: 书签操作 (bookmark)
封装了书签的创建、查找、填充操作, 用于"模板填充"场景:
Word 里排版好模板并插入书签, Python 打开模板后按书签名填入数据。

使用示例:
    from docx import Document
    from office_utils import fill_bookmark, fill_bookmarks, set_run_font

    doc = Document("template.docx")

    # 单个填充 (默认清除旧内容)
    fill_bookmark(doc, "reporter", "王五", font="微软雅黑")

    # 批量填充 (只遍历一次 XML, 比循环调 fill_bookmark 快)
    fill_bookmarks(doc, {
        "reporter": "王五",
        "report_date": "2026-08-14",
        "result": "合格",
    }, font="微软雅黑")

    # 追加模式 (不清除旧内容, 新文字追加在后面)
    fill_bookmark(doc, "note", "补充说明", clear=False)

    doc.save("output.docx")
"""
from .bookmark import (
    set_run_font,
    set_style_font,
    add_bookmark,
    find_bookmark,
    find_bookmarks,
    fill_bookmark,
    fill_bookmarks,
)

__all__ = [
    "set_run_font",
    "set_style_font",
    "add_bookmark",
    "find_bookmark",
    "find_bookmarks",
    "fill_bookmark",
    "fill_bookmarks",
]
