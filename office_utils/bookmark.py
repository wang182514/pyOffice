# -*- coding: utf-8 -*-
"""
书签操作模块 —— 封装 python-docx 没有原生支持的书签功能

【背景: docx 书签的 XML 结构】
Word 书签存储为一对 XML 元素, 包住一段内容:

    <w:p>                                              ← 段落
        <w:pPr/>                                       ← 段落属性(必须是最前面的子元素)
        <w:r><w:t>报告人：</w:t></w:r>                  ← 前缀 run(书签外, 不参与替换)
        <w:bookmarkStart w:id="100" w:name="reporter"/> ← 书签起点(名字在这里)
        <w:r><w:t>【待填写】</w:t></w:r>                ← 占位符 run(书签包住的内容)
        <w:bookmarkEnd w:id="100"/>                    ← 书签终点(用相同 id 配对)
    </w:p>

典型工作流: 在 Word 里排版模板 → 插入书签 → Python 打开 → 定位 → 填数据 → 另存。
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# =================================================================
# 书签创建
# =================================================================
def add_bookmark(run, name, bookmark_id):
    """给某个 run 打上书签 —— 书签只包住这个 run, 不影响段落里的其他文字

    两个注意点(都是调试中踩过的坑):
    1. 书签要包"占位符 run"而不是整个段落 —— 否则替换时会把段落里
       "报告人：" 这样的前缀文字一起删掉;
    2. 书签元素不能插到段落最前面 —— 段落第一个子元素必须永远是
       <w:pPr>(段落属性), 插到它前面属于非法 XML 顺序。

    Args:
        run:         python-docx 的 Run 对象 (要被书签包住的占位文字)
        name:        书签名 (Word 里"插入→书签"给的位置起的名字, 查找就靠它)
        bookmark_id: 书签编号 (文档内唯一整数, 可自编, 无特殊含义)
    """
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    run._r.addprevious(start)    # lxml 方法: 插到该 run 之前
    run._r.addnext(end)          # lxml 方法: 插到该 run 之后
    # 结果: start ... run ... end, 书签正好包住这个 run


# =================================================================
# 书签查找
# =================================================================
def find_bookmark(doc, name):
    """按名字查找书签, 返回 (bookmarkStart, bookmarkEnd) 两个 XML 元素

    原理: doc.element.body 是正文的根 XML 节点,
    .iter() 深度优先遍历整棵树(段落里/表格单元格里/页眉里的书签都能找到)。

    Args:
        doc:  python-docx 的 Document 对象
        name: 书签名

    Returns:
        (bookmarkStart, bookmarkEnd) 元组; 找不到返回 (None, None)
    """
    end_map = {}
    for el in doc.element.body.iter():
        if el.tag == qn("w:bookmarkEnd"):
            end_map[el.get(qn("w:id"))] = el
    for el in doc.element.body.iter():
        if el.tag == qn("w:bookmarkStart") and el.get(qn("w:name")) == name:
            return el, end_map.get(el.get(qn("w:id")))
    return None, None


def find_bookmarks(doc):
    """一次遍历, 返回文档中所有书签的映射

    相比循环调用 find_bookmark(doc, name), 只遍历一次 XML 树,
    当书签数量多时(N 个) 性能从 O(N * 文档节点数) 降到 O(文档节点数)。

    Args:
        doc: python-docx 的 Document 对象

    Returns:
        dict: {书签名: (bookmarkStart, bookmarkEnd), ...}
    """
    starts = {}
    end_map = {}
    for el in doc.element.body.iter():
        if el.tag == qn("w:bookmarkStart"):
            bid = el.get(qn("w:id"))
            bname = el.get(qn("w:name"))
            starts[bid] = (el, bname)
        elif el.tag == qn("w:bookmarkEnd"):
            end_map[el.get(qn("w:id"))] = el
    return {
        bname: (start_el, end_map.get(bid))
        for bid, (start_el, bname) in starts.items()
    }


def get_bookmark_text(doc, name):
    """读取书签包住的内容(纯文本)

    用途: 填充后验证结果 / 读取已有 docx 的某个字段值

    Args:
        doc:  python-docx 的 Document 对象
        name: 书签名

    Returns:
        书签内文字; 找不到书签返回 None
    """
    start, _ = find_bookmark(doc, name)
    if start is None:
        return None
    # 书签包住的内容是 start 之后、end 之前的所有 <w:t> 文本节点
    texts = []
    for el in start.itersiblings():
        if el.tag == qn("w:bookmarkEnd"):
            break
        for t in el.iter(qn("w:t")):
            if t.text:
                texts.append(t.text)
    return "".join(texts)


# =================================================================
# 书签填充
# =================================================================
def _build_run(text, font=None):
    """内部: 构建一个 <w:r> XML 节点, 带可选字体"""
    run = OxmlElement("w:r")
    if font:
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font)
        rFonts.set(qn("w:hAnsi"), font)
        rFonts.set(qn("w:eastAsia"), font)
        rPr.append(rFonts)
        run.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    return run


def _clear_bookmark_content(start, end):
    """内部: 删除书签之间的旧内容(占位符文字)"""
    if end is None:
        return
    el = start.getnext()
    while el is not None and el is not end:
        nxt = el.getnext()
        el.getparent().remove(el)
        el = nxt


def fill_bookmark(doc, name, text, font=None, *, clear=True):
    """按书签填充: 在书签位置插入新文字

    Args:
        doc:   python-docx 的 Document 对象
        name:  书签名
        text:  要写入的新文字
        font:  可选字体名, 如 "微软雅黑"
        clear: 是否清除书签内的旧内容, 默认 True (覆盖语义)
               False 为追加语义, 旧内容保留、新文字插在书签起点后

    Raises:
        ValueError: 书签不存在时抛出
    """
    start, end = find_bookmark(doc, name)
    if start is None:
        raise ValueError(f"书签 '{name}' 不存在")

    if clear:
        _clear_bookmark_content(start, end)
    start.addnext(_build_run(text, font))


def fill_bookmarks(doc, data, font=None, *, clear=True):
    """批量按书签填充: 一次遍历查找所有书签, 然后逐个替换

    相比循环调用 fill_bookmark, 查找阶段只遍历一次 XML 树,
    N 个书签的性能从 O(2N * 节点数) 降到 O(节点数)。

    Args:
        doc:   python-docx 的 Document 对象
        data:  dict, {书签名: 要写入的文字, ...}
        font:  可选字体名
        clear: 是否清除旧内容, 默认 True

    Raises:
        ValueError: 有书签不存在时抛出 (一次性列出所有缺失的书签名)

    Example:
        fill_bookmarks(doc, {
            "reporter": "王五",
            "report_date": "2026-08-14",
        }, font="微软雅黑")
    """
    all_bookmarks = find_bookmarks(doc)

    missing = [name for name in data if name not in all_bookmarks]
    if missing:
        raise ValueError(
            f"以下书签不存在: {missing}  (文档中现有书签: {list(all_bookmarks.keys())})"
        )

    for name, text in data.items():
        if clear:
            start, end = all_bookmarks[name]
            _clear_bookmark_content(start, end)
            start.addnext(_build_run(text, font))
        else:
            start, end = find_bookmark(doc, name)
            start.addnext(_build_run(text, font))
