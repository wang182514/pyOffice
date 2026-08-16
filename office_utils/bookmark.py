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
# 字体设置
# =================================================================
def set_run_font(run, name):
    """给单个 run 设置字体 —— 西文和中文要分别设置

    Word 的字体分 4 个槽位:
      w:ascii    英文半角字符 (a,b,c,1,2,3)
      w:hAnsi    高 ANSI 字符 (带音标的欧洲文字等)
      w:eastAsia 中日韩字符   ← 中文看这个!
      w:cs       复杂文种 (阿拉伯文等)

    run.font.name = "微软雅黑" 只会填 ascii + hAnsi,
    所以中文必须再手动补 eastAsia 槽位, 否则打开文档中文显示"方框叉"。

    Args:
        run:  python-docx 的 Run 对象
        name: 字体名, 如 "微软雅黑" / "宋体" / "等线"
    """
    run.font.name = name                                  # 西文 (ascii/hAnsi)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)   # 中文 (eastAsia)


def set_style_font(style, font_name):
    """给样式设置字体 —— 一次设置, 所有使用该样式的文字全部生效

    直接格式(加在单个 run 上)只影响那一段文字;
    样式是命名的格式预设, 相当于改了 Word 样式面板里的定义,
    影响所有使用该样式的段落。

    Args:
        style: python-docx 的 Style 对象 (如 doc.styles["Normal"])
        font_name:  字体名
    """
    style.font.name = font_name
    style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


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
    # 第一遍: 收集所有终点, 建立 id → bookmarkEnd 的映射
    end_map = {}
    for el in doc.element.body.iter():
        if el.tag == qn("w:bookmarkEnd"):
            end_map[el.get(qn("w:id"))] = el
    # 第二遍: 找名字匹配的起点, 顺带查出它的终点
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
    # 拼成 {name: (start, end)} 字典
    return {
        bname: (start_el, end_map.get(bid))
        for bid, (start_el, bname) in starts.items()
    }


# =================================================================
# 书签填充
# =================================================================
def _build_run(text, font=None):
    """内部: 构建一个 <w:r> XML 节点, 带可选字体

    XML 结构:
      <w:r>
        <w:rPr><w:rFonts .../></w:rPr>   ← 可选, 有 font 才加
        <w:t>文字内容</w:t>
      </w:r>
    """
    run = OxmlElement("w:r")
    if font:
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font)      # 西文槽位
        rFonts.set(qn("w:hAnsi"), font)      # 西文槽位
        rFonts.set(qn("w:eastAsia"), font)   # 中文槽位
        rPr.append(rFonts)
        run.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")       # 保留首尾空格
    t.text = text
    run.append(t)
    return run


def _clear_bookmark_content(start, end):
    """内部: 删除书签之间的旧内容(占位符文字)

    start 和 end 互为兄弟节点(同一个父节点下), 从 start 的下一个兄弟开始删,
    删到 end 为止(end 本身保留)。
    """
    if end is None:
        return
    el = start.getnext()                     # lxml: 取下一个兄弟节点
    while el is not None and el is not end:
        nxt = el.getnext()                   # 先存下一个, 再删当前的
        el.getparent().remove(el)            # lxml: 找到父节点再 remove
        el = nxt


def fill_bookmark(doc, name, text, font=None, *, clear=True):
    """按书签填充: 在书签位置插入新文字

    分"清旧"和"插新"两步(仅 clear=True 时清旧):
    ① 清旧: 删掉 bookmarkStart 和 bookmarkEnd 之间的所有 run;
    ② 插新: 构建带字体的 <w:r>, 插到 bookmarkStart 正后方。

    Args:
        doc:   python-docx 的 Document 对象
        name:  书签名
        text:  要写入的新文字
        font:  可选字体名, 如 "微软雅黑" (建议总是传, 避免中文显示问题)
        clear: 是否清除书签内的旧内容, 默认 True
               False 时新文字追加到旧内容后面, 适合追加数据的场景

    Raises:
        ValueError: 书签不存在时抛出
    """
    start, end = find_bookmark(doc, name)
    if start is None:
        raise ValueError(f"书签 '{name}' 不存在")

    if clear:
        _clear_bookmark_content(start, end)   # ① 删旧
    start.addnext(_build_run(text, font))     # ② 插新


def fill_bookmarks(doc, data, font=None, *, clear=True):
    """批量按书签填充: 一次遍历查找所有书签, 然后逐个替换

    相比循环调用 fill_bookmark(doc, name, text, font),
    查找阶段只遍历一次 XML 树, N 个书签的性能从 O(2N * 节点数) 降到 O(节点数)。

    注意: 因为查找只做一次, 后续逐个填充时直接用书签名再查找(单次 O(节点数)
    的开销可以忽略), 避免了"修改树导致旧引用失效"的问题。

    Args:
        doc:   python-docx 的 Document 对象
        data:  dict, {书签名: 要写入的文字, ...}
        font:  可选字体名
        clear: 是否清除书签内的旧内容, 默认 True (同 fill_bookmark)

    Raises:
        ValueError: 有书签不存在时抛出 (一次性列出所有缺失的书签名)

    Example:
        fill_bookmarks(doc, {
            "reporter": "王五",
            "report_date": "2026-08-14",
            "result": "合格",
        }, font="微软雅黑")

        # 追加模式: 不清除旧内容, 新文字追加在后面
        fill_bookmarks(doc, {"note": "补充说明"}, clear=False)
    """
    all_bookmarks = find_bookmarks(doc)      # 一次遍历, 拿到全部书签

    # 检查缺失的书签, 一次性报错而不是填到一半才报
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
            # 不清除时, 用书签名重新查找以获取最新引用
            start, end = find_bookmark(doc, name)
            start.addnext(_build_run(text, font))
