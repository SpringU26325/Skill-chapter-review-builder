# -*- coding: utf-8 -*-
"""crb_docxkit.py — python-docx 的排版原语

段落底纹与边框、run 样式、大纲级别、书签、内部链接、单元格边框与底纹、
表格边距、图片显示宽度。这些函数只碰 docx 对象，不涉及公式。
"""
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from crb_style import BODY, LINK, CN_FONT

PPR_SEQ = ('w:shd','w:tabs','w:suppressAutoHyphens','w:kinsoku','w:wordWrap',
           'w:overflowPunct','w:topLinePunct','w:autoSpaceDE','w:autoSpaceDN',
           'w:bidi','w:adjustRightInd','w:snapToGrid','w:spacing','w:ind',
           'w:contextualSpacing','w:mirrorIndents','w:suppressOverlap','w:jc',
           'w:textDirection','w:textAlignment','w:textboxTightWrap',
           'w:outlineLvl','w:divId','w:cnfStyle','w:rPr','w:sectPr','w:pPrChange')

def shade_para(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    pPr.insert_element_before(shd, *PPR_SEQ)

def border_para(p, side, sz, color, space=6):
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.insert_element_before(pBdr, *PPR_SEQ)
    el = OxmlElement('w:' + side)
    el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
    el.set(qn('w:space'), str(space)); el.set(qn('w:color'), color)
    pBdr.append(el)

def set_run(run, size=10.5, bold=False, color=BODY, font=CN_FONT, italic=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    return run

_U_AFTER = ('w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl',
            'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')

_P_AFTER_OUTLINE = ('w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')

def set_outline(p, lvl):
    """给段落挂大纲级别：导航窗格可见，Word 的折叠三角也依赖它。"""
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn('w:outlineLvl'))
    if old is not None:
        pPr.remove(old)
    el = OxmlElement('w:outlineLvl')
    el.set(qn('w:val'), str(lvl))
    pPr.insert_element_before(el, *_P_AFTER_OUTLINE)

def add_bookmark(p, name, bid):
    """书签范围 = 整段。bookmarkStart 紧跟 pPr，bookmarkEnd 收在段尾。"""
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bid))
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bid))
    pPr = p._p.find(qn('w:pPr'))
    if pPr is not None:
        pPr.addnext(start)
    else:
        p._p.insert(0, start)
    p._p.append(end)

def _underline(r):
    el = getattr(r, '_element', r)          # Run 与裸 w:r 都能传
    rPr = el.get_or_add_rPr()
    u = rPr.find(qn('w:u'))
    if u is None:
        u = OxmlElement('w:u')
        rPr.insert_element_before(u, *_U_AFTER)
    u.set(qn('w:val'), 'single')

def add_link(p, text, anchor, size=8.5, color=LINK, bold=False, font=CN_FONT):
    """追加一段可点击文字。anchor 为书签名（同文档内跳转）。"""
    r = p.add_run(text)
    set_run(r, size, bold, color, font)
    _underline(r)
    h = OxmlElement('w:hyperlink')
    h.set(qn('w:anchor'), anchor)
    h.set(qn('w:history'), '1')
    r._element.addprevious(h)
    h.append(r._element)
    return h

def link_runs(p, anchor):
    """把段内已有的连续 run 包成超链接，保持与公式的原有先后顺序。

    逐个连续组包，不整段搬运：脉络行里夹着 oMath，数学对象不在 run 层级，
    整段包进去会打乱「文字—公式—文字」的顺序。同一段可以有多个 hyperlink。
    """
    kids = list(p._p)
    group = []

    def flush():
        nonlocal group
        if not group:
            return
        h = OxmlElement('w:hyperlink')
        h.set(qn('w:anchor'), anchor)
        h.set(qn('w:history'), '1')
        group[0].addprevious(h)
        for r in group:
            _underline(r)
            h.append(r)
        group = []

    for k in kids:
        if k.tag == qn('w:r'):
            group.append(k)
        else:
            flush()
    flush()

IMG_DPI = 300.0         # 裁图脚本固定出 300dpi；宽度换算与它绑定，改一处要改两处

def _png_px_width(path):
    """读 PNG 的 IHDR 宽（像素）；读不到返回 None。jpg 不解析 SOF，不做半吊子兼容。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(33)
    except OSError:
        return None
    if head[:8] == b'\x89PNG\r\n\x1a\n' and head[12:16] == b'IHDR':
        return int.from_bytes(head[16:20], 'big') or None
    return None

def img_width_cm(path, override=None):
    """显示宽度（cm）：显式 {w=} 优先；否则按 300dpi 还原主干上的物理尺寸，夹在 3–16.6。

    裁图脚本按裁切框的 pt 尺寸乘 300dpi 出图，所以「像素 ÷ 300 × 2.54」正好等于这张图
    在教材上原本多大——比手写一个 cm 数可靠。jpg 读不到像素时退 12cm。
    """
    if override:
        try:
            return float(override)
        except ValueError:
            pass
    w = _png_px_width(path)
    if not w:
        return 12.0
    return min(max(w / IMG_DPI * 2.54, 3.0), 16.6)

def set_cell_border(cell, color, left_sz=24, others_none=True):
    tcPr = cell._tc.get_or_add_tcPr()
    tcB = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:' + side)
        if side == 'left':
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(left_sz))
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        else:
            el.set(qn('w:val'), 'nil' if others_none else 'single')
            if not others_none:
                el.set(qn('w:sz'), '4'); el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'D8DEE6')
        tcB.append(el)
    tcPr.append(tcB)

def set_cell_shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def set_table_layout(tbl, mar_l=140, mar_r=140, mar_t=62, mar_b=62):
    tblPr = tbl._tbl.tblPr
    S_W = ('w:jc','w:tblCellSpacing','w:tblInd','w:tblBorders','w:shd','w:tblLayout',
           'w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange')
    S_B = ('w:shd','w:tblLayout','w:tblCellMar','w:tblLook','w:tblCaption',
           'w:tblDescription','w:tblPrChange')
    S_L = ('w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange')
    S_M = ('w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange')

    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct')
    tblPr.insert_element_before(tblW, *S_W)

    borders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side); e.set(qn('w:val'), 'nil'); borders.append(e)
    tblPr.insert_element_before(borders, *S_B)

    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    tblPr.insert_element_before(layout, *S_L)

    mar = OxmlElement('w:tblCellMar')
    for side, w in (('top', mar_t), ('left', mar_l), ('bottom', mar_b), ('right', mar_r)):
        e = OxmlElement('w:' + side); e.set(qn('w:w'), str(w)); e.set(qn('w:type'), 'dxa'); mar.append(e)
    tblPr.insert_element_before(mar, *S_M)
