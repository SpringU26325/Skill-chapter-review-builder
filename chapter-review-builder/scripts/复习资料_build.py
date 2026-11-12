# -*- coding: utf-8 -*-
"""复习资料 DOCX 生成器（各科共用）：MD -> DOCX，公式走 LaTeX->MathML->OMML 原生公式

用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]
符号表用于自动注入「记号约定」框；不传则不注入。
"""
import re, copy, sys, io, unicodedata, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import latex2mathml.converter as l2m
from lxml import etree
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]
MD     = sys.argv[1] if len(sys.argv) > 1 else ""
OUT    = sys.argv[2] if len(sys.argv) > 2 else ""
HEADER = sys.argv[3] if len(sys.argv) > 3 else ""
NOTATION_MD = sys.argv[4] if len(sys.argv) > 4 else ""
from _env import find_xsl

XSL = find_xsl()          # 环境变量 → 包内 assets → Office 安装目录，三级探测

# ---------- 配色 ----------
INK     = "1B3A5C"   # 主色 深靛蓝
ACCENT  = "B03A2E"   # 警示 砖红
GREEN   = "0E6E5C"   # 结论 墨绿
GOLD    = "9C7415"   # 注释 赭金
GREY    = "8A94A0"
BODY    = "1F2933"

FILL = {
    'goal': "E8EFF7",
    'box':  "F3F7F6",
    'note': "FDF8EA",
    'warn': "FBEFEE",
    'code': "F4F6F8",
    'quiz': "F2F1F9",
    'flow': "F3F7F6",
}
BAR = {
    'goal': INK, 'box': GREEN, 'note': GOLD, 'warn': ACCENT, 'code': GREY,
    'quiz': "5B4A9B", 'flow': GREEN,
}

CN_FONT = "等线"
HD_FONT = "微软雅黑"
MONO    = "Consolas"

# ---------- 公式 ----------
_XSLT = etree.XSLT(etree.parse(XSL))
NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
MAT_RE = re.compile(r'\\begin\{(bmatrix|pmatrix|vmatrix)\}(.*?)\\end\{\1\}', re.S)
DELIM = {'bmatrix': ('[', ']'), 'pmatrix': ('(', ')'), 'vmatrix': ('|', '|')}

TEXT_RE = re.compile(r'\\text\{([^}]*)\}')

def _prep(tex):
    # \| 在 XSL 里会被误判成 mfenced 分隔符，吞掉内容 -> 换 \Vert
    tex = tex.replace(r'\|', r'\Vert')
    # \text{} 内的空格会被转成 NBSP（伪换行），中文标签不需要空格 -> 全部去掉
    tex = TEXT_RE.sub(lambda m: r'\text{' + m.group(1).replace(' ', '').replace('\u00a0', '') + '}', tex)
    return tex

def _conv(tex):
    mml = l2m.convert(tex)
    return _XSLT(etree.fromstring(mml.encode("utf-8"))).getroot()

def _mk(tag):
    return etree.Element('{%s}%s' % (NS_M, tag))

def _sub(parent, tag, **attrs):
    e = etree.SubElement(parent, '{%s}%s' % (NS_M, tag))
    for k, v in attrs.items():
        e.set('{%s}%s' % (NS_M, k), v)
    return e

def _spacer(n):
    r"""\qquad / \quad 的 mspace 会被 XSL 丢弃，这里用不换行空格补回来"""
    r = _mk('r')
    t = _sub(r, 't')
    t.text = '\u2003' * n
    t.set(XML_SPACE, 'preserve')
    return r

def _thin():
    r"""相邻矩阵之间补一个窄空格，否则两个 bmatrix 会贴在一起"""
    r = _mk('r')
    t = _sub(r, 't')
    t.text = '\u2009'
    t.set(XML_SPACE, 'preserve')
    return r

def _build_matrix(env, body):
    """自建 m:d + m:m，绕开 XSL 丢失定界符的问题"""
    beg, end = DELIM[env]
    d = _mk('d')
    dPr = _sub(d, 'dPr')
    _sub(dPr, 'begChr', val=beg)
    _sub(dPr, 'sepChr', val='')
    _sub(dPr, 'endChr', val=end)
    e = _sub(d, 'e')
    m = _sub(e, 'm')
    mPr = _sub(m, 'mPr')
    _sub(mPr, 'baseJc', val='center')
    _sub(mPr, 'plcHide', val='on')
    mcs = _sub(mPr, 'mcs')
    rows = [r for r in body.split(r'\\')]
    ncol = max(len(r.split('&')) for r in rows)
    mc = _sub(mcs, 'mc')
    mcp = _sub(mc, 'mcPr')
    _sub(mcp, 'count', val=str(ncol))
    _sub(mcp, 'mcJc', val='center')
    for r in rows:
        mr = _sub(m, 'mr')
        for cell in r.split('&'):
            ce = _sub(mr, 'e')
            if cell.strip():
                for ch in list(_conv(cell.strip())):
                    ce.append(copy.deepcopy(ch))
    return d

TOKEN_RE = re.compile(
    r'(\\qquad|\\quad|\\begin\{(bmatrix|pmatrix|vmatrix)\}(.*?)\\end\{\2\})', re.S)

def tex2omml(tex):
    tex = _prep(tex.strip())
    parts, pos = [], 0
    for m in TOKEN_RE.finditer(tex):
        if m.start() > pos:
            parts.append(('t', tex[pos:m.start()]))
        g = m.group(1)
        if g == r'\qquad':
            parts.append(('sp', 4))
        elif g == r'\quad':
            parts.append(('sp', 2))
        else:
            parts.append(('m', m.group(2), m.group(3)))
        pos = m.end()
    if pos < len(tex):
        parts.append(('t', tex[pos:]))

    if not any(p[0] in ('m', 'sp') for p in parts):
        return _conv(tex)

    root = etree.Element('{%s}oMath' % NS_M, nsmap={'m': NS_M})
    prev = None
    for p in parts:
        if p[0] == 't':
            txt = p[1]
            if txt.strip():
                # 矩阵后面紧跟字母/数字（如 [..]S）时补窄空格
                if prev == 'm' and txt.strip()[:1].isalnum():
                    root.append(_thin())
                for ch in list(_conv(txt)):
                    root.append(copy.deepcopy(ch))
                prev = 't'
        elif p[0] == 'sp':
            root.append(_spacer(p[1])); prev = 'sp'
        else:
            if prev == 'm':
                root.append(_spacer(1))   # 两个矩阵相乘，留 1em
            root.append(_build_matrix(p[1], p[2])); prev = 'm'
    return root

# ---------- 低层工具 ----------
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

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
CODE_MARK_RE = re.compile(r'(_\{[^}]*\}|\^\{[^}]*\})')
CODE_INLINE_RE = re.compile(r'(\$[^$\n]+?\$)')
TREE_LINE_RE = re.compile(r'^\s*[│├└]')                 # 有树符就当结构块
TREE_PREFIX_RE = re.compile(r'^([│ ]*)([├└]─[ \t]*)?')

def _force_bar(el):
    """把 m:type=lin 的分式改回 bar。

    latex2mathml 会把 `1/2`、`1/n`、`x/2` 这类**斜杠除法**转成 MathML 分式，
    MML2OMML 再标成 `lin`（线性分式）——Word 渲染成压扁的斜杠，不是上下堆叠。
    复习资料要的是「分数线在中间、分子分母各一边」，所以一律提升为 bar。
    """
    n = 0
    for f in el.iter('{%s}f' % NS_M):
        fPr = f.find('{%s}fPr' % NS_M)
        if fPr is None:
            continue
        ty = fPr.find('{%s}type' % NS_M)
        if ty is not None and ty.get('{%s}val' % NS_M) == 'lin':
            ty.set('{%s}val' % NS_M, 'bar')
            n += 1
    return n


_W_RPR_AFTER = ('w:szCs', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position',
                'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText',
                'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
                'w:eastAsianLayout', 'w:specVanish')


def _shrink_omml(el, pt):
    """把 OMML 内每个 m:r 的字号压到 pt。

    OMML 默认继承正文 10.5pt，直接塞进 8.6pt 的树行会比周围大一圈。
    """
    sz = str(int(round(pt * 2)))
    for r in el.iter('{%s}r' % NS_M):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            t = r.find('{%s}t' % NS_M)
            if t is not None:
                t.addprevious(rPr)
            else:
                r.append(rPr)
        for tag in ('w:sz', 'w:szCs'):
            e = rPr.find(qn(tag))
            if e is None:
                e = OxmlElement(tag)
            else:
                rPr.remove(e)
            e.set(qn('w:val'), sz)
            if tag == 'w:sz':
                first = next((c for c in rPr if c.tag in {qn(x) for x in _W_RPR_AFTER}), None)
                if first is not None:
                    first.addprevious(e)
                else:
                    rPr.append(e)
            else:
                szel = rPr.find(qn('w:sz'))
                if szel is not None:
                    szel.addnext(e)
                else:
                    rPr.append(e)


def add_math(p, tex, size=None):
    try:
        el = copy.deepcopy(tex2omml(tex))
    except Exception as e:
        r = p.add_run(tex); set_run(r, color=ACCENT)
        print(f"  !! 公式失败: {tex[:50]} -> {e}")
        return
    for t in el.iter('{%s}t' % NS_M):
        t.set(XML_SPACE, 'preserve')
    _force_bar(el)
    if not p.runs:
        # 段落里只有公式时，Word 会当成独立公式居中显示（表格单元格因此对不齐）；
        # 垫一个零宽字符把它压回行内，左对齐交给段落自身的 jc 决定。
        set_run(p.add_run('\u200b'), 10.5)
    p._p.append(el)

INLINE_RE = re.compile(r'(\$[^$]+?\$|\*\*.+?\*\*|`[^`]+`)', re.S)

def add_inline(p, text, size=10.5, color=BODY, bold=False, font=CN_FONT):
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if len(part) > 1 and part[0] == '$' and part[-1] == '$':
            add_math(p, part[1:-1])
        elif len(part) > 4 and part.startswith('**') and part.endswith('**'):
            # 粗体内部可能含公式，递归解析
            add_inline(p, part[2:-2], size, color, True, font)
        elif len(part) > 2 and part[0] == '`' and part[-1] == '`':
            r = p.add_run(part[1:-1]); set_run(r, size - 0.5, bold, GREEN, MONO)
        else:
            r = p.add_run(part); set_run(r, size, bold, color, font)

# ---------- 容器（单格表格） ----------
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

# ---------- 块渲染 ----------
class Builder:
    def __init__(self):
        self.doc = Document()
        self._page()
        self._styles()

    def _page(self):
        s = self.doc.sections[0]
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.top_margin, s.bottom_margin = Cm(1.9), Cm(1.8)
        s.left_margin, s.right_margin = Cm(2.2), Cm(2.2)
        s.header_distance, s.footer_distance = Cm(1.0), Cm(1.0)

    def _styles(self):
        st = self.doc.styles['Normal']
        st.font.name = CN_FONT
        st.font.size = Pt(10.5)
        st.font.color.rgb = RGBColor.from_string(BODY)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), CN_FONT)
        pf = st.paragraph_format
        pf.line_spacing = 1.32
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)

    # 页眉页脚
    def header_footer(self):
        sec = self.doc.sections[0]
        header = sec.header
        p0 = header.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = Pt(1)
        set_run(p0.add_run(""), 1)

        tbl = header.add_table(rows=1, cols=2, width=Cm(16.6))
        p0._p.addprevious(tbl._tbl)
        tbl.autofit = False
        set_table_layout(tbl, mar_l=0, mar_r=0, mar_t=0, mar_b=20)

        tblPr = tbl._tbl.tblPr
        old = tblPr.find(qn('w:tblBorders'))
        if old is not None:
            tblPr.remove(old)
        b = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'right', 'insideV', 'insideH'):
            e = OxmlElement('w:' + side); e.set(qn('w:val'), 'nil'); b.append(e)
        e = OxmlElement('w:bottom')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '6')
        e.set(qn('w:space'), '0'); e.set(qn('w:color'), 'C9D2DC')
        b.append(e)
        tblPr.insert_element_before(b, 'w:shd','w:tblLayout','w:tblCellMar','w:tblLook',
                                    'w:tblCaption','w:tblDescription','w:tblPrChange')

        for j, w in enumerate([Cm(10.4), Cm(6.2)]):
            tbl.columns[j].width = w
            for c in tbl.columns[j].cells:
                c.width = w

        c0, c1 = tbl.cell(0, 0), tbl.cell(0, 1)
        pl = c0.paragraphs[0]
        pl.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pl.paragraph_format.space_before = Pt(0); pl.paragraph_format.space_after = Pt(0)
        set_run(pl.add_run(HEADER or "复习资料"), 8.5, False, GREY, CN_FONT)

        pr = c1.paragraphs[0]
        pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pr.paragraph_format.space_before = Pt(0); pr.paragraph_format.space_after = Pt(0)

        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = fp.add_run(); set_run(r, 9, False, GREY)
        fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE')
        fp._p.append(fld)

    def h1(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.55
        p.paragraph_format.left_indent = Cm(0.15)
        shade_para(p, INK)
        r = p.add_run(text)
        set_run(r, 21, True, "FFFFFF", HD_FONT)

    def sub(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        add_inline(p, text, size=9, color=GREY, font=CN_FONT)

    def h2(self, text, page_break=False):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        if page_break:
            pf.page_break_before = True
        pf.space_before = Pt(12); pf.space_after = Pt(6); pf.line_spacing = 1.22
        pf.left_indent = Cm(0.1)
        pf.keep_with_next = True
        shade_para(p, "EDF2F8")
        border_para(p, 'left', 30, INK, 8)
        add_inline(p, text, size=14.5, color=INK, bold=True, font=HD_FONT)

    def h3(self, text):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(9); pf.space_after = Pt(3); pf.line_spacing = 1.3
        pf.keep_with_next = True
        r = p.add_run("▸ ")
        set_run(r, 11, True, GREEN, CN_FONT)
        add_inline(p, text, size=11.5, color=INK, bold=True, font=HD_FONT)

    def para(self, text, size=10.5):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        add_inline(p, text, size=size)
        return p

    def math(self, tex):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(5); pf.space_after = Pt(6)
        pf.line_spacing = 1.15
        add_math(p, tex)

    def bullet(self, text, idx=None):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.75)
        pf.first_line_indent = Cm(-0.45)
        pf.space_after = Pt(2.5)
        pf.line_spacing = 1.35
        mark = f"{idx}. " if idx else "▪ "
        r = p.add_run(mark)
        set_run(r, 10.5, True, GREEN if not idx else INK, CN_FONT)
        add_inline(p, text)

    def container(self, kind, title, inner):
        fill, bar = FILL[kind], BAR[kind]
        tbl = self.doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_layout(tbl)
        cell = tbl.cell(0, 0)
        cell.width = Cm(16.6)
        set_cell_border(cell, bar, 26 if kind != 'code' else 16)
        set_cell_shade(cell, fill)
        # 整块不跨页，避免框体被腰斩
        if kind != 'flow':
            trPr = tbl.rows[0]._tr.get_or_add_trPr()
            trPr.append(OxmlElement('w:cantSplit'))
        first = cell.paragraphs[0]
        if title:
            first.paragraph_format.space_before = Pt(1)
            first.paragraph_format.space_after = Pt(4)
            add_inline(first, title, size=10.5, color=bar, bold=True, font=HD_FONT)
            target = None
        else:
            target = first

        if kind == 'code':
            self._render_code(cell, inner, first)
        elif kind == 'quiz':
            self._render_quiz(cell, inner, target)
        else:
            self._render_blocks(inner, cell, target)

        # 容器后留白
        sp = self.doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.line_spacing = Pt(5)

    def _render_code(self, cell, lines, first):
        if any(TREE_LINE_RE.match(ln) for ln in lines):
            self._render_tree(cell, lines, first)
            return
        for i, ln in enumerate(lines):
            p = first if i == 0 else cell.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = 1.0
            txt = ln.replace("⚠", "▲").replace("\t", "    ")
            self._code_line(p, txt if txt else " ")

    def _render_tree(self, cell, lines, first):
        """结构块（原「拓扑树」）：去掉 │ ├ └ ─，用缩进表达层级，行内支持 $…$。

        层级 = 该行前导竖线个数 + 是否有 ├/└；没有树符的缩进续行沿用上一行层级。
        纯竖线行（只用来画主干）直接丢弃——主干已经由缩进取代。
        旧写法 `_{..}` / `^{..}` 仍然有效，不用改源稿。
        """
        p, prev_depth, wrote = first, 0, False
        for ln in lines:
            s = ln.replace("⚠", "▲").replace("\t", "    ")
            m = TREE_PREFIX_RE.match(s)
            bars = m.group(1).count("│")
            branch = m.group(2) is not None
            txt = s[m.end():].strip()
            if not txt:
                continue
            depth = bars + (1 if branch else 0)
            if depth == 0 and bars == 0 and not branch and wrote:
                depth = prev_depth                      # 无树符的续行
            prev_depth = depth
            if wrote:
                p = cell.add_paragraph()
            wrote = True
            pf = p.paragraph_format
            pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = 1.14
            pf.left_indent = Cm(0.58 * depth)
            if depth == 0:
                self._code_line(p, txt, 9.0, True, INK, CN_FONT)
            elif depth == 1:
                self._code_line(p, txt, 8.8, True, INK, CN_FONT)
            else:
                self._code_line(p, txt, 8.6, False, "2C3E50", CN_FONT)

    def _code_line(self, p, txt, size=8.6, bold=False, color="2C3E50", font=MONO):
        """等宽/结构行：支持 $…$ 行内公式与 _{..}/^{..}，避免 Eij 这种"无下标"写法"""
        for part in CODE_INLINE_RE.split(txt):
            if len(part) > 2 and part[0] == '$' and part[-1] == '$':
                add_math(p, part[1:-1])
                _shrink_omml(p._p[-1], size)            # 刚追加的 oMath 是末位子元素
            elif part:
                self._code_marks(p, part, size, bold, color, font)

    def _code_marks(self, p, txt, size=8.6, bold=False, color="2C3E50", font=MONO):
        pos = 0
        for m in CODE_MARK_RE.finditer(txt):
            if m.start() > pos:
                self._code_run(p, txt[pos:m.start()], size, bold, color, font)
            body = m.group(0)[2:-1]
            r = self._code_run(p, body, size, bold, color, font)
            r.font.subscript = (m.group(0)[0] == '_')
            r.font.superscript = (m.group(0)[0] == '^')
            pos = m.end()
        if pos < len(txt):
            self._code_run(p, txt[pos:], size, bold, color, font)

    def _code_run(self, p, txt, size=8.6, bold=False, color="2C3E50", font=MONO):
        r = p.add_run(txt)
        set_run(r, size, bold, color, font)
        _t = r._r.find(qn('w:t'))
        if _t is not None:
            _t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def _render_quiz(self, cell, lines, target=None):
        """题卡：题干 / 提示 / 答案 三行式，提示与答案用色区分"""
        marks = (
            ('**题**',   '题',   BAR['quiz'], 9.8, 9.6),
            ('**提示**', '提示', GOLD,        8.8, 9.0),
            ('**答案**', '答案', GREEN,       9.4, 9.4),
        )
        started = target is not None
        for ln in lines:
            s = ln.strip()
            if not s:
                continue
            p = target if (not started and target is not None) else cell.add_paragraph()
            started = True
            pf = p.paragraph_format
            pf.space_before = Pt(0.5); pf.space_after = Pt(0.5); pf.line_spacing = 1.2
            pf.left_indent = Cm(0.62); pf.first_line_indent = Cm(-0.62)
            for key, mark, color, msize, tsize in marks:
                if s.startswith(key):
                    r = p.add_run(mark + " ")
                    set_run(r, msize, True, color, HD_FONT)
                    add_inline(p, s[len(key):].strip(), size=tsize)
                    break
            else:
                add_inline(p, s, size=9.8)

    def _render_blocks(self, lines, cell, target=None):
        """在单元格内渲染块序列"""
        buf = []
        def flush():
            if not buf: return
            txt = "".join(buf).strip()
            buf.clear()
            if not txt: return
            p = cell.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.4
            add_inline(p, txt, size=10)
        i = 0
        while i < len(lines):
            s = lines[i].strip()
            if s.startswith('$$'):
                flush()
                tex = s.strip('$')
                p = cell.add_paragraph()
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)
                add_math(p, tex)
            elif s.startswith('- '):
                flush()
                p = cell.add_paragraph()
                pf = p.paragraph_format
                pf.left_indent = Cm(0.55); pf.first_line_indent = Cm(-0.38)
                pf.space_after = Pt(2); pf.line_spacing = 1.36
                r = p.add_run("▪ "); set_run(r, 9.5, True, GREEN, CN_FONT)
                add_inline(p, s[2:], size=10)
            elif re.match(r'^\d+\.\s', s):
                flush()
                m = re.match(r'^(\d+)\.\s+(.*)$', s)
                p = cell.add_paragraph()
                pf = p.paragraph_format
                pf.left_indent = Cm(0.55); pf.first_line_indent = Cm(-0.38)
                pf.space_after = Pt(2); pf.line_spacing = 1.36
                r = p.add_run(m.group(1) + ". "); set_run(r, 9.5, True, INK, CN_FONT)
                add_inline(p, m.group(2), size=10)
            elif s == '':
                flush()
            else:
                buf.append(s)
            i += 1
        flush()

    def md_table(self, rows):
        header, body = rows[0], rows[2:]
        ncol = len(header)
        tbl = self.doc.add_table(rows=1 + len(body), cols=ncol)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_layout(tbl, mar_l=110, mar_r=110, mar_t=55, mar_b=55)

        RATIO = {2: [0.40, 0.60], 3: [0.20, 0.34, 0.46], 4: [0.17, 0.31, 0.25, 0.27]}
        ratios = RATIO.get(ncol, [1.0 / ncol] * ncol)
        for j, rt in enumerate(ratios):
            w = Cm(16.6 * rt)
            tbl.columns[j].width = w
            for c in tbl.columns[j].cells:
                c.width = w

        # 三线：外框 + 表头下横线
        tblPr = tbl._tbl.tblPr
        old = tblPr.find(qn('w:tblBorders'))
        if old is not None:
            tblPr.remove(old)
        borders = OxmlElement('w:tblBorders')
        for side, sz in (('top', 12), ('bottom', 12)):
            e = OxmlElement('w:' + side)
            e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), str(sz))
            e.set(qn('w:space'), '0'); e.set(qn('w:color'), INK)
            borders.append(e)
        for side in ('left', 'right', 'insideV'):
            e = OxmlElement('w:' + side); e.set(qn('w:val'), 'nil'); borders.append(e)
        e = OxmlElement('w:insideH')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:space'), '0'); e.set(qn('w:color'), 'DCE3EA')
        borders.append(e)
        tblPr.insert_element_before(borders, 'w:shd','w:tblLayout','w:tblCellMar','w:tblLook',
                                    'w:tblCaption','w:tblDescription','w:tblPrChange')

        for j, h in enumerate(header):
            c = tbl.cell(0, j)
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
            c.paragraphs[0].paragraph_format.space_before = Pt(0)
            add_inline(c.paragraphs[0], h.strip(), size=9.5, color=INK, bold=True, font=HD_FONT)
            set_cell_shade(c, "EDF2F8")

        for i, row in enumerate(body, start=1):
            for j in range(ncol):
                txt = row[j].strip() if j < len(row) else ""
                c = tbl.cell(i, j)
                p = c.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.line_spacing = 1.25
                add_inline(p, txt, size=9.5)

        sp = self.doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing = Pt(5)

# ---------- MD 解析 ----------
def split_containers(lines):
    """把 :::kind title ... ::: 解析成块"""
    blocks, i = [], 0
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()
        m = re.match(r'^:::(\w+)\s*(.*)$', s)
        if m:
            kind, title = m.group(1), m.group(2).strip()
            inner, i = [], i + 1
            while i < len(lines) and lines[i].strip() != ':::':
                inner.append(lines[i]); i += 1
            i += 1
            if kind == 'code':
                blocks.append(('code', title, inner))
            else:
                blocks.append(('box', kind, title, inner))
            continue
        blocks.append(('raw', ln))
        i += 1
    return blocks

def render(b: Builder, lines):
    blocks = split_containers(lines)
    i = 0
    buf = []
    def flush():
        if not buf: return
        txt = " ".join(x.strip() for x in buf).strip()
        buf.clear()
        if txt:
            b.para(txt)

    while i < len(blocks):
        blk = blocks[i]
        if blk[0] == 'code':
            flush(); b.container('code', '', blk[2]); i += 1; continue
        if blk[0] == 'box':
            flush(); b.container(blk[1], blk[2], blk[3]); i += 1; continue

        ln = blk[1]
        s = ln.strip()

        if s.startswith('@header '):
            i += 1; continue          # 页眉元数据，只被 main() 读取，不进正文
        elif s.startswith('# '):
            flush(); b.h1(s[2:].strip())
        elif s.startswith('@sub '):
            flush(); b.sub(s[5:].strip())
        elif s.startswith('## '):
            flush()
            title = s[3:].strip()
            b.h2(title, page_break=(title in ('一页速查', '附录 · 题库')))
        elif s.startswith('### '):
            flush(); b.h3(s[4:].strip())
        elif s.startswith('$$'):
            flush(); b.math(s.strip('$'))
        elif s.startswith('|') and i + 1 < len(blocks) and blocks[i+1][0] == 'raw' and re.match(r'^\|[\s:|-]+\|$', blocks[i+1][1].strip()):
            flush()
            rows = []
            j = i
            while j < len(blocks) and blocks[j][0] == 'raw' and blocks[j][1].strip().startswith('|'):
                rows.append([c for c in blocks[j][1].strip().strip('|').split('|')])
                j += 1
            b.md_table(rows)
            i = j
            continue
        elif s.startswith('- '):
            flush(); b.bullet(s[2:].strip())
        elif re.match(r'^\d+\.\s', s):
            flush()
            m = re.match(r'^(\d+)\.\s+(.*)$', s)
            b.bullet(m.group(2).strip(), idx=m.group(1))
        elif s == '':
            flush()
        else:
            buf.append(s)
        i += 1
    flush()

# ---------- 公式后处理：字体 + Unicode 归一化 ----------
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MATH_FONT = "Cambria Math"

def _char_style(ch):
    """Unicode 数学字母 -> m:sty（b / i / bi），其余返回 None"""
    cp = ord(ch)
    def R(a, b): return a <= cp <= b
    if R(0x1D468, 0x1D49B) or R(0x1D63C, 0x1D66F) or R(0x1D71C, 0x1D755) or R(0x1D790, 0x1D7C9):
        return 'bi'
    if (R(0x1D400, 0x1D433) or R(0x1D4D0, 0x1D503) or R(0x1D56C, 0x1D59F)
            or R(0x1D5D4, 0x1D607) or R(0x1D6A8, 0x1D6E1) or R(0x1D756, 0x1D78F)
            or R(0x1D7CA, 0x1D7CB) or R(0x1D7CE, 0x1D7FF)):
        return 'b'
    if R(0x1D434, 0x1D467) or R(0x1D608, 0x1D63B) or R(0x1D6E2, 0x1D71B):
        return 'i'
    return None

def _mk_math_run(text, sty=None, nor=False, sz=None):
    r = etree.Element('{%s}r' % NS_M, nsmap={'m': NS_M})
    rPr = etree.SubElement(r, '{%s}rPr' % NS_M)
    if nor:
        etree.SubElement(rPr, '{%s}nor' % NS_M)
    if sty:
        e = etree.SubElement(rPr, '{%s}sty' % NS_M)
        e.set('{%s}val' % NS_M, sty)
    wRPr = etree.SubElement(r, '{%s}rPr' % W_NS)
    f = etree.SubElement(wRPr, '{%s}rFonts' % W_NS)
    if nor:
        for k in ('ascii', 'hAnsi', 'eastAsia'):
            f.set('{%s}%s' % (W_NS, k), CN_FONT)
    else:
        for k in ('ascii', 'hAnsi', 'cs'):
            f.set('{%s}%s' % (W_NS, k), MATH_FONT)
    if sz:
        # 结构块里的公式被压到 8.6pt，重建 run 时必须带上，否则字号被打回默认
        for tag in ('sz', 'szCs'):
            e = etree.SubElement(wRPr, '{%s}%s' % (W_NS, tag))
            e.set('{%s}val' % W_NS, str(sz))
    t = etree.SubElement(r, '{%s}t' % NS_M)
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return r

def polish_math(doc):
    """每个 m:r 指定 Cambria Math；Unicode 数学字母 NFKC 还原为 ASCII 并补 m:sty。
    不指定数学字体时，公式会继承 Normal 的「等线」——字形回退、字距松散，
    肉眼看起来就像从教材上截下来的一块图。"""
    body = doc.element.body
    for r in list(body.iter('{%s}r' % NS_M)):
        rPr = r.find('{%s}rPr' % NS_M)
        nor = rPr is not None and rPr.find('{%s}nor' % NS_M) is not None
        wRPr = r.find('{%s}rPr' % W_NS)
        sz = None
        if wRPr is not None:
            _sz = wRPr.find('{%s}sz' % W_NS)
            if _sz is not None:
                sz = _sz.get('{%s}val' % W_NS)     # 结构块里预设的小字号要带过去
        ts = r.findall('{%s}t' % NS_M)
        if not ts:
            continue
        text = ''.join(t.text or '' for t in ts)
        if not text:
            # 空 run（多出现在 m:sSub 的空底数里，视觉不显示）：换成带字体的空 run，
            # 否则"缺 Cambria Math 的 run = 0"这条硬指标会被这 3 个空壳打破。
            parent, idx = r.getparent(), r.getparent().index(r)
            parent.insert(idx, _mk_math_run('', None, nor, sz))
            parent.remove(r)
            continue
        parent, idx = r.getparent(), None
        idx = parent.index(r)
        if nor:
            parent.insert(idx, _mk_math_run(text, nor=True, sz=sz))
            parent.remove(r)
            continue
        base = None
        if rPr is not None:
            st = rPr.find('{%s}sty' % NS_M)
            if st is not None:
                base = st.get('{%s}val' % NS_M)
        groups = []
        for ch in text:
            cp = ord(ch)
            sty = _char_style(ch) or base
            if cp == 0xA0:
                ch = '\u2009'          # 数学模式下残留的 NBSP 换成 thin space，避免"伪换行"标记
            out = unicodedata.normalize('NFKC', ch) if 0x1D400 <= cp <= 0x1D7FF else ch
            if groups and groups[-1][0] == sty:
                groups[-1][1].append(out)
            else:
                groups.append((sty, [out]))
        for k, (sty, chars) in enumerate(groups):
            parent.insert(idx + k, _mk_math_run(''.join(chars), sty, sz=sz))
        parent.remove(r)


def inject_notation(lines):
    """记号框只存一处（notation.md），构建时注入；正文里已有则不重复。"""
    if any('记号约定' in ln for ln in lines):
        return lines
    try:
        txt = open(NOTATION_MD, encoding='utf-8').read()
        box = re.search(r'<!-- inject:compact -->(.*?)<!-- /inject:compact -->', txt, re.S).group(1).strip()
    except Exception:
        return lines
    out, done = [], False
    for ln in lines:
        out.append(ln)
        if not done and ln.startswith('@sub '):
            out.extend(box.split('\n'))
            out.append('')
            done = True
    return out


def trim_trailing_empty(doc):
    """清理文末的空段落，并给"以表格结尾"的文档补一个 1pt 段落。

    每个容器（表格）后都会补一个 5pt 的空段作为间隔；页面刚好填满时它会被推到
    下一页，产生一整页只有页眉页码的空白（高数第 6 章下就这样多出第 16 页）。
    但文档又不能以表格结尾——Word 会隐式补一段，同样落到下一页。所以：删掉多余
    空段，只留一个 1pt 的，既合法又不占高度。"""
    body = doc.element.body
    removed = 0
    for el in list(body)[::-1]:
        if el.tag == qn('w:sectPr'):
            continue
        if el.tag == qn('w:p') and not ''.join(el.itertext()).strip():
            body.remove(el)
            removed += 1
            continue
        break

    blocks = [e for e in body if e.tag != qn('w:sectPr')]
    if blocks and blocks[-1].tag == qn('w:tbl'):
        p = OxmlElement('w:p')
        pPr = p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:line'), '20')          # 1pt（单位 twips），固定行距
        sp.set(qn('w:lineRule'), 'exact')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '0')
        pPr.append(sp)
        body.insert(list(body).index(blocks[-1]) + 1, p)
        removed -= 1                        # 净删数：补了一个
    return removed


def pick_header(lines, md_path):
    """页眉优先级：md 里的 @header > 「文件名前缀 · 首行标题」> 文件名前缀。

    让页眉跟着 md 走（单一真相），外部就不需要维护清单文件——
    多科目并发时也就没有共享可变文件可抢。"""
    for ln in lines:
        if ln.startswith('@header '):
            return ln[8:].strip()
    base = os.path.basename(md_path)
    subj = base.split('_')[0] if '_' in base else ''
    for ln in lines:
        if ln.startswith('# '):
            title = ln[2:].strip()
            return f'{subj} · {title}' if subj else title
    return subj or '复习资料'


def main():
    global HEADER
    if not MD or not OUT:
        print('用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]')
        sys.exit(2)
    src = open(MD, encoding='utf-8').read()
    lines = inject_notation(src.split('\n'))
    if not HEADER:
        HEADER = pick_header(lines, MD)
    b = Builder()
    b.header_footer()
    render(b, lines)
    polish_math(b.doc)
    trim_trailing_empty(b.doc)
    import os
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    try:
        b.doc.save(OUT)
        print("saved:", OUT)
    except PermissionError:
        # 正式路径被预览面板/Word 占用：降级写到 md 同目录，交付时再覆盖。
        # 正常情况下不产生任何 tmp 副本，tmp 只留源稿与脚本。
        alt = os.path.join(os.path.dirname(MD), os.path.basename(OUT))
        b.doc.save(alt)
        print("LOCKED:", OUT)
        print("saved:", alt, "（正式文件被占用，关闭预览后覆盖）")
        # 响亮失败（2026-11-10 定）：退出码非 0，调用方无法误判为交付成功。
        # 正式文件未更新时流程必须停下，查明持有者：
        #    python "<本脚本同目录>\_wholock.py" <docx>
        print("WARNING: 正式文件未更新！退出码 3。先解除占用再重跑，禁止当作已交付。")
        raise SystemExit(3)

if __name__ == '__main__':
    main()
