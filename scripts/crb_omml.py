# -*- coding: utf-8 -*-
"""crb_omml.py — LaTeX → MathML → OMML，以及 OMML 后处理

对外只暴露三样：`tex2omml(tex)` 出 OMML 元素、`polish_math(doc)` 做整册收尾、
`NS_M` / `XML_SPACE` 等命名空间常量。后处理函数（`_fix_*` / `_bold_greek` …
全部在 add_math() 里按顺序调用，调用顺序由 复习资料_build.py 掌握。
依赖自检由入口脚本先做，本模块只 import。
"""
import copy
import re
import unicodedata

from lxml import etree
import latex2mathml.converter as l2m
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from _env import find_xsl
from crb_style import CN_FONT, MATH_FONT

XSL = find_xsl()          # 环境变量 → 包内 assets → Office 安装目录，三级探测
_GREEK_LO, _GREEK_HI = 0x03B1, 0x03C9   # α–ω，_bold_greek 的替换区间

_XSLT = etree.XSLT(etree.parse(XSL))

NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

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

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

def _run_upright(r):
    rPr = r.find('{%s}rPr' % NS_M)
    if rPr is None:
        return False
    sty = rPr.find('{%s}sty' % NS_M)
    return sty is not None and sty.get('{%s}val' % NS_M) == 'p'

def _part_upright(part):
    """分子或分母是否「全正体」——用来认单位，不靠字母白名单。

    `\\mathrm{m/s}` 经 latex2mathml 出来带 mathvariant=normal，MML2OMML 标成
    m:sty=p；真正的代数分式（`1/2`、`x/2`、`v/u`、`\\mathrm{P}/T`）没有这个标记。
    上下标里的指数数字（s² 的 2）不带 sty，不算反例，跳过不看。
    """
    if part is None:
        return False
    runs, stack = [], [part]
    while stack:
        cur = stack.pop()
        for ch in cur:
            if not isinstance(ch.tag, str):
                continue
            ln = etree.QName(ch).localname
            if ln == 'r':
                runs.append(ch)
            elif ln in ('sup', 'sub', 'fPr'):
                continue
            else:
                stack.append(ch)
    return bool(runs) and all(_run_upright(r) for r in runs)

def _force_bar(el):
    """lin 分式分两种：真除法提成 bar，单位斜杠保持 lin。

    latex2mathml 会把 `1/2`、`x/2` 这类**斜杠除法**转成 MathML 分式，MML2OMML
    标成 `lin`（线性分式）——Word 渲染成压扁的斜杠，不是上下堆叠，所以要提成 bar。
    但单位 `\\mathrm{m/s}` 走的是同一条链路，它本来就该是斜杠形式（m/s，不是 m 压在
    s 上）。用「分子分母是否全正体」区分：单位是正体，代数式是斜体。
    """
    n_bar = 0
    for f in el.iter('{%s}f' % NS_M):
        fPr = f.find('{%s}fPr' % NS_M)
        if fPr is None:
            continue
        ty = fPr.find('{%s}type' % NS_M)
        if ty is None or ty.get('{%s}val' % NS_M) != 'lin':
            continue
        if _part_upright(f.find('{%s}num' % NS_M)) and \
                _part_upright(f.find('{%s}den' % NS_M)):
            continue            # 单位：保留 lin，斜杠形式
        ty.set('{%s}val' % NS_M, 'bar')
        n_bar += 1
    return n_bar

_ACC_MAP = {'\u02d9': '\u0307',      # ˙ -> 组合上点
            '\u00a8': '\u0308',      # ¨ -> 组合分音符
            '~': '\u0303',           # ~ -> 组合波浪
            '\u2192': '\u20d7',      # → -> 组合右箭
            '\u2190': '\u20d6'}      # ← -> 组合左箭

_NARROW_ACC = ('\u02d9', '\u00a8')

def _acc_el(chrval, base):
    """造一个 m:acc，base 是现成的 m:e（被搬进新节点）。"""
    acc = etree.Element('{%s}acc' % NS_M)
    accPr = etree.SubElement(acc, '{%s}accPr' % NS_M)
    chr_el = etree.SubElement(accPr, '{%s}chr' % NS_M)
    chr_el.set('{%s}val' % NS_M, chrval)
    acc.append(base)
    return acc

def _fix_accents(el):
    """把 limUpp 的重音改写为 m:acc，让重音贴在字母上。

    叠加重音按「窄内宽外」重排：\\dot{\\hat r} 输出成帽在外、点在内。
    """
    n = 0
    for up in list(el.iter('{%s}limUpp' % NS_M)):
        lim = up.find('{%s}lim' % NS_M)
        base = up.find('{%s}e' % NS_M)
        if lim is None or base is None:
            continue
        txt = ''.join(t.text or '' for t in lim.iter('{%s}t' % NS_M))
        if len(txt) != 1 or txt not in _ACC_MAP:
            continue
        parent = up.getparent()
        acc = _acc_el(_ACC_MAP[txt], base)
        inner = None
        if txt in _NARROW_ACC and len(base) == 1:
            inner = base.find('{%s}acc' % NS_M)
        inner_e = inner.find('{%s}e' % NS_M) if inner is not None else None
        if inner_e is not None and len(inner_e):
            # 我们的点比里层那个宽重音更窄：点进里层，宽重音留在外层
            letter = inner_e[0]
            inner_e.remove(letter)
            # 原来的 m:e 里装着外层那个 acc，不能直接搬；换成只装字母的新 m:e
            old_e = acc.find('{%s}e' % NS_M)
            acc.remove(old_e)
            new_e = etree.SubElement(acc, '{%s}e' % NS_M)
            new_e.append(letter)
            inner_e.append(acc)
            parent.replace(up, inner)
            n += 1
            continue
        parent.replace(up, acc)
        n += 1
    return n

_GREEK_BOLD_OFF = 0x1D6C2 - 0x03B1

def _bold_greek(el):
    """m:sty=b/bi 的希腊小写换成数学粗体码位。"""
    n = 0
    for r in el.iter('{%s}r' % NS_M):
        rPr = r.find('{%s}rPr' % NS_M)
        if rPr is None:
            continue
        sty = rPr.find('{%s}sty' % NS_M)
        if sty is None or sty.get('{%s}val' % NS_M) not in ('b', 'bi'):
            continue
        t = r.find('{%s}t' % NS_M)
        if t is None or not t.text or len(t.text) != 1:
            continue
        ch = ord(t.text)
        if not (_GREEK_LO <= ch <= _GREEK_HI):
            continue
        t.text = chr(ch + _GREEK_BOLD_OFF)
        rPr.remove(sty)
        n += 1
    return n

_FUNC_NAMES = ('lim', 'limsup', 'liminf', 'max', 'min', 'sup', 'inf', 'det', 'gcd',
               'Pr', 'dim', 'ker', 'arg', 'deg', 'exp',
               'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
               'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
               'ln', 'log', 'lg')

_LIMIT_NAMES = ('lim', 'limsup', 'liminf', 'max', 'min', 'sup', 'inf',
                'det', 'gcd', 'Pr')

def _run_text(r):
    """取一个 m:r 的文本（可能被拆成多个 m:t）。"""
    return ''.join(t.text or '' for t in r.iter('{%s}t' % NS_M))

def _set_plain(r):
    """给 run 打 m:sty=p，按正体排；m:rPr 必须是 m:r 的第一个子元素。"""
    rPr = r.find('{%s}rPr' % NS_M)
    if rPr is None:
        rPr = etree.Element('{%s}rPr' % NS_M)
        r.insert(0, rPr)
    sty = rPr.find('{%s}sty' % NS_M)
    if sty is None:
        sty = etree.SubElement(rPr, '{%s}sty' % NS_M)
    sty.set('{%s}val' % NS_M, 'p')

def _is_var_start(c):
    r"""判断字符是否为变量名开头（ASCII 字母数字 / 希腊 / 数学字母）。

    算子名后面碰到这种字符就要补细空格：'det'+'A' 排出来是 detA，'det'+'('
    则是标准的 det(A)。数学字母区单列，因为 \mathbf A 出来是 U+1D400 而不是 'A'。
    """
    return bool(c) and ((c.isascii() and c.isalnum())
                        or '\u0370' <= c <= '\u03ff'
                        or '\U0001d400' <= c <= '\U0001d7ff')

def _pad_func_before_var(r):
    r"""算子独占一个 run、变量在下一个 run 时，给算子末尾补细空格。

    下一个兄弟不一定是 run：\ln\sec^2 x 里的 sec 被包在 m:sSup 里（sec 是底数），
    只看下一个 run 会漏，ln 与 sec² 粘成 lnsec²x——那是两个函数套着写，必须留缝。
    """
    t = r.find('{%s}t' % NS_M)
    if t is None or not t.text or t.text.endswith('\u2009'):
        return
    nxt = r.getnext()
    if nxt is None:
        return
    if nxt.tag == '{%s}r' % NS_M:
        nt = nxt.find('{%s}t' % NS_M)
        ok = nt is not None and _is_var_start((nt.text or '')[:1])
    else:
        # 只认「后一个算子带指数」这一种：底数是算子名才补，普通变量/数字不管
        ok = False
        if nxt.tag in ('{%s}sSup' % NS_M, '{%s}sSub' % NS_M, '{%s}sSupSub' % NS_M):
            e = nxt.find('{%s}e' % NS_M)
            er = e.find('{%s}r' % NS_M) if e is not None else None
            ok = er is not None and _run_text(er).rstrip('\u2009') in _FUNC_NAMES
    if ok:
        t.text += '\u2009'
        t.set(XML_SPACE, 'preserve')

def _split_func_run(r, names):
    r"""把 run 里出现的算子名逐段切成独立 run 并排正体，返回切出的个数。

    latex2mathml 常把 \det A+\det B 整串合成单个 m:t='detA+detB'：算子名不
    在 run 开头，按前缀切就漏（全库 128 处，线代第 5 章最密）。这里扫全文、
    按最长名贪心切，'detA+detB' 切成 det/A+/det/B 四段。
    切完给「名后紧跟变量」的补一个 U+2009 细空格：源稿写 \det A，转换时那个
    空格被吞，不补会粘成 detA。名后是括号则不加（\det(A+B) 是标准写法）。
    m:nor 是 \text 的正文（可能是 argument 这种英文词），一律跳过，切了就毁。
    已带 m:sty=p 的 run 不能一律跳过：写成 \ln\sec x 时 latex2mathml 把两个名
    并成一个 m:t='lnsec' 并直接标成正体，跳过就排成「lnsec x」，读起来像一个
    不存在的函数名。所以正体 run 只在切不出第二段时才放过。
    """
    if r.getparent() is None:
        return 0
    if r.find('{%s}rPr/{%s}nor' % (NS_M, NS_M)) is not None:
        return 0
    sty = r.find('{%s}rPr/{%s}sty' % (NS_M, NS_M))
    plain = sty is not None and sty.get('{%s}val' % NS_M) == 'p'
    txt = _run_text(r)
    if not txt:
        return 0
    hits, i = [], 0                         # hits 记每处算子名的 [起, 止)
    while i < len(txt):
        for name in names:                  # names 按长度降序，sinh 先于 sin
            if txt.startswith(name, i):
                hits.append((i, i + len(name)))
                i += len(name)
                break
        else:
            i += 1                          # 该位置不命中，右移一格再试
    if not hits:
        return 0
    if plain and len(hits) < 2:
        return 0                            # 已经正体的单名 run 不必重切
    segs, pos = [], 0                       # segs 是 (文本, 是否算子) 的切段
    for s, e in hits:
        if s > pos:
            segs.append((txt[pos:s], False))
        body = txt[s:e]
        nxt = txt[e:e + 1]
        if _is_var_start(nxt):
            body += '\u2009'                # 名后紧跟变量，补细空格防粘连
        segs.append((body, True))
        pos = e
    if pos < len(txt):
        segs.append((txt[pos:], False))
    parent = r.getparent()
    idx = parent.index(r)
    cut = []                                # 切出来的算子 run，插完再补间隙
    for k, (seg, is_func) in enumerate(segs):
        nr = copy.deepcopy(r)               # 复制保 m:rPr / w:rPr（字体字号不下滑）
        ts = nr.findall('{%s}t' % NS_M)
        ts[0].text = seg
        ts[0].set(XML_SPACE, 'preserve')    # 不写 preserve，Word 会吞首尾空白
        for extra in ts[1:]:                # 原 run 可能拆成多个 m:t，只留第一个
            nr.remove(extra)
        if is_func:
            _set_plain(nr)
            cut.append(nr)
        parent.insert(idx + k, nr)
    parent.remove(r)
    for nr in cut:                          # 变量在下一个 run 时（'ln sec'+'x'）也要留间隙
        _pad_func_before_var(nr)
    return len(hits)

def _fix_delims(el):
    r"""m:d 里的空槽：把被 XSL 当成分隔符的那个字符拿回来当正文。

    `\left(-\frac{a}{b}\right)` 的 MathML 本身没问题（fence + mo + mfrac），但
    MML2OMML.XSL 会把 m:d 内非定界符的 mo 当成 sepChr，并按它把内容切成多个 m:e；
    负号紧跟左括号时前半段就是空的，Word 于是画一个空槽方框（`( □ − a/b )`）。
    修法：紧邻空 e 的那个 e 收下该字符，清空 sepChr，删掉空 e。
    分隔符两侧都非空时（`\left(a-b\right)`）显示本来就对，不动。
    """
    n = 0
    for d in el.iter('{%s}d' % NS_M):
        pr = d.find('{%s}dPr' % NS_M)
        if pr is None:
            continue
        sep = pr.find('{%s}sepChr' % NS_M)
        ch = sep.get('{%s}val' % NS_M) if sep is not None else ''
        if not ch:
            continue
        es = [x for x in d if x.tag == '{%s}e' % NS_M]
        k = next((i for i, e in enumerate(es) if len(e) == 0), None)
        if k is None:
            continue
        # 字符属于离空槽最近的那一侧内容：优先右边（负号在括号开头），没有就看左边
        if k + 1 < len(es):
            host, at_head = es[k + 1], True
        elif k > 0:
            host, at_head = es[k - 1], False
        else:
            continue
        r = _mk('r')
        t = _sub(r, 't')
        t.text = ch
        t.set(XML_SPACE, 'preserve')
        if at_head:
            host.insert(0, r)
        else:
            host.append(r)
        sep.set('{%s}val' % NS_M, '')
        for e in es:
            if len(e) == 0:
                d.remove(e)
        n += 1
    return n

_EMPTY_SLOT_FIX = (
    ('sSup', ('e',)),
    ('sSub', ('e',)),
    ('sSubSup', ('e',)),
    ('sPre', ('e',)),
    ('nary', ('e', 'sub', 'sup')),
)

_STRUCT_TAGS = ('d', 'f', 'm', 'rad', 'nary', 'eqArr', 'func', 'groupChr',
                'limLow', 'limUpp', 'sSup', 'sSub', 'sSubSup', 'sPre',
                'acc', 'bar', 'box', 'borderBox', 'phant')

def _has_structure(e):
    """槽内是否还有嵌套公式结构（说明不是「空」，是内容丢失或粘连）。"""
    for node in e.iter():
        if node is e:
            continue
        tag = node.tag
        if isinstance(tag, str) and tag.startswith('{%s}' % NS_M) and \
                tag.split('}')[1] in _STRUCT_TAGS:
            return True
    return False

def _hidden_slots(node):
    r"""nary 的上下限是否被关掉：`supHide`/`subHide` 为 on 时 Word 不画那一格，
    空着属于正常（`\sum_i` 的 sup 就是空的），补零宽字符只是噪声。"""
    pr = node.find('{%s}naryPr' % NS_M)
    if pr is None:
        return ()
    out = []
    for slot, key in (('sub', 'subHide'), ('sup', 'supHide')):
        e = pr.find('{%s}%s' % (NS_M, key))
        if e is not None and e.get('{%s}val' % NS_M) == 'on':
            out.append(slot)
    return tuple(out)

def _fix_empty_base(el):
    r"""空参数槽填零宽字符：别让 Word 画点线方框。

    latex2mathml 把紧贴上下标的细空 `\,` 当成底数——`0\,^\circ` 里 mspace 是
    `m:e`，MML2OMML.XSL 又把 mspace 丢掉，于是 `m:sSup` 的底数成了空槽。Word
    把空槽画成点线小方框（打印时留 ≈1 个字宽的空白）。同族另有三形：数字写在
    数学外造成的空下标（`C$_{60}$`）、空 run 做底数（`$^{147}\mathrm{Sm}$`）、
    裸算子（`$-\int_a^x$` 没有被积函数）。

    修法统一：给空槽塞一个零宽空格 run。槽里有「看不见的实体」，Word 就不再
    认为它是空参数，方框与多出的间距一起消失；源稿一个字都不用改。
    """
    n = 0
    for name, slots in _EMPTY_SLOT_FIX:
        for node in el.iter('{%s}%s' % (NS_M, name)):
            hide = _hidden_slots(node) if name == 'nary' else ()
            for slot in slots:
                if slot in hide:
                    continue        # nary 的上下限被 supHide/subHide 关掉了，空着也不画
                e = node.find('{%s}%s' % (NS_M, slot))
                if e is None:
                    continue
                if ''.join(e.itertext()).strip():
                    continue        # 槽里有可见字符（含零宽字符）就不算空
                if _has_structure(e):
                    continue        # 槽里嵌着公式结构：内容被转坏了，别在这里抹平
                for c in list(e):   # 槽里只剩空 run（`<m:r><m:t/></m:r>`）也要换掉
                    e.remove(c)
                r = _mk('r')
                t = _sub(r, 't')
                t.text = '\u200b'
                t.set(XML_SPACE, 'preserve')
                e.append(r)
                n += 1
    return n

def _fix_operators(el):
    """算子名（sin、lim…）排正体，极限类的下标压到正下方。"""
    n = 0
    # 整段精确命中直接排正体；其余交给 _split_func_run 扫全文切段——算子名
    # 可能落在 run 中间（'=detA'、'detA+detB'），只认开头会漏一大半。
    names = sorted(_FUNC_NAMES, key=len, reverse=True)
    for r in list(el.iter('{%s}r' % NS_M)):
        txt = _run_text(r)
        if not txt:
            continue
        if txt in _FUNC_NAMES:
            _set_plain(r)
            _pad_func_before_var(r)          # 变量在下一个 run 时也要留间隙
            n += 1
            continue
        n += _split_func_run(r, names)
    for s in list(el.iter('{%s}sSub' % NS_M)):
        e = s.find('{%s}e' % NS_M)
        sub = s.find('{%s}sub' % NS_M)
        if e is None or sub is None:
            continue
        r = e.find('{%s}r' % NS_M)
        if r is None or _run_text(r) not in _LIMIT_NAMES:
            continue
        low = etree.Element('{%s}limLow' % NS_M)
        s.remove(e)
        low.append(e)
        lim = etree.SubElement(low, '{%s}lim' % NS_M)
        for child in list(sub):
            sub.remove(child)
            lim.append(child)
        s.getparent().replace(s, low)
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

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

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
            # 希腊小写的数学粗体码位（U+1D6C2–U+1D6DA）不能还原：Word 对
            # m:sty=b/bi 的希腊字母不照办（实测 \boldsymbol{\omega} 仍是细的），
            # 只有字形本身带粗才显示为粗。还原成 ω 再补 sty 等于把粗体丢了。
            if 0x1D6C2 <= cp <= 0x1D6DA and sty in ('b', 'bi'):
                out, sty = ch, None
            if groups and groups[-1][0] == sty:
                groups[-1][1].append(out)
            else:
                groups.append((sty, [out]))
        for k, (sty, chars) in enumerate(groups):
            parent.insert(idx + k, _mk_math_run(''.join(chars), sty, sz=sz))
        parent.remove(r)
