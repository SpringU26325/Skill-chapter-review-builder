# -*- coding: utf-8 -*-
"""复习资料 DOCX 生成器（各科共用）：MD -> DOCX

本文件是**入口与组装层**，只管四件事：命令行参数、并发锁、Markdown 解析与块渲染
（`Builder` 类）、构建收尾（统计、保存、被占用时降级）。底层能力在同目录三个模块：

    crb_style.py     颜色与字体常量
    crb_docxkit.py   排版原语：段落底纹与边框、run 样式、大纲级别、书签、链接、单元格、图片宽度
    crb_omml.py      LaTeX → OMML 与后处理：tex2omml / polish_math / _fix_* / _bold_greek

公式的后处理顺序仍由本文件的 `add_math()` 掌握：tex2omml → _force_bar → _fix_accents →
_fix_delims → _fix_empty_base → _fix_operators → _bold_greek → _shrink_omml。改顺序要先想清楚。

用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]
符号表用于自动注入「记号约定」框；不传则不注入。
"""
import re, copy, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 依赖自检必须在第三方 import 之前：否则缺依赖抛的是裸 ImportError 栈，看不出
# 「该用哪个解释器」，实测会被误判成「没装包」而去 pip install 到另一个环境。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import find_root, lock_dir, ensure as _ensure_deps
_ensure_deps()

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from crb_style import (INK, ACCENT, GREEN, GOLD, GREY, MUTED, BODY,
                       CN_FONT, HD_FONT, MONO)
from crb_omml import (NS_M, XML_SPACE, tex2omml, polish_math,
                      _force_bar, _fix_accents, _fix_delims, _fix_empty_base,
                      _fix_operators, _bold_greek, _shrink_omml)
from crb_docxkit import (shade_para, border_para, set_run, set_outline,
                         add_bookmark, add_link, link_runs, set_cell_border,
                         set_cell_shade, set_table_layout, img_width_cm)

# 用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]
MD     = sys.argv[1] if len(sys.argv) > 1 else ""
OUT    = sys.argv[2] if len(sys.argv) > 2 else ""
HEADER = sys.argv[3] if len(sys.argv) > 3 else ""
NOTATION_MD = sys.argv[4] if len(sys.argv) > 4 else ""


# ---------- 并发锁（直接调用本脚本时） ----------
def _subject_of(md):
    stem = os.path.splitext(os.path.basename(md))[0]
    return stem.split('_')[0] if '_' in stem else stem


def _hold_build_lock():
    """按科目拿 `.build_<科目>.lock`，与 复习资料_重建.py 同一把锁。

    直接跑 `复习资料_build.py a.md out.docx` 时无人加锁，两个会话同时保存同一个
    docx 就是后写覆盖前写，静默丢成果。msvcrt 是跨进程锁，子进程重复加同一把会
    自锁，所以 重建.py 通过 CRB_LOCK_HELD 告诉这里「锁在祖先手上」，跳过。
    """
    subj = _subject_of(MD)
    if not subj or os.environ.get('CRB_LOCK_HELD') == subj:
        return None
    root = find_root(os.path.dirname(os.path.abspath(OUT)) or None)
    if not root:
        print('WARN 定位不到工作区，本次不加构建锁；并行会话下可能互相覆盖。')
        return None
    import msvcrt, time
    path = os.path.join(lock_dir(root), '.build_%s.lock' % subj)
    fd = os.open(path, os.O_CREAT | os.O_RDWR)
    t0 = time.time()
    while True:
        try:
            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
            return fd
        except OSError:
            if time.time() - t0 > 900:
                os.close(fd)
                raise SystemExit('等待 %s 构建锁超时：另一个会话正在构建，稍后重试' % subj)
            time.sleep(0.5)


def _release_build_lock(fd):
    if fd is None:
        return
    import msvcrt
    try:
        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
    except OSError:
        pass
    os.close(fd)

# ---------- 配色 ----------

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


NUM_RE = re.compile(r'^[*\s]*§?\s*(\d+(?:\.\d+)*)')
ITEM_RE = re.compile(r'^\s*(?:[-*+]\s+|\d+[.)]\s+)?\**\s*【([^】]{2,60})】')

# 正文里只有这些前缀才是「题」。结构性【】（【国内考法 · …】、【代码验证 · NumPy】、
# 【三种失效的具体样子 · 自建算例】等）不是题目，不进题目索引——否则每节冒出假「对应题目」。
TASK_PREFIX = ('自拟', '课后习题', '教材例题', '课堂例题', '真题', '思考题', '总复习题')


def _is_task(label):
    return any(label.startswith(x) for x in TASK_PREFIX)

# 首页速达指向的固定块。正反两个方向共用这一份名单，避免两处各写一套后漂移。
NAV_BLOCKS = ('本章脉络', '易错点与极客备忘录', '附录 · 题库', '附录：题库', '一页速查')


def _nztxt(s):
    """标题/脉络行的比较键：去公式、去「的」「§」、统一括号冒号、去装饰横线。"""
    s = re.sub(r'\$[^$]*\$', '', s)
    s = s.replace('§', '').replace('的', '').replace('**', '')
    s = s.replace('（', '(').replace('）', ')').replace('：', ':')
    s = re.sub(r'[─—–\s]+', '', s)
    return s


def _core(s):
    """标题主干：「梯度：方向导数最大的那个方向」-> 梯度"""
    n = _nztxt(s)
    for sep in (':', '('):
        if sep in n:
            n = n.split(sep)[0]
    return n


def _item_label(text):
    m = ITEM_RE.match(text.strip())
    return m.group(1).strip() if m else ''


CODE_MARK_RE = re.compile(r'(_\{[^}]*\}|\^\{[^}]*\})')
CODE_INLINE_RE = re.compile(r'(\$[^$\n]+?\$)')
TREE_LINE_RE = re.compile(r'^\s*[│├└]')                 # 有树符就当结构块
TREE_PREFIX_RE = re.compile(r'^([│ ]*)([├└]─[ \t]*)?')


# \dot / \ddot / \vec / \tilde 在 MathML 里是 <mover>，MML2OMML 却转成 m:limUpp
# （上方极限位）——Word 把点顶到离字母很远的地方，实测 ≈ 一个行高。\hat 走的是
# m:acc（正常）。这里把单字符的 limUpp 改写回 m:acc，并用组合字符（U+0307 等）
# 而不是间距字符（U+02D9）。

# 窄重音（点、双点）。叠加重音（\dot{\hat r}）时 Word 会给第二个重音再留一个完整
# 重音槽——点坐在帽顶上还要空出 ≈1 pt，9 pt 正文里看着就是「飘」。Word 自己排
# \hat{\dot r}（窄的在内、宽的外层）是贴的：帽的两臂正好把点围住，点还贴着字母。
# 所以叠加重音一律把窄的换到里层，与 \hat{\dot r} 同形；读者看到的仍是「帽 + 点」，
# 数学含义不变。


# \mathbf{希腊字母} 只给出 m:sty=b / bi，Word 对 U+03B1–U+03C9 不照办（实测
# \boldsymbol{\omega} 渲染出来仍是细的）。数学粗体有独立码位（U+1D6C2–U+1D6DA），
# 换成那个字符才真的粗。大写区有未分配码位、偏移不连续，只处理小写。


# latex2mathml 把 \lim_{x\to0} 输出成 <msub><mo>lim</mo>…</mo></msub>，MML2OMML
# 转成 m:sSub，底下的 run 没有任何样式——Word 对普通字母 run 默认按变量排，
# lim 就成了斜体；m:sSub 又是「右侧下标」，极限落在右下角而不是正下方。
# 算子名一律补 m:sty=p 排正体；极限类再把 m:sSub 改造成 m:limLow 压到下方
# （11-18 用户定：极限放下面，行高不是问题——有极限的地方一般也有分式）。


# 会被 XSL 掏空的参数槽：父节点 -> 该父节点下允许补 ZWSP 的槽名。
# 只收「结构上允许为空、空了却难看」的槽；m:d 交给 _fix_delims，m:f/m:rad 的
# 空分子分母是真缺内容，留着让 audit 报，不在这里掩盖。

# 槽里出现这些结构元素时不再算「空」：那是内容被转坏（如 `(\sin x)'` 的括号被
# 掏空成 m:d），补零宽字符只会把坏结构抹掉，让问题更难查。


MATH_FAIL = []          # 转换失败的 tex 原文，构建结束时汇总，非空即内容有缺口
IMG_FAIL = []           # 插不进去的图片行（路径不存在、写法不合规），同样必须清零
IMG_OK = [0]            # 成功插入的图片数，用于和 audit 的 media 数对账

# ![图注](相对路径){w=6.4} —— 独占一行。路径按 md 所在目录解析；不要用带括号的文件名，
# 正则到第一个 ) 就截断了。
IMG_RE = re.compile(r'^!\[(?P<cap>[^\]]*)\]\((?P<src>[^)]+)\)\s*'
                    r'(?:\{w=(?P<w>[\d.]+)\})?\s*$')
IMG_EXT = ('.png', '.jpg', '.jpeg')


def add_math(p, tex, size=None):
    try:
        el = copy.deepcopy(tex2omml(tex))
    except Exception as e:
        r = p.add_run(tex); set_run(r, color=ACCENT)
        MATH_FAIL.append(tex)
        print(f"  !! 公式失败: {tex[:50]} -> {e}")
        return
    for t in el.iter('{%s}t' % NS_M):
        t.set(XML_SPACE, 'preserve')
    _force_bar(el)
    _fix_accents(el)
    _fix_delims(el)
    _fix_operators(el)
    _bold_greek(el)
    _fix_empty_base(el)          # 最后一步：别的修正可能整段换掉槽里的内容
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


# ---------- 块渲染 ----------
class Builder:
    def __init__(self):
        self.doc = Document()
        self._page()
        self._styles()
        # 跳转索引（render 里的 index_anchors 填充）
        self._hlist = []       # [{lv,text,name,sec}]
        self._hi = 0           # 标题消费游标
        self._bid = 100        # 书签 id 自增
        self._items = []       # 附录条目 [{label,disp,name,sec}]
        self._qi = 0
        self._qmap = {}        # '6.5' -> [item…]
        self._in_appx = False
        self._tree_links_on = False
        self._nav_done = False
        self._sec_row = {}     # 标题书签名 -> 脉络行书签名，供该节「↑ 脉络」回跳
        self._sec_row_order = []   # 脉络行书签的渲染顺序，继承回跳时按它取第一行
        self._h2_of = {}       # h3 书签名 -> 所属 h2，index_anchors 填
        self._tree_n = 0       # 脉络行书签计数
        self._nav_bm = None    # 速达行书签名，供四个固定块「↑ 速达」回跳
        self.link_stat = {'tree': 0, 'tree_miss': [], 'q': 0, 'item': 0, 'item_miss': 0,
                          'nav': 0, 'back': 0}

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

    # ---------- 跳转索引 ----------
    def _next_id(self):
        self._bid += 1
        return self._bid

    def index_anchors(self, blocks):
        """预扫描：给标题与附录条目编号，供前置的脉络块提前指向后置的小节。"""
        cur_sec, appx, cur_h2 = None, False, None
        for blk in blocks:
            if blk[0] == 'raw':
                s = blk[1].strip()
                m = re.match(r'^(#{1,3})\s+(.*)$', s)
                if m:
                    lv, text = len(m.group(1)), m.group(2).strip()
                    if lv == 2:
                        appx = ('附录' in text or '题库' in text)
                        n = NUM_RE.match(text)
                        cur_sec = n.group(1) if n else None
                    entry = {'lv': lv, 'text': text,
                             'name': '_Sec%d' % (len(self._hlist) + 1),
                             'sec': cur_sec if lv >= 2 else None,
                             'parent': cur_h2 if lv == 3 else None}
                    self._hlist.append(entry)
                    if lv == 2:
                        cur_h2 = entry['name']
                    continue
                if appx:
                    self._index_item(s)
                elif cur_sec:
                    # 正文题目行：按所在节挂靠；结构性【】由 _is_task 过滤。
                    lbl = _item_label(s)
                    if lbl and _is_task(lbl):
                        self._index_item(s, sec=cur_sec)
            elif appx:
                if blk[0] == 'box' and blk[2]:
                    self._index_item(blk[2])
                for t in blk[-1]:
                    self._index_item(t)
            elif cur_sec and blk[0] == 'box':
                # 正文题卡 :::quiz 【标签】标题：按所在节挂靠。
                if blk[1] == 'quiz' and blk[2] and _is_task(_item_label(blk[2])):
                    self._index_item(blk[2], sec=cur_sec)
        # 回跳继承用：h3 -> 所属 h2 的书签名。本节自己没有脉络行时退到它的 h3 那行。
        self._h2_of = {h['name']: h.get('parent') for h in self._hlist}

    def _index_item(self, s, sec=None):
        label = _item_label(s)
        if not label:
            return
        # sec 由调用方给定＝正文题按所在节挂靠（【自拟】无节号，标题所在节即挂靠节）。
        # 没给才从标签提取：显式 §X.Y 优先——真题没有教材题号，只能自己声明挂在哪一节；
        # 否则回到「习题 6.5 第 1 题」这种自带节号的写法。
        if sec is None:
            m = re.search(r'§\s*(\d+\.\d+)', label)
            if not m:
                # 只把「习题 6.5 第 1 题」里的 6.5 当节号；「习题 3.1」是题号，不是节号
                m = re.match(r'^[^\d]{0,12}?(\d+\.\d+)\s*第', label)
            if m:
                sec = m.group(1)
        # 显示文字与 sec 来源无关：正文题（【自拟】等）同样要拼出可读的题签。
        if '真题' in label:
            # 「第 3 题」不足以区分多套卷，保留卷名；§ 只用于挂靠，不进显示文字
            disp = re.sub(r'§\s*\d+\.\d+', '', re.sub(r'^真题\s*', '', label)).strip()
            disp = re.sub(r'\s+', ' ', disp) or label
        else:
            disp = label[label.index('第'):] if '第' in label else label
        it = {'label': label, 'disp': disp,
              'name': '_Q%d' % (len(self._items) + 1), 'sec': sec}
        self._items.append(it)
        if sec:
            self._qmap.setdefault(sec, []).append(it)

    def _next_hdr(self, lv, text):
        if self._hi < len(self._hlist):
            h = self._hlist[self._hi]
            if h['lv'] == lv and h['text'] == text:
                self._hi += 1
                return h
            for j in range(self._hi, len(self._hlist)):
                if self._hlist[j]['lv'] == lv and self._hlist[j]['text'] == text:
                    print('  !! 标题索引跳号: %s' % text)
                    self._hi = j + 1
                    return self._hlist[j]
            print('  !! 标题未入索引: %s' % text)
        return None

    @staticmethod
    def _nums(line):
        """脉络行里出现的节号：行首 `§6.5`、括号里的 `（2.3）`。

        括号只认带点的节号。`（物理页 112–117）` 里的 112 是页码，不是节号；
        若拿它去匹配，标题一旦以 112 开头就会连错。
        """
        out = []
        m = re.match(r'^[*\s]*§?\s*(\d+(?:\.\d+)*)', line)
        if m:
            out.append(m.group(1))
        m = re.search(r'[（(]\s*(\d+\.\d+(?:\.\d+)*)', line)
        if m and m.group(1) not in out:
            out.append(m.group(1))
        return out

    def _match(self, line, strict=False):
        n = _nztxt(line)
        if not n:
            return None
        exact = [h for h in self._hlist if _nztxt(h['text']) == n]
        if len(exact) == 1:
            return exact[0]['name']
        for num in self._nums(line):
            hit = [h for h in self._hlist
                   if re.match(r'^\s*' + re.escape(num) + r'(?![\d.])', h['text'])]
            if len(hit) == 1:
                return hit[0]['name']
        core = _core(line)
        if len(core) >= 2:
            hit = [h for h in self._hlist if _core(h['text']) == core]
            if len(hit) == 1:
                return hit[0]['name']
            hit = [h for h in self._hlist if len(_core(h['text'])) >= 3 and _core(h['text']) in n]
            if len(hit) == 1:
                return hit[0]['name']
            hit = [h for h in self._hlist if core in _nztxt(h['text'])]
            if len(hit) == 1 and not (strict and len(core) < 3):
                return hit[0]['name']
            if len(hit) > 1:
                # 同一主干挂多个标题时取最靠上的一层：点脉络是要落到那一节，
                # 不是落到节里的某个小标题；仍并列就放弃
                up = min(h['lv'] for h in hit)
                top = [h for h in hit if h['lv'] == up]
                if len(top) == 1:
                    return top[0]['name']
        # 末招：只取开头的汉字段（「罗尔 f(a)=f(b)…」-> 罗尔），由长到短试前缀。
        # 前缀一变短歧义概率就上升，所以一旦出现多个候选就停手，不再缩。
        cjk = re.match(r'[\u4e00-\u9fff]{2,}', core or n)
        if cjk:
            p = cjk.group(0)
            top_k = len(p) if strict else min(len(p), 8)
            for k in range(top_k, 1, -1):
                q = p[:k]
                hit = [h for h in self._hlist if _core(h['text']).startswith(q)]
                if not hit and k >= 3 and not strict:
                    hit = [h for h in self._hlist if q in _nztxt(h['text'])]
                if len(hit) == 1:
                    return hit[0]['name']
                if len(hit) > 1:
                    up = min(h['lv'] for h in hit)
                    top = [h for h in hit if h['lv'] == up]
                    return top[0]['name'] if len(top) == 1 else None
        return None

    def _resolve(self, txt, strict=False):
        """脉络行 -> 标题书签。命中唯一才连，歧义一律不连。

        先去原样匹配；不成再剥掉行首的节号重来——「§1 罗尔」的 §1 只在
        脉络里出现，标题侧写的是「罗尔中值定理」，靠剥号才能对上。
        """
        line = txt.strip()
        a = self._match(line, strict)
        if a:
            return a
        s = re.sub(r'^\**\s*§?\s*\d+(?:\.\d+)*[\s.、]*', '', line).strip()
        if s and s != line:
            return self._match(s, strict)
        return None

    def _root_anchor(self, txt):
        """脉络首行：只认精确匹配与节号，都不成才回退章标题。

        首行通常是概念式根名（「复合函数 z=f(u,v)…」），做模糊匹配会连到某个
        小节；但也存在整棵树没有根行、首行直接写节名的情况（「五种表示法（2.5）」），
        那种必须连到那一节，不能一律当章标题。
        """
        n = _nztxt(txt)
        if n:
            exact = [h for h in self._hlist if _nztxt(h['text']) == n]
            if len(exact) == 1:
                return exact[0]['name']
        for num in self._nums(txt):
            hit = [h for h in self._hlist
                   if re.match(r'^\s*' + re.escape(num) + r'(?![\d.])', h['text'])]
            if len(hit) == 1:
                return hit[0]['name']
        return self._hlist[0]['name'] if self._hlist else None

    def _tree_link(self, p, txt, depth, first_line):
        if not self._tree_links_on:
            return
        anchor = self._root_anchor(txt) if (first_line and depth == 0) \
            else self._resolve(txt, strict=(depth >= 2))
        if anchor:
            link_runs(p, anchor)
            self.link_stat['tree'] += 1
            # 反向：给这一行挂书签，供该节「↑ 脉络」点回来。同一节被多行指到
            # 只认第一行——回跳落点必须稳定，否则同一节在不同册里落点会漂。
            if anchor not in self._sec_row:
                self._tree_n += 1
                bm = '_Tree%d' % self._tree_n
                add_bookmark(p, bm, self._next_id())
                self._sec_row[anchor] = bm
                self._sec_row_order.append(anchor)
        elif depth <= 1:
            self.link_stat['tree_miss'].append(txt.strip()[:40])
        if os.environ.get('CRB_LINK_DEBUG'):
            print('    [link] %-38s -> %s' % (txt.strip()[:38], anchor or '—'))

    def _try_anchor(self, p, text):
        """附录条目挂书签：按预扫描顺序消费，标签不一致时向后容忍 2 条。"""
        if self._qi >= len(self._items):
            return
        label = _item_label(text)
        if not label or not _is_task(label):
            return
        k = self._qi
        for j in range(self._qi, min(self._qi + 3, len(self._items))):
            if _nztxt(self._items[j]['label']) == _nztxt(label):
                k = j
                break
        else:
            self.link_stat['item_miss'] += 1
            return
        add_bookmark(p, self._items[k]['name'], self._next_id())
        self._qi = k + 1
        self.link_stat['item'] += 1

    def nav_line(self):
        """首页速达：指向本篇几块固定结构，省去翻页。缺哪块就不出现哪项。"""
        got = [h for w in NAV_BLOCKS for h in self._hlist if h['text'] == w]
        if len(got) < 2:
            return
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(2)
        pf.line_spacing = 1.2
        r = p.add_run("速达 ")
        set_run(r, 8.5, True, GOLD, CN_FONT)
        for i, h in enumerate(got):
            if i:
                s = p.add_run(" · ")
                set_run(s, 8.5, False, GREY, CN_FONT)
            add_link(p, h['text'], h['name'], size=8.5)
            self.link_stat['nav'] += 1
        # 反向：四个固定块各自回跳到这里。书签盖整段，点回来落在速达行上
        self._nav_bm = '_Nav'
        add_bookmark(p, self._nav_bm, self._next_id())

    def _back_target(self, hdr):
        """本节的回跳目标：(文案, 书签)；连不上返回 None。

        先认脉络行——它才是这节在整章里的位置；固定块没有脉络行，退回速达。
        「本章脉络」本身不挂速达回跳：速达行就在它上方两三行，指过去是重复。
        """
        if not hdr:
            return None
        bm = self._sec_row.get(hdr['name'])
        if bm:
            return ('↑ 脉络', bm)
        # 本节自己没上树（脉络只写了它的知识点）：退到本节名下第一个上树的 h3 那行。
        # 落点仍在这节的子树里，读者看到的是「我这节在树里的第一处」。
        for name in self._sec_row_order:
            if self._h2_of.get(name) == hdr['name']:
                return ('↑ 脉络', self._sec_row[name])
        if hdr['text'] in NAV_BLOCKS and hdr['text'] != '本章脉络' and self._nav_bm:
            return ('↑ 速达', self._nav_bm)
        return None

    def link_line(self, sec, hdr):
        """小节标题下的链接行：题目在前（前向），回跳在后（反向）。

        没题就不出行——强制自拟要求每节有题，缺题是源稿缺陷，lint 会报。
        回跳挂在这一行尾，不另起段落：每节多一行约 14pt，二十节就是三成页。
        固定块（易错点 / 题库 / 速查）是例外：它们本来就没有「对应题目」行，
        回跳是这一行唯一的内容，不出行就等于速查点进去回不来。
        """
        items = self._qmap.get(sec) or []
        back = self._back_target(hdr)
        fixed = bool(hdr) and hdr['text'] in NAV_BLOCKS
        if not items and not (fixed and back):
            # 普通小节：缺题是源稿缺陷（lint 报 ERROR），
            # 不为回跳单造一行，免得没题的节凭空多占 1 行。
            return
        def _ord(it):
            ns = [int(x) for x in re.findall(r'\d+', it['disp'])]
            return ns or [999]
        items = sorted(items, key=_ord)
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(5)
        pf.line_spacing = 1.2
        pf.left_indent = Cm(0.1)
        # 标题的 keep_with_next 只牵到这一行；这一行若不跟正文同页，
        # 标题就会孤悬页底，正文翻到下一页
        pf.keep_with_next = True
        if items:
            r = p.add_run("对应题目 ")
            set_run(r, 8.5, True, GOLD, CN_FONT)
            for i, it in enumerate(items):
                if i:
                    s = p.add_run(" · ")
                    set_run(s, 8.5, False, GREY, CN_FONT)
                add_link(p, it['disp'], it['name'], size=8.5)
                self.link_stat['q'] += 1
        if back:
            if items:
                s = p.add_run(" · ")
                set_run(s, 8.5, False, GREY, CN_FONT)
            add_link(p, back[0], back[1], size=8.5)
            self.link_stat['back'] += 1

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
        set_run(pl.add_run(HEADER or "复习资料"), 8.5, False, MUTED, CN_FONT)

        pr = c1.paragraphs[0]
        pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pr.paragraph_format.space_before = Pt(0); pr.paragraph_format.space_after = Pt(0)

        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = fp.add_run(); set_run(r, 9, False, MUTED)
        fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE')
        fp._p.append(fld)

    def h1(self, text):
        p = self.doc.add_paragraph()
        p.style = self.doc.styles['Heading 1']
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.55
        p.paragraph_format.left_indent = Cm(0.15)
        shade_para(p, INK)
        r = p.add_run(text)
        set_run(r, 21, True, "FFFFFF", HD_FONT)
        h = self._next_hdr(1, text)
        if h:
            add_bookmark(p, h['name'], self._next_id())
            set_outline(p, 0)

    def sub(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        add_inline(p, text, size=9, color=GREY, font=CN_FONT)

    def h2(self, text, page_break=False):
        if not self._nav_done:
            self._nav_done = True
            self.nav_line()          # 速达行放在第一个二级标题之上，不隶属任何一节
        p = self.doc.add_paragraph()
        p.style = self.doc.styles['Heading 2']
        pf = p.paragraph_format
        if page_break:
            pf.page_break_before = True
        pf.space_before = Pt(12); pf.space_after = Pt(6); pf.line_spacing = 1.22
        pf.left_indent = Cm(0.1)
        pf.keep_with_next = True
        shade_para(p, "EDF2F8")
        border_para(p, 'left', 30, INK, 8)
        add_inline(p, text, size=14.5, color=INK, bold=True, font=HD_FONT)
        self._in_appx = ('附录' in text or '题库' in text)
        self._tree_links_on = (text.strip() == '本章脉络')
        h = self._next_hdr(2, text)
        if h:
            add_bookmark(p, h['name'], self._next_id())
            set_outline(p, 1)
        return h          # render 要靠它把本节的回跳目标传给 link_line

    def h3(self, text):
        p = self.doc.add_paragraph()
        p.style = self.doc.styles['Heading 3']
        pf = p.paragraph_format
        pf.space_before = Pt(9); pf.space_after = Pt(3); pf.line_spacing = 1.3
        pf.keep_with_next = True
        r = p.add_run("▸ ")
        set_run(r, 11, True, GREEN, CN_FONT)
        add_inline(p, text, size=11.5, color=INK, bold=True, font=HD_FONT)
        h = self._next_hdr(3, text)
        if h:
            add_bookmark(p, h['name'], self._next_id())
            set_outline(p, 2)

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

    def _pic(self, container, caption, rel_src, override=None, space_after=9):
        """在 container（doc 或表格单元格）里插一张图 + 图注。

        只收「本身就是图」的内容：流程图符号、结构框图、受力图、波形图。表格写成
        文字表格、公式走 LaTeX，两者都不进这条通道。

        宽度：显式 {w=…} 优先；没写就按「像素 ÷ 300dpi」还原主干上的物理尺寸——
        裁图脚本固定 300dpi 出图，所以这个换算等于原书里这张图原本多大。
        """
        src = rel_src if os.path.isabs(rel_src) else os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(MD)), rel_src))
        if not src.lower().endswith(IMG_EXT):
            IMG_FAIL.append('%s（扩展名只支持 png/jpg）' % rel_src)
            return None
        if not os.path.isfile(src):
            IMG_FAIL.append('%s（文件不存在：%s）' % (rel_src, src))
            return None
        w_cm = img_width_cm(src, override)
        if w_cm > 16.6:
            IMG_FAIL.append('%s（宽 %.1fcm 超过版心 16.6cm，用 {w=…} 收窄）' % (rel_src, w_cm))
            return None
        p = container.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(6); pf.space_after = Pt(2)
        pf.keep_with_next = True        # 图与图注不能被分页拆开
        pf.keep_together = True
        try:
            p.add_run().add_picture(src, width=Cm(w_cm))
        except Exception as e:
            IMG_FAIL.append('%s（插入失败：%s: %s）' % (rel_src, type(e).__name__, e))
            return None
        IMG_OK[0] += 1
        if not caption:
            IMG_FAIL.append('%s（图注为空）' % rel_src)
            return p
        cp = container.add_paragraph()
        cpf = cp.paragraph_format
        cpf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cpf.space_before = Pt(2); cpf.space_after = Pt(space_after)
        cpf.keep_together = True
        set_run(cp.add_run(caption), 8.0, False, MUTED, CN_FONT)
        return p

    def img(self, caption, rel_src, override=None):
        return self._pic(self.doc, caption, rel_src, override)

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
        self._try_anchor(p, text)

    def container(self, kind, title, inner):
        if kind not in FILL:
            # 未知类型以前直接 KeyError 裸崩（P1-2）；降级成 box 并喊一声，让构建跑完
            print('  !! 未知容器类型 %s，按 box 处理（可用：%s）'
                  % (kind, ' / '.join(FILL)), file=sys.stderr)
            kind = 'box'
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
            self._try_anchor(first, title)
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
            first_line = not wrote
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
            self._tree_link(p, txt, depth, first_line)

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
            if s.startswith('!['):
                # 配图题：图放在题干之前/之后均可，独立成段
                m = IMG_RE.match(s)
                if m:
                    self._pic(cell, m.group('cap').strip(), m.group('src').strip(),
                              m.group('w'), space_after=3)
                else:
                    IMG_FAIL.append('%s（题卡内图片行写法不合规）' % s[:60])
                started = True
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
            self._try_anchor(p, txt)
        i = 0
        while i < len(lines):
            s = lines[i].strip()
            if s.startswith('!['):
                flush()
                m = IMG_RE.match(s)
                if m:
                    self._pic(cell, m.group('cap').strip(), m.group('src').strip(),
                              m.group('w'), space_after=3)
                else:
                    IMG_FAIL.append('%s（容器内图片行写法不合规）' % s[:60])
                i += 1
                continue
            if s.startswith('$$'):
                tex = s.strip('$').strip()
                if not tex:
                    MATH_FAIL.append('容器内孤行 $$（引擎只认单行 $$…$$）')
                    i += 1
                    continue
                flush()
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
                self._try_anchor(p, s)
            elif re.match(r'^\d+\.\s', s):
                flush()
                m = re.match(r'^(\d+)\.\s+(.*)$', s)
                p = cell.add_paragraph()
                pf = p.paragraph_format
                pf.left_indent = Cm(0.55); pf.first_line_indent = Cm(-0.38)
                pf.space_after = Pt(2); pf.line_spacing = 1.36
                r = p.add_run(m.group(1) + ". "); set_run(r, 9.5, True, INK, CN_FONT)
                add_inline(p, m.group(2), size=10)
                self._try_anchor(p, s)
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
    b.index_anchors(blocks)
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
            hdr = b.h2(title, page_break=(title in ('一页速查', '附录 · 题库')))
            mm = NUM_RE.match(title)
            b.link_line(mm.group(1) if mm else None, hdr)
        elif s.startswith('### '):
            flush(); b.h3(s[4:].strip())
        elif s.startswith('$$'):
            tex = s.strip('$').strip()
            if not tex:
                # 孤行 $$ 的 strip('$') 是空串：add_math('') 产出空公式，公式内容静默降级成
                # 正文文本，audit 还是绿的（P0-2）。这里拦下并计入公式失败，别让它静默坏。
                flush()
                MATH_FAIL.append('孤行 $$（引擎只认单行 $$…$$）')
                i += 1; continue
            flush(); b.math(tex)
        elif s.startswith('!['):
            # 图片必须独占一行且单独成段：落进 buf 会被 flush() 用空格拼进正文段落，
            # 图就插不进来（这是本分支要放在 else 之前的原因）。
            flush()
            m = IMG_RE.match(s)
            if m:
                b.img(m.group('cap').strip(), m.group('src').strip(), m.group('w'))
            else:
                IMG_FAIL.append('%s（写法应为 ![图注](路径)，可带 {w=6.4}）' % s[:60])
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
        ln = ln.strip()               # 缩进过的 @header 以前认不出，页眉静默回退（P2-2）
        if ln.startswith('@header '):
            return ln[8:].strip()
    base = os.path.basename(md_path)
    subj = base.split('_')[0] if '_' in base else ''
    for ln in lines:
        ln = ln.strip()
        if ln.startswith('# '):
            title = ln[2:].strip()
            return f'{subj} · {title}' if subj else title
    return subj or '复习资料'


def main():
    global HEADER
    if not MD or not OUT:
        print('用法：python 复习资料_build.py <md> <out.docx> "<页眉文字>" [符号表.md]')
        sys.exit(2)
    lock_fd = _hold_build_lock()
    try:
        src = open(MD, encoding='utf-8').read()
        lines = inject_notation(src.split('\n'))
        if not HEADER:
            HEADER = pick_header(lines, MD)
        b = Builder()
        b.header_footer()
        render(b, lines)
        st = b.link_stat
        print('  跳转: 标题书签 %d | 速达 %d | 脉络链接 %d | 回跳 %d | 题目链接 %d | 附录条目 %d/%d'
              % (len(b._hlist), st['nav'], st['tree'], st['back'], st['q'],
                 st['item'], len(b._items)))
        if st['tree_miss']:
            print('  · 脉络中未连接的小节行 %d 条（无同名标题，按纯文本保留）' % len(st['tree_miss']))
            for t in st['tree_miss'][:5]:
                print('      -', t)
        if st['item_miss']:
            print('  ! 附录条目未挂书签 %d 条（标签与预扫描不一致）' % st['item_miss'])
        if IMG_OK[0] or IMG_FAIL:
            print('  插图 %d 张；失败 %d 处' % (IMG_OK[0], len(IMG_FAIL)))
        if IMG_FAIL:
            print('  !! 图片失败清单（成品里该处无图，必须修源稿后重建）：')
            for t in IMG_FAIL[:6]:
                print('      - %s' % t)
        if MATH_FAIL:
            uniq = []
            for t in MATH_FAIL:
                if t not in uniq:
                    uniq.append(t)
            print('  !! 公式失败 %d 处（去重 %d 条，成品里该处是原始 LaTeX 或空白，必须修源稿）'
                  % (len(MATH_FAIL), len(uniq)))
            for t in uniq[:5]:
                print('      - %s' % t[:60])
        polish_math(b.doc)
        trim_trailing_empty(b.doc)
        import os
        out_dir = os.path.dirname(OUT)
        if out_dir:                    # 裸文件名时 dirname 是空串，makedirs('') 会崩（P1-3）
            os.makedirs(out_dir, exist_ok=True)
        try:
            b.doc.save(OUT)
            print("saved:", OUT)
        except PermissionError:
            # 正式路径被预览面板/Word 占用：降级写到 md 同目录，交付时再覆盖。
            # 正常情况下不产生任何 tmp 副本，tmp 只留源稿与脚本。
            # 降级名带 pid：并行会话同时撞占用时，两个副本不会互相覆盖。
            stem = os.path.splitext(os.path.basename(OUT))[0]
            alt = os.path.join(os.path.dirname(MD), '%s.locked-%d.docx' % (stem, os.getpid()))
            b.doc.save(alt)
            print("LOCKED:", OUT)
            print("saved:", alt, "（正式文件被占用，关闭预览后覆盖）")
            # 响亮失败（2026-11-10 定）：退出码非 0，调用方无法误判为交付成功。
            # 正式文件未更新时流程必须停下，查明持有者：
            #    python "<本脚本同目录>\_wholock.py" <docx>
            print("WARNING: 正式文件未更新！退出码 3。先解除占用再重跑，禁止当作已交付。")
            raise SystemExit(3)
    finally:
        _release_build_lock(lock_fd)

if __name__ == '__main__':
    main()
