# -*- coding: utf-8 -*-
"""md 级前置检查（2026-11-10 新增）：写完 md、构建 docx 之前跑。

背景：11-09 一天新增 8+ 条公式坑，全部是 docx 生成后由 audit 报错再人工定位。
本脚本把已踩的坑规则化，在 md 阶段就拦截：
  - ERROR：必然产出错误结果（表格被截断、转换失败、残留反斜杠）
  - WARN ：高概率出问题，需人工确认

规则全部来自工作区的 复习资料/_踩坑清单.md，新坑先记清单、再落成规则。

用法：python 复习资料_lint_md.py <md> [md2 ...]
退出码：0 无 ERROR / 2 有 ERROR（WARN 不影响退出码，但需在自检报告说明）。
docx 生成后的后置检查仍用 复习资料_audit.py，两者不可互替。
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 下面 _empty_delim_math 会惰性 import 引擎，引擎也会包一层 sys.stdout；旧 wrapper
# 被 GC 时会关掉共享的 buffer（实测 "I/O operation on closed file"）。持有引用避免。
_KEEP_STDOUT = [sys.stdout]


def _inline_spans(line):
    """产出 (start, end, content)：一行里各个行内 $…$ 的内容区间。

    跳过 \\$ 转义与 $$ 块定界。不能靠正则的 `$…$` 配对：`$A$ | $B$`
    这种「闭合 $ + 表格竖线 + 下一个开 $」会被当成一对，是历史误报的根源。
    多行 $$ 块在本行找不到闭合就整体放弃扫描，块内内容不参与行内规则。
    """
    i, n = 0, len(line)
    while i < n:
        ch = line[i]
        if ch == '\\':
            i += 2
            continue
        if ch != '$':
            i += 1
            continue
        if i + 1 < n and line[i + 1] == '$':
            j = line.find('$$', i + 2)
            if j < 0:
                return
            i = j + 2
            continue
        j, closed = i + 1, -1
        while j < n:
            if line[j] == '\\':
                j += 2
                continue
            if line[j] == '$':
                closed = j
                break
            j += 1
        if closed < 0:
            return
        yield i, closed, line[i + 1:closed]
        i = closed + 1


def _pipe_in_table_math(line):
    """表格行里，行内公式内部出现未转义的 |：会被当成列分隔符截断整行。"""
    if not line.lstrip().startswith('|'):
        return False
    for _, _, content in _inline_spans(line):
        for k, ch in enumerate(content):
            if ch == '|' and (k == 0 or content[k - 1] != '\\'):
                return True
    return False


def _begin_in_inline_math(line):
    """行内公式里放 \\begin{…}（矩阵/环境）：必须提到块级。"""
    return any('\\begin{' in c for _, _, c in _inline_spans(line))


def _multi_dfrac(line):
    """一行里出现 ≥2 个行内 \\dfrac：每个都会挤成小号，且各自抬高行距。"""
    return sum(c.count('\\dfrac') for _, _, c in _inline_spans(line)) >= 2


def _long_inline_eq(line):
    """行内公式 ≥80 字符且含 =：Word 里容易被折断。

    不能用正则 `\\$[^$\\n]{80,}\\$`：遇到 `$$…$$` 块行，它会从第二个 `$` 配到
    闭合 `$$` 的第一个 `$`，把每个长块级公式都当成行内公式（11-11 实测 38 章
    共 460 行误报）。改用 _inline_spans，块级内容根本不进入判定。
    """
    return any(len(c) >= 80 and '=' in c for _, _, c in _inline_spans(line))


TREE_LINE_RE = re.compile(r'^\s*[│├└]')


BARE_MARK_RE = re.compile(r'[A-Za-z][_^][A-Za-z0-9]')


def _bare_math_in_tree(s):
    """树块（本章脉络）里裸写的 x_y / x^2 不会渲染，只认 $…$ 与 _{…}/^{…}。

    11-15 实测：`f_u * u_x` 原样进 Word，整块脉络没有一条公式。
    `x_{y}` / `x^{2}` 已被 _code_marks 认得，不算裸写；Unicode 下标（`xᵢ`）
    是 11-18 之前的写法，旧稿沿用，新稿优先 `$x_{y}$`。
    只查"裸下/上标"这一种，`sum=a+b` 这类 C 代码不含 _ ^，不会误伤。
    """
    outside = re.sub(r'\$[^$\n]*\$', '', s)     # 先摘掉已经包好的公式
    if '$' in s:
        if BARE_MARK_RE.search(outside):
            return [r'$…$ 外面还有裸写 x_y：并进公式，或写成 x_{y}']
        return []
    if BARE_MARK_RE.search(outside):
        return [r'树块里裸写 x_y 不会变下标，包成 $x_{y}$（Unicode xᵢ 也行，旧稿沿用）']
    return []


def _tree_rule_hits(s):
    """结构块（本章脉络）的写法检查。

    11-18 起树块公式改用 `$…$` + `\\dfrac`（横线分式、矢量加粗），行高允许不齐，
    所以不再拦「分式撑高」；但一行只能放一个分式——并排两个会把该行抬到相邻行的
    两倍高，那种参差就没法看了。
    """
    hits = []
    if any('/' in seg and '\\frac' not in seg
           for seg in re.findall(r'\$([^$\n]+)\$', s)):
        hits.append(r'树块里分式写 \dfrac（横线），斜杠会排成线性分式')
    if '\\limits' in s:
        hits.append(r'结构块里 \limits 会把上下限压到下方、撑高行距，去掉 \limits')
    hits.extend(_bare_math_in_tree(s))
    return hits


def _overline_in_math(line):
    """公式里的 \\overline：横线是 U+2015 破折号字形，会从字母中间穿过去。

    latex2mathml 把 \\overline{BC} 转成 <mo>U+2015</mo>，MML2OMML 当重音字符塞进
    m:acc；但那个字形本身排在字高中部，Word 不会把它抬成上划线——出来就是一道
    删除线。\\bar 走 m:bar（Word 自己画的线），才是真上划线。
    不带花括号的 \\overline D 同样走 m:acc（实测），不能只拦花括号形式。
    """
    return bool(re.search(r'\\overline\s*[\{A-Za-z]', line))


def _abs_with_paren(line):
    """公式里的 |( ：裸括号会被当定界符的「分隔符」，砸出一个空槽。

    |(\\Delta\\mathbf{v})_n| 里 latex2mathml 把 _n 挂在 ) 上，MML2OMML 于是把
    ( 当成 m:sepChr，拆出两个参数槽、第一个是空的——成品里 | 与内容之间空一块，
    Word 编辑态显示成虚线占位框。内层括号写成 \\left( … \\right) 即可。
    """
    segs = [c for _, _, c in _inline_spans(line)]
    segs += [m.group(1) for m in re.finditer(r'\$\$(.+?)\$\$', line)]
    return any('|(' in seg for seg in segs)


def _empty_delim_math(line):
    """括号里只有运算符（[\\cdot]、(\\times)）会让公式整个丢失，用 OMML 复算判定。

    latex2mathml 把 · 当成定界符的 sepChr，MML2OMML 于是拆出两个空槽、m:r 为 0；
    成品里这两个括号之间空一块，内容全没了。正则只挑「[ 或 ( 开头、] 或 ) 结尾」
    的行内公式当候选，真正判据是复算后的 m:r 数——[\\alpha] 这类正常公式也会被
    挑出来，但复算 m:r 非 0，不会误报。
    """
    cands = [c for _, _, c in _inline_spans(line)
             if c[:1] in '([' and c[-1:] in '])']
    if not cands:
        return False
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        if here not in sys.path:
            sys.path.insert(0, here)
        import 复习资料_build as _b
        if sys.stdout not in _KEEP_STDOUT:
            _KEEP_STDOUT.append(sys.stdout)  # 别让引擎那层 wrapper 被 GC 关掉 buffer
    except Exception:
        return False                    # 引擎不可用时不误报，交给 audit 兜底
    for c in cands:
        try:
            el = _b.tex2omml(c)
            if not list(el.iter('{%s}r' % _b.NS_M)):
                return True
        except Exception:
            pass                        # 转换异常由别的规则与 audit 负责
    return False


# (级别, 说明, 正则/判定函数)
RULES = [
    ('ERROR', r'表格行内公式里出现裸 |（行内公式外的列分隔符不算），改 \vert',
     _pipe_in_table_math),
    ('ERROR', r'\\text{…}\ （反斜杠空格）留 NBSP，改 \quad',
     re.compile(r'\\text\{[^{}]*\}\\\s')),
    ('ERROR', r'\\big\vert / \\Big\vert 带上下标漏成字面 \vert，改代入求值写法',
     re.compile(r'\\[bB]ig[a-zA-Z]*\s*\\vert\s*[_^{]')),
    ('ERROR', r'\\displaystyle 不被认，行内直接省掉',
     re.compile(r'\\displaystyle')),
    ('ERROR', r'\\bigr)\'(N) 求导记法转换后 m:r 为 0，改 \left(f^{-1}\right)\'(N)',
     re.compile(r"\\bigr\)\s*'")),
    ('ERROR', r'\\[Npt] 行距参数禁用，只写 \\，行距靠段落样式',
     re.compile(r'\\\\\[\d+(pt|ex|em|mm)\]')),
    ('ERROR', r'括号内只有运算符、没有字母（[\cdot] 类），m:r 为 0、内容整个丢失，改 \lbrack…\rbrack',
     _empty_delim_math),
    ('ERROR', r'行内 $…$ 里放 \begin{…}（矩阵/环境），提到块级公式',
     _begin_in_inline_math),
    ('ERROR', r'::: 容器未闭合或多余 :::',
     re.compile(r'^\s*:::\s*$')),  # 计数型规则，特殊处理见下
    ('ERROR', r'\\overline{…} / \\overline D 生成的是破折号字形，线会从字母中间穿过，改 \\bar',
     _overline_in_math),
    ('ERROR', r'公式里 |( 没写 \left(\right)：裸括号被当分隔符，会砸出空槽（成品里 | 与内容空一块）',
     _abs_with_paren),
    ('WARN', r'\sqrt{单符号^{…}} 上标会跑到根号外，改写（如 \sqrt{x\cdot x}）',
     re.compile(r'\\sqrt\{[A-Za-z\\]\^\{')),
    ('WARN', r'一行里 ≥2 个行内 \dfrac，各占一份行高，提到块级或拆行',
     _multi_dfrac),
    ('WARN', r'\\text{…} 超 4 字，整句移到公式外',
     re.compile(r'\\text\{([^{}]{5,})\}')),
    ('WARN', r'U+2080 区下标字符出现在正文，写 $v_1$（树块内可用，见 md-dialect 第 6 节）',
     re.compile(r'[\u2080-\u209f\u2070-\u2079]')),
    ('WARN', r'树块里连写 ≥3 个 ─，进 Word 会成一条突兀长横线，写 ── 即可',
     re.compile(r'─{3,}')),
    ('WARN', r'长行内公式含 = 连等式，易被 Word 折断，拆短用文字连',
     _long_inline_eq),
]


def _check_container_balance(text):
    """::: 开头（带类型）与单独 ::: 结尾必须配对；顺带查 :::code 嵌进别的容器。

    嵌套会静默错位：容器内第一个 `:::` 被当结束符,内外层级立刻乱，而引擎不报错。
    """
    errs = []
    opens = closes = 0
    stack = []
    for i, line in enumerate(text.splitlines(), 1):
        s = line.strip()
        m = re.match(r'^:::([a-zA-Z]\w*)\s*(.*)$', s)
        if m:
            opens += 1
            kind = m.group(1)
            if kind == 'code' and stack:
                errs.append(f'  ERROR L{i}: :::code 不能嵌在 :::{stack[-1]} 里——'
                            f'容器内第一个 ::: 会被当结束符，内外错位；移到容器外')
            stack.append(kind)
        elif s == ':::':
            closes += 1
            if stack:
                stack.pop()
    if opens != closes:
        errs.append(f'  ERROR L?: ::: 容器不配对（开头 {opens} / 结尾 {closes}）')
    return errs


def _check_container_kind(lines, code_lines):
    """::: 类型必须在引擎的 FILL 表里，否则构建直接 KeyError 裸崩（P1-2）。

    配对检查只数开合，不管类型；`:::tip 提示` 能过 lint，到构建才炸，
    traceback 里也没有「可用哪几种」的信息。
    """
    kinds = ('goal', 'box', 'note', 'warn', 'code', 'quiz', 'flow')  # 与 build.FILL 同步
    errs = []
    for i, line in enumerate(lines, 1):
        if i in code_lines:
            continue
        m = re.match(r'^:::([a-zA-Z]\w*)', line.strip())
        if m and m.group(1) not in kinds:
            errs.append(f'  ERROR L{i}: 未知容器类型 :::{m.group(1)}——构建会因 KeyError 中断；'
                        f'可用：{" / ".join(kinds)}')
    return errs


def _check_dollar_blocks(lines, code_lines):
    """`$$` 只有「单行 $$…$$」一种写法能正确成块。

    引擎逐行解析：孤行 $$ 的 strip('$') 是空串，产出空公式、公式内容降级成正文文本；
    行中 $$ 只从第二个 $ 起配对，两个孤儿 $ 落进正文。两者 audit 都查不出（P0-2 / P0-3）。
    """
    errs = []
    for i, line in enumerate(lines, 1):
        if i in code_lines:
            continue
        s = line.strip()
        if not s or '$' not in s:
            continue
        if s.startswith('$$'):
            if s == '$$':
                errs.append(f'  ERROR L{i}: 孤行 $$ 不构成公式——引擎产出空公式、内容静默降级成正文；'
                            f'改写成单行 $$…$$')
            elif not s.endswith('$$'):
                errs.append(f'  ERROR L{i}: $$ 块首尾不在同一行——引擎按单行解析，尾部会漏成正文；'
                            f'改写成单行 $$…$$')
        elif '$$' in s:
            errs.append(f'  ERROR L{i}: 行中 $$ 会把美元符印进正文；显示公式改为独占一行 $$…$$')
    return errs


def _code_blocks(lines):
    """按整块判定 :::code，与引擎 _render_code 一致：块内任一行带树符即整块算树。

    返回两个 1-based 行号集合：
      code_lines —— 块内全部行。代码是原文，反引号/反斜杠/树符都不该按正文规则查。
      tree_lines —— 块内任一行带树符时整块都要跑树规则（根行本身不带树符，也要覆盖）。
    """
    code_lines, tree_lines = set(), set()
    block, in_code = [], False
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if s.startswith(':::code'):
            in_code, block = True, []
            continue
        if in_code and s == ':::':
            rng = range(i - len(block), i)
            code_lines.update(rng)
            if any(TREE_LINE_RE.match(ln) for ln in block):
                tree_lines.update(rng)
            in_code = False
            continue
        if in_code:
            block.append(line)
    return code_lines, tree_lines


TREE_PRE_RE = re.compile(r'^([│ ]*)([├└]─[ \t]*)')


def _check_tree_depth(lines):
    """树子项只写空格、不写 │ 时，层级会静默塌回一级。

    引擎 depth = 前导 │ 个数 + 分支符。`└─ 父行` 是 1 级，其子项写成
    `    ├─` 得 0 个 │、仍是 1 级，渲染出来与父行同缩进同字重。11-15 全量
    32 份中招（线代 15、高数 13、物理拓展 3、程序设计第 4 章）。
    """
    errs, in_code, block_start, prev = [], False, 0, None
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if s.startswith(':::code'):
            in_code, block_start, prev = True, i, None
            continue
        if in_code and s == ':::':
            in_code, prev = False, None
            continue
        if not in_code:
            continue
        t = line.replace('\t', '    ')
        m = TREE_PRE_RE.match(t)
        if not m:
            if not t.strip():
                prev = None
            continue
        bars, bcol = m.group(1).count('│'), len(m.group(1))
        if not t[m.end():].strip():
            continue
        depth = bars + 1
        if prev and bars == 0 and bcol > prev[1] and depth <= prev[0]:
            errs.append(f'  ERROR L{i}: 树子项只用了空格缩进——引擎按前导 │ 计层级，'
                        f'这一行仍是 {depth} 级，与父行同级')
            errs.append(f'        | {s[:100]}')
            errs.append('        改：行首补 │（父行是 `└─` 也要补），见 md-dialect 结构块第 4 条')
        prev = (depth, bcol)
    return errs


EXT_SRC_RE = re.compile(r'【([^】]*)】')


def _check_external_src(lines):
    """外部资料标签：真题必须可唯一标识，笔记不得进附录。

    真题没有教材题号，只能靠「年份 + 卷别 + 题号」区分；缺年份或卷别时，两套卷的
    「第 3 题」在「对应题目」行里会显示成同一串文字。`【笔记 …】` 写进附录会被引擎
    当题库条目索引，显示成「第 N 页」混在题目链接里。
    """
    out = []
    seen = {}
    in_appx = False
    for i, line in enumerate(lines, 1):
        s = line.strip()
        m2 = re.match(r'^(#{1,3})\s+(.*)$', s)
        if m2:
            in_appx = len(m2.group(1)) == 2 and ('附录' in m2.group(2) or '题库' in m2.group(2))
            continue
        m = EXT_SRC_RE.search(s)
        if not m:
            continue
        label = m.group(1).strip()
        if '真题' in label:
            if not re.search(r'20\d\d', label) or '卷' not in label:
                out.append(f'  ERROR L{i}: 真题标签缺年份或卷别，多套卷之间无法唯一标识')
                out.append(f'        | {s[:100]}')
                out.append('        改：【真题 2024秋期末A卷 第3题】')
            elif '§' not in label:
                out.append(f'  WARN  L{i}: 真题未标 §X.Y，小节下不会生成「对应题目」')
                out.append(f'        | {s[:100]}')
            if label in seen:
                out.append(f'  ERROR L{i}: 真题标签与 L{seen[label]} 重复，跳转按顺序消费会错配')
                out.append(f'        | {s[:100]}')
            else:
                seen[label] = i
        elif label.startswith('笔记') and in_appx:
            out.append(f'  ERROR L{i}: 笔记条目不能放进附录区，会被索引成题库题目')
            out.append(f'        | {s[:100]}')
            out.append('        改：移到正文 :::note 内')
    return out


# ---------- 转换链不认 / 会静默出错的写法 ----------
# 这三条构建不报错、audit 也查不出（空公式里连反斜杠都没有），只有打开 docx
# 才看得见。换模型时最容易在这里翻车，所以全部拦在 lint。

def _dd_block_lines(lines):
    """多行 $$…$$ 块占用的行号（1-based）；单行 $$…$$ 不计。"""
    out, in_dd = set(), False
    for i, line in enumerate(lines, 1):
        odd = line.count('$$') % 2 == 1
        if in_dd:
            out.add(i)
            if odd:
                in_dd = False
            continue
        if odd:
            in_dd = True
            out.add(i)
    return out


def _check_engine_math(lines, code_lines):
    """矩阵简写与 \\underline：引擎不认，成品里变平铺字符或空白。"""
    errs = []
    dd = _dd_block_lines(lines)
    for i, line in enumerate(lines, 1):
        if i in code_lines:
            continue
        if '[[' in line and ('$' in line or i in dd):
            errs.append(f'  ERROR L{i}: 数学里的 `[[…]]` 不是引擎语法，只会当普通字符平铺'
                        f'（实测成品是 `[1,0.2],[0,0.8]`，相邻两个还粘连）；'
                        f'矩阵写 \\begin{{bmatrix}} 并放 $$ 块级（行内禁 \\begin），或改文字描述')
            errs.append(f'        | {line.strip()[:100]}')
        if '\\underline' in line:
            errs.append(f'  ERROR L{i}: \\underline{{…}} 经 latex2mathml 变成 munder，引擎不支持，'
                        f'成品里该公式整段为空；填空直接写 ____')
            errs.append(f'        | {line.strip()[:100]}')
    return errs


def _check_appendix(lines):
    """有题卡就必须有「附录 · 题库」，且附录里要有 `【…】` 条目。

    引擎的 appx 开关只认含「附录 / 题库」的二级标题。缺了它 _qmap 恒空：
    每节少一行「对应题目」、题卡没有 _Q 书签，构建只打印「题目链接 0」——
    不报错、不失败，跳转整体失效。
    """
    quiz = sum(1 for l in lines if l.strip().startswith(':::quiz'))
    if quiz < 4:
        return []
    head = None
    for i, l in enumerate(lines):
        m = re.match(r'^##\s+(.*)$', l.strip())
        if m and ('附录' in m.group(1) or '题库' in m.group(1)):
            head = i
            break
    if head is None:
        return [f'  ERROR L?: 有 {quiz} 道题卡却没有「## 附录 · 题库」——题库索引为空，'
                f'各节「对应题目」行与题卡书签都不会生成（构建只显示「题目链接 0」，不报错）',
                '        改：在「## 一页速查」之前补 `## 附录 · 题库`，'
                '用 `:::flow 先做题` + `1. 【课后习题 10.5 第 1 题】…` 列未配答案的题']
    body = []
    for l in lines[head + 1:]:
        if re.match(r'^##\s+', l.strip()):
            break
        body.append(l)
    if not any('【' in l for l in body):
        return [f'  ERROR L{head + 1}: 附录区没有 `【…】` 条目，题库索引仍为空，'
                f'「对应题目」行不会生成',
                '        改：每条写成 `1. 【课后习题 10.5 第 1 题】题干…`']
    return []


# 与 build.py 的 TASK_PREFIX 保持一致：正文只有这些前缀才是题，
# 【国内考法 · …】/【代码验证 · …】等结构性标签不算。
_TASK_PREFIX = ('自拟', '课后习题', '教材例题', '课堂例题', '真题', '思考题', '总复习题')
_TASK_RE = re.compile(r'【([^】]{2,60})】')
_SEC_RE = re.compile(r'^##\s+[\*\s]*§?\s*(\d+(?:\.\d+)+)')


def _check_section_tasks(lines):
    """强制自拟的可执行校验：带节号的小节至少一道题。

    引擎改成「正文题按所在节进 _qmap」后，没题的节连「对应题目」行都不出，
    回跳也因此消失。此前这种节靠静默跳过混过去（高数 5 的 5.1–5.5 全无题卡但有
    正文题，属假警报；大学物理 1.2 等是真没题）。这里按题型白名单逐节核，
    正文题或附录里带 §X.Y 的题都算这一节有题。
    """
    out, cur, has = [], None, False
    for i, line in enumerate(lines, 1):
        s = line.strip()
        m = _SEC_RE.match(s)
        if m:
            if cur and not has:
                out.append(f'  ERROR L?: 第 {cur} 节没有任何题目（正文无【自拟】/【课后习题…】，'
                           f'附录也无 §{cur} 的题）——该节不出「对应题目」行，回跳也消失')
                out.append('        改：正文用 `:::quiz 【自拟】标题` 或 `- **【自拟】** 题干…`，'
                           '附录用 `【课后习题 %s 第 N 题】`' % cur)
            cur, has = m.group(1), False
            continue
        for mm in _TASK_RE.finditer(s):
            label = mm.group(1)
            if not any(label.startswith(x) for x in _TASK_PREFIX):
                continue
            # 附录里显式标了 §X.Y 的题归那一节；其余归当前节
            sec = re.search(r'§\s*(\d+(?:\.\d+)+)', label)
            if sec:
                if sec.group(1) == cur:
                    has = True
                continue
            if cur:
                has = True
    if cur and not has:
        out.append(f'  ERROR L?: 第 {cur} 节没有任何题目（正文无【自拟】/【课后习题…】，'
                   f'附录也无 §{cur} 的题）——该节不出「对应题目」行，回跳也消失')
        out.append('        改：正文用 `:::quiz 【自拟】标题` 或 `- **【自拟】** 题干…`，'
                   '附录用 `【课后习题 %s 第 N 题】`' % cur)
    return out


IMG_RE = re.compile(r'^!\[(?P<cap>[^\]]*)\]\((?P<src>[^)]+)\)\s*'
                    r'(?:\{w=(?P<w>[\d.]+)\})?\s*$')
IMG_EXT = ('.png', '.jpg', '.jpeg')
IMG_DPI = 300.0
PAGE_W_CM = 16.6


def _png_px(path):
    """PNG 的 IHDR 像素宽；读不到返回 None（jpg 不解析）。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(33)
    except OSError:
        return None
    if head[:8] == b'\x89PNG\r\n\x1a\n' and head[12:16] == b'IHDR':
        return int.from_bytes(head[16:20], 'big') or None
    return None


def _check_images(lines, code_lines, md_path):
    """图片行：路径、图注、扩展名、宽度、登记。

    表格与公式不走这条通道（表格写成 md 表格、公式走 LaTeX），所以这里只管
    「本身就是图」的内容。图必须登记在 figures\\_图索引.md：主干换版后要靠那一行
    回查是哪一页裁的。
    """
    errs = []
    base = os.path.dirname(os.path.abspath(md_path))
    idx_path = os.path.join(base, 'figures', '_图索引.md')
    idx_txt = ''
    if os.path.isfile(idx_path):
        with open(idx_path, encoding='utf-8') as f:
            idx_txt = f.read()

    seen = {}
    for i, ln in enumerate(lines, 1):
        if i in code_lines:
            continue
        s = ln.strip()
        if not s.startswith('!['):
            continue
        m = IMG_RE.match(s)
        if not m:
            errs.append('  ERROR L%d: 图片行写法不合规，应为 ![图注](路径)，可带 {w=6.4}；'
                        '路径里不要出现括号' % i)
            continue
        cap, src, w = m.group('cap').strip(), m.group('src').strip(), m.group('w')
        if not cap:
            errs.append('  ERROR L%d: 图注为空——图下没有文字，读者与检索都认不出它是什么' % i)
        ext = os.path.splitext(src)[1].lower()
        if ext not in IMG_EXT:
            errs.append('  ERROR L%d: 只支持 png/jpg，当前是 %s' % (i, ext or '(无扩展名)'))
        full = src if os.path.isabs(src) else os.path.normpath(os.path.join(base, src))
        exists = os.path.isfile(full)
        if not exists:
            errs.append('  ERROR L%d: 图片不存在：%s（相对 md 解析为 %s）' % (i, src, full))
        if w:
            try:
                wv = float(w)
                if wv <= 0 or wv > PAGE_W_CM:
                    errs.append('  ERROR L%d: {w=%s} 超出 0–%.1fcm 版心' % (i, w, PAGE_W_CM))
            except ValueError:
                errs.append('  ERROR L%d: {w=%s} 不是数字' % (i, w))
        elif exists and ext == '.png':
            px = _png_px(full)
            if px:
                cm = px / IMG_DPI * 2.54
                if cm > PAGE_W_CM:
                    errs.append('  ERROR L%d: 不写 {w=} 时按 300dpi 折算宽 %.1fcm，'
                                '超出版心 %.1fcm；写 {w=…} 收窄' % (i, cm, PAGE_W_CM))
        if exists:
            name = os.path.basename(src)
            if not idx_txt:
                errs.append('  WARN  L%d: 有图但同目录没有 figures\\_图索引.md，'
                            '主干换版后无法回查这张图是从哪页裁的' % i)
            elif name not in idx_txt:
                errs.append('  WARN  L%d: %s 未登记在 _图索引.md' % (i, name))
        seen[src] = seen.get(src, 0) + 1
    for src, n in seen.items():
        if n > 1:
            errs.append('  WARN  : 同一张图引用了 %d 次（%s）；重复贴图多半是复制粘贴带过来的' % (n, src))
    return errs


def check_file(path):
    with open(path, encoding='utf-8') as f:
        text = f.read()
    lines = text.splitlines()
    code_lines, tree_lines = _code_blocks(lines)
    n_err = n_warn = 0
    out = [f'=== {path}']
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if i not in code_lines and s.startswith('```'):
            # 引擎只认 :::code。围栏行会被当普通段落，代码各行被 flush() 用空格拼成
            # 一整段流水文本，缩进/注释/换行全丢，而 lint 与 audit 都是绿的（11-14 实测）。
            n_err += 1
            out.append(f'  ERROR L{i}: ``` 围栏不被引擎识别（只认 :::code），'
                       f'围栏会变正文、代码被拼成一段流水文本')
            out.append(f'        | {s[:100]}')
            out.append('        改：:::code … :::（块内保留原缩进）')
            continue
        for level, msg, pat in RULES:
            if 'U+2080' in msg and i in code_lines:
                # 下标字符在 :::code 树块里是唯一能表达下标的纯文本手段（$…$ 在树里
                # 会堆叠、字体也不统一），实测等宽与正文字体都正常显示；这条只管正文。
                continue
            if '连写 ≥3 个 ─' in msg and i not in tree_lines:
                # 只对结构块生效：正文里的 ─ 是正常标点。
                continue
            if '长行内公式含 =' in msg and i in tree_lines:
                # 树块的缩进已经限定了换行，长公式不会被 Word 折成两行。
                continue
            if callable(pat):
                hit = pat(line)
            else:
                if pat.pattern.startswith(r'^\s*:::'):
                    continue  # 计数型规则单独跑
                hit = bool(pat.search(line))
            if hit:
                tag = 'ERROR' if level == 'ERROR' else 'WARN '
                if level == 'ERROR':
                    n_err += 1
                else:
                    n_warn += 1
                out.append(f'  {tag} L{i}: {msg}')
                out.append(f'        | {line.strip()[:100]}')
        if i in tree_lines:
            for msg in _tree_rule_hits(s):
                n_warn += 1
                out.append(f'  WARN  L{i}: {msg}')
                out.append(f'        | {s[:100]}')
    for e in _check_container_balance(text):
        n_err += 1
        out.append(e)
    for e in _check_container_kind(lines, code_lines) + _check_dollar_blocks(lines, code_lines):
        n_err += 1
        out.append(e)
    for e in _check_tree_depth(lines):
        if e.lstrip().startswith('ERROR'):
            n_err += 1
        out.append(e)
    for e in _check_external_src(lines):
        if e.lstrip().startswith('ERROR'):
            n_err += 1
        elif e.lstrip().startswith('WARN'):
            n_warn += 1
        out.append(e)
    for e in _check_engine_math(lines, code_lines) + _check_appendix(lines) + _check_section_tasks(lines):
        if e.lstrip().startswith('ERROR'):
            n_err += 1
        out.append(e)
    for e in _check_images(lines, code_lines, path):
        if e.lstrip().startswith('ERROR'):
            n_err += 1
        elif e.lstrip().startswith('WARN'):
            n_warn += 1
        out.append(e)
    out.append(f'  小结：ERROR {n_err} / WARN {n_warn}')
    return out, n_err, n_warn


def main():
    args = [a for a in sys.argv[1:] if a.strip()]
    if not args:
        print(__doc__)
        raise SystemExit(1)
    total_err = total_warn = 0
    for p in args:
        out, e, w = check_file(p)
        print('\n'.join(out))
        total_err += e
        total_warn += w
    print(f'\n总计：ERROR {total_err} / WARN {total_warn}')
    if total_err:
        print('存在 ERROR：先修 md 再构建。规则依据见 复习资料\\_生产\\_踩坑清单.md')
    raise SystemExit(2 if total_err else 0)


if __name__ == '__main__':
    main()
