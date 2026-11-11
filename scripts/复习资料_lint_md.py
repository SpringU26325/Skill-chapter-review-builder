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
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


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


def _frac_in_math(line):
    r"""结构块里会撑高单行的写法，只查 $…$ 内部：\frac/\dfrac、斜杠除法。
    不加 $ 的纯文本 a/b 不走公式转换，不报。"""
    for _, _, c in _inline_spans(line):
        if re.search(r'\\(d)?frac\b', c) or '/' in c:
            return True
    return False


TREE_LINE_RE = re.compile(r'^\s*[│├└]')


def _tree_rule_hits(s):
    """结构块（本章脉络）里会破坏行距节奏的写法。

    树块现在支持 $…$，公式进树后行高会随内容变化；堆叠分式与 \\limits
    都会把单行撑高，整块缩进节奏立刻不齐。
    """
    hits = []
    if _frac_in_math(s):
        hits.append(r'结构块 $…$ 里的分式会变堆叠、撑高该行：改成不加 $ 的纯文本 a/b，或移出树')
    if '\\limits' in s:
        hits.append(r'结构块里 \limits 会把上下限压到下方、撑高行距，去掉 \limits')
    return hits


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
    ('ERROR', r'行内 $…$ 里放 \begin{…}（矩阵/环境），提到块级公式',
     _begin_in_inline_math),
    ('ERROR', r'::: 容器未闭合或多余 :::',
     re.compile(r'^\s*:::\s*$')),  # 计数型规则，特殊处理见下
    ('WARN', r'\sqrt{单符号^{…}} 上标会跑到根号外，改写（如 \sqrt{x\cdot x}）',
     re.compile(r'\\sqrt\{[A-Za-z\\]\^\{')),
    ('WARN', r'一行里 ≥2 个行内 \dfrac，各占一份行高，提到块级或拆行',
     _multi_dfrac),
    ('WARN', r'\\text{…} 超 4 字，整句移到公式外',
     re.compile(r'\\text\{([^{}]{5,})\}')),
    ('WARN', r'U+2080 区下标字符出现在代码块/正文，写 v_1 或 v1',
     re.compile(r'[\u2080-\u209f\u2070-\u2079]')),
    ('WARN', r'长行内公式含 = 连等式，易被 Word 折断，拆短用文字连',
     _long_inline_eq),
]


def _check_container_balance(text):
    """::: 开头（带类型）与单独 ::: 结尾必须配对。返回错误列表。"""
    errs = []
    opens = closes = 0
    for i, line in enumerate(text.splitlines(), 1):
        s = line.strip()
        if re.match(r'^:::[a-zA-Z]', s):
            opens += 1
        elif re.match(r'^:::$', s) or s == ':::':
            closes += 1
    if opens != closes:
        errs.append(f'  ERROR L?: ::: 容器不配对（开头 {opens} / 结尾 {closes}）')
    return errs


def _tree_code_lines(lines):
    """按整块判定结构块，与引擎 _render_code 一致：块内任一行带树符即整块算树。
    返回需要跑树规则的行号集合（1-based）；根行不带树符，也要覆盖到。"""
    hit = set()
    block, in_code = [], False
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if s.startswith(':::code'):
            in_code, block = True, []
            continue
        if in_code and s == ':::':
            if any(TREE_LINE_RE.match(ln) for ln in block):
                hit.update(range(i - len(block), i))
            in_code = False
            continue
        if in_code:
            block.append(line)
    return hit


def check_file(path):
    with open(path, encoding='utf-8') as f:
        text = f.read()
    lines = text.splitlines()
    tree_lines = _tree_code_lines(lines)
    n_err = n_warn = 0
    out = [f'=== {path}']
    for i, line in enumerate(lines, 1):
        s = line.strip()
        for level, msg, pat in RULES:
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
        print('存在 ERROR：先修 md 再构建。规则依据见 复习资料\\_踩坑清单.md')
    raise SystemExit(2 if total_err else 0)


if __name__ == '__main__':
    main()
