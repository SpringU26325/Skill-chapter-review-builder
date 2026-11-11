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
    ('WARN', r'一行内多个行内 \dfrac 会被压成上标大小，答案分式改斜杠 A/B',
     re.compile(r'^[^\\]*\$[^$\n]*\\dfrac')),
    ('WARN', r'\\text{…} 超 4 字，整句移到公式外',
     re.compile(r'\\text\{([^{}]{5,})\}')),
    ('WARN', r'U+2080 区下标字符出现在代码块/正文，写 v_1 或 v1',
     re.compile(r'[\u2080-\u209f\u2070-\u2079]')),
    ('WARN', r'长行内公式含 = 连等式，易被 Word 折断，拆短用文字连',
     re.compile(r'\$[^$\n]{80,}\$')),
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


def check_file(path):
    with open(path, encoding='utf-8') as f:
        text = f.read()
    lines = text.splitlines()
    n_err = n_warn = 0
    out = [f'=== {path}']
    for i, line in enumerate(lines, 1):
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
