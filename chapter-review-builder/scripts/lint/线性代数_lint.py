# -*- coding: utf-8 -*-
"""线性代数_lint.py — 线性代数记号与断言的可执行检查

设计原则：约定只存一处（notation.md），正文靠脚本兜底，不靠记忆。
每条规则对应一类"读者一眼看出、脚本也能看出"的错误。

用法：
    python 线性代数_lint.py 线性代数_第1章.md 线性代数_第2章上.md 线性代数_第2章下.md
退出码 0 = 全过；1 = 有 ERROR（不许交付）；WARN 不阻断但要在自检报告里说明。
"""
import re
import sys
import io
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))        # scripts/，便于 import _env
import _env

ERRORS, WARNS = [], []


def err(f, ln, msg):
    ERRORS.append(f'{f}:{ln}  ERROR  {msg}')


def warn(f, ln, msg):
    WARNS.append(f'{f}:{ln}  WARN   {msg}')


def notation_path():
    """记号表在工作区 `复习资料\\_源稿\\线性代数_notation.md`，不在包内。"""
    root = _env.find_root()
    if not root:
        raise SystemExit('定位不到复习资料工作区（当前目录及上级都没有「复习资料」目录）。请在复习资料工作区里执行。')
    p = os.path.join(_env.src_dir(root), '线性代数_notation.md')
    if not os.path.exists(p):
        raise SystemExit('找不到记号表：%s\n本 lint 依赖它，缺失时先补这份文件。' % p)
    return p


def load_notation():
    with open(notation_path(), encoding='utf-8') as fh:
        return fh.read()


def check_lines(fname, lines, notation):
    unregistered = set()
    for i, raw in enumerate(lines, 1):
        ln = raw.strip()
        if not ln:
            continue

        # R1 说"行向量"必须自证形状：1×n 或横排方括号
        if '行向量' in ln:
            ok = ('1\\times n' in ln) or ('1×n' in ln) or ('[\\,' in ln) or ('[\\ ' in ln)
            if not ok:
                err(fname, i, '出现"行向量"但没给形状（缺 1\\times n 或横排方括号）')

        # R2 圆括号元组被直接断言成行向量（"$(..)$ 是 … 行向量"）
        for m in re.finditer(r'\$\([^)]*\)\$[^。；]{0,14}?行向量', ln):
            if '列向量' in m.group(0):
                continue
            err(fname, i, '圆括号元组被断言成行向量（圆括号=列向量，行向量用横排方括号）')

        # R3 粗体符号必须在符号表登记。
        # 判据必须比对 `\mathbf{name}` 原文 —— 早期写成 `name not in notation`，
        # 单字母名在表里几乎必然出现（任意位置），实测 25 个名字报 0 个，等于没查。
        for m in re.finditer(r'\\mathbf\{([^}]+)\}', ln):
            name = m.group(1)
            if f'\\mathbf{{{name}}}' not in notation:
                unregistered.add(name)

        # R4 无下标写法：Eij / lij / Pij
        for m in re.finditer(r'(?<![\\A-Za-z])([ElP])(ij|ji|kl)\b', ln):
            err(fname, i, f'"{m.group(0)}" 缺下标，应写 {m.group(1)}_{{{m.group(2)}}}')

        # R5 代码块里画矩阵（用分号或方括号堆数字）
        if re.search(r'\[\s*[\d\-]+\s+[\d\-]+', ln) and not ln.startswith('$$'):
            warn(fname, i, '疑似在正文/代码块里用字符画矩阵，应改用 bmatrix')

    if unregistered:
        # 聚合报一条：逐行报警会把整份文档刷满，反而没人看
        warn(fname, 0,
             '这些 \\mathbf{…} 没在 notation.md 登记：' + '、'.join(sorted(unregistered))
             + '（新增符号先加表再用）')


def check_structure(fname, text):
    # 每份文档必须有且仅有一个「记号约定」框
    n = text.count('记号约定')
    if n == 0:
        warn(fname, 0, '正文未写记号框（构建时从 线性代数_notation.md 注入，需在 docx 里复核）')
    elif n > 1:
        warn(fname, 0, f'「记号约定」出现 {n} 次，可能重复注入')


def main(paths):
    notation = load_notation()
    for p in paths:
        with open(p, encoding='utf-8') as fh:
            text = fh.read()
        check_lines(os.path.basename(p), text.split('\n'), notation)
        check_structure(os.path.basename(p), text)

    for line in ERRORS:
        print(line)
    for line in WARNS:
        print(line)
    print(f'\n合计 ERROR {len(ERRORS)} / WARN {len(WARNS)}')
    return 1 if ERRORS else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
