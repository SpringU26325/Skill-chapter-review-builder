# -*- coding: utf-8 -*-
"""空工作区建骨架：只建约定目录、台账与配置，不碰素材、不建 .venv、不装库。

用法：
    "<py>" "<skill>\\scripts\\复习资料_init.py" [科目 ...] [--root <工作区根>] [--dry]

默认以当前目录为工作区根；若当前目录在既有复习资料工作区内，自动用那个根。
科目名可有可无：给了就顺手建 `复习资料\\<科目>\\` 供用户放原件，没给就只建生产骨架。

幂等：台账、配置已存在时一律保留，绝不覆盖；`--dry` 只打印将要建什么。
"""
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import _env


def parse_args(argv):
    dry = '--dry' in argv
    argv = [a for a in argv if a != '--dry']
    root = None
    subs = []
    i = 0
    while i < len(argv):
        if argv[i] == '--root':
            if i + 1 >= len(argv):
                raise SystemExit('--root 需要一个参数')
            root = os.path.abspath(argv[i + 1])
            i += 2
        else:
            subs.append(argv[i].strip())
            i += 1
    for s in subs:
        if not s or s in ('.', '..') or os.path.isabs(s) or '/' in s or '\\' in s:
            raise SystemExit('科目名只能是一层目录名：%r' % s)
    return root, subs, dry


def main():
    root_arg, subs, dry = parse_args(sys.argv[1:])
    if root_arg:
        if not os.path.isdir(root_arg):
            raise SystemExit('指定的工作区根不存在：%s' % root_arg)
        root = root_arg
    else:
        root = _env.find_root() or os.getcwd()

    ws = os.path.join(root, _env.WORKSPACE_MARK)
    prod = os.path.join(ws, _env.PROD_DIR)
    dirs = [os.path.join(prod, x) for x in ('_源稿', '_locks', '_render')]
    dirs += [os.path.join(ws, s) for s in subs]
    ledger = os.path.join(prod, '_台账.md')
    template = os.path.join(SKILL_DIR, 'assets', '台账_template.md')
    # 台账与配置成对生成：缺一份就只剩半套骨架（配置说现状、台账说过往），
    # 所以并进一张种子表，省掉两份模板各写一遍拷贝逻辑。
    seeds = [
        (ledger, template),
        (os.path.join(prod, '_配置.md'), os.path.join(SKILL_DIR, 'assets', '配置_template.md')),
    ]

    print('工作区根：%s' % root)
    for d in dirs:
        print('  [%s] %s' % ('有' if os.path.isdir(d) else '建', os.path.relpath(d, root)))
    for dst, tpl in seeds:
        if os.path.isfile(dst):
            print('  [留] %s（已存在，不覆盖）' % os.path.relpath(dst, root))
        else:
            print('  [建] %s（取自 assets\\%s）' % (os.path.relpath(dst, root), os.path.basename(tpl)))

    if dry:
        print('\n--dry：只预演，未写任何文件。')
        return

    for d in dirs:
        os.makedirs(d, exist_ok=True)
    for dst, tpl in seeds:
        if os.path.isfile(dst):
            continue  # 幂等：既有台账/配置绝不用模板覆盖，用户可能已经写过内容
        if not os.path.isfile(tpl):
            raise SystemExit('模板缺失：%s' % tpl)
        with io.open(tpl, 'rb') as f:
            data = f.read()
        with io.open(dst, 'wb') as f:
            f.write(data)

    print('\n骨架就绪。下一步：')
    print('  1) 请用户指定主干资料（教材 / 课件 / 讲义 / 题库）。没指定就先问；没有就不开工。')
    print('  2) 主干与补充统一放 复习资料\\<科目>\\ 根层（能被 原件.py 盘点）；用户已有书库')
    print('     （如按学期分的 Ebooks\\）可原地只读取用，把位置记进 _生产\\_配置.md 的素材库。')
    print('  3) 主干定下后写进 _生产\\_配置.md 的「主干」表：文件、位置、类型、编号单位、页码换算。')
    print('     台账只记进度，不重复这些字段；材料在系统别处时位置写绝对路径，不改工作区外的原件。')
    print('  4) 登记原件："<py>" "<skill>\\scripts\\复习资料_原件.py" <科目> --write')
    print('  5) md 写进 复习资料\\_生产\\_源稿\\<科目>_第X章.md（骨架见 assets\\chapter_template.md），')
    print('     跑 lint_md，再跑 复习资料_重建.py。')
    print('工作区没有 .venv 时先问用户用哪个解释器：本脚本不建虚拟环境、不装任何库。')


if __name__ == '__main__':
    main()
