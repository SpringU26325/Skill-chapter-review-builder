# -*- coding: utf-8 -*-
"""建骨架：只建约定目录、台账、配置与踩坑清单，不碰素材、不建 .venv、不装库。

用法：
    "<py>" "<skill>\\scripts\\复习资料_init.py" [科目 ...] [--root <工作区根>] [--dry]

默认以当前目录为工作区根；若当前目录在既有复习资料工作区内，自动用那个根。
科目名可有可无：给了就顺手建 `复习资料\\<科目>\\` 供用户放原件，没给就只建生产骨架。

**也用于补建配置**：已有一堆源稿/成品、却没 `_生产\\_配置.md` 的老工作区同样跑这条命令——
新建配置时会扫描既有源稿的 `@sub`，把教材版本与页码换算预填成候选行，供用户确认。

幂等：台账、配置、踩坑清单已存在时一律保留，绝不覆盖；`--dry` 只打印将要建什么。
"""
import collections
import io
import os
import re
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


# ---- 既有源稿扫描：从 @sub 推断主干候选 ----
# 为什么需要：老工作区（已有一批源稿与成品）常常没有 _配置.md，而 init 以前
# 只在「没有 复习资料\」时才会被调用，配置就一直补不上，每个新会话都从
# 「请指定主干」重来。这里让 init 顺手把 @sub 里的教材版本与页码换算读出来，
# 预填成候选行——教材名可信（来自已交付章节），文件名/位置仍由用户确认。
_RANGE = re.compile(
    r'(物理页|印刷页?|原书)[^0-9（）()]{0,8}(?:p|p\.)?\s*(\d+)\s*[–\-—]\s*(\d+)')


def _sub_head(path):
    """读一份源稿 md 的 @sub 行；约定写在文件头 60 行内，没有就返回 None。"""
    try:
        with io.open(path, encoding='utf-8') as f:
            for i, ln in enumerate(f):
                if i > 60:
                    break
                s = ln.strip()
                if s.startswith('@sub'):
                    return s[4:].strip()
    except (OSError, UnicodeDecodeError):
        return None      # 单份坏文件不该毁掉整次扫描，跳过即可
    return None


def _infer_primary(src_dir):
    """按科目聚合 @sub，返回 [(科目, 教材, 编号单位, 页码换算, 册数)]。

    分两遍：先逐册抽成记录，再按科目取「出现最多的书名与编号单位」——首册常把
    课件一起写进 @sub（`谭浩强《C 程序设计》第五版 p1–14（…）· 第六版课件 p3–23`），
    取众数才不会让第一册的写法污染整科。页码换算只在同一册里同时出现「物理页」
    与「印刷页/原书」时才算得出；多册不一致就写「各册不一致」，不猜。
    """
    if not os.path.isdir(src_dir):
        return []
    records = []
    for name in sorted(os.listdir(src_dir)):
        if not name.lower().endswith('.md'):
            continue
        head = _sub_head(os.path.join(src_dir, name))
        if not head:
            continue
        # 分隔符是空格包着的「·」：人名里的「李忠·周建莹」不带空格，不会被切错
        segs = [s.strip() for s in head.split(' · ') if s.strip()]
        if not segs:
            continue
        subject = re.split(r'[（(]', segs[0], maxsplit=1)[0].strip()
        if not subject:
            continue
        # 第一段自带括号版本（如「高等数学（李忠·周建莹 第三版 上册）」）时它本身就是书名；
        # 否则第一段只是科目名，书名在第二段
        book = segs[0] if re.search(r'[（(]', segs[0]) else (segs[1] if len(segs) >= 2 else segs[0])
        # 页码整行扫，但先剔掉课件/辅导段——「《学习辅导》第 5 章物理页 47–63」不是主干页码
        page_text = ' · '.join(s for s in segs if not re.search(r'课件|辅导|习题', s))
        rng = {}
        for m in _RANGE.finditer(page_text):
            key = '印刷页' if m.group(1).startswith('印刷') else m.group(1)
            rng.setdefault(key, (int(m.group(2)), int(m.group(3))))
        unit = ('幻灯片号 pNN' if '幻灯片' in page_text
                else ('原书物理页' if rng else '—'))
        off = None
        base = rng.get('印刷页') or rng.get('原书')
        if '物理页' in rng and base:
            diff = rng['物理页'][0] - base[0]
            off = diff if diff > 0 else None   # 差值为负说明写法反常，宁可留待确认
        records.append((subject, book, unit, off))

    rows = []
    for subject in dict.fromkeys(r[0] for r in records):   # 保序去重
        rs = [r for r in records if r[0] == subject]
        book = collections.Counter(r[1] for r in rs).most_common(1)[0][0]
        units = [r[2] for r in rs if r[2] != '—']
        unit = collections.Counter(units).most_common(1)[0][0] if units else '—'
        offs = sorted({r[3] for r in rs if r[3] is not None})
        if not offs:
            conv = '—（待确认）'
        elif len(offs) == 1:
            conv = '物理页 = 印刷页 + %d（推断，待确认）' % offs[0]
        else:
            conv = '各册不一致（%s），待确认' % '、'.join('+%d' % o for o in offs)
        rows.append((subject, book, unit, conv, len(rs)))
    return rows


def _fill_primary(text, rows):
    """把候选行插进配置模板的主干表，返回 (新文本, 插入行数)。"""
    note = ('> 以下行由 `init` 扫描既有源稿的 `@sub` 预填：科目与教材版本来自已交付章节，'
            '**文件名、位置与换算请核对后补齐**，拿不准的留「待确认」。')
    eol = '\r\n' if '\r\n' in text else '\n'   # 模板是 CRLF，别把行尾混成两种
    lines = text.split(eol)
    idx = None
    for i, ln in enumerate(lines):
        if ln.strip().startswith('|---'):
            idx = i
            break
    if idx is None:
        return text, 0
    payload = ['| %s | %s | — | 教材 | %s | %s |' % (s, b, u, c) for s, b, u, c, _ in rows]
    payload += ['', note]
    lines[idx + 1:idx + 1] = payload
    return eol.join(lines), len(rows)


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
    # 台账、配置、踩坑清单成套生成：缺一份就只剩半套骨架（配置说现状、台账说过往、
    # 踩坑清单装着「改引擎前必读」的既成事实），所以并进一张种子表，省掉三份模板各写一遍拷贝逻辑。
    cfg = os.path.join(prod, '_配置.md')
    pitfall = os.path.join(prod, '_踩坑清单.md')
    seeds = [
        (ledger, template),
        (cfg, os.path.join(SKILL_DIR, 'assets', '配置_template.md')),
        # SKILL.md 把踩坑清单写成「改引擎、写公式、调 COM 前必读」，sources.md 又让人
        # 「见踩坑清单顶部」找 WPS 禁令——新工作区不生成这个文件，这两处就是指向不存在的路。
        (pitfall, os.path.join(SKILL_DIR, 'assets', '踩坑清单_template.md')),
    ]
    # 只在「配置还不存在」时扫描：已有配置说明主干定过了，不能拿推断值去覆盖
    rows = [] if os.path.isfile(cfg) else _infer_primary(os.path.join(prod, '_源稿'))

    print('工作区根：%s' % root)
    for d in dirs:
        print('  [%s] %s' % ('有' if os.path.isdir(d) else '建', os.path.relpath(d, root)))
    for dst, tpl in seeds:
        if os.path.isfile(dst):
            print('  [留] %s（已存在，不覆盖）' % os.path.relpath(dst, root))
        else:
            tail = '，并从既有源稿预填 %d 科主干候选' % len(rows) if dst == cfg and rows else ''
            print('  [建] %s（取自 assets\\%s%s）'
                  % (os.path.relpath(dst, root), os.path.basename(tpl), tail))
    if rows:
        print('\n既有源稿里读到的主干候选：')
        for subject, book, unit, conv, n in rows:
            print('  %s ← %s（%d 册；%s；%s）' % (subject, book, n, unit, conv))

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
        if dst == cfg and rows:
            # 台账原样拷贝；只有配置要加工——把候选填进主干表，省掉用户从零填一遍
            text, n_fill = _fill_primary(data.decode('utf-8'), rows)
            with io.open(dst, 'w', encoding='utf-8', newline='') as f:
                f.write(text)
            print('  [填] %s：预填 %d 科主干候选（待用户确认）'
                  % (os.path.relpath(dst, root), n_fill))
            continue
        with io.open(dst, 'wb') as f:
            f.write(data)

    print('\n骨架就绪。下一步：')
    if rows:
        print('  1) 「主干」表已按既有源稿预填 %d 科：先把候选念给用户确认（回复「按现有配置继续」即可），'
              % len(rows))
        print('     再补文件名、位置与换算；用户否认就按第 2 条重定。')
    else:
        print('  1) 请用户指定主干资料（教材 / 课件 / 讲义 / 题库）。没指定就先问；没有就不开工。')
    print('  2) 主干与补充统一放 复习资料\\<科目>\\ 根层（能被 原件.py 盘点）；用户已有书库')
    print('     （如 Ebooks\\）可原地只读取用，把位置记进 _生产\\_配置.md 的素材库。')
    print('  3) 主干定下后写进 _生产\\_配置.md 的「主干」表：文件、位置、类型、编号单位、页码换算。')
    print('     台账只记进度，不重复这些字段；材料在系统别处时位置写绝对路径，不改工作区外的原件。')
    print('  4) 登记原件："<py>" "<skill>\\scripts\\复习资料_原件.py" <科目> --write')
    print('  5) md 写进 复习资料\\_生产\\_源稿\\<科目>_第X章.md（骨架见 assets\\chapter_template.md），')
    print('     跑 lint_md，再跑 复习资料_重建.py。')
    print('工作区没有 .venv 时先问用户用哪个解释器：本脚本不建虚拟环境、不装任何库。')


if __name__ == '__main__':
    main()
