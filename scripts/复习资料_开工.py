# -*- coding: utf-8 -*-
"""开工简报：一条命令替代「读配置 + 读台账 + 扫踩坑 + 数缺陷」。

为什么要有这个脚本
    台账 75 KB、踩坑 44 KB、缺陷 12 KB，模型接手时整读一遍约 6 万 token；
    这些内容还会被后续每一轮对话重发。这里只输出决策真正要用的几百字：
    主干是谁、做到哪、下一章、哪些坑、几条未修。
    要细节时再按下面的行号去 grep 或 Read 局部，不要整读原文件。

用法（在复习资料工作区根目录执行，脚本自动上溯定位）：
    "<py>" "<skill>\\scripts\\复习资料_开工.py"             # 人读，默认 ≤900 字
    "<py>" "<skill>\\scripts\\复习资料_开工.py" 线性代数     # 只看某一科
    "<py>" "<skill>\\scripts\\复习资料_开工.py" --json       # 结构化，供别的脚本消费
    "<py>" "<skill>\\scripts\\复习资料_开工.py" --full       # 附进度表全部行

退出码：0 正常；2 缺 `_台账.md` / `_配置.md`（先跑 `复习资料_init.py`）。
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env  # noqa: E402


# ---------- 读取工具：只取需要的片段，不整读 ----------

def _read(path):
    try:
        with open(path, encoding='utf-8', newline='') as f:
            return f.read().replace('\r\n', '\n')
    except OSError:
        return ''


def _section(text, *titles):
    """取 `## <标题>` 到下一个 `## ` 之间的正文（不含标题行）。

    `titles` 是别名组，顺序即优先级。**先全组精确等同，再全组前缀匹配**——
    否则「进度」会先前缀命中旧台账的「## 进度摘要」而漏掉真正的「## 已完成明细」。
    """
    lines = text.split('\n')
    heads = [(i, s[3:].strip()) for i, s in enumerate(lines) if s.startswith('## ')]

    def grab(k):
        start = heads[k][0] + 1
        end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
        return lines[start:end]

    for want in titles:
        for k, (_, name) in enumerate(heads):
            if name == want:
                return grab(k)
    for want in titles:
        for k, (_, name) in enumerate(heads):
            if name.startswith(want):
                return grab(k)
    return []


def _col_map(head):
    """表头 → 列下标。按列名取，列序变了也不会读错位。"""
    return {h: i for i, h in enumerate(head)}


def _cell(row, mp, *names):
    """按候选列名取一格；缺列返回空串，不抛。"""
    for n in names:
        i = mp.get(n)
        if i is not None and i < len(row):
            return row[i]
    return ''


def _bullets(lines, limit=8):
    """片段里的一级 bullet，压缩成短句。"""
    out = []
    for s in lines:
        t = s.strip()
        if t.startswith('- ') or t.startswith('* '):
            out.append(t[2:].strip())
        if len(out) >= limit:
            break
    return out


def _table(lines):
    """解析 Markdown 表格，返回 (表头, 数据行列表)。"""
    rows, head = [], None
    for s in lines:
        t = s.strip()
        if not t.startswith('|'):
            continue
        cells = [c.strip() for c in t.strip('|').split('|')]
        if head is None:
            head = cells
        elif set(''.join(cells)) <= set('-: '):
            continue
        else:
            rows.append(cells)
    return head or [], rows


# ---------- 各数据源 ----------

def read_config(prod):
    """主干表：科目 / 主干文件 / 位置 / 类型 / 编号单位 / 页码换算。"""
    text = _read(os.path.join(prod, '_配置.md'))
    head, rows = _table(_section(text, '主干'))
    mp = _col_map(head)
    out = []
    for r in rows:
        subj = _cell(r, mp, '科目')
        if not subj or subj == '科目':
            continue
        out.append({'科目': subj,
                    '主干': _cell(r, mp, '主干文件', '主干').strip('`'),
                    '位置': _cell(r, mp, '位置').strip('`'),
                    '类型': _cell(r, mp, '类型'),
                    '编号单位': _cell(r, mp, '编号单位'),
                    '页码换算': _cell(r, mp, '页码换算')})
    return out


def read_ledger(prod):
    """进度表 + 未做与下一章 / 待办 / 最近交付。

    2.9.0 前的旧台账没有「## 进度」，只有「## 进度摘要（叙述）」与
    「## 已完成明细（六列，列名与语义都不同）」。这里按列名适配旧表
    （`章节 / 册`、`原文页码范围`、`成品页数`），并把 legacy 报给上层——
    宁可让模型看见「这是旧格式」，也不要静默给出空表。
    """
    text = _read(os.path.join(prod, '_台账.md'))
    body = _section(text, '进度', '已完成明细')
    legacy = not any(s.startswith('## 进度') and not s.startswith('## 进度摘要')
                     for s in text.split('\n'))
    head, rows = _table(body)
    mp = _col_map(head)
    chapters = []
    for r in rows:
        subj = _cell(r, mp, '科目')
        if not subj or subj == '科目':
            continue
        pages = _cell(r, mp, '成品页', '成品页数')
        chapters.append({'科目': subj,
                         '章节': _cell(r, mp, '章节', '章节 / 册'),
                         '页码': _cell(r, mp, '页码', '原文页码范围'),
                         '分册': _cell(r, mp, '分册'),
                         '册数': max(len([x for x in pages.split('/') if x.strip()]), 1)})
    m = re.search(r'\*\*最近交付\*\*[：:]\s*(.+)', text)
    return {
        'chapters': chapters,
        'legacy': legacy,
        '最近': m.group(1).strip() if m else '',
        '未做': _bullets(_section(text, '未做')),
        '待办': _bullets(_section(text, '待办', '已知待办')),
    }


def read_pitfalls(prod):
    """只取节名 + 起止行 + 字数，供按需局部读。"""
    path = os.path.join(prod, '_踩坑清单.md')
    text = _read(path)
    lines = text.split('\n')
    idx = [(i + 1, s[3:].strip()) for i, s in enumerate(lines)
           if s.startswith('## ') and not s.startswith('### ')]
    out = []
    for k, (ln, name) in enumerate(idx):
        end = idx[k + 1][0] - 1 if k + 1 < len(idx) else len(lines)
        chars = sum(len(x) for x in lines[ln:end])
        out.append({'节': name, '起': ln, '止': end, '字': chars})
    return out


def read_defects(prod):
    """从头部「**状态：N 条未修**（P0-4、…）」抽未修清单。

    措辞要放宽：`状态：` 常被加粗，`条未修` 与括号之间会夹 `**`
    （`> **状态：4 条未修**（P0-4…）`）。严格贴括号的正则会一条都读不到，
    而这个数字正是最该保住的信息——宁可放宽，读到 0 也要当可疑。
    """
    text = _read(os.path.join(prod, '_待修缺陷.md'))
    head = text.split('\n## ')[0][:2000]     # 只取首节，不整读
    m = re.search(r'(\d+)\s*条未修', head)
    if not m:
        return {'数量': 0, '条目': ''}
    b = re.search(r'条未修[*\s]*[（(]([^）)]*)[）)]', head)
    items = [x.strip() for x in b.group(1).split('、') if x.strip()] if b else []
    return {'数量': int(m.group(1)), '条目': '、'.join(items)}


def read_pending(prod):
    """配置缺失时给一句准话。"""
    # 只把「没有就跑不动」的三份算缺；历史与缺陷文件是可选的
    return {n: os.path.isfile(os.path.join(prod, n))
            for n in ('_配置.md', '_台账.md', '_踩坑清单.md')}


# ---------- 输出 ----------

def _subj_match(subj, name):
    """科目名对得上：子串或字集包含。「线代」靠字集命中「线性代数」。"""
    if not subj:
        return True
    if subj in name or name in subj:
        return True
    return set(subj) <= set(name) or set(name) <= set(subj)


def _short(p, root, width=52):
    p = p.replace('/', '\\')
    if root and p.lower().startswith(root.replace('/', '\\').lower()):
        p = p[len(root) + 1:]
    return p if len(p) <= width else '…' + p[-width + 1:]


def _name(p, width=30):
    """主干名只留书一部分：去掉作者/译者括号与 .pdf，超长再截。"""
    n = os.path.basename(p.replace('/', '\\'))
    for sep in (' (', '（'):
        if sep in n:
            n = n.split(sep)[0].strip()
            break
    n = re.sub(r'\s*\.pdf$', '', n, flags=re.I)
    return n if len(n) <= width else n[:width - 1] + '…'


def render(data, root, full=False):
    out = []
    cfg, led, pits, dfc = data['config'], data['ledger'], data['pitfalls'], data['defects']
    out.append('=== 复习资料 · 开工简报 ===')
    out.append(f'工作区 {root}｜解释器 {_short(data["python"], root, 40)}')
    if data['missing']:
        out.append(f'[!] 缺文件：{" / ".join(data["missing"])}——先跑 复习资料_init.py')

    out.append('')
    out.append('【主干】以 _配置.md 为准，用户没改口就不换')
    for c in cfg:
        conv = c['页码换算'].split('（')[0].strip()
        extra = '' if conv in ('—', '-', '') else '  ' + conv
        out.append(f'  {c["科目"]}：{c["类型"]}｜{_name(c["主干"])}{extra}')

    out.append('')
    out.append('【进度】章数 / 册数（册数按成品页列拆 /）')
    if led['legacy']:
        out.append('  [!] 台账是 2.9.0 之前的格式（无「## 进度」节），下面按旧表列名尽力读——'
                   '可能不准。按 RELEASE-2.9.0.md 的迁移段整理一次再接手。')
    by = {}
    for c in led['chapters']:
        g = by.setdefault(c['科目'], {'章': set(), '册': 0})
        g['章'].add(re.sub(r'^(.*?)\s.*', r'\1', c['章节']))
        g['册'] += c['册数']
    if by:
        out.append('  ' + '｜'.join(f'{s} {len(g["章"])} 章 {g["册"]} 册' for s, g in by.items())
                   + (f'；最近 {led["最近"][:34]}' if led['最近'] else ''))
    else:
        out.append('  暂无已交付章节。')
    if full:
        for c in led['chapters']:
            out.append(f'   {c["科目"]} {c["章节"]} p{c["页码"]} {c["分册"]}')

    for key, title in (('未做', '【未做与下一章】'), ('待办', '【待办】')):
        if led[key]:
            out.append('')
            out.append(title + '（全文见 _台账.md）')
            for b in led[key][:5]:
                out.append(f'  · {b[:80]}')

    if dfc['数量']:
        out.append(f'\n【未修引擎缺陷】{dfc["数量"]} 条（{dfc["条目"]}）——撞上了按 _待修缺陷.md 的症状索引查')

    if pits:
        out.append('')
        out.append('【踩坑】选一节按行号局部读，不要整份读')
        for p in pits:
            if p['节'].startswith('目录'):
                continue
            size = f'{p["字"] // 1000}k' if p['字'] >= 1000 else f'{p["字"]}字'
            out.append(f'  {p["节"][:30]}  L{p["起"]}-{p["止"]} {size}')

    out.append('')
    out.append('【下一步】章节与主干未指定先问用户；进度只写 _台账.md。')
    return '\n'.join(out)


def main():
    args = [a for a in sys.argv[1:] if a.strip()]
    as_json = '--json' in args
    full = '--full' in args
    subj = next((a for a in args if not a.startswith('-')), None)

    root = _env.find_root()
    if not root:
        print('找不到工作区：cwd 往上 8 层内没有名为「复习资料」的目录。'
              'cd 到工作区根，或设 CRB_ROOT。')
        raise SystemExit(2)
    prod = _env.prod_dir(root)

    missing = [n for n, ok in read_pending(prod).items() if not ok]
    if missing and ('_台账.md' in missing or '_配置.md' in missing):
        print(f'缺 {", ".join(missing)}：先跑 `复习资料_init.py` 建骨架，不要猜。')
        raise SystemExit(2)

    py = _env.venv_python(root) or '未找到（先问用户，不要自建 .venv）'
    data = {
        'config': read_config(prod),
        'ledger': read_ledger(prod),
        'pitfalls': read_pitfalls(prod),
        'defects': read_defects(prod),
        'python': py,
        'missing': missing,
    }
    if subj:
        data['config'] = [c for c in data['config'] if _subj_match(subj, c['科目'])]
        data['ledger']['chapters'] = [c for c in data['ledger']['chapters']
                                      if _subj_match(subj, c['科目'])]
        data['ledger']['最近'] = ''      # 最近交付是全量口径，按科目看时不显示，免得错认

    if as_json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        print(render(data, root, full=full))


if __name__ == '__main__':
    main()
