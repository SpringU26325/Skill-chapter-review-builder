# -*- coding: utf-8 -*-
"""包内唯一的机器相关逻辑：XSL 定位、工作区根、锁目录、渲染工作目录。

设计约束（改这个文件前先读）：
  1. 包里不出现任何宿主目录名、不出现本机绝对路径；
  2. 所有路径由「自身位置」或「工作区结构」推导，必要时用环境变量覆盖；
  3. 找不到就明确报错，给出可执行的定位方法，绝不静默降级。
"""
import os
import glob

# scripts/ 的上一级 = 包根
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL_DIR, 'scripts')
ASSETS = os.path.join(SKILL_DIR, 'assets')

# 工作区标志目录：向上找它来定位根
WORKSPACE_MARK = '复习资料'

# 生产目录：源稿、渲染过程、锁、迁移备份与三份台账全部收在 复习资料\_生产\ 下，
# 复习资料\ 根目录只留科目文件夹（科目根层放原件，成品进 _成品\）。
PROD_DIR = '_生产'
OUT_SUB = '_成品'

# 渲染工作目录是否允许中文路径。
# 2026-11-11 P4-b 实测：源与目标都含中文时 ExportAsFixedFormat 正常（10.3s / 8 页）。
# 历史上认定的「中文路径导出静默挂住」不成立，真因是 render.ps1 在创建 Word 实例之后
# 才读取清单，导致导出阻塞；清单改为在 New-Object 之前读取并归一化后即恢复正常。
# 因此默认允许中文路径：渲染中间件留在工作区 复习资料\_生产\_render，不向工作区外写文件。
ALLOW_CJK_RENDER_DIR = True


def _is_ascii(s):
    try:
        s.encode('ascii')
        return True
    except UnicodeEncodeError:
        return False


# ---------- XSL ----------

def xsl_candidates():
    """按优先级产出候选 XSL 路径：环境变量 → 包内 assets → Office 安装目录。"""
    env = os.environ.get('MML2OMML_XSL')
    if env:
        yield env
    yield os.path.join(ASSETS, 'MML2OMML.XSL')
    roots = [os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)'),
             r'C:\Program Files', r'C:\Program Files (x86)']
    seen = set()
    for root in roots:
        if not root or root in seen:
            continue
        seen.add(root)
        for pat in ('Microsoft Office/root/Office*/MML2OMML.XSL',
                    'Microsoft Office/Office*/MML2OMML.XSL'):
            for p in sorted(glob.glob(os.path.join(root, pat))):
                yield p


def find_xsl():
    """返回第一个存在的 XSL 路径；都不存在则报错并给出三种解法。"""
    for p in xsl_candidates():
        if p and os.path.exists(p):
            return p
    raise RuntimeError(
        '找不到 MML2OMML.XSL（公式转换必需）。三种解法任选其一：\n'
        '  1) 把 MML2OMML.XSL 放进包的 assets\\ 目录；\n'
        '  2) 设环境变量 MML2OMML_XSL=<xls 绝对路径>；\n'
        '  3) 确认本机已安装 Microsoft Office（XML 通常在 '
        'Program Files\\Microsoft Office\\root\\OfficeNN\\ 下）。')


# ---------- 工作区根 ----------

def find_root(start=None, max_up=8):
    """从 start（默认当前目录）向上找含「复习资料」目录的祖先，找不到返回 None。"""
    env = os.environ.get('CRB_ROOT')
    if env and os.path.isdir(env):
        return os.path.abspath(env)
    p = os.path.abspath(start or os.getcwd())
    for _ in range(max_up):
        if os.path.isdir(os.path.join(p, WORKSPACE_MARK)):
            return p
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return None


def require_root(start=None):
    root = find_root(start)
    if not root:
        raise SystemExit(
            '定位不到复习资料工作区：当前目录及其上级都没有「复习资料」目录。\n'
            '若这是空工作区，先建骨架： "<py>" "<skill>\\scripts\\复习资料_init.py" [科目]\n'
            '（init 只建目录、台账与配置，不建 .venv、不装库；可用 --dry 预演。）\n'
            '若资料在别处，设环境变量 CRB_ROOT=<工作区根>（所有脚本都认），或 cd 到工作区根再执行。\n'
            '注意 --root 参数只有 复习资料_重建.py 认，其余脚本会把它当文件名（P2-1）。')
    return root


def prod_dir(root):
    """生产目录：复习资料\\_生产\\。跨科目的工程文件都在这里。"""
    return os.path.join(root, WORKSPACE_MARK, PROD_DIR)


def src_dir(root):
    return os.path.join(prod_dir(root), '_源稿')


def out_dir(root, subj):
    """成品目录：复习资料\\<科目>\\_成品\\。科目根层留给原件。"""
    return os.path.join(root, WORKSPACE_MARK, subj, OUT_SUB)


def lock_dir(root=None, create=True):
    """锁目录：工作区内 复习资料\\_生产\\_locks；定位不到工作区时退到系统临时目录。"""
    if root:
        d = os.path.join(prod_dir(root), '_locks')
    else:
        d = os.path.join(os.environ.get('TEMP') or os.environ.get('TMP') or '.', 'crb_locks')
    if create:
        os.makedirs(d, exist_ok=True)
    return d


# ---------- 渲染工作目录 ----------

def render_candidates(root):
    """渲染中间件候选目录（存放 pN.docx / pN.pdf / PNG，都是过程文件）。"""
    env = os.environ.get('CRB_RENDER_DIR')
    if env:
        yield os.path.abspath(env)
    if root:
        in_ws = os.path.join(prod_dir(root), '_render')
        if ALLOW_CJK_RENDER_DIR or _is_ascii(in_ws):
            yield in_ws
    for base in (os.environ.get('TEMP'), os.environ.get('TMP')):
        if base:
            yield os.path.join(base, 'crb_render')
    # 不再兜底写盘符根目录。历史版本在最后一档 yield 了一个写死的根路径，导致工作区
    # 可写失败时把数百 MB 过程文件倒进盘符根——外部可见、换机器可能无权限，且实例
    # 挂死后那里的脚本会被反复拉起。宁可报错，见下方 raise。


def render_dir(root, create=True):
    """取第一个可写候选；都不行则报错要求设 CRB_RENDER_DIR。

    探测文件必须带 pid：固定名 `.w` 在并行会话下会互相踩——A 建、B 建、A 删、
    B 删时 FileNotFoundError，被 except 吞掉后 B **静默退到下一个候选**，于是
    `--prepare` 落在 `复习资料\\_生产\\_render`、`--render` 却去 `TEMP\\crb_render` 找清单，
    报「先 --prepare」。写权限探测本身不能让调用方换目录。
    """
    for d in render_candidates(root):
        if not _is_ascii(d) and not ALLOW_CJK_RENDER_DIR and not os.environ.get('CRB_RENDER_DIR'):
            continue
        try:
            if create:
                os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, '.w.%d' % os.getpid())
            try:
                with open(probe, 'w') as f:
                    f.write('ok')
            finally:
                try:
                    os.remove(probe)
                except OSError:
                    pass
            return d
        except OSError:
            continue
    raise SystemExit(
        '找不到可写的渲染工作目录。已依次尝试：CRB_RENDER_DIR → 工作区 复习资料\\_生产\\_render'
        ' → 系统 TEMP\\crb_render，均不可写。\n'
        '请设环境变量 CRB_RENDER_DIR 指向一个可写目录（含中文也可以，实测无碍）。')


# ---------- 依赖自检 ----------

REQUIRED = ('docx', 'latex2mathml', 'lxml')

# import 名 → pip 名，只有不一致的才列出来
PIP_NAME = {'docx': 'python-docx'}


def venv_python(root=None):
    """项目级解释器的路径；都没有返回 None。

    优先级：CRB_PYTHON（显式覆盖，工作区没有 .venv 时也靠它指定）→ 工作区
    .venv\\Scripts\\python.exe。PATH 里的 python 不算候选——那通常是宿主自带的
    另一个解释器，依赖不共享。
    """
    env = os.environ.get('CRB_PYTHON')
    if env and os.path.isfile(env):
        return os.path.abspath(env)
    r = root or find_root()
    if not r:
        return None
    p = os.path.join(r, '.venv', 'Scripts', 'python.exe')
    return p if os.path.exists(p) else None


def ensure(need_pymupdf=False):
    """缺依赖时给一条能直接照做的报错，而不是 ImportError 栈。

    两种误判都发生过：把「解释器用错」当成「没装包」，pip 装进了另一套环境；
    以及在没有 .venv 的机器上直接把包装进系统解释器。所以这里只陈述事实与选项，
    不替用户决定——安装命令一律标注「先征得用户同意」，缺 .venv 时要求先去问。
    """
    import importlib
    import sys
    need = list(REQUIRED) + (['pymupdf'] if need_pymupdf else [])
    miss = []
    for m in need:
        try:
            importlib.import_module(m)
        except ImportError:
            miss.append(m)
    if not miss:
        return
    cur = sys.executable
    vp = venv_python()
    pkgs = ' '.join(PIP_NAME.get(m, m) for m in miss)
    same = bool(vp) and os.path.normcase(os.path.abspath(vp)) == \
        os.path.normcase(os.path.abspath(cur))
    lines = ['缺少 Python 依赖：%s' % '、'.join(miss),
             '当前解释器：%s' % cur]
    if not vp:
        lines += ['工作区里没有 .venv，也没有设 CRB_PYTHON。',
                  '先问用户用哪个解释器、是否新建 .venv——不要自行 python -m venv，',
                  '也不要直接把包装进当前解释器。用户指定后可固定为：',
                  '  CRB_PYTHON=<解释器绝对路径>']
    elif not same:
        lines += ['这不是项目级解释器，两套依赖不共享——装到另一边不影响这里。',
                  '改用项目级解释器重跑（通常已装齐，不必再装任何东西）：',
                  '  "%s" <脚本> ...' % vp]
    else:
        lines += ['当前就是项目级解释器，缺包是环境本身不全。']
    lines += ['确需安装时先征得用户同意（说明装哪个包、装进哪套 site-packages、',
              '为什么现有依赖不够），再执行：',
              '  "%s" -m pip install %s' % (cur, pkgs)]
    raise SystemExit('\n'.join(lines))
