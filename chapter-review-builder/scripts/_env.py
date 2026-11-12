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

# 渲染工作目录是否允许中文路径。
# 2026-11-11 P4-b 实测：源与目标都含中文时 ExportAsFixedFormat 正常（10.3s / 8 页）。
# 历史上认定的「中文路径导出静默挂住」不成立，真因是 render.ps1 在创建 Word 实例之后
# 才读取清单，导致导出阻塞；清单改为在 New-Object 之前读取并归一化后即恢复正常。
# 因此默认允许中文路径：渲染中间件留在工作区 复习资料\_render，不向工作区外写文件。
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
            '请在复习资料工作区里执行，或用 --root <工作区根> 指定。')
    return root


def src_dir(root):
    return os.path.join(root, WORKSPACE_MARK, '_源稿')


def out_dir(root, subj):
    return os.path.join(root, WORKSPACE_MARK, subj)


def lock_dir(root=None, create=True):
    """锁目录：工作区内 复习资料\\_locks；定位不到工作区时退到系统临时目录。"""
    if root:
        d = os.path.join(root, WORKSPACE_MARK, '_locks')
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
        in_ws = os.path.join(root, WORKSPACE_MARK, '_render')
        if ALLOW_CJK_RENDER_DIR or _is_ascii(in_ws):
            yield in_ws
    for base in (os.environ.get('TEMP'), os.environ.get('TMP')):
        if base:
            yield os.path.join(base, 'crb_render')
    # 不再兜底写盘符根目录。历史版本在最后一档 yield 了一个写死的根路径，导致工作区
    # 可写失败时把数百 MB 过程文件倒进盘符根——外部可见、换机器可能无权限，且实例
    # 挂死后那里的脚本会被反复拉起。宁可报错，见下方 raise。


def render_dir(root, create=True):
    """取第一个可写候选；都不行则报错要求设 CRB_RENDER_DIR。"""
    for d in render_candidates(root):
        if not _is_ascii(d) and not ALLOW_CJK_RENDER_DIR and not os.environ.get('CRB_RENDER_DIR'):
            continue
        try:
            if create:
                os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, '.w')
            with open(probe, 'w') as f:
                f.write('ok')
            os.remove(probe)
            return d
        except OSError:
            continue
    raise SystemExit(
        '找不到可写的渲染工作目录。已依次尝试：CRB_RENDER_DIR → 工作区 复习资料\\_render'
        ' → 系统 TEMP\\crb_render，均不可写。\n'
        '请设环境变量 CRB_RENDER_DIR 指向一个可写目录（含中文也可以，实测无碍）。')


# ---------- 依赖自检 ----------

REQUIRED = ('docx', 'latex2mathml', 'lxml')


def ensure(need_pymupdf=False):
    """缺依赖时给一条能直接照做的报错，而不是 ImportError 栈。"""
    import importlib
    miss = []
    for m in REQUIRED + (('pymupdf',) if need_pymupdf else ()):
        try:
            importlib.import_module(m)
        except ImportError:
            miss.append(m)
    if miss:
        raise SystemExit(
            '缺少 Python 依赖：%s\n'
            '在工作区 .venv 里装：python -m pip install %s'
            % ('、'.join(miss), ' '.join(miss)))
