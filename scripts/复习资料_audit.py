# -*- coding: utf-8 -*-
"""交付前审计（各科通用）：md 公式预检 + docx 四项硬指标。

用法：python 复习资料_audit.py <md> <docx>
科目专属的记号规则（行/列向量、符号表登记）在 <科目>_lint.py 里，不在这里。
"""
import re, sys, io, zipfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import latex2mathml.converter as l2m
from lxml import etree

from _env import find_xsl

T = etree.XSLT(etree.parse(find_xsl()))
NS = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

MD, DOCX = sys.argv[1], sys.argv[2]

# ---------- 1. md 侧公式预检 ----------
src = open(MD, encoding='utf-8').read()
lines, keep, incode = src.split('\n'), [], False
for ln in lines:
    s = ln.strip()
    if s.startswith(':::code'):
        incode = True; keep.append(''); continue
    if incode and s == ':::':
        incode = False; keep.append(''); continue
    keep.append('' if incode else ln)
body = '\n'.join(keep)

display = re.findall(r'\$\$(.+?)\$\$', body, re.S)
inline = re.findall(r'(?<!\$)\$([^$\n]+?)\$(?!\$)', re.sub(r'\$\$.+?\$\$', '', body, flags=re.S))

bad = 0
for kind, items in (('DISPLAY', display), ('INLINE', inline)):
    for tex in items:
        t = tex.strip()
        try:
            mml = l2m.convert(t)
            omml = T(etree.fromstring(mml.encode('utf-8')))
            s = etree.tostring(omml, encoding='unicode')
            if 'oMath' not in s or ('m:r' not in s and 'm:t' not in s):
                print(f"[FAIL] {kind}: {t[:70]}"); bad += 1
        except Exception as e:
            print(f"[ERROR] {kind}: {t[:70]} -> {type(e).__name__}: {e}"); bad += 1
print(f"公式总数 {len(display) + len(inline)}（独立 {len(display)} / 行内 {len(inline)}），转换失败 {bad}")

# 图片引用（code 块内已在上面的 body 里置空，不会误计）
img_refs = re.findall(r'^!\[[^\]]*\]\(([^)\s]+)\)', body, re.M)

# ---------- 2. docx 侧四项硬指标 ----------
z = zipfile.ZipFile(DOCX)
names = z.namelist()
media = [n for n in names if n.startswith('word/media/')]
doc = z.read('word/document.xml').decode('utf-8')

draw = doc.count('<w:drawing') + doc.count('<w:pict') + doc.count('<w:object')
mruns = doc.count('<m:r>') + doc.count('<m:r ')
cambria = doc.count('Cambria Math')
mts = re.findall(r'<m:t[^>]*>(.*?)</m:t>', doc, re.S)
allmt = ''.join(mts)
uni_math = sum(1 for ch in allmt if 0x1D400 <= ord(ch) <= 0x1D7FF)
# 希腊小写粗体（U+1D6C2–U+1D6DA）是 `_bold_greek()` 的产物：Word 不认 m:sty=b/bi 的
# 希腊字母，只能靠码位本身粗。其余数学字母码位仍应被 polish_math 还原成 ASCII + sty。
greek_bold = sum(1 for ch in allmt if 0x1D6C2 <= ord(ch) <= 0x1D6DA)
uni_math_bad = uni_math - greek_bold
backslash = allmt.count('\\'); dollar = allmt.count('$')
nbsp = allmt.count('\u00a0'); emsp = allmt.count('\u2003')
mtext = doc.count('<m:mtext')
nary = doc.count('<m:nary')
d = doc.count('<m:d>')
mm = doc.count('<m:m>')

ndraw = doc.count('<w:drawing')
npict = doc.count('<w:pict') + doc.count('<w:object')
print(f"图片：md 引用 {len(img_refs)} 张；media {len(media)} 个（同图复用会去重）"
      f"；drawing {ndraw} 个；pict/object {npict} 个（应为 0）")
if ndraw != len(img_refs):
    print(f"[FAIL] drawing {ndraw} ≠ md 图片引用 {len(img_refs)}：有图没插进去，或多插了")
if npict:
    print(f"[FAIL] pict/object {npict} 个：旧式图形或嵌入对象，一律不允许")
print(f"m:r {mruns} 个；Cambria Math {cambria} 次（应≈m:r）")
print(f"m:t 中 U+1D400–1D7FF 字符 {uni_math} 个"
      f"（希腊粗体 {greek_bold} 个属预期，其余 {uni_math_bad} 个应为 0）")
if uni_math_bad:
    print(f"[FAIL] 非希腊粗体的数学字母码位 {uni_math_bad} 个：polish_math 应还原成 ASCII + m:sty")
print(f"m:t 中反斜杠 {backslash}、美元符 {dollar}（应为 0）；m:mtext {mtext}（应为 0）")
print(f"m:t 中 NBSP {nbsp}（应为 0）、EM SPACE {emsp}（\\qquad 正常产物，不计入违规）")
print(f"结构计数：m:d {d}、m:m {mm}、m:nary {nary}、m:eqArr {doc.count('<m:eqArr')}")

# 分式类型：lin = Word 的「线性分式」（渲染成斜杠），bar = 上下堆叠。
# 引擎把**真除法**（`1/2`、`x/2`、`v/u`）提升为 bar，**单位**（`\mathrm{m/s}`）刻意
# 保留 lin——单位本来就该写成 m/s，不该排成 m 压在 s 上。所以判据不是「lin 必须为 0」，
# 而是对账：每个 lin 的分子分母必须全是正体（m:sty=p）；出现斜体的 lin 说明有真除法
# 漏成了压扁的斜杠。
MNS = '{%s}' % NS['m']
_ftypes = {}
_lin_bad = 0
_lin_unit = 0


def _all_upright(part):
    """分子/分母里除上下标内的指数数字外，是否全是正体 run。"""
    if part is None:
        return False
    runs, stack = [], [part]
    while stack:
        cur = stack.pop()
        for ch in cur:
            if not isinstance(ch.tag, str):
                continue
            ln = etree.QName(ch).localname
            if ln == 'r':
                runs.append(ch)
            elif ln in ('sup', 'sub', 'fPr'):
                continue
            else:
                stack.append(ch)
    if not runs:
        return False
    for r in runs:
        sty = r.find(MNS + 'rPr/' + MNS + 'sty')
        if sty is None or sty.get(MNS + 'val') != 'p':
            return False
    return True


for _f in etree.fromstring(doc.encode('utf-8')).iter(MNS + 'f'):
    _ty = _f.find(MNS + 'fPr/' + MNS + 'type')
    _val = _ty.get(MNS + 'val') if _ty is not None else '缺 fPr(默认 bar)'
    _ftypes[_val] = _ftypes.get(_val, 0) + 1
    if _val == 'lin':
        if _all_upright(_f.find(MNS + 'num')) and _all_upright(_f.find(MNS + 'den')):
            _lin_unit += 1
        else:
            _lin_bad += 1
summary = '、'.join(f'{k} {v}' for k, v in sorted(_ftypes.items())) or '无分式'
print(f"分式类型：{summary}（其中 lin 单位 {_lin_unit} 条，走 m/s 斜杠形式，属预期）")
if _lin_bad:
    print(f"[FAIL] {_lin_bad} 个 lin 分式的分子或分母含斜体：真除法被排成了压扁的斜杠")
