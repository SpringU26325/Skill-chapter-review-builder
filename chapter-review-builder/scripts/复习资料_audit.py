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
backslash = allmt.count('\\'); dollar = allmt.count('$')
nbsp = allmt.count('\u00a0'); emsp = allmt.count('\u2003')
mtext = doc.count('<m:mtext')
nary = doc.count('<m:nary')
d = doc.count('<m:d>')
mm = doc.count('<m:m>')

print(f"media 文件 {len(media)} 个；drawing/pict/object {draw} 个（应为 0）")
print(f"m:r {mruns} 个；Cambria Math {cambria} 次（应≈m:r）")
print(f"m:t 中 U+1D400–1D7FF 字符 {uni_math} 个（应为 0）")
print(f"m:t 中反斜杠 {backslash}、美元符 {dollar}（应为 0）；m:mtext {mtext}（应为 0）")
print(f"m:t 中 NBSP {nbsp}（应为 0）、EM SPACE {emsp}（\\qquad 正常产物，不计入违规）")
print(f"结构计数：m:d {d}、m:m {mm}、m:nary {nary}、m:eqArr {doc.count('<m:eqArr')}")

# 分式类型：lin = Word 的「线性分式」，渲染成压扁的斜杠，不是上下堆叠。
# 引擎在 add_math 里把 lin 统一提升为 bar，所以成品里 lin 必须为 0。
ftypes = {}
for body in re.findall(r'<m:f>(.*?)</m:f>', doc, re.S):
    mt = re.search(r'<m:type m:val="(\w+)"', body)
    k = mt.group(1) if mt else '缺 fPr(默认 bar)'
    ftypes[k] = ftypes.get(k, 0) + 1
lin = ftypes.get('lin', 0)
summary = '、'.join(f'{k} {v}' for k, v in sorted(ftypes.items())) or '无分式'
print(f"分式类型：{summary}（lin 应为 0；非 0 说明有分式被渲染成压扁的斜杠）")
