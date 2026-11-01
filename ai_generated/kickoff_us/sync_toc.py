"""キックオフ資料の目次・ページ番号・章ラベルを、今のページ並びに合わせて直す。
使い方: python3 sync_toc.py 入力.pptx 出力.pptx
- 章は、各ページ左上ラベルの「0N 章名  ｜  」で判定する。
- ラベルに章が無いページ(手で足したページ)は、直前のページと同じ章とみなし、ラベルに章名を付け足す。
- 中扉や QA のように、章名そのものが書かれたページは、その章とみなす。
- 最後の章より後ろのページ(メモなど)は、番号を振らない。
"""
import re, sys
from pptx import Presentation
from pptx.util import Emu

src, out = sys.argv[1], sys.argv[2]
prs = Presentation(src)
slides = list(prs.slides)
LBL = re.compile(r"^(0\d) (.+?)  ｜  ")
PG = re.compile(r"^\d\d / \d\d$")

def boxes(s):
    return [sh for sh in s.shapes if sh.has_text_frame]
def label_box(s):
    for sh in boxes(s):
        if abs(Emu(sh.top).inches - 0.28) < 0.05 and abs(Emu(sh.left).inches - 0.62) < 0.06 and sh.text_frame.text.strip():
            return sh
    return None
def is_toc(s):
    return any(sh.text_frame.text.strip() == "目次" for sh in boxes(s))

toc = next((s for s in slides if is_toc(s)), None)
# 目次の行(章番号・章名・ページ欄)を読む
rows = []
if toc:
    items = [(Emu(sh.top).inches, sh) for sh in boxes(toc)]
    nums = sorted([(y, sh) for y, sh in items if re.fullmatch(r"0\d", sh.text_frame.text.strip())], key=lambda x: x[0])
    for y, nb in nums:
        same = [sh for yy, sh in items if abs(yy - y) < 0.2 and sh is not nb]
        name = next((sh for sh in same if sh.text_frame.paragraphs[0].runs and sh.text_frame.paragraphs[0].runs[0].font.bold and not sh.text_frame.text.startswith("p.")), None)
        pg = next((sh for sh in same if re.fullmatch(r"p\.\d+(-\d+)?|—", sh.text_frame.text.strip())), None)
        rows.append((nb.text_frame.text.strip(), name.text_frame.text.strip() if name else "", pg))
names = {n: t for n, t, _ in rows}

sec_of = []
cur = None
for i, s in enumerate(slides):
    if s is toc or i == 0:
        sec_of.append(None); continue
    lb = label_box(s); sec = None
    if lb:
        m = LBL.match(lb.text_frame.text)
        if m: sec = m.group(1)
    if not sec:
        texts = {sh.text_frame.text.strip() for sh in boxes(s)}
        sec = next((n for n, t in names.items() if t in texts), None)
    if not sec and cur and any(PG.match(sh.text_frame.text.strip()) for sh in boxes(s)):
        sec = cur
        if lb and not LBL.match(lb.text_frame.text):
            r = lb.text_frame.paragraphs[0].runs[0]; r.text = f"{cur} {names.get(cur,'')}  ｜  " + r.text
    sec_of.append(sec)
    if sec: cur = sec

last = max(i for i, s in enumerate(sec_of) if s) + 1
for i, s in enumerate(slides[:last], 1):
    for sh in boxes(s):
        if PG.match(sh.text_frame.text.strip()):
            rs = sh.text_frame.paragraphs[0].runs; rs[0].text = f"{i:02d} / {last:02d}"
            for r in rs[1:]: r.text = ""
rng = {}
for i, sec in enumerate(sec_of, 1):
    if sec: rng.setdefault(sec, [i, i])[1] = i
for n, t, pg in rows:
    if pg is None: continue
    a = rng.get(n)
    txt = "—" if not a else (f"p.{a[0]}" if a[0] == a[1] else f"p.{a[0]}-{a[1]}")
    rs = pg.text_frame.paragraphs[0].runs; rs[0].text = txt
    for r in rs[1:]: r.text = ""
prs.save(out)
for n, t, _ in rows:
    a = rng.get(n); print(n, t, "→", "ページなし" if not a else f"p.{a[0]}-{a[1]}")
print("番号を振ったページ:", last, "枚 / 全", len(slides), "枚")
