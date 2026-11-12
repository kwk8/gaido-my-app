import re, sys
from pptx import Presentation
from pptx.util import Inches, Emu
SRC,OUT=sys.argv[1],sys.argv[2]
prs=Presentation(SRC); S=list(prs.slides); TOT=len(S)
def settext(sh,t):
    rs=sh.text_frame.paragraphs[0].runs; rs[0].text=t
    for r in rs[1:]: r.text=""
for i,s in enumerate(S,1):
    for sh in s.shapes:
        if sh.has_text_frame and re.fullmatch(r"\s*\d+\s*/\s*\d+\s*",sh.text_frame.text): settext(sh,f"{i:02d} / {TOT:02d}")
toc=S[1]; by=lambda i: next(sh for sh in toc.shapes if sh.shape_id==i)
pages={12:"p.3-9",17:"p.10",22:"p.11",28:"p.12-20",34:"p.21-22",40:"p.23-26",45:"p.27",51:"p.28-31"}
for k,v in pages.items(): settext(by(k),v)
prs.save(OUT); print("saved",TOT)
