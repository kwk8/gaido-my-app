import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
SRC,OUT=sys.argv[1],sys.argv[2]
NAVY,ORANGE,TEAL,GRAY,LINE,WHITE="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","FFFFFF"
FONT="Noto Sans CJK JP"; rgb=lambda h: RGBColor.from_string(h)
prs=Presentation(SRC); s=prs.slides[19]
by=lambda i: next(sh for sh in s.shapes if sh.shape_id==i)
def tb(x,y,w,h,lines,size=12,c=NAVY,bold=False):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True
    for k,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if k==0 else tf.add_paragraph()
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold); r.font.color.rgb=rgb(kw.get("c",c))
    return sh
# タイトル
t=by(3).text_frame.paragraphs[0]; t.runs[0].text="確定フライト"
for r in t.runs[1:]: r.text=""
# 東京発の4行を詰める
groups=[(6,[7,8,9,10,11]),(14,[15,16,17,18,19,39]),(22,[23,24,25,26,27,42]),(30,[31,32,33,34,35])]
for k,(card,items) in enumerate(groups):
    c=by(card); oldy=Emu(c.top).inches; newy=1.78+k*0.8
    c.top=Inches(newy); c.height=Inches(0.68)
    for i in items:
        sh=by(i); sh.top=Inches(Emu(sh.top).inches-oldy+newy-0.07 if i not in (39,42) else Emu(sh.top).inches-oldy+newy-0.14)
tb(0.62,1.44,4,0.3,"東京発",size=12,c=TEAL,bold=True)
# 大阪発
tb(0.62,4.98,4,0.3,"大阪発（1名）",size=12,c=TEAL,bold=True)
card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.62),Inches(5.28),Inches(12.05),Inches(1.05))
card.fill.solid(); card.fill.fore_color.rgb=rgb(WHITE); card.line.color.rgb=rgb(LINE); card.line.width=Pt(1)
tb(0.88,5.38,6.9,0.9,[("往路",{"bold":True,"c":TEAL,"size":11}),"1/5 火　NH 7070　KIX 18:35 → SFO 11:50（UA運航）","1/5 火　NH 7308　SFO 14:20 → LAS 15:59（UA運航）※東京発と同じ便"],size=12)
tb(7.9,5.38,4.6,0.9,[("1/9 以降・復路",{"bold":True,"c":TEAL,"size":11}),"1/9 〜 1/14 羽田着までは東京発と同じ便","1/14 木　NH 013　HND 07:00 → ITM 08:10"],size=12)
prs.save(OUT); print("saved",OUT)
