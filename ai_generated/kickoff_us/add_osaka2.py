import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
SRC,OUT=sys.argv[1],sys.argv[2]
NAVY,ORANGE,TEAL,GRAY,LINE,WHITE,RED="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","FFFFFF","AA463C"
FONT="Noto Sans CJK JP"; rgb=lambda h: RGBColor.from_string(h)
prs=Presentation(SRC); s=prs.slides[19]
by=lambda i: next(sh for sh in s.shapes if sh.shape_id==i)
WN=by(24).text_frame.text.strip()          # 資料の便名をそのまま使う
t=by(3).text_frame.paragraphs[0]; t.runs[0].text="確定フライト"
for r in t.runs[1:]: r.text=""
for sh in list(s.shapes):                   # 中身の行を消して描き直す(見出し・注記・フッターは残す)
    if 1.4<Emu(sh.top).inches<6.4 and sh.shape_id not in (38,): sh._element.getparent().remove(sh._element)
def tb(x,y,w,h,text,size=16,c=NAVY,bold=True,align=PP_ALIGN.LEFT):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align; r=p.add_run(); r.text=text; r.font.name=FONT; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=rgb(c)
def card(x,y,w,h=0.72):
    c=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
    c.fill.solid(); c.fill.fore_color.rgb=rgb(WHITE); c.line.color.rgb=rgb(LINE); c.line.width=Pt(1)
def half(x,y,d,fl,dep,arr,sub=None):
    card(x,y,5.95); tb(x+0.2,y+0.2,1.0,0.3,d,14,TEAL); tb(x+1.15,y+0.19,1.2,0.3,fl)
    tb(x+2.3,y+0.19,1.5,0.3,dep); tb(x+3.75,y+0.17,0.35,0.3,"→",18,ORANGE); tb(x+4.1,y+0.19,1.8,0.3,arr)
    if sub: tb(x+4.1,y+0.46,1.8,0.22,sub,9,GRAY,False)
def full(y,d,fl,dep,arr,note=None):
    card(0.62,y,12.05); tb(0.88,y+0.2,1.15,0.3,d,14,TEAL); tb(2.05,y+0.19,1.3,0.3,fl)
    tb(3.5,y+0.19,2.2,0.3,dep); tb(5.75,y+0.17,0.45,0.3,"→",18,ORANGE); tb(6.25,y+0.19,2.8,0.3,arr)
    if note: tb(8.82,y+0.24,3.6,0.3,note,11,RED)
L,R=0.62,6.72
tb(L,1.45,4,0.28,"東京発",12,TEAL); tb(R,1.45,4,0.28,"大阪発（1名）",12,TEAL)
half(L,1.75,"1/5 火","NH 008","NRT 17:00","SFO 09:20")
half(R,1.75,"1/5 火","NH 7070","KIX 18:35","SFO 11:50","UA運航")
half(L,2.57,"1/5 火","UA 661","SFO 14:20","LAS 15:59")
half(R,2.57,"1/5 火","NH 7308","SFO 14:20","LAS 15:59","UA運航（東京発と同じ便）")
tb(L,3.33,12,0.25,"※UA便：受託手荷物有料／合計158cm（62インチ）以内",11,RED)
tb(L,3.65,6,0.28,"1/9 以降（全員同じ便）",12,TEAL)
full(3.95,"1/9 土",WN,"LAS 09:30","SJC 11:05","WN便は受託手荷物条件を別途確認")
full(4.77,"1/13 水","NH 107","SFO 00:20","HND 04:50（1/14着）")
tb(R,5.55,4,0.28,"大阪発のみ",12,TEAL)
half(R,5.83,"1/14 木","NH 013","HND 07:00","ITM 08:10")
prs.save(OUT); print("saved",OUT,"WN=",WN)
