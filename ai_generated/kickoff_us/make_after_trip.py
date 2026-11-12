from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
BASE="uploads/01a11e4a-daa2-7d0a-b74b-992e50475894/情通_若手海外US研修_CES2027_キックオフ_20261021.pptx"
OUT="ai_generated/kickoff_us/帰国後のやることリスト_キックオフ書式.pptx"
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF"
FONT="Noto Sans CJK JP"; rgb=lambda h: RGBColor.from_string(h)
prs=Presentation(BASE); lst=prs.slides._sldIdLst
for e in list(lst): prs.part.drop_rel(e.rId); lst.remove(e)
s=prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid(); s.background.fill.fore_color.rgb=rgb(CREAM)
def tb(x,y,w,h,lines,size=14,c=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold); r.font.color.rgb=rgb(kw.get("c",c))
def rect(x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
def badge(x,y,w,h,t,fill,size=11):
    rect(x,y,w,h,fill,fill); tb(x,y,w,h,t,size,WHITE,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
tb(0.62,0.28,9.8,0.28,"AFTER THE TRIP · 帰国後",10,TEAL)
tb(0.62,0.68,11.6,0.55,"帰国後のやることリスト",28,NAVY,True)
tb(11.85,0.35,0.85,0.25,"00 / 00",9)
rect(0.62,1.35,12.08,0.03,LINE,shape=MSO_SHAPE.RECTANGLE)
rows=[("01","出張報告",["出張報告は、各自がそれぞれの部で行います。","報告の開催方法は、各自で所属の部に確認してください。"],"部署ごと・各自で確認"),
      ("02","アンケート回答",["アンケートは、後日 川北から配信する予定です。","回答期限までに回答をお願いします。"],"回答期限：1/22（金）"),
      ("03","出張精算",["経費精算の方法は、川北から順次案内する予定です。","帰国後1週間以内（1/21（木）まで）の完了が目安です。"],"1月中に完了")]
for i,(n,t,d,when) in enumerate(rows):
    y=1.7+i*1.5
    rect(0.62,y,12.05,1.3,WHITE,LINE)
    badge(0.9,y+0.22,0.55,0.32,n,ORANGE)
    tb(1.65,y+0.15,3.0,0.45,t,20,NAVY,True)
    tb(1.65,y+0.62,7.6,0.65,d,13,GRAY)
    rect(9.55,y+0.42,2.85,0.46,TEAL if i!=1 else ORANGE,None)
    tb(9.55,y+0.42,2.85,0.46,when,13,WHITE,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
rect(0.62,6.35,12.05,0.42,NAVY)
tb(0.85,6.35,11.6,0.42,"※ 帰国後すぐに経費精算をお願いします。部のアシスタントさんが海外出張の経費に対応できるか、早めに確認してください。",12,CREAM,True,anchor=MSO_ANCHOR.MIDDLE)
tb(0.62,7.14,7.0,0.2,"情通 若手海外（US）研修｜CES 2027",8)
prs.save(OUT); print(OUT)
