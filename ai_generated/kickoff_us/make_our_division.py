from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
SRC="ai_generated/presentations/情通ウェイ発表/template.pptx"
OUT="ai_generated/kickoff_us/ベンダー訪問用_Our Division（本部紹介）.pptx"
NAVY=RGBColor(0x00,0x3A,0x8C); GRAY=RGBColor(0x7F,0x7F,0x7F); DARK=RGBColor(0x26,0x26,0x26); LIGHT=RGBColor(0xE6,0xEC,0xF5); WHITE=RGBColor(255,255,255); RED=RGBColor(0xC0,0x00,0x00)
prs=Presentation(SRC); lst=prs.slides._sldIdLst
for e in list(lst): prs.part.drop_rel(e.rId); lst.remove(e)
L={l.name:l for l in prs.slide_masters[1].slide_layouts}
def drop(s,*idx):
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx in idx: ph._element.getparent().remove(ph._element)
def tb(s,x,y,w,h,lines,size=14,c=DARK,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold)
        r.font.color.rgb=kw.get("c",GRAY if txt.strip().startswith("[") else c)
    return sh
def box(s,x,y,w,h,fill=WHITE,line=LIGHT):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); sh.adjustments[0]=0.06
    sh.fill.solid(); sh.fill.fore_color.rgb=fill; sh.line.color.rgb=line; sh.line.width=Pt(1); sh.shadow.inherit=False; return sh
def head(s,x,y,w,t):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(0.38)); sh.fill.solid(); sh.fill.fore_color.rgb=NAVY; sh.line.fill.background(); sh.shadow.inherit=False
    tb(s,x+0.12,y,w-0.2,0.38,[(t,{"c":WHITE})],15,WHITE,True,anchor=MSO_ANCHOR.MIDDLE)
def photo(s,x,y,w,h,label="[Photo]"):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=LIGHT; sh.line.color.rgb=GRAY; sh.line.dash_style=4; sh.shadow.inherit=False
    tb(s,x,y,w,h,label,12,GRAY,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def body(title,note):
    s=prs.slides.add_slide(L["本文ページ（White）_1"]); s.placeholders[16].text=title; drop(s,15)
    s.notes_slide.notes_text_frame.text=note; return s
def guide(s,text):
    tb(s,0.89,6.62,11.55,0.3,"【記入ガイド】"+text,10,RED)
from pptx.enum.shapes import MSO_SHAPE as _S
def sbox(s,x,y,w,h,fill,line):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); sh.adjustments[0]=0.1
    sh.fill.solid(); sh.fill.fore_color.rgb=fill; sh.line.color.rgb=line; sh.line.width=Pt(1); sh.shadow.inherit=False; return sh
s=body("Our Division","本部紹介（左右2分割の1枚）。左：ベンダーとフロント部署の間に立つ主管部署／右：事業開発（Physical AI と AI Lab）。")
tb(s,0.89,0.95,11.55,0.4,"Telecommunication Business Strategy & Planning Division",14,GRAY)
NOTE_JP="""【話す内容（日本語メモ）】
左：私たちの本部は、ベンダーとフロント部署（お客様を担当する部署）の間に立つ主管部署として仕事をしています。
右：もう一つの役割は事業開発です。Physical AI への取り組みと、AI Lab をつくって様々な検証を行っています。
【英語の例】Our division stands between vendors and our front divisions, and works as the primary division for vendors. We also work on business development, including Physical AI and an AI Lab for various validations."""
s.notes_slide.notes_text_frame.text=NOTE_JP
# ---- 左：主管部署
head(s,0.89,1.45,5.6,"1   Between Vendors and Front Divisions"); box(s,0.89,1.83,5.6,3.98)
xs=[0.99,2.84,4.69]
labels=[("Vendors",None,WHITE,NAVY),("Our Division","Primary division",NAVY,WHITE),("Front Divisions","Customer-facing",WHITE,NAVY)]
for x,(t,sub,fill,fc) in zip(xs,labels):
    sbox(s,x,2.05,1.5,1.25,fill,NAVY)
    lines=[(t,{"size":15,"bold":True,"c":fc})]+([(sub,{"size":11,"c":fc})] if sub else [])
    tb(s,x,2.05,1.5,1.25,lines,15,fc,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
for x in (2.52,4.37):
    a=s.shapes.add_shape(_S.LEFT_RIGHT_ARROW,Inches(x),Inches(2.53),Inches(0.29),Inches(0.28)); a.fill.solid(); a.fill.fore_color.rgb=RED if False else NAVY; a.line.fill.background(); a.shadow.inherit=False
tb(s,1.15,3.55,5.15,2.2,[("• Stands between vendors and front divisions",{"size":18}),"",("• Works as the primary division for vendors",{"size":18}),"",("[e.g. Coordinating evaluation, support, and business with vendors]",{"size":14})],18)
# ---- 右：事業開発
head(s,6.84,1.45,5.6,"2   Business Development")
sbox(s,6.84,1.83,5.6,1.9,WHITE,LIGHT)
tb(s,7.05,1.9,5.2,1.8,[("Physical AI",{"size":20,"bold":True,"c":NAVY}),("• Initiatives toward Physical AI",{"size":18}),("[Specific initiative / target area]",{"size":14})],18)
sbox(s,6.84,3.83,5.6,1.98,WHITE,LIGHT)
tb(s,7.05,3.9,5.2,1.9,[("AI Lab",{"size":20,"bold":True,"c":NAVY}),("• Building an AI Lab",{"size":18}),("• Running various validations",{"size":18}),("[e.g. what we validate]",{"size":14})],18)
# ---- 下：まとめ
bx=sbox(s,0.89,5.98,11.55,0.52,LIGHT,LIGHT)
tb(s,1.1,5.98,11.1,0.52,[("We connect vendors and front divisions, and create new business.",{"c":NAVY})],17,NAVY,True,anchor=MSO_ANCHOR.MIDDLE)
guide(s,"[ ]は具体的な内容を記入します。英語の部門名は組織図から入れたので、正式名称を確認してください。")
prs.save(OUT); print(OUT,len(prs.slides))
