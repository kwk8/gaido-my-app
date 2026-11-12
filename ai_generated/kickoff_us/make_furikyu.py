import io, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image
BASE="uploads/01a11e4a-daa2-7d0a-b74b-992e50475894/情通_若手海外US研修_CES2027_キックオフ_20261021.pptx"
SRC=sys.argv[1]; OUT=sys.argv[2]
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF"
FONT="Noto Sans CJK JP"; FOOT="情通 若手海外（US）研修｜CES 2027"; TOT=4
rgb=lambda h: RGBColor.from_string(h)
GUIDE="https://ctcg.sharepoint.com/sites/biz_00034/SiteAssets/%E5%8A%B4%E5%8B%99%E7%AE%A1%E7%90%86%E3%83%BB%E4%BC%91%E6%9A%87/romukanri/kanri_rule/kyuufuri-guide.pdf"
ETIME="https://ctcg.sharepoint.com/sites/biz_00034/SiteAssets/roumu/etime/manual_ippan.pdf"
src=Presentation(SRC); imgs={}
for n,s in enumerate(src.slides,1):
    for sh in s.shapes:
        try: b=sh.image.blob
        except Exception: continue
        im=Image.open(io.BytesIO(b)); w,h=im.size
        im=im.crop((int(w*sh.crop_left),int(h*sh.crop_top),int(w*(1-sh.crop_right)),int(h*(1-sh.crop_bottom))))
        bio=io.BytesIO(); im.convert("RGB").save(bio,"PNG"); imgs.setdefault(n,[]).append(bio.getvalue())
prs=Presentation(BASE); lst=prs.slides._sldIdLst
for e in list(lst): prs.part.drop_rel(e.rId); lst.remove(e)
BLANK=prs.slide_layouts[6]
def tb(s,x,y,w,h,lines,size=14,color=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold); r.font.color.rgb=rgb(kw.get("color",color))
        if kw.get("link"): r.hyperlink.address=kw["link"]; r.font.underline=True
    return sh
def rect(s,x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
    return sh
def badge(s,x,y,w,h,text,fill=ORANGE,size=10):
    rect(s,x,y,w,h,fill,fill); tb(s,x,y,w,h,text,size=size,color=WHITE,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def slide(label,title,pg,dark=False,tsize=28):
    s=prs.slides.add_slide(BLANK); s.background.fill.solid(); s.background.fill.fore_color.rgb=rgb(NAVY if dark else CREAM)
    if not dark:
        tb(s,0.62,0.28,9.8,0.28,label,size=10,color=TEAL); tb(s,0.62,0.68,11.9,0.55,title,size=tsize,color=NAVY,bold=True)
        tb(s,11.85,0.35,0.85,0.25,f"{pg:02d} / {TOT:02d}",size=9); rect(s,0.62,1.35,12.08,0.03,LINE,shape=MSO_SHAPE.RECTANGLE); tb(s,0.62,7.14,7.0,0.2,FOOT,size=8)
    return s
def picture(s,blob,x,y,w,h):
    rect(s,x,y,w,h,WHITE,LINE); im=Image.open(io.BytesIO(blob)); iw,ih=im.size; pad=0.15
    bw,bh=w-2*pad,h-2*pad; sc=min(bw/iw,bh/ih); pw,ph=iw*sc,ih*sc
    s.shapes.add_picture(io.BytesIO(blob),Inches(x+pad+(bw-pw)/2),Inches(y+pad+(bh-ph)/2),Inches(pw),Inches(ph))
# 1 表紙
s=slide("","",1,dark=True)
tb(s,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(s,11.0,0.45,1.6,0.3,"REFERENCE",size=11,color=CREAM,bold=True)
tb(s,0.65,1.25,8.6,0.55,"情通 若手海外（US）研修",size=28,color=CREAM,bold=True)
tb(s,0.65,1.95,11,0.8,"CES 休日の振替申請方法",size=34,color=CREAM,bold=True)
# 2 振休の考え方
s=slide("HOLIDAY · 振替休日","1月11日祝日だが勤務する必要があるため1月14日を振休とする",2,tsize=24)
picture(s,imgs[2][1],0.62,1.6,4.3,5.35)
picture(s,imgs[2][0],5.1,1.6,4.0,4.3)
rect(s,9.28,1.6,3.39,4.3,NAVY,NAVY)
badge(s,9.53,1.85,1.3,0.32,"理由")
tb(s,9.53,2.4,2.95,3.3,["1/14 早朝4時着の便であるため","","1/14 を振休とする"],size=17,color=CREAM,bold=True)
tb(s,5.1,6.05,7.57,0.5,"左：休日の振替ガイドライン（抜粋）　中央：2027年1月のカレンダー",size=10)
# 3 ガイドライン
s=slide("GUIDELINE · 休日の振替ガイドライン","休日の振替ガイドライン",3)
rect(s,0.62,1.6,5.95,1.35,WHITE,LINE); badge(s,0.87,1.82,1.5,0.32,"詳細",ORANGE,size=11)
tb(s,0.87,2.25,5.5,0.6,[("休日の振替ガイドライン",{"link":GUIDE,"color":TEAL,"bold":True,"size":15}),("詳細については、こちらをクリック",{"size":11})])
rect(s,6.72,1.6,5.95,1.35,WHITE,LINE); badge(s,6.97,1.82,1.5,0.32,"申請方法",TEAL,size=11)
tb(s,6.97,2.25,5.5,0.6,[("eTimeマニュアル（一般社員用）2007年4月",{"link":ETIME,"color":TEAL,"bold":True,"size":15}),("休日振替申請は、この資料の 6.（5）休日の振替の申請 を参照",{"size":11})])
picture(s,imgs[3][0],0.62,3.15,12.05,3.75)
# 4 申請の手順
s=slide("HOW TO APPLY · 申請方法","休日の振替の申請（eTimeマニュアル 6.（5））",4)
picture(s,imgs[4][0],0.62,1.6,8.7,5.35)
rect(s,9.5,1.6,3.17,2.2,WHITE,LINE)
tb(s,9.72,1.8,2.8,1.9,[("詳しいルールは",{"size":13}),("休日の振替ガイドライン",{"link":GUIDE,"color":TEAL,"bold":True,"size":15}),("こちらをクリック",{"size":13})],color=GRAY)
prs.save(OUT); print("saved",OUT)
