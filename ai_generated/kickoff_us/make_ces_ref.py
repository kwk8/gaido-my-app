import io, glob
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

BASE="uploads/01a11e4a-daa2-7d0a-b74b-992e50475894/情通_若手海外US研修_CES2027_キックオフ_20261021.pptx"
SRC=glob.glob("ai_generated/input/*/情報通信ビジネス企画統括本部/若手向け海外選抜研修/FY26/US/CES参考情報_2026年9月8日時点.pptx")[0]
OUT="ai_generated/kickoff_us/CES参考情報_2026年9月8日時点_キックオフ書式.pptx"
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE,LIGHT="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF","BECAD7"
FONT="Noto Sans CJK JP"; FOOT="情通 若手海外（US）研修｜CES 2027"; TOT=4
rgb=lambda h: RGBColor.from_string(h)

src=Presentation(SRC)
imgs=[]
for s in src.slides:
    for sh in s.shapes:
        try: imgs.append(sh.image.blob)
        except Exception: pass
prs=Presentation(BASE); lst=prs.slides._sldIdLst
for e in list(lst): prs.part.drop_rel(e.rId); lst.remove(e)
BLANK=prs.slide_layouts[6]

def tb(s,x,y,w,h,lines,size=14,color=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,link=None):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold)
        r.font.color.rgb=rgb(kw.get("color",color))
        if kw.get("link"): r.hyperlink.address=kw["link"]; r.font.underline=True
    return sh
def rect(s,x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
    return sh
def badge(s,x,y,w,h,text,fill=ORANGE,size=10,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    rect(s,x,y,w,h,fill,fill,shape); tb(s,x,y,w,h,text,size=size,color=WHITE,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def slide(label,title,pg,dark=False):
    s=prs.slides.add_slide(BLANK); s.background.fill.solid(); s.background.fill.fore_color.rgb=rgb(NAVY if dark else CREAM)
    if not dark:
        tb(s,0.62,0.28,9.8,0.28,label,size=10,color=TEAL); tb(s,0.62,0.68,11.6,0.55,title,size=28,color=NAVY,bold=True)
        tb(s,11.85,0.35,0.85,0.25,f"{pg:02d} / {TOT:02d}",size=9); rect(s,0.62,1.35,12.08,0.03,LINE,shape=MSO_SHAPE.RECTANGLE)
        tb(s,0.62,7.14,7.0,0.2,FOOT,size=8)
    return s
def picture(s,blob,x,y,w,h):
    rect(s,x,y,w,h,WHITE,LINE)
    im=Image.open(io.BytesIO(blob)); iw,ih=im.size; pad=0.12
    bw,bh=w-2*pad,h-2*pad; sc=min(bw/iw,bh/ih); pw,ph=iw*sc,ih*sc
    s.shapes.add_picture(io.BytesIO(blob),Inches(x+pad+(bw-pw)/2),Inches(y+pad+(bh-ph)/2),Inches(pw),Inches(ph))
def bar(s,y,lines,h=0.42):
    rect(s,0.62,y,12.05,h,NAVY,NAVY); tb(s,0.85,y,11.6,h,lines,size=13,color=CREAM,bold=True,anchor=MSO_ANCHOR.MIDDLE)

# 1 表紙
s=slide("","",1,dark=True)
tb(s,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(s,11.0,0.45,1.6,0.3,"REFERENCE",size=11,color=CREAM,bold=True)
tb(s,0.65,1.25,8.6,0.55,"情通 若手海外（US）研修",size=28,color=CREAM,bold=True)
tb(s,0.65,1.95,11,0.8,"CES 2027 参加登録の参考情報",size=34,color=CREAM,bold=True)
tb(s,0.68,3.0,9.8,0.4,"登録の手順と、参考になる情報をまとめました",size=16,color=CREAM)
tb(s,0.68,6.52,8,0.34,"2026年9月8日時点の情報",size=15,color=ORANGE,bold=True)
# 2 手順①
s=slide("REGISTRATION · STEP 1","参加登録の手順①　公式サイトで登録を始める",2)
picture(s,imgs[0],0.62,1.6,7.4,5.35)
rect(s,8.2,1.6,4.47,2.4,WHITE,LINE); badge(s,8.45,1.85,0.55,0.32,"01")
tb(s,9.15,1.8,3.4,0.4,"「SIGN UP NOW」をクリック",size=16,color=NAVY,bold=True)
tb(s,8.45,2.4,4.0,1.5,"CES 公式サイトの画面右上にある「SIGN UP NOW」をクリックして、登録画面へ進みます。",size=13)
rect(s,8.2,4.2,4.47,2.75,WHITE,LINE); badge(s,8.45,4.45,0.55,0.32,"02",TEAL)
tb(s,9.15,4.4,3.4,0.4,"登録開始の通知を設定",size=16,color=NAVY,bold=True)
tb(s,8.45,5.0,4.0,1.9,[("登録が始まったときに、メールで通知が届くように設定できます。",{}),("CES 2027 Registration Notification",{"link":"https://www.ces.tech/attend/registration-notification/","color":TEAL})],size=13)
# 3 手順②
s=slide("REGISTRATION · STEP 2","参加登録の手順②　登録フォームの入力",3)
picture(s,imgs[1],0.62,1.6,7.4,4.6)
pts=[("01","参加区分","メディア関係者ではないので、通常の参加者（Industry attendee）を選択"),
     ("02","メールアドレス","会社のメールアドレスで登録（会社以外のメールは不可）"),
     ("03","送信","入力が済んだら「SUBMIT（提出）」をクリック")]
for i,(n,k,v) in enumerate(pts):
    y=1.6+i*1.55
    rect(s,8.2,y,4.47,1.4,WHITE,LINE); badge(s,8.45,y+0.2,0.55,0.32,n,ORANGE if i<2 else TEAL)
    tb(s,9.15,y+0.15,3.4,0.4,k,size=16,color=NAVY,bold=True)
    tb(s,8.45,y+0.65,4.05,0.75,v,size=12)
bar(s,6.45,"※ 2026年9月8日時点では、まだ登録できません。登録開始の通知を設定しておいてください。")
# 4 お役立ち情報
s=slide("TIPS · お役立ち情報","お役立ち情報",4)
rect(s,0.62,1.7,12.05,2.0,WHITE,LINE); badge(s,0.86,1.95,1.6,0.32,"参考ブログ",TEAL,size=11)
tb(s,0.86,2.45,11.6,0.5,[("【最新】CES2026参加登録レジストレーションの方法",{"link":"https://www.prismsol.com/ja/post/%E3%80%90%E6%9C%80%E6%96%B0%E3%80%91ces2026%E5%8F%82%E5%8A%A0%E3%83%AC%E3%82%B8%E3%82%B9%E3%83%88%E3%83%AC%E3%83%BC%E3%82%B7%E3%83%A7%E3%83%B3%E3%81%AE%E6%96%B9%E6%B3%95","color":TEAL})],size=18,bold=True)
tb(s,0.86,3.05,11.6,0.5,"2026年9月8日時点で、CES 2027 の登録がまだ始まっていなかったため、前年（CES 2026）の登録方法を解説したブログを参考にしています。",size=13)
bar(s,4.0,"登録画面や手順は、CES 2027 で変わる可能性があります。実際の画面で確認してください。")
prs.save(OUT); print(OUT,len(prs.slides),"枚, 画像",len(imgs))
