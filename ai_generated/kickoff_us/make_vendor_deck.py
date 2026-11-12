from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
SRC="ai_generated/presentations/情通ウェイ発表/template.pptx"
OUT="ai_generated/kickoff_us/ベンダー訪問用_発表資料テンプレート.pptx"
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
# 0 使い方(配布用。発表前に削除)
s=body("テンプレートの使い方（発表前に削除）","配布用の説明ページです。発表資料からは削除してください。")
rows=[("ページの順番","① 自己紹介 → ② 本部・顧客紹介 → ③ 実績 → ④ 関心・課題"),
      ("枚数","各ページは必要に応じて複製して増やしてください（自己紹介は1人1枚、実績は1件1枚）"),
      ("言語","資料・説明は英語。前に立って話す前提で、文字は大きく（20pt前後）、1項目1行"),
      ("画像","実績ページは3つのレイアウトから選べます。画像の配置は自由に変えて構いません"),
      ("社外に出す情報","訪問先に見せる資料です。顧客名・数字・画像は公開情報の範囲にし、迷うものは上長と事務局に確認"),
      ("お手本","FY25 台湾の各チーム資料（Box：FY25台湾／Quanta用_各本部資料）")]
for i,(k,v) in enumerate(rows):
    y=1.45+i*0.82
    box(s,0.89,y,11.55,0.7); tb(s,1.1,y+0.12,2.5,0.5,k,15,NAVY,True); tb(s,3.6,y+0.12,8.7,0.5,v,14)
# 1 自己紹介
s=body("Member Self-Introduction","1人1枚。人数分複製します。30秒〜1分で話せる量にします。写真は任意です。")
box(s,0.89,1.45,7.4,4.95)
tb(s,1.2,1.75,6.9,0.8,[("[Name in English]",{"size":32,"bold":True,"c":NAVY})])
tb(s,1.2,2.55,6.9,0.5,"（[氏名]）",18)
rows=[("Division","[e.g. Telecom Systems Division No.1]"),("Role","[e.g. Network Engineer (3rd year)]"),("Favorite [ ]","[e.g. Favorite Food: Ramen]")]
for i,(k,v) in enumerate(rows):
    y=3.35+i*0.85
    tb(s,1.2,y,2.3,0.6,k,20,NAVY,True); tb(s,3.5,y,4.6,0.6,v,20)
photo(s,8.55,1.45,3.89,4.95)
guide(s,"「Favorite ~」は、好きなもの（食べ物・スポーツ・場所・技術など）を1つ選び、「Favorite Food: Ramen」のように短く書きます。")
# 2 本部・顧客紹介
s=body("Our Division & Customer","所属本部と担当顧客の紹介。1項目1行の短い言葉にします。顧客情報は公開情報の範囲で書きます。必要に応じて複製します。")
head(s,0.89,1.45,5.6,"Our Division"); box(s,0.89,1.83,5.6,4.6)
tb(s,1.15,2.05,5.1,4.3,[("[Division name]",{"size":22,"bold":True,"c":NAVY}),"",("• [Main business]",{"size":20}),"",("• [Main customers]",{"size":20}),"",("• [Key technologies]",{"size":20})],20)
head(s,6.84,1.45,5.6,"Our Customer"); box(s,6.84,1.83,5.6,4.6)
tb(s,7.1,2.05,5.1,4.3,[("[Customer name]",{"size":22,"bold":True,"c":NAVY}),"",("• [Industry / position in Japan]",{"size":20}),"",("• [Main services]",{"size":20}),"",("• [Scale (public info)]",{"size":20})],20)
guide(s,"1項目1行で短く書きます。顧客の規模などの数字は公開情報だけを使います。社内の組織図は不要です。")
# 3A 実績: 左に文字・右に画像
s=body("Contributions and Achievements","レイアウトA（左に文字・右に画像）。実績は1件1枚。使わないレイアウトのページは削除します。")
head(s,0.89,1.45,11.55,"[Project name]")
box(s,0.89,1.83,4.9,4.6)
tb(s,1.12,2.0,4.5,4.3,[("What CTC did",{"bold":True,"c":NAVY,"size":18}),("[e.g. Built a virtualized core network]",{"size":20}),"",("Result",{"bold":True,"c":NAVY,"size":18}),("[e.g. Faster deployment, higher availability]",{"size":20})],20)
photo(s,5.95,1.83,6.49,4.6,"[Image: system diagram / photo / product]")
guide(s,"レイアウトA：説明と画像を半分ずつ。構成図や製品写真に向いています。")
# 3B 実績: 上に文字・下に大きな画像
s=body("Contributions and Achievements","レイアウトB（上に文字・下に大きな画像）。実績は1件1枚。")
head(s,0.89,1.45,11.55,"[Project name]")
box(s,0.89,1.83,11.55,1.15)
tb(s,1.12,1.88,11.1,1.05,[("What CTC did: [e.g. Built a virtualized core network]",{"size":18,"c":NAVY,"bold":True}),("Result: [e.g. Faster deployment, higher availability]",{"size":18})],18,anchor=MSO_ANCHOR.MIDDLE)
photo(s,0.89,3.1,11.55,3.35,"[Image: wide system diagram / timeline]")
guide(s,"レイアウトB：横に長い構成図や、年表・ネットワーク図に向いています。")
# 3C 実績: 画像中心+一言
s=body("Contributions and Achievements","レイアウトC（画像中心＋一言）。写真やイメージ図で見せたいときに使います。")
photo(s,0.89,1.45,11.55,4.15,"[Image: photo / illustration]")
box(s,0.89,5.7,11.55,0.8,LIGHT,LIGHT)
tb(s,1.1,5.7,11.2,0.8,[("[Project name]: [one-line message, e.g. Supporting Japan's largest mobile network for over 20 years]",{"c":NAVY})],18,NAVY,True,anchor=MSO_ANCHOR.MIDDLE)
guide(s,"レイアウトC：写真やイメージ図を大きく見せ、説明は口頭で補います。")
# 4 関心・課題
s=body("Customer interests and issues","担当顧客の関心ごと（Interests）と課題（Issues）。それぞれ1〜2行に絞ります。必要に応じて複製します。")
for i,(t,ph) in enumerate([("Interests",["[e.g. Using generative AI in daily operations]","[ ]"]),("Issues",["[e.g. Data security concerns / GPU cost]","[ ]"])]):
    y=1.45+i*2.55
    head(s,0.89,y,11.55,t); box(s,0.89,y+0.38,11.55,1.95)
    tb(s,1.15,y+0.38,11.0,1.95,["• "+x for x in ph],22,anchor=MSO_ANCHOR.MIDDLE)
guide(s,"課題は、訪問先の製品・サービスと結びつく内容を選ぶと、ディスカッションが広がります。")
prs.save(OUT); print(OUT,len(prs.slides))
