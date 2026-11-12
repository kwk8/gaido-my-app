from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
SRC="ai_generated/presentations/情通ウェイ発表/template.pptx"
OUT="ai_generated/kickoff_us/ベンダー訪問用_本部・顧客紹介テンプレート_v3full.pptx"
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
# 1 表紙
s=prs.slides.add_slide(L["表紙（White）"])
s.placeholders[17].text="ITOCHU Techno-Solutions Corporation (CTC)"
s.placeholders[0].text="CTC Introduction\n[Division name]"
s.placeholders[15].text="[Customer name] team"
s.placeholders[18].text="January [11 / 12], 2027"
s.placeholders[13].text="[Presenter name]"
drop(s,14)
s.notes_slide.notes_text_frame.text="表紙。部名・担当顧客・訪問日・発表者を記入します。資料・説明は英語で作成します。"
# 2 目次
s=prs.slides.add_slide(L["目次（White）"])
s.placeholders[18].text="Contents"; s.placeholders[15].text="01\n02\n03\n04\n05"
s.placeholders[17].text="Member Self-Introduction\nOur Division & Customer\nContributions and Achievements\nCustomer Interests and Issues\nQuestions for [Vendor name]"
s.notes_slide.notes_text_frame.text="目次。使わない項目は消してください。"
# 3 自己紹介
s=body("Member Self-Introduction","1人1枚。30秒〜1分で話せる量にします。部署・担当顧客は次のページで話すので、ここでは名前・役割・人柄が伝わる程度にします。写真は任意です。")
box(s,0.89,1.45,7.4,4.95)
tb(s,1.2,1.75,6.9,0.8,[("[Name in English]",{"size":32,"bold":True,"c":NAVY})])
tb(s,1.2,2.55,6.9,0.5,"（[氏名]）",18)
rows=[("Role","[e.g. Network Engineer (3rd year)]"),("Interests","[e.g. AI infrastructure]"),("Hobby","[e.g. Traveling, Baseball]")]
for i,(k,v) in enumerate(rows):
    y=3.35+i*0.85
    tb(s,1.2,y,2.0,0.6,k,20,NAVY,True); tb(s,3.2,y,4.9,0.6,v,20)
photo(s,8.55,1.45,3.89,4.95)
guide(s,"名前と役割（何年目）を必ず入れます。部署と担当顧客は、次の「Our Division & Customer」で説明します。")
# 4 本部・顧客紹介
s=body("Our Division & Customer","所属本部と担当顧客の紹介。前に立って話す前提で、1項目1行の短い言葉にします。顧客情報は公開情報（IR、プレスリリース等）の範囲で書きます。")
head(s,0.89,1.45,5.6,"Our Division"); box(s,0.89,1.83,5.6,4.6)
tb(s,1.15,2.05,5.1,4.3,[("[Division name]",{"size":22,"bold":True,"c":NAVY}),"",("• [Main business]",{"size":20}),"",("• [Main customers]",{"size":20}),"",("• [Key technologies]",{"size":20})],20)
head(s,6.84,1.45,5.6,"Our Customer"); box(s,6.84,1.83,5.6,4.6)
tb(s,7.1,2.05,5.1,4.3,[("[Customer name]",{"size":22,"bold":True,"c":NAVY}),"",("• [Industry / position in Japan]",{"size":20}),"",("• [Main services]",{"size":20}),"",("• [Scale (public info)]",{"size":20})],20)
guide(s,"1項目1行で短く書きます。顧客の規模などの数字は公開情報だけを使います。社内の組織図は不要です。")
# 5 実績(1件1ページ)
s=body("Contributions and Achievements","実績は1件1ページ。画像（構成図・写真・製品）を使って説明します。実績が多い場合は、このページを複製します。")
head(s,0.89,1.45,11.55,"[Project name]")
box(s,0.89,1.83,4.9,4.6)
tb(s,1.12,2.0,4.5,4.3,[("What CTC did",{"bold":True,"c":NAVY,"size":18}),("[e.g. Built a virtualized core network]",{"size":20}),"",("Result",{"bold":True,"c":NAVY,"size":18}),("[e.g. Faster deployment, higher availability]",{"size":20})],20)
photo(s,5.95,1.83,6.49,4.6,"[Image: system diagram / photo / product]")
guide(s,"画像は、構成図・写真・製品の画像など、話の内容が一目で分かるものを選びます。顧客名や社外秘の情報は、上長と事務局に確認してから載せます。")
# 6 顧客の関心・課題
s=body("Customer interests and issues","担当顧客の関心ごと（Interests）と課題（Issues）。それぞれ1〜2行に絞ります。")
for i,(t,ph) in enumerate([("Interests",["[e.g. Using generative AI in daily operations]","[ ]"]),("Issues",["[e.g. Data security concerns / GPU cost]","[ ]"])]):
    y=1.45+i*2.55
    head(s,0.89,y,11.55,t); box(s,0.89,y+0.38,11.55,1.95)
    tb(s,1.15,y+0.38,11.0,1.95,["• "+x for x in ph],22,anchor=MSO_ANCHOR.MIDDLE)
guide(s,"課題は、訪問先の製品・サービスと結びつく内容を選ぶと、ディスカッションが広がります。")
# 7 質問
s=body("Questions for [Vendor name]","各ベンダーへの質問を2〜3問。キックオフ資料の質問例（日本企業での導入課題、競合との差別化、実績・ユースケース）を参考にします。")
for i in range(3):
    y=1.55+i*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.95),Inches(y+0.3),Inches(0.6),Inches(0.6)); sh.fill.solid(); sh.fill.fore_color.rgb=NAVY; sh.line.fill.background(); sh.shadow.inherit=False
    tb(s,0.95,y+0.3,0.6,0.6,f"Q{i+1}",14,WHITE,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
    box(s,1.75,y,10.69,1.25)
    tb(s,1.95,y+0.12,10.3,1.05,["[Question]  e.g. What are the main challenges when Japanese companies adopt your solution?","[Background: why you ask (related customer issue)]"],13)
guide(s,"質問の背景（担当顧客のどの課題と関係するか）を添えると、相手も答えやすくなります。")
# 8 裏表紙
prs.slides.add_slide(L["裏表紙（White）"])
prs.save(OUT); print(OUT,len(prs.slides))
