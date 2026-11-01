import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC="uploads/01a11e4a-daa2-7d0a-b74b-992e50475894/情通_若手海外US研修_CES2027_キックオフ_20261021_1.pptx"
OUT="ai_generated/kickoff_us/情通_若手海外US研修_CES2027_キックオフ_構成見直し版_v2.pptx"
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE,LIGHT="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF","BECAD7"
FONT="Noto Sans CJK JP"; FOOT="情通 若手海外（US）研修｜CES 2027"
rgb=lambda h: RGBColor.from_string(h)
prs=Presentation(SRC); BLANK=prs.slide_layouts[6]
O=list(prs.slides._sldIdLst)  # 元の22枚(1始まりで O[n-1])

def tb(s,x,y,w,h,lines,size=14,color=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        txt,kw=(l if isinstance(l,tuple) else (l,{}))
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold); r.font.color.rgb=rgb(kw.get("color",color))
    return sh
def rect(s,x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
    return sh
def badge(s,x,y,w,h,text,fill=ORANGE,size=10,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    rect(s,x,y,w,h,fill,fill,shape); tb(s,x,y,w,h,text,size=size,color=WHITE,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def content(label,title):
    s=prs.slides.add_slide(BLANK); s.background.fill.solid(); s.background.fill.fore_color.rgb=rgb(CREAM)
    tb(s,0.62,0.28,9.8,0.28,label,size=10,color=TEAL); tb(s,0.62,0.68,11.6,0.55,title,size=28,color=NAVY,bold=True)
    tb(s,11.85,0.35,0.85,0.25,"00 / 00",size=9); rect(s,0.62,1.35,12.08,0.03,LINE,shape=MSO_SHAPE.RECTANGLE); tb(s,0.62,7.14,7.0,0.2,FOOT,size=8)
    return s

SEC=[("01","研修概要説明","目的、研修の目的とゴール、旅程とテーマ、持ち帰ってほしいもの、3つの心構え、研修スケジュール、ベンダー訪問日程"),
     ("02","大橋さんからのお言葉",""),
     ("03","研修前課題","CES視察・ベンダー訪問の事前準備、参加者全員にお願いすること、CES 2027 参加登録の手順"),
     ("04","参加メンバー自己紹介",""),
     ("05","今後のスケジュール・準備状況確認","準備状況の確認、確定フライト、ESTA の申請"),
     ("06","現地での行動と報告","現地でのお願い事項、活動記録の共通フォーマット"),
     ("07","参考情報","宿泊・滞在拠点"),
     ("08","QA","")]
# ---- 新しい目次・目的・QA
toc=content("CONTENTS · 目次","目次")
obj=content("OBJECTIVES · 募集要項より抜粋","目的")
cards=[("01","中期経営計画キーアジェンダ「人的資本:グローバル人材の獲得・育成」に資する施策の実施"),("02","情報通信事業Gの次世代を担う人材を育成"),
       ("03","海外未経験の社員が海外ベンダーからビジネスユースケースを聞く機会をつくる"),("04","将来、グローバル人材として活躍するための意識をつくる")]
for i,(n,t) in enumerate(cards):
    x=0.62+(i%2)*6.1; y=1.75+(i//2)*2.3; dark=(i in (1,2))
    rect(obj,x,y,5.95,2.05,NAVY if dark else WHITE,NAVY if dark else LINE)
    badge(obj,x+0.3,y+0.3,0.55,0.32,n,ORANGE if i%2==0 else TEAL)
    tb(obj,x+0.3,y+0.8,5.35,1.1,t,size=17,color=CREAM if dark else NAVY,bold=True)
tb(obj,0.62,6.5,12.05,0.3,"募集要項資料より抜粋",size=10,align=PP_ALIGN.RIGHT)
esta=content("ESTA · 入国手続き","ESTA の申請（11/30 までに）")
rows=[("申請先","米国政府（CBP）の公式サイト：esta.cbp.dhs.gov"),("料金","40.62 ドル（2026/10/16 から）"),
      ("審査","支払い後、最大72時間で結果が出ます"),("有効期間","承認から2年間。パスポートの期限が先に切れる場合は、その日まで"),
      ("持っている方","有効期間内の ESTA があれば、申請し直す必要はありません")]
for i,(k,v) in enumerate(rows):
    y=1.65+i*0.86
    rect(esta,0.62,y,7.9,0.72,WHITE,LINE); badge(esta,0.86,y+0.2,1.55,0.32,k,TEAL if i%2 else ORANGE,size=11)
    tb(esta,2.65,y+0.1,5.75,0.55,v,size=13,color=NAVY,anchor=MSO_ANCHOR.MIDDLE)
rect(esta,8.7,1.65,3.97,4.16,NAVY,NAVY)
tb(esta,8.95,1.85,3.5,3.9,[("期限：11/30（月）まで",{"bold":True,"size":17,"color":ORANGE}),"",
  ("注意",{"bold":True,"color":CREAM}),"・代行サイトは高額な手数料がかかります。必ず公式サイトで申請してください","",
  "・入力内容は、パスポートと同じ表記にしてください","",
  "・経費精算の可否：【要確認】"],size=12,color=CREAM)
rect(esta,0.62,6.1,12.05,0.42,NAVY,NAVY)
tb(esta,0.85,6.1,11.6,0.42,"※料金などは2026年10月時点の情報です。申請時に公式サイトで最新の情報を確認してください。",size=12,color=CREAM,bold=True,anchor=MSO_ANCHOR.MIDDLE)
qa=prs.slides.add_slide(BLANK); qa.background.fill.solid(); qa.background.fill.fore_color.rgb=rgb(NAVY)
tb(qa,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(qa,0.65,2.9,11.5,0.9,"QA",size=40,color=CREAM,bold=True)
tb(qa,0.68,3.95,9,0.4,"ご質問・ご不明点をお聞かせください",size=16,color=LIGHT)
N=list(prs.slides._sldIdLst)[len(O):]
TOC,OBJ,ESTA,QA=N
# ---- 並び(台湾版の流れ) と 各ページの章
plan=[(O[0],None),(TOC,None),
 (OBJ,"01"),(O[4],"01"),(O[3],"01"),(O[15],"01"),(O[5],"01"),(O[6],"01"),(O[9],"01"),
 (O[10],"02"),
 (O[13],"03"),(O[14],"03"),(O[18],"03"),(O[19],"03"),(O[20],"03"),
 (O[11],"04"),
 (O[12],"05"),(O[7],"05"),(ESTA,"05"),
 (O[16],"06"),(O[17],"06"),
 (O[8],"07"),
 (QA,"08"),
 (O[21],None)]
lst=prs.slides._sldIdLst
for e in (O[1],O[2]): prs.part.drop_rel(e.rId); lst.remove(e)
for e in list(lst): lst.remove(e)
for e,_ in plan: lst.append(e)
TOT=len(plan)-1  # メモを除く
secname={k:n for k,n,_ in SEC}
first={}
slides=list(prs.slides)
for i,(s,(_,sec)) in enumerate(zip(slides,plan),1):
    if sec and sec not in first: first[sec]=i
    if sec: last=i
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text.strip()
        if re.fullmatch(r"\d\d / \d\d",t):
            rs=sh.text_frame.paragraphs[0].runs; rs[0].text=f"{i:02d} / {TOT:02d}"
            for r in rs[1:]: r.text=""
        elif sec and abs(Emu(sh.top).inches-0.28)<0.04 and abs(Emu(sh.left).inches-0.62)<0.05 and sh.text_frame.paragraphs[0].runs:
            r=sh.text_frame.paragraphs[0].runs[0]; r.text=f"{sec} {secname[sec]}  ｜  "+r.text
    if sec and sec!="02" and sec!="08" and not any(sh.has_text_frame and sh.text_frame.text.strip()==FOOT for sh in s.shapes):
        tb(s,0.62,7.14,7.0,0.2,FOOT,size=8)
# 章ごとのページ範囲
rng={}
for i,(_,sec) in enumerate(plan,1):
    if sec: rng.setdefault(sec,[i,i])[1]=i
for i,(n,t,d) in enumerate(SEC):
    y=1.6+i*0.66
    rect(toc,0.62,y,12.05,0.56,WHITE,LINE); badge(toc,0.86,y+0.12,0.55,0.32,n,ORANGE if i<4 else TEAL)
    tb(toc,1.65,y+0.08,4.6,0.4,t,size=16,color=NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
    if d: tb(toc,6.0,y+0.1,5.6,0.38,d,size=11,anchor=MSO_ANCHOR.MIDDLE)
    a,b=rng[n]; tb(toc,11.65,y+0.1,0.95,0.38,f"p.{a}" if a==b else f"p.{a}-{b}",size=12,color=TEAL,bold=True,align=PP_ALIGN.RIGHT,anchor=MSO_ANCHOR.MIDDLE)
prs.save(OUT); print(OUT,len(prs.slides),"枚",rng)
