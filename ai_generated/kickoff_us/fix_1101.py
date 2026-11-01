import re, copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC=sys.argv[1]; OUT=sys.argv[2]
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF"
FONT="Noto Sans CJK JP"; rgb=lambda h: RGBColor.from_string(h)
prs=Presentation(SRC); S=list(prs.slides)
by=lambda s,i: next(sh for sh in s.shapes if sh.shape_id==i)
def white_card(sh):
    sh.fill.solid(); sh.fill.fore_color.rgb=rgb(WHITE); sh.line.color.rgb=rgb(LINE); sh.line.width=Pt(1)
def color(sh,c):
    for p in sh.text_frame.paragraphs:
        for r in p.runs: r.font.color.rgb=rgb(c)
def settext(sh,t):
    rs=sh.text_frame.paragraphs[0].runs; rs[0].text=t
    for r in rs[1:]: r.text=""
def tb(s,x,y,w,h,text,size=14,c=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    p=tf.paragraphs[0]; p.alignment=align; r=p.add_run(); r.text=text; r.font.name=FONT; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=rgb(c); return sh
def rect(s,x,y,w,h,fill,line=None):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
    return sh
def move(s,ids,dx,dy):
    for i in ids:
        sh=by(s,i); sh.left=sh.left+Inches(dx); sh.top=sh.top+Inches(dy)

# ---- p3 目的: 4枚とも白背景・同じ文字色・番号はオレンジ
s=S[2]
for sh in s.shapes:
    if sh.shape_type==1 and Emu(sh.top).inches>1.5 and Emu(sh.height).inches>1.5: white_card(sh)
    elif sh.shape_type==1 and Emu(sh.height).inches<0.4 and Emu(sh.top).inches>1.5:
        sh.fill.solid(); sh.fill.fore_color.rgb=rgb(ORANGE); sh.line.color.rgb=rgb(ORANGE)
    elif sh.has_text_frame and Emu(sh.top).inches>1.5 and Emu(sh.top).inches<6.4 and not re.fullmatch(r"0\d",sh.text_frame.text.strip()):
        color(sh,NAVY)
# ---- p6 CTCの事業機会を白背景に
s=S[5]; white_card(by(s,16)); color(by(s,19),NAVY); color(by(s,20),GRAY)
# ---- p7 心構え 02 を白背景に
s=S[6]; white_card(by(s,10)); color(by(s,12),NAVY)
# ---- p9 ベンダー訪問日程
s=S[8]
rows=[(8,9,10,11,"午前",ORANGE),(12,13,14,15,"昼頃",TEAL),(20,21,22,23,"午後",ORANGE)]
for k,(bg,bt,ti,de,t,c) in enumerate(rows):
    y=2.42+k*0.76
    b=by(s,bg); b.top=Inches(y); b.fill.solid(); b.fill.fore_color.rgb=rgb(c); b.line.color.rgb=rgb(c)
    by(s,bt).top=Inches(y); settext(by(s,bt),t)
    by(s,ti).top=Inches(y-0.02); by(s,de).top=Inches(y+0.31)
settext(by(s,14),"移動＋昼食")
white_card(by(s,24))
for i in (28,32,36,40,44): color(by(s,i),NAVY)
for i in (29,33,37,41,45): color(by(s,i),GRAY)
# ---- p12 出張前課題: CES を左、ベンダーを右
s=S[11]; settext(by(s,3),"出張前課題")
G={"L1":(6,7,8,9),"R1":(10,11,12,13),"L2":(14,15,16,17),"R2":(18,19,20,21),"L3":(22,23,24,25),"R3":(26,27,28,29)}
pos={k:(Emu(by(s,v[0]).left).inches,Emu(by(s,v[0]).top).inches) for k,v in G.items()}
target={"L1":"L1","R1":"L2","L2":"L3","R2":"R1","L3":"R2","R3":"R3"}  # 元の位置→新しい位置
for src,dst in target.items():
    (x0,y0),(x1,y1)=pos[src],pos[dst]; move(s,G[src],x1-x0,y1-y0)
tb(s,0.62,1.38,5.86,0.25,"CES に関する課題",size=11,c=TEAL,bold=True)
tb(s,6.74,1.38,5.86,0.25,"ベンダー訪問に関する課題",size=11,c=TEAL,bold=True)
# ---- p14 全員が自己紹介を準備 を白背景に
s=S[13]; white_card(by(s,10)); color(by(s,12),NAVY)
src_r=by(s,9).text_frame.paragraphs[0].runs
for p in by(s,13).text_frame.paragraphs:
    for r in p.runs:
        rp=r._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
        if rp is not None:
            for f in rp.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill'): rp.remove(f)
        if src_r and src_r[0].font.color and src_r[0].font.color.type is not None:
            r.font.color.rgb=src_r[0].font.color.rgb

# ---- 章・目次・ページ番号
SEC=[("01","研修概要説明","目的、研修の目的とゴール、旅程とテーマ、持ち帰ってほしいもの、3つの心構え、研修スケジュール、ベンダー訪問日程",3,9),
     ("02","大橋さんからのお言葉","",10,10),
     ("03","参加メンバー自己紹介","",11,11),
     ("04","出張前課題","出張前課題、CES視察・ベンダー訪問の事前準備、CES 2027 参加登録の手順",12,16),
     ("05","現地での行動と報告","現地でのお願い事項、活動記録の共通フォーマット",17,18),
     ("06","今後のスケジュール・準備状況確認","準備状況の確認、確定フライト、宿泊・滞在拠点、ESTA の申請",19,22),
     ("07","QA","",23,23),
     ("08","参考情報","訪問先ベンダーの会社概要（Exaforce、MinIO、vCluster Labs）",24,27)]
TOT=len(S)
secof={}
for n,name,_,a,b in SEC:
    for i in range(a,b+1): secof[i]=(n,name)
PFX=re.compile(r"^(0\d [^｜]+?\s*｜\s*)+")
for i,s in enumerate(S,1):
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text.strip()
        if re.fullmatch(r"\d\d / \d\d",t): settext(sh,f"{i:02d} / {TOT:02d}")
        elif i in secof and i>2 and abs(Emu(sh.top).inches-0.28)<0.05 and abs(Emu(sh.left).inches-0.62)<0.06 and t and sh.text_frame.paragraphs[0].runs:
            para=sh.text_frame.paragraphs[0]; full="".join(r.text for r in para.runs); n,name=secof[i]
            para.runs[0].text=f"{n} {name}  ｜  "+PFX.sub("",full)
            for r in para.runs[1:]: r.text=""
# 目次を描き直す
toc=S[1]
for sh in list(toc.shapes):
    if 1.5<Emu(sh.top).inches<7.0: sh._element.getparent().remove(sh._element)
for k,(n,name,d,a,b) in enumerate(SEC):
    y=1.6+k*0.66
    rect(toc,0.62,y,12.05,0.56,WHITE,LINE); rect(toc,0.86,y+0.12,0.55,0.32,ORANGE,ORANGE)
    tb(toc,0.86,y+0.12,0.55,0.32,n,size=10,c=WHITE,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    tb(toc,1.65,y+0.08,4.6,0.4,name,size=16,c=NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
    if d: tb(toc,6.0,y+0.1,5.6,0.38,d,size=11,anchor=MSO_ANCHOR.MIDDLE)
    tb(toc,11.65,y+0.1,0.95,0.38,f"p.{a}" if a==b else f"p.{a}-{b}",size=12,c=TEAL,bold=True,align=PP_ALIGN.RIGHT,anchor=MSO_ANCHOR.MIDDLE)
prs.save(OUT); print("saved",OUT,TOT)
