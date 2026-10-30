import copy, re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC="uploads/01a11e4a-daa2-7d0a-b74b-992e50475894/情通_若手海外US研修_CES2027_キックオフ_20261021.pptx"
OUT_ALL="ai_generated/kickoff_us/情通_若手海外US研修_CES2027_キックオフ_追加ページ統合版_v2.pptx"
OUT_NEW="ai_generated/kickoff_us/情通_若手海外US研修_CES2027_キックオフ_追加ページのみ_v2.pptx"
NAVY,ORANGE,TEAL,GRAY,LINE,CREAM,WHITE,LIGHT="16243F","C46B3D","3E7B97","4A5A75","E8ECEF","F5F0E6","FFFFFF","BECAD7"
FONT="Noto Sans CJK JP"; FOOT="情通 若手海外（US）研修｜CES 2027"
rgb=lambda h: RGBColor.from_string(h)

prs=Presentation(SRC); BLANK=prs.slide_layouts[6]
orig=list(prs.slides._sldIdLst)

def tb(s,x,y,w,h,lines,size=14,color=GRAY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,spacing=None,mark=True):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=tf.margin_right=Inches(0.05)
    for i,l in enumerate(lines if isinstance(lines,list) else [lines]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        if spacing: p.space_after=Pt(spacing)
        if isinstance(l,tuple): txt,kw=l
        else: txt,kw=l,{}
        r=p.add_run(); r.text=txt; r.font.name=FONT; r.font.size=Pt(kw.get("size",size)); r.font.bold=kw.get("bold",bold)
        r.font.color.rgb=rgb(kw.get("color",ORANGE if "【" in txt and mark else color))
    return sh
def rect(s,x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=rgb(fill)
    if line: sh.line.color.rgb=rgb(line); sh.line.width=Pt(1)
    else: sh.line.fill.background()
    return sh
def badge(s,x,y,w,h,text,fill=ORANGE,size=10,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    rect(s,x,y,w,h,fill,fill,shape)
    tb(s,x,y,w,h,text,size=size,color=WHITE,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def slide(label,title,dark=False):
    s=prs.slides.add_slide(BLANK)
    s.background.fill.solid(); s.background.fill.fore_color.rgb=rgb(NAVY if dark else CREAM)
    if not dark:
        tb(s,0.62,0.28,9.8,0.28,label,size=10,color=TEAL)
        tb(s,0.62,0.68,11.6,0.55,title,size=28,color=NAVY,bold=True)
        s._pg=tb(s,11.85,0.35,0.85,0.25,"00 / 00",size=9,color=GRAY)
        rect(s,0.62,1.35,12.08,0.03,LINE,shape=MSO_SHAPE.RECTANGLE)
        tb(s,0.62,7.14,7.0,0.2,FOOT,size=8,color=GRAY)
    return s
def bar(s,y,text,fill=NAVY,h=0.42):
    rect(s,0.62,y,12.05,h,fill,fill)
    tb(s,0.78,y,11.7,h,text,size=13,color=CREAM,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
def table(s,x,y,w,data,widths,size=11,rh=0.36,hdr=NAVY):
    t=s.shapes.add_table(len(data),len(data[0]),Inches(x),Inches(y),Inches(w),Inches(rh*len(data))).table
    for j,cw in enumerate(widths): t.columns[j].width=Inches(cw)
    for i,row in enumerate(data):
        t.rows[i].height=Inches(rh)
        for j,v in enumerate(row):
            c=t.cell(i,j); c.text=""; c.vertical_anchor=MSO_ANCHOR.MIDDLE
            c.margin_left=c.margin_right=Inches(0.08); c.margin_top=c.margin_bottom=Inches(0.02)
            for k,line in enumerate(v.split("\n")):
                p=c.text_frame.paragraphs[0] if k==0 else c.text_frame.add_paragraph()
                r=p.add_run(); r.text=line; r.font.name=FONT; r.font.size=Pt(size)
                if i==0: r.font.bold=True; r.font.color.rgb=rgb(WHITE); p.alignment=PP_ALIGN.CENTER
                else: r.font.color.rgb=rgb(ORANGE if "【" in line else NAVY)
            c.fill.solid(); c.fill.fore_color.rgb=rgb(hdr if i==0 else (WHITE if i%2 else "FAF7F1"))
    return t
NEW={}
# ---------- 目次
s=slide("CONTENTS · 目次","目次")
items=[("01","研修概要","目的とゴール、旅程とテーマ、研修の背景、研修テーマ、3つの心構え"),
       ("02","スケジュール・手配状況","研修スケジュール、確定フライト、宿泊、ベンダー訪問日程"),
       ("03","【　】さんからのお言葉",""),
       ("04","研修前の準備","CES視察の事前準備、ベンダー訪問の事前準備"),
       ("05","参加メンバー自己紹介",""),
       ("06","今後のスケジュール・出張前対応","準備状況の確認、e申請、PC保険"),
       ("07","現地での行動と報告","持ち帰ってほしいもの、現地でのお願い事項、活動記録、お願いすること"),
       ("08","参考情報・QA","訪問先、問い合わせ先、Appendix")]
for i,(n,t,d) in enumerate(items):
    y=1.62+i*0.66
    rect(s,0.62,y,12.05,0.56,WHITE,LINE)
    badge(s,0.86,y+0.12,0.55,0.32,n,ORANGE if i<4 else TEAL)
    tb(s,1.65,y+0.08,4.4,0.4,t,size=16,color=NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
    if d: tb(s,6.0,y+0.1,6.5,0.38,d,size=12,anchor=MSO_ANCHOR.MIDDLE)
NEW["toc"]=s
# ---------- 旅程とテーマ
s=slide("OVERVIEW · 募集要項より抜粋","旅程とテーマ")
rows=[("ルート","US ラスベガス＆シリコンバレー：CES への参加とベンダーへの訪問"),
      ("主な研修テーマ","AI、ロボティクス、サイバーセキュリティ"),
      ("実施日程","1月：7泊9日想定。1/5〜13を予定（確定した旅程は 1/5〜1/14。研修スケジュールを参照）")]
for i,(k,v) in enumerate(rows):
    y=1.65+i*0.72
    rect(s,0.62,y,12.05,0.6,WHITE,LINE); badge(s,0.86,y+0.14,1.9,0.32,k,TEAL,size=11)
    tb(s,3.0,y+0.08,9.5,0.45,v,size=14,color=NAVY,anchor=MSO_ANCHOR.MIDDLE)
y=3.85; rect(s,0.62,y,12.05,2.55,NAVY,NAVY); badge(s,0.86,y+0.22,1.9,0.32,"学習・見学のポイント",ORANGE,size=11)
tb(s,3.0,y+0.18,9.4,2.3,["●  世界最先端のテック・イノベーションの視察","●  AI、ロボティクス、サイバーセキュリティ等の注力分野における最新トレンドを把握",
   "●  グローバルのマーケットでの各ベンダーの影響力を視察","●  シリコンバレーのベンダーを直接訪問",
   "●  USベンダー本社もしくはスタートアップ各社の最新ソリューション・共同開発の可能性を探索"],size=14,color=CREAM,spacing=4)
tb(s,0.62,6.6,12.05,0.3,"募集要項資料より抜粋",size=10,color=GRAY,align=PP_ALIGN.RIGHT)
NEW["route"]=s
# ---------- 研修の背景
s=slide("BACKGROUND · 研修の背景","研修の背景")
blocks=[("市場・技術の動向","【US研修の背景となる市場・技術の動向を記入】"),("お客様の動向","【担当顧客の動向・ニーズを記入】"),
        ("対策方針","【CTCとしての対策方針を記入】"),("具体的な取り組み","【本研修を含む具体的な取り組みを記入】")]
for i,(k,v) in enumerate(blocks):
    x=0.62+(i%2)*6.1; y=1.65+(i//2)*2.45
    rect(s,x,y,5.95,2.25,WHITE if i!=3 else NAVY,LINE if i!=3 else NAVY)
    badge(s,x+0.25,y+0.25,2.2,0.32,k,ORANGE if i<2 else TEAL,size=11)
    tb(s,x+0.25,y+0.75,5.45,1.3,v,size=14,color=NAVY if i!=3 else CREAM)
NEW["bg"]=s
# ---------- 研修テーマ
s=slide("THEME · 研修テーマ","研修テーマ")
th=["世界最先端のテック・イノベーションを視察する","AI、ロボティクス、サイバーセキュリティ等の注力分野の最新トレンドを把握する",
    "グローバルのマーケットでの各ベンダーの影響力を視察する","シリコンバレーのベンダーを直接訪問する",
    "USベンダー本社・スタートアップ各社の最新ソリューションと、共同開発の可能性を探る"]
for i,t in enumerate(th):
    y=1.72+i*0.92
    rect(s,0.62,y,12.05,0.76,WHITE,LINE); badge(s,0.86,y+0.19,0.38,0.38,str(i+1),ORANGE if i<3 else TEAL,size=11,shape=MSO_SHAPE.OVAL)
    tb(s,1.5,y+0.17,10.9,0.42,t,size=16,color=NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
tb(s,0.62,6.6,12.05,0.3,"募集要項資料（学習・見学のポイント）より",size=10,color=GRAY,align=PP_ALIGN.RIGHT)
NEW["theme"]=s
# ---------- 3つの心構え
s=slide("MINDSET · 研修の心構え","実りある研修にするための「3つの心構え」")
mm=["健康と安全を第一に","各訪問先で一つは質問をする","自分なりの成果を一つ持って帰る"]
for i,t in enumerate(mm):
    x=0.62+i*4.07
    rect(s,x,1.9,3.9,3.6,NAVY if i==1 else WHITE,NAVY if i==1 else LINE)
    tb(s,x+0.3,2.15,1.0,0.6,f"0{i+1}",size=28,color=ORANGE,bold=True)
    tb(s,x+0.3,3.1,3.3,1.6,t,size=22,color=CREAM if i==1 else NAVY,bold=True)
bar(s,6.0,"3つを意識して、研修に臨んでください。")
NEW["mind"]=s
# ---------- 中扉
s=slide("","",dark=True)
tb(s,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(s,0.65,2.9,11.5,0.9,"【　】さんからのお言葉",size=34,color=CREAM,bold=True,mark=False)
tb(s,0.68,3.9,9,0.4,"MESSAGE",size=14,color=LIGHT)
NEW["msg"]=s
# ---------- 自己紹介
s=slide("MEMBERS · 自己紹介","参加メンバー自己紹介")
mem=[("情報通信第1本部","テクニカルデザイン第4部","神 崇人",""),("情報通信第1本部","通信ビジネス営業第2部","正木 洋晃",""),
("情報通信第3本部","営業第1部","平松 優季",""),("モバイル＆サービス本部","デジタルテクノロジー第3部","岡崎 勇人","主任"),
("SP・メディア技術第2部","メディア技術第2本部","菅野 海咲",""),("情報通信ビジネス企画統括本部","情報通信事業開発部","川北 敬仁",""),
("情報通信ビジネス企画統括本部","情報通信事業開発部","王 帆","主任")]
table(s,0.62,1.62,12.05,[["#","本部","部署","名前","役職"]]+[[str(i+1),*m] for i,m in enumerate(mem)],[0.6,3.7,3.9,2.25,1.6],size=13,rh=0.5)
rect(s,0.62,5.85,12.05,0.95,NAVY,NAVY)
tb(s,0.85,5.9,11.6,0.85,[("以下の内容を含め、1分程度でお願いします",{"bold":True,"size":14}),"①名前・所属・社会人何年目か　②普段の仕事内容　③今回の研修について思うこと"],size=13,color=CREAM,anchor=MSO_ANCHOR.MIDDLE)
NEW["mem"]=s
# ---------- 今後のスケジュール
s=slide("TIMELINE · 準備状況の確認","今後のスケジュール・準備状況確認")
tb(s,0.62,1.42,12,0.28,"出張申請（e申請）、航空券予約、PC海外保険申請の準備は進んでいますか？",size=12,color=ORANGE,bold=True,mark=False)
for x,w,h in ((0.85,1.6,"時期"),(2.55,2.3,"イベント"),(5.0,6.4,"参加者がやること"),(11.5,1.1,"状況")):
    tb(s,x,1.72,w,0.24,h,size=10,color=TEAL,bold=True)
rows=[("10月",None,None,"e申請、航空券の手配、ホテル予約（Mandalay Bay）、パスポート情報の送付","✓ 済"),
("11/5（木）","キックオフ",ORANGE,"注目テーマ3つの提出",""),
("～11/19（木）","CES事前ミーティング",TEAL,"CES 参加登録、セッション・ブースのピックアップ、当日の動きの確定",""),
("～11/30（月）","資料レビュー①",TEAL,"ESTA の申請、会社・本部紹介の資料作成",""),
("～12/7（月）","資料レビュー②",TEAL,"紹介資料の仕上げ",""),
("12/21（月）","直前ガイダンス",ORANGE,"保険・eSIM・Uber の準備、申請状況の確認（会場：N14X）",""),
("～12/29（火）",None,None,"PC保険の申請（遅くとも出発1週間前）",""),
("1/5～1/14","US出張",NAVY,"CES（1/6～1/8）、ベンダー訪問（1/11～1/12）",""),
("～1/21（木）","精算","5F768F","経費精算、e申請の再回付（帰国後1週間以内）、アンケート回答",""),
("～1/29頃","報告","5F768F","出張報告書の作成（報告会の日程は別途案内）","")]
for i,(d,ev,col,todo,st) in enumerate(rows):
    y=1.98+i*0.44; done=(st!="")
    rect(s,0.62,y,12.05,0.38,"EFEBE3" if done else WHITE,LINE)
    tb(s,0.85,y,1.65,0.38,d,size=12,color=GRAY if done else NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
    if ev: badge(s,2.55,y+0.06,2.25,0.26,ev,col,size=10)
    tb(s,5.0,y,6.4,0.38,todo,size=12,color=GRAY if done else NAVY,anchor=MSO_ANCHOR.MIDDLE)
    if done: tb(s,11.5,y,1.1,0.38,st,size=12,color=TEAL,bold=True,anchor=MSO_ANCHOR.MIDDLE)
bar(s,6.48,"※帰国後すぐに経費精算をお願いします。部のアシスタントさんが海外出張の経費に対応できるか、早めに確認してください。",h=0.4)
NEW["sched"]=s
# ---------- e申請 1/2
s=slide("PRE-TRIP 1 · e申請","出張前対応の確認　① e申請 1/2")
pts=[("申請先","e申請 > 書類作成 > 就業系 > 海外出張申請書（兼 手配依頼書）"),
("出張期間","出発日＝自宅（オフィス）の最寄駅を出る日、帰着日＝最寄駅に着く日"),
("同行者","同行者の氏名を記載し、同行者リストを添付"),
("クラス","プレミアムエコノミーの場合、理由欄に「海外旅費規程第7条第1項に基づく（日本発・アメリカ渡航の往復旅程のため）」"),
("宿泊先","乗り継ぎ＝「（乗り継ぎ）」、帰国便＝「帰宅」、連絡先は「*****」"),
("付帯リクエスト","ホテルは各自で予約（個人手配）。海外旅行保険はJTB-CWTで手配")]
for i,(k,v) in enumerate(pts):
    y=1.6+i*0.8
    rect(s,0.62,y,12.05,0.68,WHITE,LINE); badge(s,0.86,y+0.18,1.7,0.32,k,TEAL if i%2 else ORANGE,size=11)
    tb(s,2.8,y+0.08,9.7,0.55,v,size=13,color=NAVY,anchor=MSO_ANCHOR.MIDDLE)
bar(s,6.5,"期限：出発の1か月前（12/4）まで。承認後、JTB-CWTへ申請書と手配依頼書を送付")
NEW["e1"]=s
# ---------- e申請 2/2
s=slide("PRE-TRIP 1 · e申請","出張前対応の確認　① e申請 2/2")
rect(s,0.62,1.6,7.6,5.2,WHITE,LINE)
tb(s,0.9,3.6,7.0,1.0,["【記入例のスクリーンショットを貼り付け】","（事務局の申請画面：旅程表と経費負担部課コード欄）"],size=14,color=GRAY,align=PP_ALIGN.CENTER)
rect(s,8.4,1.6,4.27,5.2,NAVY,NAVY)
tb(s,8.65,1.8,3.8,4.9,[("費用負担の部課コード",{"bold":True,"size":16,"color":ORANGE}),"",
 "本部によって、本部・部・課のどの単位かが異なる可能性があります。一般的には課の単位です。","",
 "所属の本部で確認したうえで、ご自身のe申請に反映してください。","",
 "情報通信ビジネス企画統括本部のメンバーは、自分の課のコードを入力してください。","",
 ("部課コードは精算時に上長へ要確認",{"bold":True,"color":ORANGE})],size=13,color=CREAM)
NEW["e2"]=s
# ---------- PC保険
s=slide("PRE-TRIP 2 · PC保険","出張前対応の確認　② PC保険")
rect(s,0.62,1.55,7.7,5.3,WHITE,LINE)
mail=["送付先：ctc_order_rentec@orix.jp, pc-desk-dm@ctc-g.co.jp","CC：自分の上長、アシスタントさん","件名：海外出張保険申請_対象機器OIRxxxxx",
"――――――――――――――――","オリックス・レンテック株式会社　ご担当者様","いつもお世話になっております。CTC ○○です。","海外出張のため、以下PCの海外保険付保をお願いいたします。",
"■対象機器情報","・OIR：　・仕向地の国名：アメリカ合衆国","・海外持ち出し期間：10日間（2027年1月5日（火）～1月14日（木））",
"■使用者情報","・住所：　・会社名：伊藤忠テクノソリューションズ株式会社","・部署名：　・ご使用者様名（フルネーム）：","・電話番号：　・メールアドレス：",
"■請求書郵送の送付先","・住所：　・会社名：伊藤忠テクノソリューションズ株式会社","・部署名：　・ご担当者様名（フルネーム）：","・電話番号：　・メールアドレス："]
tb(s,0.85,1.65,7.3,5.1,mail,size=10,color=NAVY)
rect(s,8.5,1.55,4.17,5.3,NAVY,NAVY)
tb(s,8.75,1.75,3.7,5.0,[("左記の文面で各自申請",{"bold":True,"size":16,"color":ORANGE}),"",
 ("対象",{"bold":True}),"SecurePC・シンクラ専用端末（オリックス・レンテックから借用のPC）","",
 ("期限",{"bold":True}),"遅くとも出発1週間前（12/29）まで。早めの申請をお願いします","",
 ("注意",{"bold":True}),"送付先アドレスは、社内マニュアル（2024年版）の記載と異なります。【送付前に最新の宛先を確認】"],size=12,color=CREAM)
NEW["pc"]=s
# ---------- 訪問先
s=slide("REFERENCE 1 · 訪問先","参考情報　訪問先（仮）")
vv=[("1/6～1/8","CES 2027","会場：【Las Vegas Convention Center ほか。確定後に記入】"),
("1/11（月）","CTCA","住所：【　】"),("1/11（月）","Forward Networks","住所：【　】"),("1/11（月）","Cisco","住所：【　】"),
("1/12（火）","Exaforce","住所：【　】"),("1/12（火）","MinIO","住所：【　】（Redwood City 付近）"),("1/12（火）","vCluster","住所：【　】（SF）")]
for i,(d,n,a) in enumerate(vv):
    y=1.58+i*0.72
    rect(s,0.62,y,12.05,0.6,WHITE,LINE); badge(s,0.86,y+0.14,1.45,0.32,d,ORANGE if i==0 else (TEAL if i<4 else NAVY),size=10)
    tb(s,2.55,y+0.1,3.2,0.42,n,size=15,color=NAVY,bold=True,anchor=MSO_ANCHOR.MIDDLE)
    tb(s,5.8,y+0.1,6.7,0.42,a,size=12,anchor=MSO_ANCHOR.MIDDLE)
tb(s,0.62,6.72,12.05,0.3,"訪問先・時間はベンダー訪問日程（UPDATED）に基づく。確定後に住所を記入してください。",size=10,color=GRAY)
NEW["visit"]=s
# ---------- 問い合わせ先
s=slide("REFERENCE 2 · 問い合わせ先","参考情報　問い合わせ先")
rect(s,0.62,1.6,5.95,5.0,WHITE,LINE); badge(s,0.87,1.85,2.2,0.32,"出張までの相談役",ORANGE,size=11)
tb(s,0.87,2.4,5.5,4.0,[("情報通信ビジネス企画統括本部",{"bold":True,"color":NAVY}),"・【氏名】［【部署】］","・【氏名】［【部署】］","",
 ("相談できること",{"bold":True,"color":NAVY}),"・出張手配関連","・資料レビューの内容、日程など","・質疑内容の相談"],size=14)
rect(s,6.72,1.6,5.95,5.0,NAVY,NAVY); badge(s,6.97,1.85,1.8,0.32,"問い合わせ先",TEAL,size=11)
tb(s,6.97,2.4,5.5,4.0,[("① 研修用トカログループ",{"bold":True}),"「【グループ名】」","",
 ("② Adminメーリス",{"bold":True}),"telbiz-global-admin@ctc-g.co.jp","【US研修でも同じメーリスを使うか確認】"],size=14,color=CREAM)
NEW["contact"]=s
# ---------- Appendix 1
s=slide("APPENDIX 1 · 資料レビュー／直前ガイダンス","Appendix 1：資料レビュー・直前ガイダンス")
for i,(k,d,c) in enumerate([("【仮】資料レビュー","日程：【日時】（3日分を確保。2回になる可能性あり）","CESの準備、サンノゼで行うベンダー訪問の準備"),
                             ("【仮】直前ガイダンス","日程：【日時】＠【会議室】","出張前の注意事項、保険加入、各種申請が済んでいるかの確認")]):
    y=1.7+i*2.2
    rect(s,0.62,y,12.05,1.95,WHITE if i==0 else NAVY,LINE if i==0 else NAVY)
    tb(s,0.9,y+0.2,6,0.45,k,size=20,color=NAVY if i==0 else CREAM,bold=True,mark=False)
    tb(s,0.9,y+0.8,11.5,0.4,d,size=14,color=NAVY if i==0 else CREAM)
    tb(s,0.9,y+1.25,11.5,0.4,"内容："+c,size=14,color=GRAY if i==0 else LIGHT)
tb(s,0.62,6.2,12.05,0.6,["※予定は事務局から仮登録済みです。詳細はトカログループでアナウンスします。"],size=12,color=GRAY)
NEW["app1"]=s
# ---------- Appendix 2
s=slide("APPENDIX 2 · 現地連絡","Appendix 2：現地連絡・写真共有用のグループ")
rect(s,0.62,1.7,7.6,4.6,WHITE,LINE)
tb(s,0.9,2.0,7.1,4.0,[("グループ名",{"bold":True,"color":NAVY}),"【グループ名】","",("参加方法（To join）",{"bold":True,"color":NAVY}),"【参加用URL】","",
 "現地での連絡と、写真の共有に使います。出発前に参加してください。"],size=15)
rect(s,8.45,1.7,4.22,4.6,WHITE,LINE)
tb(s,8.6,3.6,3.9,0.8,["【QRコードを貼り付け】"],size=14,color=GRAY,align=PP_ALIGN.CENTER)
NEW["app2"]=s
# ---------- QA
s=slide("","",dark=True)
tb(s,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(s,0.65,2.9,11.5,0.9,"QA",size=40,color=CREAM,bold=True)
tb(s,0.68,3.95,9,0.4,"ご質問・ご不明点をお聞かせください",size=16,color=LIGHT)
NEW["qa"]=s
# ---------- 裏表紙
s=slide("","",dark=True)
tb(s,0.65,0.45,8,0.3,"INFORMATION & COMMUNICATIONS · US TRAINING",size=11,color=ORANGE)
tb(s,0.65,3.0,12,0.6,"情通 若手海外（US）研修｜CES 2027",size=24,color=CREAM,bold=True,align=PP_ALIGN.CENTER)
tb(s,0.65,3.75,12,0.4,"2027.01.05 出発  →  2027.01.14 帰国",size=15,color=ORANGE,bold=True,align=PP_ALIGN.CENTER)
NEW["back"]=s

# ---------- 並べ替え（②の元スライド o1..o13）
lst=prs.slides._sldIdLst; allids=list(lst)
o=orig; n=allids[len(o):]
key=list(NEW.keys()); N={k:n[i] for i,k in enumerate(key)}
order=[o[0],N["toc"],o[1],N["route"],N["bg"],N["theme"],N["mind"],o[2],o[3],o[4],o[5],N["msg"],o[6],o[7],N["mem"],N["sched"],N["e1"],N["e2"],N["pc"],
       o[8],o[9],o[10],o[11],N["visit"],N["contact"],N["app1"],N["app2"],N["qa"],N["back"],o[12]]
for e in list(lst): lst.remove(e)
for e in order: lst.append(e)
# ページ番号を振り直し（裏表紙とメモを除く29枚中、表紙を含む28枚を総数に）
slides=list(prs.slides); TOT=len(order)-2
for i,s in enumerate(slides,1):
    for sh in s.shapes:
        if sh.has_text_frame and re.fullmatch(r"\d\d / \d\d",sh.text_frame.text.strip()):
            r=sh.text_frame.paragraphs[0].runs; r[0].text=f"{i:02d} / {TOT:02d}"
            for x in r[1:]: x.text=""
prs.save(OUT_ALL)
# 追加ページのみ
p2=Presentation(OUT_ALL); lst=p2.slides._sldIdLst
origset={e.rId for e in o}
for e in list(lst):
    if e.rId in origset: p2.part.drop_rel(e.rId); lst.remove(e)
p2.save(OUT_NEW)
print("統合版",len(order),"枚 / 追加",len(NEW),"枚")
