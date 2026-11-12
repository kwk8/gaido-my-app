from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
OUT="ai_generated/kickoff_us/Itinerary_US出張_20270105-14_草案.xlsx"
FN="Meiryo UI"; thin=Side(style="thin",color="A6A6A6"); BD=Border(left=thin,right=thin,top=thin,bottom=thin)
HDR=PatternFill("solid",fgColor="D9E1F2"); DAY=PatternFill("solid",fgColor="FCE4D6"); TBD=Font(name=FN,size=10,color="C00000")
M=["神","正木","平松","岡崎","菅野","川北","王","新島"]
wb=Workbook(); ws=wb.active; ws.title="itinerary（US草案）"
ws["B1"]="Itinerary（US出張・草案）"; ws["B1"].font=Font(name=FN,size=14,bold=True); ws["E1"]="2027/1/5-1/14　※未確定。赤字【 】は要確認"; ws["E1"].font=TBD
head=["date","Start","end","ACT","内容","Agenda / Venue","ベンダー担当者情報","Status"]+M+["移動手段","Gift","Address"]
for j,h in enumerate(head,2):
    c=ws.cell(3,j,h); c.font=Font(name=FN,size=10,bold=True); c.fill=HDR; c.border=BD; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
ALL="all"; TKY="tokyo"; OSA="osaka"; VEN="vendor"; VEN2="vendor?"
R=[ # (date, start, end, act, 内容, agenda, 担当者, status, 参加, 移動, gift, address)
("1/5（火）","17:00","09:20","フライト","NH008 成田→サンフランシスコ（東京発）","","","確定",TKY,"","","【成田のターミナル】"),
("","18:35","11:50","フライト","NH7070 関西→サンフランシスコ（大阪発・UA運航）","","","確定",OSA,"","",""),
("","14:20","15:59","フライト","UA661／NH7308 サンフランシスコ→ラスベガス（同じ便）","","","確定",ALL,"","","※受託手荷物有料／合計158cm以内"),
("","【 】","【 】","移動","ラスベガス空港→ホテル","","","【未定】",ALL,"【手段】","",""),
("","","","HOTEL","Mandalay Bay（1/5 チェックイン〜1/9 チェックアウト）","","","確定",ALL,"","","【住所・電話】"),
("1/6（水）","【 】","【 】","視察","CES 2027（1日目）","【セッション・ブース】","","【未定】",ALL,"【ホテル→会場の手段】","","【会場：LVCCほか】"),
("1/7（木）","【 】","【 】","視察","CES 2027（2日目）","【セッション・ブース】","","【未定】",ALL,"【手段】","","【会場】"),
("1/8（金）","【 】","【 】","視察","CES 2027（3日目）","【セッション・ブース】","","【未定】",ALL,"【手段】","","【会場】"),
("1/9（土）","【 】","【 】","移動","ホテルチェックアウト　ホテル→ラスベガス空港","","","【未定】",ALL,"【手段】","",""),
("","09:30","11:05","フライト","WN2503 ラスベガス→サンノゼ","","","確定（便名は要確認：2503/2203）",ALL,"","","※WN便の受託手荷物条件を確認"),
("","【 】","【 】","移動","サンノゼ空港→ホテル","","","【未定】",ALL,"【手段】","",""),
("","","","HOTEL","AC Hotel Sunnyvale Cupertino（1/9 チェックイン〜1/12 チェックアウト）","","","確定",ALL,"","","【住所・電話】"),
("1/10（日）","","","調整日","休息・翌日以降の準備","【過ごし方】","","",ALL,"","",""),
("1/11（月）","午前","","ベンダー訪問","CTCA","全体ブリーフィングを兼ねる","【先方出席者】","調整中",VEN,"貸切車","【お土産】","【住所】"),
("","昼頃","","移動・昼食","CTCA→Cisco（移動＋昼食）","Santa Clara内、移動10〜15分程度","","【昼食場所】",VEN,"貸切車","",""),
("","午後","","ベンダー訪問","Cisco","訪問＋dinner","【先方出席者】（CTC側窓口：椿さん）","調整中（部屋確保できそう）",VEN,"貸切車","【お土産】","【住所】"),
("","【 】","【 】","会食","w/ Cisco（dinner）","【会場】","","【未定】",VEN2,"【手段】","","【住所】"),
("1/12（火）","【 】","【 】","移動","ホテルチェックアウト　ホテル→Exaforce","","","【未定】",VEN2,"貸切車","","【1/12夜は宿泊なし（深夜便）】"),
("","10:30","12:00","ベンダー訪問","Exaforce","訪問・ディスカッション","Madhukar 様","OK（先方から昼食の提案あり）",VEN,"","【お土産】","【住所】"),
("","12:00","13:00","移動・昼食","北へ移動＋昼食","移動約40分／Redwood City付近","","【昼食場所】",VEN,"貸切車","",""),
("","13:00","14:30","ベンダー訪問","MinIO","訪問・ディスカッション","Mike 様（global partner team）","OK",VEN,"","【お土産】","【住所】（Redwood City付近）"),
("","14:30","15:30","移動","MinIO→サンフランシスコ","移動約40〜50分","","",VEN,"貸切車","",""),
("","15:30","17:00","ベンダー訪問","vCluster（SF）","訪問・ディスカッション","Matt 様","返信待ち",VEN,"","【お土産】","【住所】"),
("","【 】","【 】","会食・移動","【夕食】→サンフランシスコ空港","","","【未定】",ALL,"【手段】","",""),
("1/13（水）","00:20","04:50","フライト","NH107 サンフランシスコ→羽田（1/14 着）","","","確定",ALL,"","",""),
("1/14（木）","04:50","","解散","羽田空港着・解散（東京発）","","","",TKY,"","",""),
("","07:00","08:10","フライト","NH013 羽田→伊丹（大阪発）","","","確定",OSA,"","",""),
]
r=4
for row in R:
    d,s,e,act,cont,ag,pic,st,who,mv,gift,adr=row
    if d: r+=1
    vals=[d,s,e,act,cont,ag,pic,st]
    base=["●"]*7
    if who==ALL: marks=base+[""]
    elif who==VEN: marks=base+["●"]
    elif who==VEN2: marks=base+["【 】"]
    elif who==TKY: marks=["●","","●","●","●","●","●",""]; vals[4]=cont+"　※正木さんを除く"
    else: marks=["","●","","","","","",""]; vals[4]=cont+"　※正木さんのみ"
    vals+=marks+[mv,gift,adr]
    for j,v in enumerate(vals,2):
        c=ws.cell(r,j,v); c.border=BD; c.alignment=Alignment(vertical="center",wrap_text=True)
        c.font=TBD if isinstance(v,str) and "【" in v else Font(name=FN,size=10)
        if d and j==2: c.fill=DAY
    r+=1
for col,w in zip("BCDEFGHIJKLMNOPQRST",[11,7,7,11,44,30,26,22,5,5,5,5,5,5,5,9,10,10,26]): ws.column_dimensions[col].width=w
ws.freeze_panes="F4"
# 要確認事項
q=wb.create_sheet("要確認事項")
items=[("参加者","新島さんの合流場所・時刻（1/11 朝）と、1/12 朝の移動・Cisco の dinner に同行するか"),
("フライト","1/9 の WN 便名（キックオフ資料は 2503、大阪発の予約は 2203）"),("フライト","成田のターミナル、空港での集合時刻・場所"),
("宿泊","Mandalay Bay、AC Hotel Sunnyvale Cupertino の住所・電話"),("宿泊","1/12 チェックアウト後、1/13 00:20 発までの荷物と過ごし方"),
("移動","ラスベガス空港→ホテル、ホテル→CES 会場、サンノゼ空港→ホテルの手段"),("移動","ベンダー訪問の貸切車（手配先、車種、人数、集合場所・時刻）"),
("CES","3日間の行動（セッション・ブース）と、会場・集合時刻"),
("ベンダー訪問","各社の確定時刻（CTCA・Cisco は午前／午後のみ）"),("ベンダー訪問","各社の住所、先方出席者（氏名・役職）、アジェンダ、言語、入館手続き"),
("ベンダー訪問","vCluster の返信（訪問可否）"),("食事","1/11 昼食・Cisco の dinner、1/12 昼食（Exaforce の提案を受けるか）・夕食の場所"),
("お土産","各社へのお土産（WBS では当日の朝に購入予定）"),("その他","1/10（調整日）の過ごし方")]
for j,h in enumerate(["分類","確認すること","回答・状況"],2):
    c=q.cell(2,j,h); c.font=Font(name=FN,size=10,bold=True); c.fill=HDR; c.border=BD
for i,(a,b) in enumerate(items,3):
    for j,v in enumerate([a,b,""],2):
        c=q.cell(i,j,v); c.font=Font(name=FN,size=10); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical="center")
q.column_dimensions["B"].width=12; q.column_dimensions["C"].width=70; q.column_dimensions["D"].width=40
wb.save(OUT); print(OUT)
