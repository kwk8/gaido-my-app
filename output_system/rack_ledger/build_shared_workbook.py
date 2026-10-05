#!/usr/bin/env python3
"""SharePoint(Teams) 上で共同編集する前提の、Excel だけで完結するラック搭載図ブックを作る。

台帳シートを1行編集すると、搭載図・電力集計・点検がすべて数式で追随する。
Python もマクロも使わないため、ブラウザの Excel でも動く。
"""
import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MAXROW = 400            # 台帳の数式を張る範囲（増設ぶんの余白込み）
LED = "機器台帳"
RACKDEF = "ラック定義"
CIRDEF = "回路定義"
ELEV = "搭載図"
SUMMARY = "電力集計"
HOWTO = "使い方"

HEAD = PatternFill("solid", fgColor="DDEBF7")
AUTO = PatternFill("solid", fgColor="F2F2F2")     # 自動計算列
WARN = PatternFill("solid", fgColor="FFF2CC")
NGF = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

USAGE_FILL = {
    "Server": "DAE8FC", "Storage": "FFE6CC", "SW": "D5E8D4",
    "PDU": "F8CECC", "Firewall": "E1D5E7", "FW": "E1D5E7",
}
NON_POWERED_WORDS = ["棚板", "棚", "アングル", "ｱﾝｸﾞﾙ", "使用不可", "ケーブル", "ｹｰﾌﾞﾙ",
                     "PatchPnael", "パッチ", "引き出し", "引出", "ブランク", "ﾌﾞﾗﾝｸ",
                     "マネージ", "ﾏﾈｰｼﾞ", "PDU", "ONU"]
POWER_OFF_WORDS = ["電源オフ", "未投入", "停止", "未使用", "撤去"]

LEDGER_COLS = [
    ("ラック", 12), ("面", 7), ("開始U", 7), ("U数", 6), ("終了U", 7), ("機器名", 26),
    ("資産番号", 11), ("システム", 10), ("用途", 9), ("重量kg", 8), ("電圧V", 7), ("系統数", 7),
    ("回路1", 14), ("電流1A", 8), ("口数1", 7),
    ("回路2", 14), ("電流2A", 8), ("口数2", 7),
    ("回路3", 14), ("電流3A", 8), ("口数3", 7),
    ("マウントキット", 12), ("耐震ベルト", 10), ("備考", 28),
    ("回路数", 7), ("消費電力W", 10), ("両系1A", 8), ("両系2A", 8), ("両系3A", 8), ("点検", 46),
]
C = {name: get_column_letter(i) for i, (name, _) in enumerate(LEDGER_COLS, start=1)}
AUTO_COLS = ["終了U", "回路数", "消費電力W", "両系1A", "両系2A", "両系3A", "点検"]
HIDDEN_COLS = ["両系1A", "両系2A", "両系3A"]


def arr(words, cell):
    """機器名などに語が含まれるかを判定する式の断片"""
    lst = ",".join(f'"{w}"' for w in words)
    return f'SUMPRODUCT(--ISNUMBER(SEARCH({{{lst}}},{cell}&"")))>0'


def header(ws, cols):
    for i, (name, width) in enumerate(cols, start=1):
        c = ws.cell(1, i, name)
        c.font = Font(bold=True)
        c.fill = AUTO if name in AUTO_COLS else HEAD
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def build_ledger(wb, racks_json):
    ws = wb.create_sheet(LED)
    header(ws, LEDGER_COLS)
    from build_ledger_xlsx import merge_faces
    r = 2
    warnings = []
    for rack in racks_json:
        mg, warns = merge_faces(rack)
        warnings += warns
        for m in mg:
            d, attr = m["device"], m["attr"]
            cir = attr.get("circuits", [])[:3]
            vals = {
                "ラック": rack["rack"], "面": m["face"], "開始U": d["u_bottom"], "U数": d["u_size"],
                "機器名": d["name"], "資産番号": d["asset_no"],
                "システム": attr.get("system") or "", "用途": attr.get("usage") or "",
                "重量kg": attr.get("weight"), "電圧V": attr.get("volt"), "系統数": attr.get("systems"),
                "マウントキット": attr.get("mount_kit") or "", "耐震ベルト": attr.get("belt") or "",
                "備考": attr.get("note") or "",
            }
            for i in range(3):
                if i < len(cir):
                    vals[f"回路{i+1}"] = cir[i]["label"]
                    vals[f"電流{i+1}A"] = cir[i]["current"]
                    vals[f"口数{i+1}"] = cir[i]["ports"]
            for name, _ in LEDGER_COLS:
                if name in vals:
                    ws[f"{C[name]}{r}"] = vals[name]
            r += 1
    last_data = r - 1

    for row in range(2, MAXROW + 1):
        e = C["電流1A"], C["電流2A"], C["電流3A"]
        ws[f'{C["終了U"]}{row}'] = f'=IF({C["開始U"]}{row}="","",{C["開始U"]}{row}+{C["U数"]}{row}-1)'
        ws[f'{C["回路数"]}{row}'] = (
            f'=COUNT({e[0]}{row},{e[1]}{row},{e[2]}{row})')
        ws[f'{C["消費電力W"]}{row}'] = (
            f'=IF(OR({C["電圧V"]}{row}="",NOT(ISNUMBER({C["電圧V"]}{row}))),"",'
            f'{C["電圧V"]}{row}*SUM({e[0]}{row},{e[1]}{row},{e[2]}{row}))')
        for i in range(3):
            ws[f'{C[f"両系{i+1}A"]}{row}'] = (
                f'=IF(OR({C["回路数"]}{row}=0,{e[i]}{row}=""),"",'
                f'{e[i]}{row}/{C["回路数"]}{row})')

        np_ = arr(NON_POWERED_WORDS, f'${C["機器名"]}{row}')
        off = arr(POWER_OFF_WORDS, f'${C["備考"]}{row}')
        overlap = (
            f'SUMPRODUCT(({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW}=${C["ラック"]}{row})*'
            f'(({LED}!${C["面"]}$2:${C["面"]}${MAXROW}=${C["面"]}{row})+'
            f'({LED}!${C["面"]}$2:${C["面"]}${MAXROW}="両面")+(${C["面"]}{row}="両面")>0)*'
            f'({LED}!${C["開始U"]}$2:${C["開始U"]}${MAXROW}<=${C["開始U"]}{row}+${C["U数"]}{row}-1)*'
            f'({LED}!${C["開始U"]}$2:${C["開始U"]}${MAXROW}+'
            f'{LED}!${C["U数"]}$2:${C["U数"]}${MAXROW}-1>=${C["開始U"]}{row})*'
            f'({LED}!${C["機器名"]}$2:${C["機器名"]}${MAXROW}<>""))-1')
        checks = [
            (f'AND({C["回路数"]}{row}=0,NOT({np_}),NOT({off}))', "電力が未記入"),
            (f'AND(NOT(ISNUMBER({C["重量kg"]}{row})),NOT({np_}))', "重量が未記入"),
            (f'AND({C["電圧V"]}{row}<>"",NOT(ISNUMBER({C["電圧V"]}{row})))', "電圧が数値でない"),
            (f'AND({C["電流1A"]}{row}<>"",NOT(ISNUMBER({C["電流1A"]}{row})))', "電流が数値でない"),
            (f'AND({C["開始U"]}{row}<>"",{C["開始U"]}{row}<1)', "開始Uが1未満"),
            (f'AND({C["終了U"]}{row}<>"",{C["終了U"]}{row}>'
             f'IFERROR(VLOOKUP({C["ラック"]}{row},{RACKDEF}!$A:$B,2,FALSE),47))', "総Uをはみ出す"),
            (f'AND({C["ラック"]}{row}<>"",COUNTIF({RACKDEF}!$A:$A,{C["ラック"]}{row})=0)',
             "ラックが定義にない"),
            (f'AND({overlap}>0,NOT({np_}))', "同じUに他の機器(横並びなら正常)"),
        ]
        for i in range(3):
            checks.append(
                (f'AND({C[f"回路{i+1}"]}{row}<>"",COUNTIFS({CIRDEF}!$A:$A,{C["ラック"]}{row},'
                 f'{CIRDEF}!$B:$B,{C[f"回路{i+1}"]}{row})=0)', f"回路{i+1}が定義にない"))
        # TEXTJOIN は Excel のバージョンによって使えないため、IF と & だけで組み立てる
        body = "&".join(f'IF({cond},"／{msg}","")' for cond, msg in checks)
        ws[f'{C["点検"]}{row}'] = (
            f'=IF({C["機器名"]}{row}="","",IF(({body})="","",MID({body},2,300)))')

        for name in AUTO_COLS:
            ws[f"{C[name]}{row}"].fill = AUTO
        for i in range(1, len(LEDGER_COLS) + 1):
            ws.cell(row, i).border = BOX

    for name in HIDDEN_COLS:
        ws.column_dimensions[C[name]].hidden = True

    dv = DataValidation(type="list", formula1='"前面,背面,両面"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{C["面"]}2:{C["面"]}{MAXROW}')

    rng = f'{C["点検"]}2:{C["点検"]}{MAXROW}'
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(${C["点検"]}2<>"",ISNUMBER(SEARCH("未記入",${C["点検"]}2)))'], fill=NGF))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(${C["点検"]}2<>"",NOT(ISNUMBER(SEARCH("未記入",${C["点検"]}2))))'], fill=WARN))
    ws.auto_filter.ref = f"A1:{C['点検']}{MAXROW}"
    return last_data, warnings


def build_defs(wb, racks_json, demand, warn):
    ws = wb.create_sheet(RACKDEF)
    header(ws, [("ラック", 14), ("総U", 7), ("寸法", 26), ("需要率", 9), ("警告閾値", 10)])
    for i, rack in enumerate(racks_json, start=2):
        m = re.search(r"\((\d+)u\)", rack["spec"] or "", re.I)
        ws.cell(i, 1, rack["rack"])
        ws.cell(i, 2, int(m.group(1)) if m else 47)
        ws.cell(i, 3, rack["spec"])
        ws.cell(i, 4, demand)
        ws.cell(i, 5, warn)
        ws.cell(i, 5).number_format = "0%"

    ws2 = wb.create_sheet(CIRDEF)
    header(ws2, [("ラック", 14), ("回路", 16), ("電圧V", 8), ("ブレーカA", 10)])
    r = 2
    pairs = []
    for rack in racks_json:
        for ci in rack["circuits"]:
            ws2.cell(r, 1, rack["rack"])
            ws2.cell(r, 2, ci["label"])
            ws2.cell(r, 3, ci["volt"])
            ws2.cell(r, 4, ci["breaker_a"])
            pairs.append((rack["rack"], ci["label"]))
            r += 1
    return pairs


def build_elevation(wb, racks_json):
    ws = wb.create_sheet(ELEV)
    ws["A1"] = "台帳を直すとこの図も変わります。図のセルは数式なので、直接書き換えないでください。"
    ws["A1"].font = Font(bold=True, color="C00000")
    ws.column_dimensions["A"].width = 5
    ws["A3"] = "U"
    ws["A3"].font = Font(bold=True)
    ws["A3"].fill = HEAD

    col = 2
    for rack in racks_json:
        name = rack["rack"]
        m = re.search(r"\((\d+)u\)", rack["spec"] or "", re.I)
        total_u = int(m.group(1)) if m else 47
        for face in ("前面", "背面"):
            show = get_column_letter(col)
            hide = get_column_letter(col + 1)
            ws.column_dimensions[show].width = 30
            ws.column_dimensions[hide].hidden = True
            ws[f"{show}2"] = name
            ws[f"{show}2"].font = Font(bold=True)
            ws[f"{show}2"].fill = HEAD
            ws[f"{show}2"].alignment = Alignment(horizontal="center")
            ws[f"{show}3"] = face
            ws[f"{show}3"].font = Font(bold=True)
            ws[f"{show}3"].fill = HEAD
            ws[f"{show}3"].alignment = Alignment(horizontal="center")
            for i, u in enumerate(range(total_u, 0, -1)):
                row = 4 + i
                ws[f"A{row}"] = u
                ws[f"A{row}"].alignment = Alignment(horizontal="right")
                cond = (
                    f'({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW}="{name}")*'
                    f'(({LED}!${C["面"]}$2:${C["面"]}${MAXROW}="{face}")+'
                    f'({LED}!${C["面"]}$2:${C["面"]}${MAXROW}="両面")>0)*'
                    f'({LED}!${C["機器名"]}$2:${C["機器名"]}${MAXROW}<>"")*'
                    f'({LED}!${C["開始U"]}$2:${C["開始U"]}${MAXROW}<={u})*'
                    f'({LED}!${C["開始U"]}$2:${C["開始U"]}${MAXROW}+'
                    f'{LED}!${C["U数"]}$2:${C["U数"]}${MAXROW}-1>={u})')
                # INDEX(...,0) で配列評価させる形。Ctrl+Shift+Enter 不要で
                # Excel / Excel Online / LibreOffice のどれでも同じ結果になる
                idx = f'MATCH(1,INDEX({cond},0),0)'
                cnt = f'SUMPRODUCT({cond})'
                nm = f'INDEX({LED}!${C["機器名"]}$2:${C["機器名"]}${MAXROW},{idx})'
                ast = f'INDEX({LED}!${C["資産番号"]}$2:${C["資産番号"]}${MAXROW},{idx})'
                usg = f'INDEX({LED}!${C["用途"]}$2:${C["用途"]}${MAXROW},{idx})'
                nc = f'INDEX({LED}!${C["回路数"]}$2:${C["回路数"]}${MAXROW},{idx})'
                ws[f"{show}{row}"] = (
                    f'=IF({cnt}=0,"",{nm}&IF({ast}=""," "," "&{ast})'
                    f'&IF({cnt}>1," 他"&({cnt}-1)&"台",""))')
                ws[f"{hide}{row}"] = (
                    f'=IF({cnt}=0,"",{usg}&IF({nc}=0,"|未記入",""))')
                ws[f"{show}{row}"].border = BOX
                ws[f"{show}{row}"].alignment = Alignment(horizontal="center", vertical="center")
                ws[f"{show}{row}"].font = Font(size=9)
            rng = f"{show}4:{show}{3 + total_u}"
            for usage, color in USAGE_FILL.items():
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[f'ISNUMBER(SEARCH("{usage}",${hide}4))'],
                    fill=PatternFill("solid", fgColor=color)))
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'ISNUMBER(SEARCH("未記入",${hide}4))'],
                font=Font(color="C00000", bold=True)))
            col += 2
    ws.freeze_panes = "B4"
    return ws


def build_summary(wb, pairs):
    ws = wb.create_sheet(SUMMARY)
    header(ws, [("ラック", 14), ("回路", 16), ("電圧V", 8), ("ブレーカA", 10), ("定格電流A", 10),
                ("両系実効A", 10), ("片系実効A", 10), ("片系/ブレーカ", 12), ("警告閾値", 10),
                ("判定", 22)])
    n = C["電流1A"], C["電流2A"], C["電流3A"]
    b = C["両系1A"], C["両系2A"], C["両系3A"]
    m = C["回路1"], C["回路2"], C["回路3"]
    for i, (rack, cir) in enumerate(pairs, start=2):
        ws.cell(i, 1, rack)
        ws.cell(i, 2, cir)
        ws.cell(i, 3, f'=IFERROR(INDEX({CIRDEF}!$C:$C,MATCH(1,INDEX(({CIRDEF}!$A:$A=$A{i})*'
                      f'({CIRDEF}!$B:$B=$B{i}),0),0)),"")')
        ws.cell(i, 4, f'=IFERROR(INDEX({CIRDEF}!$D:$D,MATCH(1,INDEX(({CIRDEF}!$A:$A=$A{i})*'
                      f'({CIRDEF}!$B:$B=$B{i}),0),0)),"")')
        ws.cell(i, 5, "=" + "+".join(
            f'SUMIFS({LED}!${x}$2:${x}${MAXROW},{LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{i},'
            f'{LED}!${y}$2:${y}${MAXROW},$B{i})' for x, y in zip(n, m)))
        ws.cell(i, 6, "=(" + "+".join(
            f'SUMIFS({LED}!${x}$2:${x}${MAXROW},{LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{i},'
            f'{LED}!${y}$2:${y}${MAXROW},$B{i})' for x, y in zip(b, m))
            + f')*IFERROR(VLOOKUP($A{i},{RACKDEF}!$A:$D,4,FALSE),0.6)')
        ws.cell(i, 7, f'=$E{i}*IFERROR(VLOOKUP($A{i},{RACKDEF}!$A:$D,4,FALSE),0.6)')
        ws.cell(i, 8, f'=IFERROR($G{i}/$D{i},"")')
        ws.cell(i, 9, f'=IFERROR(VLOOKUP($A{i},{RACKDEF}!$A:$E,5,FALSE),0.8)')
        ws.cell(i, 10, f'=IF($D{i}="","",IF($G{i}>$D{i}*$I{i},"要対応（片系で閾値超過）","余裕あり"))')
        ws.cell(i, 8).number_format = "0%"
        ws.cell(i, 9).number_format = "0%"
        for c in range(1, 11):
            ws.cell(i, c).border = BOX
    last = len(pairs) + 1
    ws.conditional_formatting.add(f"A2:J{last}", FormulaRule(
        formula=['$J2="要対応（片系で閾値超過）"'], fill=NGF))

    r = last + 3
    ws.cell(r, 1, "ラック単位").font = Font(bold=True)
    r += 1
    for j, h in enumerate(["ラック", "機器総重量kg", "定格消費電力W", "予測実効W",
                           "使用U数(前面)", "電力が未記入の機器", "点検の指摘がある行"], start=1):
        c = ws.cell(r, j, h)
        c.font = Font(bold=True)
        c.fill = HEAD
        c.border = BOX
    racks = []
    for rack, _ in pairs:
        if rack not in racks:
            racks.append(rack)
    for k, rack in enumerate(racks, start=1):
        rr = r + k
        ws.cell(rr, 1, rack)
        ws.cell(rr, 2, f'=SUMIF({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["重量kg"]}$2:${C["重量kg"]}${MAXROW})')
        ws.cell(rr, 3, f'=SUMIF({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["消費電力W"]}$2:${C["消費電力W"]}${MAXROW})')
        ws.cell(rr, 4, f'=$C{rr}*IFERROR(VLOOKUP($A{rr},{RACKDEF}!$A:$D,4,FALSE),0.6)')
        ws.cell(rr, 5, f'=SUMIFS({LED}!${C["U数"]}$2:${C["U数"]}${MAXROW},'
                       f'{LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["面"]}$2:${C["面"]}${MAXROW},"前面")+'
                       f'SUMIFS({LED}!${C["U数"]}$2:${C["U数"]}${MAXROW},'
                       f'{LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["面"]}$2:${C["面"]}${MAXROW},"両面")')
        ws.cell(rr, 6, f'=COUNTIFS({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["点検"]}$2:${C["点検"]}${MAXROW},"*電力が未記入*")')
        ws.cell(rr, 7, f'=COUNTIFS({LED}!${C["ラック"]}$2:${C["ラック"]}${MAXROW},$A{rr},'
                       f'{LED}!${C["点検"]}$2:${C["点検"]}${MAXROW},"?*")')
        for c in range(1, 8):
            ws.cell(rr, c).border = BOX
    return ws


def build_howto(wb, demand, warn):
    ws = wb.create_sheet(HOWTO, 0)
    ws.column_dimensions["A"].width = 112
    rows = [
        ("ラック搭載図（台帳版）の使い方", True),
        ("", False),
        ("■ 直すのは「機器台帳」シートだけです", True),
        ("　機器を足す: 空いている行に1行書きます。ラック・面・開始U・U数・機器名は必ず入れます。", False),
        ("　機器を外す: その行を削除します（行ごと消してください。値だけ消すと点検に残ります）。", False),
        ("　機器を移す: 「開始U」を書き換えます。図は自動で動きます。", False),
        ("　2U以上の機器: 「U数」に2以上を入れます。行を増やす必要はありません。", False),
        ("　前面と背面の両方から見える機器: 「面」を『両面』にします。両方の図に出ます。", False),
        ("", False),
        ("■ 図形は使っていません", True),
        ("　「搭載図」シートは数式で機器名を引いています。図のセルを直接書き換えないでください"
         "（次に誰かが台帳を直すと消えます）。", False),
        ("　図形が無いので、Teams で同時に開いても壊れません。", False),
        ("", False),
        ("■ 電力の書き方", True),
        ("　電圧V・電流A・口数は数値で入れます。『100V』『1系統』のように単位を付けると計算できません。", False),
        ("　電流は「その回路が受け持つ電流」を入れます。1台が2回路に繋がるなら回路1・回路2に分けます。", False),
        ("　消費電力Wは 電圧×電流の合計 で自動計算します。手で入れる欄はありません。", False),
        ("", False),
        ("■ 回路の余裕の見かた（「電力集計」シート）", True),
        ("　両系実効A: 正常時。負荷が繋がっている回路の数で等分されるとみなした値です。", False),
        ("　片系実効A: 片方の回路が落ちたとき。残った回路がその機器の全負荷を受けた値です。", False),
        (f"　どちらも需要率（既定 {demand}）を掛けます。"
         f"片系実効値がブレーカ定格の {warn:.0%} を超えると「要対応」と赤くなります。", False),
        ("　需要率と閾値は「ラック定義」シートで変えられます。数式には埋め込んでいません。", False),
        ("", False),
        ("■ 記入漏れの見つけかた", True),
        ("　「機器台帳」のいちばん右に「点検」列があります。問題のある行だけ色が付きます。", False),
        ("　　赤: 電力や重量が未記入。その機器は電力集計に入っていません。", False),
        ("　　黄: 同じUに他の機器がある、回路名が定義にない、など確認が要るもの。", False),
        ("　「電力集計」シートの下に、ラックごとの未記入件数をまとめています。", False),
        ("", False),
        ("■ ファイルの置きかた", True),
        ("　このブックは1つだけ置き、全員が同じものを直します。", False),
        ("　いつ誰が何を変えたかは SharePoint のバージョン履歴に残ります"
         "（ファイル名に rev日付を付けて別ファイルを作る必要はありません）。", False),
        ("　個人の OneDrive にコピーを置くと、そちらだけ古くなります。コピーは作らないでください。", False),
    ]
    for i, (text, bold) in enumerate(rows, start=1):
        c = ws.cell(i, 1, text)
        if bold:
            c.font = Font(bold=True, size=13 if i == 1 else 11)
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ledger-json", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--demand-factor", type=float, default=0.6)
    ap.add_argument("--warn-threshold", type=float, default=0.8)
    a = ap.parse_args()

    racks = json.loads(Path(a.ledger_json).read_text(encoding="utf8"))
    wb = Workbook()
    wb.remove(wb.active)
    build_howto(wb, a.demand_factor, a.warn_threshold)
    n, warns = build_ledger(wb, racks)
    pairs = build_defs(wb, racks, a.demand_factor, a.warn_threshold)
    build_elevation(wb, racks)
    build_summary(wb, pairs)
    wb.calculation.fullCalcOnLoad = True
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    print(f"共有用ブックを書き出しました: {a.out}")
    print(f"  機器 {n - 1} 行 / ラック {len(racks)} 本 / 回路 {len(pairs)} 本")
    print(f"  取り込み時の要確認 {len(warns)} 件")


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    main()
