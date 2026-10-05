#!/usr/bin/env python3
"""抽出した ledger.json から、編集用の機器台帳 Excel を作る。

1行=1機器。U位置は「開始U」＋「U数」で持つため、多U機器も1行で足りる。
前面・背面の両方に描かれていた同一機器は「面=両面」の1行へまとめる。
"""
import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

LEDGER_COLS = [
    ("ラック", 12), ("面", 7), ("開始U", 7), ("U数", 6), ("機器名", 26), ("資産番号", 11),
    ("システム", 10), ("用途", 9), ("重量kg", 8),
    ("電圧V", 7), ("系統数", 7),
    ("回路1", 14), ("電流1A", 8), ("口数1", 7),
    ("回路2", 14), ("電流2A", 8), ("口数2", 7),
    ("回路3", 14), ("電流3A", 8), ("口数3", 7),
    ("マウントキット", 12), ("耐震ベルト", 10), ("備考", 30),
]
RACK_COLS = [
    ("ラック", 12), ("シート", 20), ("寸法", 24), ("総U", 6),
    ("回路", 14), ("電圧V", 7), ("ブレーカA", 10), ("需要率", 8), ("警告閾値", 9),
]
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
CALC_FILL = PatternFill("solid", fgColor="FFF2CC")


def has_power(a):
    return bool(a) and any(c["current"] not in (None, "") for c in a.get("circuits", []))


def has_weight(a):
    return bool(a) and isinstance(a.get("weight"), (int, float))


TEXT_KEYS = ("system", "usage", "note", "mount_kit", "belt")


def merge_faces(rack):
    """同一機器の前面/背面を1件へまとめ、属性はU位置で引き当てる。

    元Excelは属性を「面」ではなく「U位置」で管理していた（背面だけに載る機器の電力が前面側の
    行に書かれている例がある）。このため前面・背面の両方の行から探す。1つの属性行を2台で
    取り合わないよう、電力・重量を引き当てた行は使用済みにする。

    どの機器にも紐づかなかった電力行は、負荷を落とさないため「要確認」の行として残す。
    """
    rowmap = {(r["face"], r["row"]): r for r in rack["rows"]}
    by_u = defaultdict(list)
    for r in rack["rows"]:
        by_u[r["u"]].append(r)
    used_power, used_weight = set(), set()

    groups = defaultdict(list)
    for d in rack["devices"]:
        key = d["asset_no"] if d["asset_no"] else re.sub(r"\s+", "", d["name"])
        groups[key].append(d)

    merged, warnings = [], []
    # 電力を持つ機器を先に処理し、属性行の取り合いを避ける
    order = sorted(groups.items(), key=lambda kv: -max(x["u_size"] for x in kv[1]))
    for key, items in order:
        by_face = defaultdict(list)
        for d in items:
            by_face[d["face"]].append(d)
        n = max(len(v) for v in by_face.values())
        for i in range(n):
            picked = {f: v[i] for f, v in by_face.items() if i < len(v)}
            if not picked:
                continue
            base = picked.get("前面") or picked.get("背面")
            faces = sorted(picked)
            face = "両面" if len(faces) == 2 else faces[0]
            if len(faces) == 2 and (picked["前面"]["u_bottom"] != picked["背面"]["u_bottom"]
                                    or picked["前面"]["u_size"] != picked["背面"]["u_size"]):
                warnings.append(
                    f"要確認 {rack['rack']} {base['name']}: 前面は "
                    f"{picked['前面']['u_bottom']}U/{picked['前面']['u_size']}U だが背面は "
                    f"{picked['背面']['u_bottom']}U/{picked['背面']['u_size']}U（前面の位置を採用）")
            # 候補行: 自分の面の行 -> 占有U（両面ぶん）の行
            cands = [rowmap.get((f, picked[f]["row"])) for f in ("前面", "背面") if f in picked]
            us = set()
            for f, d in picked.items():
                us |= set(range(d["u_bottom"], d["u_top"] + 1))
            for u in sorted(us, reverse=True):
                cands += by_u.get(u, [])
            attr = {}
            for a in cands:
                if has_power(a) and id(a) not in used_power:
                    used_power.add(id(a))
                    attr["circuits"] = [c for c in a["circuits"] if c["current"] not in (None, "")]
                    for k in ("volt", "systems"):
                        if a.get(k) not in (None, ""):
                            attr[k] = a[k]
                    break
            for a in cands:
                if has_weight(a) and id(a) not in used_weight:
                    used_weight.add(id(a))
                    attr["weight"] = a["weight"]
                    break
            for a in cands:
                if not a:
                    continue
                for k in TEXT_KEYS:
                    if a.get(k) not in (None, "") and attr.get(k) in (None, ""):
                        attr[k] = a[k]
            merged.append({"device": base, "face": face, "attr": attr})

    # 紐づかなかった電力行（機器が描かれていない／U位置がずれている）
    for r in rack["rows"]:
        if not has_power(r) or id(r) in used_power:
            continue
        merged.append({
            "device": {"u_bottom": r["u"], "u_top": r["u"], "u_size": 1,
                       "name": f"要確認: 原本 行{r['row']} の電力",
                       "asset_no": "", "row": r["row"], "source": "orphan"},
            "face": r["face"],
            "attr": {"circuits": [c for c in r["circuits"] if c["current"] not in (None, "")],
                     "volt": r.get("volt"), "systems": r.get("systems"),
                     "weight": r.get("weight") if has_weight(r) else None,
                     "system": r.get("system"), "usage": r.get("usage"),
                     "note": "原本に電力の記入はあるが機器が描かれていない行"},
        })
        warnings.append(f"要確認 {rack['rack']} {r['face']} {r['u']}U: "
                        f"原本 {rack['sheet']}!行{r['row']} に電力の記入があるが機器が描かれていない"
                        f"（負荷を落とさないため台帳に『要確認』行として残した）")

    merged.sort(key=lambda m: (m["face"] != "両面", -m["device"]["u_top"]))
    return merged, warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ledger-json", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--demand-factor", type=float, default=0.6)
    ap.add_argument("--warn-threshold", type=float, default=0.8)
    a = ap.parse_args()

    racks = json.loads(Path(a.ledger_json).read_text(encoding="utf8"))
    wb = Workbook()

    ws = wb.active
    ws.title = "機器台帳"
    ws.append([c[0] for c in LEDGER_COLS])
    for i, (name, width) in enumerate(LEDGER_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
        cell = ws.cell(1, i)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "A2"

    total = 0
    all_warnings = []
    for rack in racks:
        mg, warns = merge_faces(rack)
        all_warnings += warns
        for m in mg:
            d, attr = m["device"], m["attr"]
            cir = attr.get("circuits", [])[:3]
            row = [rack["rack"], m["face"], d["u_bottom"], d["u_size"], d["name"], d["asset_no"],
                   attr.get("system") or "", attr.get("usage") or "", attr.get("weight"),
                   attr.get("volt"), attr.get("systems")]
            for i in range(3):
                if i < len(cir):
                    row += [cir[i]["label"], cir[i]["current"], cir[i]["ports"]]
                else:
                    row += ["", None, None]
            row += [attr.get("mount_kit") or "", attr.get("belt") or "", attr.get("note") or ""]
            ws.append(row)
            total += 1

    dv_face = DataValidation(type="list", formula1='"前面,背面,両面"', allow_blank=False)
    ws.add_data_validation(dv_face)
    dv_face.add(f"B2:B{ws.max_row}")

    ws2 = wb.create_sheet("ラック定義")
    ws2.append([c[0] for c in RACK_COLS])
    for i, (name, width) in enumerate(RACK_COLS, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = width
        cell = ws2.cell(1, i)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws2.freeze_panes = "A2"
    for rack in racks:
        m = re.search(r"\((\d+)u\)", rack["spec"] or "", re.I)
        total_u = int(m.group(1)) if m else 47
        if not rack["circuits"]:
            ws2.append([rack["rack"], rack["sheet"], rack["spec"], total_u, "", None, None,
                        a.demand_factor, a.warn_threshold])
        for ci in rack["circuits"]:
            ws2.append([rack["rack"], rack["sheet"], rack["spec"], total_u,
                        ci["label"], ci["volt"], ci["breaker_a"], a.demand_factor, a.warn_threshold])

    ws3 = wb.create_sheet("使い方")
    guide = [
        ["この台帳の使い方"],
        [],
        ["1. 機器を増やす・減らす・移す"],
        ["   「機器台帳」に1行足す（または消す・開始Uを書き換える）だけです。図形を動かす必要はありません。"],
        ["   2U以上の機器は「U数」に2以上を入れます。行を増やす必要はありません。"],
        ["   前面と背面の両方に見える機器は「面」を『両面』にします。前面図と背面図の両方へ自動で描きます。"],
        [],
        ["2. 図と集計を作り直す"],
        ["   build_rack_outputs.py を実行すると、ラック搭載図（drawio/PNG）と電力集計が作り直されます。"],
        ["   台帳が正本です。図と集計は毎回作り直す前提なので、図を直接編集しないでください。"],
        [],
        ["3. 電力の書き方"],
        ["   電圧V・電流A・口数はすべて数値で入れます（『100V』『1系統』のような文字列は計算に使えません）。"],
        ["   電流は『その回路が受け持つ電流』を入れます。消費電力Wは 電圧×電流 の合計で自動計算するため、"],
        ["   台帳には持ちません。"],
        ["   1台が複数の回路に繋がる場合は、回路1〜3に分けて入れます。"],
        [],
        ["4. 判定基準（「ラック定義」で変えられます）"],
        [f"   需要率 {a.demand_factor}: 定格電流にこの係数を掛けたものを予測実効値とします。"],
        [f"   警告閾値 {a.warn_threshold:.0%}: 片系ダウン時に残った回路が全負荷を受けた値が、"],
        ["   ブレーカ定格のこの割合を超えたら警告します。"],
    ]
    for r in guide:
        ws3.append(r)
    ws3.column_dimensions["A"].width = 110
    ws3["A1"].font = Font(bold=True, size=13)

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    print(f"台帳を書き出しました: {a.out}（機器 {total} 行 / ラック定義 {ws2.max_row - 1} 行）")
    if all_warnings:
        print(f"取り込み時の要確認 {len(all_warnings)}件:")
        for w in all_warnings:
            print("  " + w)


if __name__ == "__main__":
    main()
