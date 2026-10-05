#!/usr/bin/env python3
"""ラック実装図 Excel の機械点検。

数式・記載・集計の不備を、セル番地つきで列挙する。判断が必要なものは「要判断」として分類し、
勝手に直さない。
"""
import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))
from extract_ledger import extract, resolve_columns, find_bands, norm, CIRCUIT_RE

ERR_VALUES = ("#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!")
# 電力・重量を持たないのが正常な非電源品（棚・アングル・パッチパネル等）
NON_POWERED = re.compile(r"ｹｰﾌﾞﾙ|ケーブル|PatchPnael|パッチ|使用不可|ﾌﾞﾗﾝｸ|ブランク|棚板|棚|ｱﾝｸﾞﾙ|アングル|引き出し|引出|ﾏﾈｰｼﾞﾒﾝﾄ|PDU|ONU")
CELLREF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx, data_only=True)
    wbf = openpyxl.load_workbook(a.xlsx, data_only=False)
    racks = extract(a.xlsx)
    rackmap = {r["sheet"]: r for r in racks}

    findings = defaultdict(list)  # category -> list[str]

    for ws in wb.worksheets:
        wsf = wbf[ws.title]
        sh = ws.title
        rk = rackmap[sh]
        circuits, cols = resolve_columns(ws)
        bands = find_bands(ws)
        band_rows = set()
        for _, t, b in bands:
            band_rows |= set(range(t, b + 1))

        # 1) エラー値が残っているセル
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in ERR_VALUES:
                    f = wsf[c.coordinate].value
                    findings["集計が壊れている"].append(
                        f"{sh}!{c.coordinate} が {c.value.strip()}（式: {f}）")

        # 2) 自分の行以外を参照しているヘルパー式（行ズレのコピペ）
        for row in wsf.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                if c.row not in band_rows:
                    continue
                refs = {int(m.group(2)) for m in CELLREF_RE.finditer(v)}
                bad = {r for r in refs if r != c.row}
                if bad and len(refs) > 0:
                    findings["集計が壊れている"].append(
                        f"{sh}!{c.coordinate} が自分の行({c.row})でなく行{sorted(bad)}を参照（式: {v}）")

        # 3) 電流の記入があるのに、両系集計のヘルパー式が無い行
        helper_cols = []
        for row in wsf.iter_rows(min_row=1, max_row=1):
            pass
        # ヘルパー列は「=IF(V…」形式の式が並ぶ列として検出
        helper_by_col = defaultdict(set)
        for row in wsf.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=IF(") and "*(" in c.value:
                    helper_by_col[c.column].add(c.row)
        hcols = sorted(helper_by_col)
        missing = []
        for _, top, bot in bands:
            for r in range(top, bot + 1):
                has_current = any(
                    ws.cell(r, ci["col_current"]).value not in (None, "") for ci in circuits)
                if has_current and hcols and not any(r in helper_by_col[hc] for hc in hcols):
                    missing.append(r)
        if missing:
            findings["集計が壊れている"].append(
                f"{sh}: 電流を記入した行に両系集計の式が無い（行 {missing}）"
                f"→ この行の電力は『回路合計』に入らない")

        # 4) 集計ラベルと値の行ズレ
        for r in range(bands[0][2] + 1, bands[0][2] + 12):
            lab = norm(ws.cell(r, 10).value)
            if not lab:
                continue
            vals = [ws.cell(r, c).value for c in range(19, 30)]
            has_val = any(v not in (None, "") for v in vals)
            if lab == "機器総重量" and not has_val:
                findings["集計が壊れている"].append(
                    f"{sh}!J{r} のラベル『機器総重量』に対する値が同じ行に無い"
                    f"（値が1行下にずれている可能性）")

        # 5) 型の不一致（数値であるべき欄が文字列）
        for _, top, bot in bands:
            for r in range(top, bot + 1):
                for key, label in (("volt", "電圧"), ("systems", "系統")):
                    c = cols.get(key)
                    if not c:
                        continue
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip():
                        findings["表記が統一されていない"].append(
                            f"{sh}!{get_column_letter(c)}{r} の{label}が文字列『{v}』（数値でないため計算に使えない）")
                for ci in circuits:
                    for cc, lb in ((ci["col_current"], "電流"), (ci["col_ports"], "口数")):
                        v = ws.cell(r, cc).value
                        if isinstance(v, str) and v.strip():
                            findings["表記が統一されていない"].append(
                                f"{sh}!{get_column_letter(cc)}{r} の{lb}が文字列『{v}』")

        # 6) 消費電力(W) と 電圧×電流 の不整合
        wcol = cols.get("watt")
        vcol = cols.get("volt")
        if wcol and vcol:
            for _, top, bot in bands:
                for r in range(top, bot + 1):
                    w = ws.cell(r, wcol).value
                    volt = ws.cell(r, vcol).value
                    if not isinstance(w, (int, float)) or not isinstance(volt, (int, float)):
                        continue
                    tot_i = sum(ws.cell(r, ci["col_current"]).value
                                for ci in circuits
                                if isinstance(ws.cell(r, ci["col_current"]).value, (int, float)))
                    ports = sum(ws.cell(r, ci["col_ports"]).value
                                for ci in circuits
                                if isinstance(ws.cell(r, ci["col_ports"]).value, (int, float)))
                    if tot_i == 0:
                        continue
                    per_cord = volt * tot_i
                    if abs(w - per_cord) > 1:
                        findings["消費電力の計算根拠がそろっていない"].append(
                            f"{sh}!{get_column_letter(wcol)}{r} = {w}W だが "
                            f"電圧{volt:g}V×電流計{tot_i:g}A = {per_cord:g}W（口数計{ports:g}）")

        # 7) 機器はあるのに電力・重量が空
        rowattr = {(x["face"], x["row"]): x for x in rk["rows"]}
        for d in rk["devices"]:
            if d["face"] != "前面":
                continue
            nm = d["name"]
            if NON_POWERED.search(nm):
                continue
            at = rowattr.get(("前面", d["row"]), {})
            has_pw = any(c["current"] not in (None, "") for c in at.get("circuits", []))
            has_wt = at.get("weight") not in (None, "")
            if not has_pw or not has_wt:
                lack = []
                if not has_pw:
                    lack.append("電力")
                if not has_wt:
                    lack.append("重量")
                findings["機器の記載が抜けている"].append(
                    f"{sh} {d['u_bottom']}U {nm}"
                    f"{'（' + d['asset_no'] + '）' if d['asset_no'] else ''}: "
                    f"{'・'.join(lack)}が未記入（行{d['row']}）")

        # 8) 同一Uに複数台
        occ = defaultdict(list)
        for d in rk["devices"]:
            if NON_POWERED.search(d["name"]):
                continue
            for u in range(d["u_bottom"], d["u_top"] + 1):
                occ[(d["face"], u)].append(d["name"])
        for (face, u), names in sorted(occ.items()):
            if len(names) > 1:
                findings["同じUに複数台（横並びなら正常・要確認）"].append(
                    f"{sh} {face} {u}U に {len(names)}台: {' / '.join(names)}")

        # 9) 機器名の持ち方（図形 or セル）
        n_shape = len([d for d in rk["devices"]])
        n_cell = len([x for x in rk["rows"] if x.get("cell_text")])
        findings["表記が統一されていない"].append(
            f"{sh}: 機器名は図形{n_shape}件 / セル{n_cell}件（シート間で持ち方が違うと一括更新できない）")

        # 9b) 需要率（予測実効値の係数）が式に直書きされている
        for row in wsf.iter_rows(min_row=bands[0][2] + 1, max_row=bands[0][2] + 12):
            for c in row:
                if isinstance(c.value, str) and re.search(r"\*\s*0?\.\d+", c.value):
                    findings["需要率と判定基準が文書化されていない"].append(
                        f"{sh}!{c.coordinate} が需要率を式に直書き（式: {c.value}）"
                        f"→ 根拠と変更箇所がシート上に無い")
                    break
            else:
                continue
            break

        # 10) 回路容量に対する余裕
        for ci in circuits:
            tot = 0.0
            for _, top, bot in bands:
                for r in range(top, bot + 1):
                    v = ws.cell(r, ci["col_current"]).value
                    if isinstance(v, (int, float)):
                        tot += v
            ratio = tot / ci["breaker_a"] * 100 if ci["breaker_a"] else 0
            findings["回路容量の余裕（参考値）"].append(
                f"{sh} {ci['label']}: 定格電流計 {tot:g}A / ブレーカ {ci['breaker_a']:g}A = {ratio:.0f}%")

    out = ["# ラック実装図 点検結果", "",
           f"対象: `{Path(a.xlsx).name}`  ラック{len(racks)}本  機器{sum(len(r['devices']) for r in racks)}件", ""]
    order = ["集計が壊れている", "機器の記載が抜けている", "消費電力の計算根拠がそろっていない",
             "需要率と判定基準が文書化されていない", "表記が統一されていない",
             "同じUに複数台（横並びなら正常・要確認）", "回路容量の余裕（参考値）"]
    for cat in order:
        items = findings.get(cat, [])
        out.append(f"## {cat}（{len(items)}件）")
        out.append("")
        for it in items:
            out.append(f"- {it}")
        out.append("")
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text("\n".join(out), encoding="utf8")
    for cat in order:
        print(f"{cat}: {len(findings.get(cat, []))}件", file=sys.stderr)


if __name__ == "__main__":
    main()
