#!/usr/bin/env python3
"""ラック実装図 Excel から機器台帳（1行=1機器）を抽出する。

機器名・型番・資産番号はセルではなく浮動図形（テキストボックス）に入っているため、
図形のアンカー行から搭載U位置を割り出し、同じ行にある電源・重量の属性と結合する。
"""
import argparse
import csv
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

CIRCUIT_RE = re.compile(r"(\d+)\s*V\s*(\d+(?:\.\d+)?)\s*A")
ASSET_RE = re.compile(r"[(（]\s*(H\d{6,})\s*[)）]")
ZEN = str.maketrans("０１２３４５６７８９．", "0123456789.")
NUMTAIL_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*(?:V|A|W|kg|ｋｇ|口|系統|系|ボルト|アンペア)?$", re.I)
NORMALIZED = []


def as_number(v, where, label):
    """『100V』『1系統』『１口』のような文字列を数値へ直す。直せないものは None を返す。"""
    if v is None or isinstance(v, (int, float)):
        return v
    t = str(v).strip().translate(ZEN)
    m = NUMTAIL_RE.match(t)
    if not m:
        return None
    n = float(m.group(1))
    n = int(n) if n == int(n) else n
    NORMALIZED.append(f"{where} の{label}『{v}』を {n} として読み替えた")
    return n


def norm(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def sheet_drawing_map(path):
    """シート名 -> drawing XML パス"""
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2target = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for sh in wb.find("main:sheets", NS):
        name = sh.get("name")
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rid2target[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        relpath = target.rsplit("/", 1)[0] + "/_rels/" + target.rsplit("/", 1)[1] + ".rels"
        if relpath not in z.namelist():
            continue
        srels = ET.fromstring(z.read(relpath))
        for r in srels:
            if r.get("Type", "").endswith("/drawing"):
                dpath = r.get("Target").replace("../", "xl/")
                out[name] = dpath
    return z, out


def shapes_of(z, dpath):
    """最上位アンカーのうちテキストを持つものを (from_row, to_row, text) で返す（行は0起点）"""
    root = ET.fromstring(z.read(dpath))
    res = []
    for anch in root:
        f = anch.find("xdr:from", NS)
        t = anch.find("xdr:to", NS)
        if f is None:
            continue
        fr = int(f.find("xdr:row", NS).text)
        tr = int(t.find("xdr:row", NS).text) if t is not None else fr
        texts = [e.text for e in anch.iter(f"{{{NS['a']}}}t") if e.text and e.text.strip()]
        if not texts:
            continue
        res.append((fr, tr, "".join(texts)))
    return res


def find_bands(ws):
    """U番号列(C)から 前面/背面 の行帯を検出する。返り値 [(face, top_row, bottom_row, u_top)]"""
    bands = []
    r = 1
    while r <= ws.max_row:
        if ws.cell(r, 3).value == 47:
            top = r
            rr = r
            while rr <= ws.max_row and isinstance(ws.cell(rr, 3).value, (int, float)):
                rr += 1
            bands.append((top, rr - 1))
            r = rr
        else:
            r += 1
    out = []
    for i, (top, bot) in enumerate(bands):
        out.append(("前面" if i == 0 else "背面", top, bot))
    return out


def resolve_columns(ws):
    """row1 の回路ラベルと row2 の見出しから列を解決する"""
    circuits = []
    for c in range(1, ws.max_column + 2):
        v = norm(ws.cell(1, c).value)
        m = CIRCUIT_RE.search(v)
        if m:
            circuits.append({
                "label": v,
                "volt": float(m.group(1)),
                "breaker_a": float(m.group(2)),
                "col_current": c,
                "col_ports": c + 1,
            })
    heads = {}
    for c in range(1, ws.max_column + 2):
        v = norm(ws.cell(2, c).value)
        if v:
            heads.setdefault(v, c)

    def pick(*names):
        for n in names:
            for k, c in heads.items():
                if k.startswith(n):
                    return c
        return None

    cols = {
        "mount_kit": pick("ﾏｳﾝﾄｷｯﾄ", "マウントキット"),
        "belt": pick("耐震"),
        "weight": pick("重量"),
        "volt": pick("電圧"),
        "systems": pick("系統"),
        "watt": pick("消費電力"),
        "system": pick("システム"),
        "usage": pick("用途"),
        "note": pick("備考"),
    }
    return circuits, cols


def extract(path):
    z, dmap = sheet_drawing_map(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path, data_only=False)
    racks = []
    for ws in wb.worksheets:
        wsf = wbf[ws.title]
        circuits, cols = resolve_columns(ws)
        bands = find_bands(ws)
        shapes = shapes_of(z, dmap[ws.title]) if ws.title in dmap else []
        # 機器名セル（D列）の縦結合 -> 占有行数
        span_of = {}
        for mg in ws.merged_cells.ranges:
            if mg.min_col <= 4 <= mg.max_col:
                span_of[mg.min_row] = mg.max_row - mg.min_row + 1
        rack = {
            "rack": norm(ws.cell(1, 1).value),
            "sheet": ws.title,
            "spec": norm(ws.cell(2, 1).value),
            "circuits": circuits,
            "bands": [{"face": f, "top": t, "bottom": b} for f, t, b in bands],
            "devices": [],
            "rows": [],
        }
        for face, top, bot in bands:
            u_of_row = {r: int(ws.cell(r, 3).value) for r in range(top, bot + 1)}
            # 図形 -> 機器
            for fr, tr, text in shapes:
                r_from = fr + 1
                r_to = max(tr, fr + 1)  # to は排他境界
                if not (top <= r_from <= bot):
                    continue
                u_hi = u_of_row.get(r_from)
                u_lo = u_of_row.get(min(r_to, bot), u_hi)
                if u_hi is None:
                    continue
                u_lo = min(u_lo, u_hi)
                asset = ASSET_RE.search(text)
                name = ASSET_RE.sub("", text).strip()
                rack["devices"].append({
                    "face": face,
                    "u_top": u_hi,
                    "u_bottom": u_lo,
                    "u_size": u_hi - u_lo + 1,
                    "text": text,
                    "name": name,
                    "asset_no": asset.group(1) if asset else "",
                    "row": r_from,
                    "source": "shape",
                })
            # 機器名がセル（D列）に書かれている行も機器として拾う
            shape_rows = {fr + 1 for fr, _, _ in shapes}
            for r in range(top, bot + 1):
                dtxt = norm(ws.cell(r, 4).value)
                if not dtxt or r in shape_rows:
                    continue
                n = span_of.get(r, 1)
                u_hi = u_of_row.get(r)
                if u_hi is None:
                    continue
                u_lo = u_of_row.get(min(r + n - 1, bot), u_hi)
                asset = ASSET_RE.search(dtxt)
                rack["devices"].append({
                    "face": face,
                    "u_top": u_hi,
                    "u_bottom": min(u_lo, u_hi),
                    "u_size": u_hi - min(u_lo, u_hi) + 1,
                    "text": dtxt,
                    "name": ASSET_RE.sub("", dtxt).strip(),
                    "asset_no": asset.group(1) if asset else "",
                    "row": r,
                    "source": "cell",
                })

            # 行属性
            for r in range(top, bot + 1):
                rec = {"face": face, "u": u_of_row[r], "row": r}
                any_val = False
                for key, c in cols.items():
                    if c is None:
                        continue
                    v = ws.cell(r, c).value
                    if key in ("weight", "volt", "systems", "watt"):
                        cell = f"{ws.title}!{get_column_letter(c)}{r}"
                        nv = as_number(v, cell, key)
                        rec[key] = nv if nv is not None else (norm(v) if isinstance(v, str) else v)
                    else:
                        rec[key] = norm(v) if isinstance(v, str) else v
                    if v not in (None, ""):
                        any_val = True
                rec["circuits"] = []
                for ci in circuits:
                    cur = ws.cell(r, ci["col_current"]).value
                    ports = ws.cell(r, ci["col_ports"]).value
                    cur_n = as_number(cur, f"{ws.title}!{get_column_letter(ci['col_current'])}{r}", "電流")
                    ports_n = as_number(ports, f"{ws.title}!{get_column_letter(ci['col_ports'])}{r}", "口数")
                    cur = cur_n if cur_n is not None else cur
                    ports = ports_n if ports_n is not None else ports
                    rec["circuits"].append({"label": ci["label"], "current": cur, "ports": ports})
                    if cur not in (None, "") or ports not in (None, ""):
                        any_val = True
                # 機器名セル（D列など）
                dtxt = norm(ws.cell(r, 4).value)
                rec["cell_text"] = dtxt
                if dtxt:
                    any_val = True
                if any_val:
                    rack["rows"].append(rec)
        racks.append(rack)
    return racks


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--out-json", required=True)
    ap.add_argument("--out-csv", required=True)
    a = ap.parse_args()
    racks = extract(a.xlsx)
    Path(a.out_json).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out_json).write_text(json.dumps(racks, ensure_ascii=False, indent=1), encoding="utf8")

    # 台帳CSV: 図形由来の機器に、同一行の属性を結合
    fields = ["ラック", "面", "開始U", "終了U", "U数", "機器名", "資産番号", "システム", "用途",
              "重量kg", "電圧V", "系統数", "回路", "電流A", "口数", "消費電力W", "マウントキット", "耐震ベルト", "備考"]
    with open(a.out_csv, "w", newline="", encoding="utf8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader()
        for rk in racks:
            rowmap = {(r["face"], r["row"]): r for r in rk["rows"]}
            for d in sorted(rk["devices"], key=lambda x: (x["face"], -x["u_top"])):
                attr = rowmap.get((d["face"], d["row"]), {})
                cir = [c for c in attr.get("circuits", []) if c["current"] not in (None, "")]
                w.writerow({
                    "ラック": rk["rack"], "面": d["face"], "開始U": d["u_bottom"], "終了U": d["u_top"],
                    "U数": d["u_size"], "機器名": d["name"], "資産番号": d["asset_no"],
                    "システム": attr.get("system") or "", "用途": attr.get("usage") or "",
                    "重量kg": attr.get("weight") or "", "電圧V": attr.get("volt") or "",
                    "系統数": attr.get("systems") or "",
                    "回路": " / ".join(c["label"] for c in cir),
                    "電流A": " / ".join(str(c["current"]) for c in cir),
                    "口数": " / ".join(str(c["ports"]) for c in cir),
                    "消費電力W": attr.get("watt") or "",
                    "マウントキット": attr.get("mount_kit") or "", "耐震ベルト": attr.get("belt") or "",
                    "備考": attr.get("note") or "",
                })
    print(f"抽出: ラック{len(racks)}本 機器{sum(len(r['devices']) for r in racks)}件", file=sys.stderr)
    if NORMALIZED:
        print(f"文字列を数値へ読み替えた箇所 {len(NORMALIZED)}件:", file=sys.stderr)
        for n in NORMALIZED:
            print("  " + n, file=sys.stderr)
    for rk in racks:
        print(f"  {rk['rack']}: 機器{len(rk['devices'])}件 電源/重量の記入行{len(rk['rows'])}行 "
              f"回路{[c['label'] for c in rk['circuits']]}", file=sys.stderr)


if __name__ == "__main__":
    main()
