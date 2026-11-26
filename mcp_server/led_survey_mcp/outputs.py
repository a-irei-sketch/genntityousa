# -*- coding: utf-8 -*-
"""帳票の作成：器具台帳Excel・写真帳PDF・簡易配置図(DXF/PDF)・網羅チェック・報告地点GeoJSON。
すべてPC側で作成する（ArcGISのレポート機能は使わない＝クレジット不要）。"""
from __future__ import annotations

import csv
import json
import math
import os
from collections import Counter, defaultdict
from pathlib import Path

from .data import CHOICES, facility_master, master_dir, media_for

_HERE = Path(__file__).parent


# ---------------------------------------------------------------- マスタ
def load_fixture_master() -> dict:
    p = master_dir() / "fixture_types.csv"
    with open(p, encoding="utf-8-sig") as fp:
        return {r["name"]: r for r in csv.DictReader(fp)}


def load_led_rules() -> dict:
    p = master_dir() / "led_rules.csv"
    with open(p, encoding="utf-8-sig") as fp:
        return {r["key"]: r for r in csv.DictReader(fp)}


def label(list_name: str, code):
    if code is None:
        return ""
    lst = CHOICES.get(list_name, {})
    if code in lst:
        return lst[code]
    parts = str(code).replace(",", " ").split()  # select_multiple は空白区切り
    if len(parts) > 1 or (parts and parts[0] in lst):
        return "、".join(lst.get(x, x) for x in parts)
    return code


def _num(v, default=None):
    try:
        if v is None or v == "":
            return default
        f = float(v)
        return default if math.isnan(f) else f
    except (TypeError, ValueError):
        return default


def _floor_order(f):
    order = [c for c in CHOICES["floors"]]
    return order.index(f) if f in order else 999


def rooms(ds) -> list[dict]:
    rs = [r for r in ds["main"] if r.get("rec_type") == "room"]
    return sorted(rs, key=lambda r: (r.get("building") or "", _floor_order(r.get("floor")), str(r.get("room_no") or "")))


def fixtures_by_room(ds) -> dict:
    d = defaultdict(list)
    for f in ds["fixtures"]:
        d[f.get("parentglobalid")].append(f)
    return d


def enrich_fixture(f: dict, fm: dict, rules: dict) -> dict:
    """フォーム側で計算済みの値を優先し、欠けていればマスタから補う。"""
    t, m = f.get("fx_type"), f.get("fx_mount")
    x = dict(f)
    qty = _num(f.get("fx_qty"))
    if qty is None:
        qty = (_num(f.get("fx_cols"), 0) * _num(f.get("fx_rows"), 0)) if f.get("fx_method") == "array" else _num(f.get("fx_qty_in"), 0)
    x["qty"] = int(qty or 0)
    x["type_label"] = fm.get(t, {}).get("label", t or "")
    x["shape"] = fm.get(t, {}).get("shape", "line")
    x["w_old"] = _num(f.get("fx_w_old"), _num(fm.get(t, {}).get("watt")))
    rule = rules.get(f"{t}|{m}", {})
    x["led"] = f.get("fx_led") or rule.get("led_code", "")
    x["w_led"] = _num(f.get("fx_w_led"), _num(rule.get("led_watt")))
    return x


# ---------------------------------------------------------------- 網羅チェック
def coverage(ds) -> list[dict]:
    reg = defaultdict(lambda: {"rooms": 0, "done": 0, "check": 0, "check_list": []})
    for r in rooms(ds):
        k = (r.get("building") or "", r.get("floor") or "")
        reg[k]["rooms"] += 1
        st = r.get("room_status")
        reg[k]["done" if st == "done" else "check"] += 1
        if st != "done":
            reg[k]["check_list"].append(r.get("room_no"))
    fl = {}
    for r in ds["main"]:
        if r.get("rec_type") == "floor":
            fl[(r.get("building") or "", r.get("floor") or "")] = r
    out = []
    for k in sorted(set(reg) | set(fl), key=lambda k: (k[0], _floor_order(k[1]))):
        if k[1] == "OUT" and k not in fl:  # 屋外は扉の数による漏れチェックの対象外
            continue
        g, f = reg[k], fl.get(k, {})
        doors, declared = _num(f.get("door_count")), _num(f.get("declared_rooms"))
        if not f:
            verdict = "階の確認が未記録"
        elif doors is not None and doors > g["rooms"]:
            verdict = f"部屋が{int(doors - g['rooms'])}室不足の可能性"
        elif declared is not None and declared > g["rooms"]:
            verdict = f"申告より{int(declared - g['rooms'])}室少ない"
        else:
            verdict = "OK"
        out.append({"building": k[0], "floor": k[1], "registered": g["rooms"], "done": g["done"],
                    "check": g["check"], "check_rooms": g["check_list"],
                    "door_count": doors, "declared": declared, "verdict": verdict})
    return out


def summary(ds) -> dict:
    fm, rules = load_fixture_master(), load_led_rules()
    rm = rooms(ds)
    room_idx = {r["globalid"]: r for r in rm}
    qty = Counter()
    kwh_old = kwh_new = 0.0
    for f in ds["fixtures"]:
        x = enrich_fixture(f, fm, rules)
        r = room_idx.get(f.get("parentglobalid"))
        if not r:
            continue
        qty[x["type_label"]] += x["qty"]
        h = _num(r.get("annual_h"), _num(r.get("hours_day"), 0) * _num(r.get("days_year"), 0))
        kwh_old += (x["w_old"] or 0) * x["qty"] * h / 1000
        kwh_new += (x["w_led"] or 0) * x["qty"] * h / 1000
    reps = ds["reports"]
    return {
        "source": ds["meta"].get("source"),
        "buildings": sorted({r.get("building") for r in rm if r.get("building")}),
        "rooms": len(rm),
        "rooms_need_check": [r.get("room_no") for r in rm if r.get("room_status") != "done"],
        "fixtures_total": sum(qty.values()),
        "fixtures_by_type": dict(qty.most_common()),
        "kwh_existing": round(kwh_old, 1), "kwh_led": round(kwh_new, 1),
        "kwh_saving": round(kwh_old - kwh_new, 1),
        "reports": len(reps),
        "reports_action": sum(1 for r in reps if r.get("rp_cat") == "action"),
        "photos": len(ds["media"]),
        "coverage_alerts": [c for c in coverage(ds) if c["verdict"] != "OK"],
    }


# ---------------------------------------------------------------- 器具台帳 Excel
def build_ledger(ds, out_path: Path, unit_price: float = 30.0, co2_factor: float = 0.434) -> Path:
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    fm, rules = load_fixture_master(), load_led_rules()
    F = "Arial"
    head = PatternFill("solid", fgColor="1F3864")
    thin = Side(style="thin", color="BFBFBF")
    wb = Workbook()

    def header(ws, cols, widths):
        ws.append(cols)
        for i, c in enumerate(ws[1], 1):
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = head
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    def finish(ws, fmts=None):
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name=F)
                c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        for col, fmt in (fmts or {}).items():
            for c in ws[col][1:]:
                c.number_format = fmt
        ws.auto_filter.ref = ws.dimensions

    # 条件
    wc = wb.active
    wc.title = "条件"
    wc["A1"] = "LED現況調査 算定条件"
    wc["A1"].font = Font(name=F, bold=True, size=14)
    wc["A3"], wc["B3"] = "電力量単価（円/kWh）", unit_price
    wc["A4"], wc["B4"] = "CO2排出係数（kg-CO2/kWh）", co2_factor
    for c in ("B3", "B4"):
        wc[c].font = Font(name=F, color="0000FF")
        wc[c].fill = PatternFill("solid", fgColor="FFFF00")
    wc["B3"].comment = Comment("仮値。契約電力の実単価に置き換えてください（黄色セルは入力欄）", "LED調査MCP")
    wc["B4"].comment = Comment("仮値。契約電力会社の公表係数（環境省 電気事業者別排出係数）に置き換えてください", "LED調査MCP")
    wc["A6"] = "黄色のセルを書き換えると、器具台帳・省エネ算定の金額とCO2が再計算されます。"
    wc["A7"] = f"データ取得元：{ds['meta'].get('source')}（{ds['meta'].get('loaded_at')}）"
    wc["A8"] = "既設・LEDの消費電力（W/台）は器具マスタ（fixture_types.csv / led_rules.csv）の値です。マスタが仮値の間は算定結果も参考値です。"
    for c in ("A3", "A4", "A6", "A7", "A8"):
        wc[c].font = Font(name=F)
    wc.column_dimensions["A"].width = 34
    wc.column_dimensions["B"].width = 14

    # 部屋一覧
    wr = wb.create_sheet("部屋一覧")
    header(wr, ["部屋キー", "施設名", "階", "通し番号", "部屋名", "用途", "縦(m)", "横(m)", "面積(m²)", "天井高(m)",
                "1日点灯(h)", "年間日数", "年間点灯(h)", "器具台数", "非常灯・誘導灯", "照度(lx)", "調査状態",
                "調査者", "調査日時", "写真枚数", "照明状況", "施設利用上・更新上の課題"],
           [16, 22, 6, 10, 18, 12, 8, 8, 9, 9, 9, 8, 10, 9, 10, 8, 9, 10, 17, 8, 10, 40])
    rm = rooms(ds)
    for i, r in enumerate(rm, 2):
        key = f"{r.get('building')}|{r.get('room_no')}"
        use = label("uses", r.get("room_use"))
        if r.get("room_use") == "other" and r.get("room_use_other"):
            use = f"その他（{r['room_use_other']}）"
        wr.append([key, r.get("building"), r.get("floor"), r.get("room_no"), r.get("room_name"), use,
                   _num(r.get("depth_m")), _num(r.get("width_m")),
                   f'=IF(AND(G{i}<>"",H{i}<>""),G{i}*H{i},"")', _num(r.get("ceil_h_m")),
                   _num(r.get("hours_day")), _num(r.get("days_year")), f"=K{i}*L{i}",
                   f"=SUMIFS(器具台帳!$J:$J,器具台帳!$A:$A,A{i})",
                   label("yn3", r.get("emergency")), _num(r.get("lux")), label("statuses", r.get("room_status")),
                   r.get("reporter"), (r.get("survey_dt") or "")[:16].replace("T", " "),
                   len([m for m in ds["media"] if m["parent"] == r.get("globalid")]),
                   label("light_levels", r.get("room_light")), r.get("room_issue")])
    finish(wr, {"I": "0.0", "M": "#,##0"})

    # 器具台帳
    wf = wb.create_sheet("器具台帳")
    header(wf, ["部屋キー", "施設名", "階", "通し番号", "部屋名", "器具コード", "器具タイプ", "数え方", "取付方式",
                "台数", "点灯状況", "メーカー・型番", "既設(W/台)", "LED候補品番", "LED(W/台)", "年間点灯(h)",
                "既設(kWh/年)", "LED(kWh/年)", "削減(kWh/年)", "削減額(円/年)", "削減CO2(t/年)", "特記",
                "形状", "ランプ区分", "ランプ仕様", "銘板の消費電力(W)", "設置高さ(m)", "高所作業", "撤去条件", "撤去条件の補足",
                "老朽度・劣化", "劣化の内容"],
           [16, 22, 6, 10, 18, 12, 26, 10, 8, 7, 9, 18, 9, 20, 9, 10, 11, 11, 11, 12, 11, 24,
            22, 16, 26, 10, 9, 14, 28, 20, 10, 24])
    room_idx = {r["globalid"]: r for r in rm}
    i = 2
    for r in rm:
        for f in fixtures_by_room(ds).get(r["globalid"], []):
            x = enrich_fixture(f, fm, rules)
            method = label("count_methods", f.get("fx_method"))
            if f.get("fx_method") == "array":
                method = f"配列 {int(_num(f.get('fx_cols'), 0))}×{int(_num(f.get('fx_rows'), 0))}"
            wf.append([f"{r.get('building')}|{r.get('room_no')}", r.get("building"), r.get("floor"), r.get("room_no"),
                       r.get("room_name"), f.get("fx_type"), x["type_label"], method, label("mounts", f.get("fx_mount")),
                       x["qty"], label("lit_states", f.get("fx_lit")), f.get("fx_model"), x["w_old"], x["led"], x["w_led"],
                       f"=IFERROR(INDEX(部屋一覧!$M:$M,MATCH(A{i},部屋一覧!$A:$A,0)),0)",
                       f"=M{i}*J{i}*P{i}/1000", f"=O{i}*J{i}*P{i}/1000", f"=Q{i}-R{i}",
                       f"=S{i}*条件!$B$3", f"=S{i}*条件!$B$4/1000", f.get("fx_note"),
                       label("shapes", f.get("fx_shape")), label("lamp_kinds", f.get("fx_lamp_kind")), f.get("fx_lamp_spec"),
                       _num(f.get("fx_w_field")), _num(f.get("fx_height_m")), label("aerial", f.get("fx_aerial")),
                       label("removal", f.get("fx_removal")), f.get("fx_removal_note"),
                       label("conditions", f.get("fx_cond")), label("defects", f.get("fx_defects"))])
            i += 1
    last_fx = i - 1
    finish(wf, {"P": "#,##0", "Q": "#,##0.0", "R": "#,##0.0", "S": "#,##0.0", "T": "#,##0", "U": "0.000"})

    # 数量集計
    wq = wb.create_sheet("数量集計")
    header(wq, ["器具コード", "器具タイプ", "取付方式", "台数", "LED候補品番", "削減(kWh/年)"], [12, 28, 10, 8, 20, 12])
    combos = sorted({(f.get("fx_type"), label("mounts", f.get("fx_mount"))) for f in ds["fixtures"]
                     if f.get("parentglobalid") in room_idx}, key=lambda c: (str(c[0]), str(c[1])))
    for j, (t, m) in enumerate(combos, 2):
        code = next((k for k, v in CHOICES["mounts"].items() if v == m), m)
        x = enrich_fixture({"fx_type": t, "fx_mount": code}, fm, rules)
        wq.append([t, x["type_label"], m, f"=SUMIFS(器具台帳!$J:$J,器具台帳!$F:$F,A{j},器具台帳!$I:$I,C{j})",
                   x["led"], f"=SUMIFS(器具台帳!$S:$S,器具台帳!$F:$F,A{j},器具台帳!$I:$I,C{j})"])
    n = len(combos) + 1
    wq.append(["合計", "", "", f"=SUM(D2:D{n})", "", f"=SUM(F2:F{n})"])
    for c in wq[n + 1]:
        c.font = Font(name=F, bold=True)
    wq.append([])
    start = n + 3
    wq.append(["施設名", "階", "", "台数", "", "削減(kWh/年)"])
    for c in wq[start]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = head
    floors = sorted({(r.get("building"), r.get("floor")) for r in rm}, key=lambda k: (str(k[0]), _floor_order(k[1])))
    for j, (b, fl) in enumerate(floors, start + 1):
        wq.append([b, fl, "", f"=SUMIFS(器具台帳!$J:$J,器具台帳!$B:$B,A{j},器具台帳!$C:$C,B{j})", "",
                   f"=SUMIFS(器具台帳!$S:$S,器具台帳!$B:$B,A{j},器具台帳!$C:$C,B{j})"])
    finish(wq, {"F": "#,##0.0"})

    # 省エネ算定
    we = wb.create_sheet("省エネ算定")
    header(we, ["項目", "値", "単位"], [30, 16, 12])
    rows = [("既設 年間消費電力量", "=SUM(器具台帳!Q:Q)", "kWh/年", "#,##0"),
            ("LED化後 年間消費電力量", "=SUM(器具台帳!R:R)", "kWh/年", "#,##0"),
            ("年間削減電力量", "=B2-B3", "kWh/年", "#,##0"),
            ("削減率", '=IF(B2>0,B4/B2,"")', "", "0.0%"),
            ("年間削減額", "=B4*条件!B3", "円/年", "#,##0"),
            ("年間CO2削減量", "=B4*条件!B4/1000", "t-CO2/年", "0.00"),
            ("対象器具台数", "=SUM(器具台帳!J:J)", "台", "#,##0")]
    for k, (a, b, c, fmt) in enumerate(rows, 2):
        we.append([a, b, c])
        we[f"B{k}"].number_format = fmt
    finish(we)

    # 網羅チェック
    wk = wb.create_sheet("網羅チェック")
    header(wk, ["施設名", "階", "登録部屋数", "調査済", "要確認", "扉の数", "申告部屋数", "差（扉−登録）", "判定", "要確認の部屋"],
           [14, 6, 10, 8, 8, 8, 10, 12, 22, 30])
    for j, c in enumerate(coverage(ds), 2):
        wk.append([c["building"], c["floor"],
                   f"=COUNTIFS(部屋一覧!$B:$B,A{j},部屋一覧!$C:$C,B{j})",
                   f'=COUNTIFS(部屋一覧!$B:$B,A{j},部屋一覧!$C:$C,B{j},部屋一覧!$Q:$Q,"調査済")',
                   f"=C{j}-D{j}", c["door_count"], c["declared"],
                   f'=IF(F{j}="","",F{j}-C{j})',
                   f'=IF(AND(F{j}="",G{j}=""),"階の確認が未記録",IF(N(F{j})>C{j},"部屋が不足の可能性",IF(N(G{j})>C{j},"申告より少ない","OK")))',
                   "、".join(str(x) for x in c["check_rooms"] if x)])
    finish(wk)

    # 施設一覧（依頼対象施設ごとの調査状況）
    wfac = wb.create_sheet("施設一覧", 1)
    header(wfac, ["コード", "施設名", "延床面積(m²)", "屋内・屋外", "登録部屋数", "要確認の部屋", "器具台数",
                  "既設(kWh/年)", "削減(kWh/年)", "報告事項", "調査状況"], [7, 30, 11, 9, 10, 11, 9, 12, 12, 9, 12])
    fac = {k: v for k, v in facility_master().items() if k != "other"}
    names = [v["label"] for v in fac.values()]
    extra = sorted({r.get("building") for r in rm if r.get("building") and r.get("building") not in names})
    rows_fac = [(k, v["label"], _num(v.get("floor_area_m2")), v.get("kind")) for k, v in fac.items()] + \
               [("", b, None, "") for b in extra]
    by_gid = {r.get("globalid"): r for r in ds["main"]}
    rep_count = Counter(by_gid.get(rp.get("parentglobalid"), {}).get("building") for rp in ds["reports"])
    for j, (code, name, area, kind) in enumerate(rows_fac, 2):
        wfac.append([code, name, area, kind, f"=COUNTIFS(部屋一覧!$B:$B,B{j})",
                     f'=COUNTIFS(部屋一覧!$B:$B,B{j},部屋一覧!$Q:$Q,"要確認")',
                     f"=SUMIFS(器具台帳!$J:$J,器具台帳!$B:$B,B{j})", f"=SUMIFS(器具台帳!$Q:$Q,器具台帳!$B:$B,B{j})",
                     f"=SUMIFS(器具台帳!$S:$S,器具台帳!$B:$B,B{j})", rep_count.get(name, 0),
                     f'=IF(E{j}=0,"未調査",IF(F{j}>0,"要確認あり","調査済"))'])
    n2 = len(rows_fac) + 1
    wfac.append(["", "合計", f"=SUM(C2:C{n2})", "", f"=SUM(E2:E{n2})", f"=SUM(F2:F{n2})", f"=SUM(G2:G{n2})",
                 f"=SUM(H2:H{n2})", f"=SUM(I2:I{n2})", f"=SUM(J2:J{n2})", ""])
    finish(wfac, {"C": "#,##0", "H": "#,##0", "I": "#,##0"})
    for c in wfac[n2 + 1]:
        c.font = Font(name=F, bold=True)

    # 報告事項
    wp = wb.create_sheet("報告事項")
    header(wp, ["区分", "施設名", "階", "部屋・対象地", "現場メモ", "コメント", "緯度", "経度", "写真", "音声", "報告者", "日時"],
           [10, 14, 6, 18, 50, 30, 11, 11, 30, 20, 10, 17])
    main_idx = {r.get("globalid"): r for r in ds["main"]}
    for rp in ds["reports"]:
        p = main_idx.get(rp.get("parentglobalid"), {})
        place = p.get("site_name") or " ".join(str(x) for x in (p.get("room_no"), p.get("room_name")) if x)
        photos = [Path(x).name for x in media_for(ds, rp.get("globalid"), "rp_photo")]
        audio = [Path(x).name for x in media_for(ds, rp.get("globalid"), "rp_audio")]
        wp.append([label("rp_cats", rp.get("rp_cat")), p.get("building"), p.get("floor"), place, rp.get("rp_memo"),
                   rp.get("rp_comment"), rp.get("lat"), rp.get("lon"), "\n".join(photos), "\n".join(audio),
                   p.get("reporter"), (p.get("survey_dt") or "")[:16].replace("T", " ")])
    finish(wp, {"G": "0.000000", "H": "0.000000"})
    for row in wp.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------- 写真帳 PDF
_FONT = None


_FONT_CANDIDATES = [  # (パス, ttcの番号) 埋め込み可能なTrueType系の日本語フォント
    (r"C:\Windows\Fonts\msgothic.ttc", 0), (r"C:\Windows\Fonts\meiryo.ttc", 0),
    (r"C:\Windows\Fonts\YuGothM.ttc", 0), ("/Library/Fonts/Osaka.ttf", 0),
    ("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", 0), ("/usr/share/fonts/truetype/fonts-japanese-gothic.ttf", 0),
]


def _jp_font():
    """日本語フォントをPDFに埋め込む（どのPCで開いても文字化けしない）。見つからなければCIDフォント。"""
    global _FONT
    if _FONT is None:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        env = os.environ.get("LED_SURVEY_FONT")
        cands = ([(env, 0)] if env else []) + _FONT_CANDIDATES
        for path, idx in cands:
            if path and os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont("JP", path, subfontIndex=idx))
                    _FONT = "JP"
                    break
                except Exception:
                    continue
        if _FONT is None:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
            _FONT = "HeiseiKakuGo-W5"
    return _FONT


def _image_reader(path, max_px=1400):
    from PIL import Image, ImageOps
    from reportlab.lib.utils import ImageReader
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im.thumbnail((max_px, max_px))
    return ImageReader(im), im.size


def _wrap(text, width_chars):
    out = []
    for para in str(text or "").splitlines() or [""]:
        while len(para) > width_chars:
            out.append(para[:width_chars])
            para = para[width_chars:]
        out.append(para)
    return out


def build_photo_album(ds, out_path: Path, title: str = "LED現況調査 写真帳") -> Path:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    fm, rules = load_fixture_master(), load_led_rules()
    font = _jp_font()
    W, H = A4
    c = canvas.Canvas(str(out_path), pagesize=A4)
    c.setTitle(title)
    margin = 15 * mm
    cell_w, cell_h = (W - 2 * margin - 6 * mm) / 2, 72 * mm
    page = [0]

    def new_page(head):
        if page[0]:
            c.showPage()
        page[0] += 1
        c.setFont(font, 9)
        c.drawString(margin, H - 10 * mm, title)
        c.drawRightString(W - margin, H - 10 * mm, f"{page[0]}")
        c.setFont(font, 13)
        c.drawString(margin, H - 18 * mm, head)
        c.line(margin, H - 20 * mm, W - margin, H - 20 * mm)
        return H - 24 * mm

    def grid(items, head, info_lines):
        y = new_page(head)
        c.setFont(font, 9)
        for ln in [x for l0 in info_lines for x in _wrap(l0, 60)][:14]:
            c.drawString(margin, y, ln)
            y -= 5 * mm
        y -= 2 * mm
        col = 0
        for path, cap in items:
            if y - cell_h < 15 * mm:
                y = new_page(head + "（続き）")
                col = 0
            x = margin + col * (cell_w + 6 * mm)
            try:
                img, (iw, ih) = _image_reader(path)
                s = min(cell_w / iw, (cell_h - 10 * mm) / ih)
                c.drawImage(img, x + (cell_w - iw * s) / 2, y - (cell_h - 10 * mm) + ((cell_h - 10 * mm) - ih * s) / 2,
                            iw * s, ih * s)
            except Exception as e:  # 画像が壊れていても帳票は止めない
                c.rect(x, y - (cell_h - 10 * mm), cell_w, cell_h - 10 * mm)
                c.drawString(x + 3 * mm, y - 10 * mm, f"画像を開けません: {Path(path).name} ({e.__class__.__name__})")
            c.setFont(font, 8.5)
            for k, ln in enumerate(_wrap(cap, 42)[:2]):
                c.drawString(x, y - cell_h + 6 * mm - k * 4 * mm, ln)
            c.setFont(font, 9)
            col += 1
            if col == 2:
                col = 0
                y -= cell_h + 2 * mm
        if not items:
            c.drawString(margin, y, "（写真なし）")

    qlabel = {"room_photo": "室名札・入口", "ceil_photo": "天井見上げ", "ceil_photo2": "天井見上げ(2)",
              "fx_photo": "器具", "fx_plate": "銘板", "rp_photo": "報告写真", "floor_photo": "案内図",
              "floor_sketch": "簡易スケッチ"}
    fx_room = fixtures_by_room(ds)
    for r in rooms(ds):
        ql = dict(qlabel, **({"room_photo": "全景", "ceil_photo": "照明全景", "ceil_photo2": "照明全景(2)"}
                             if r.get("room_use") == "outdoor" else {}))
        items = [(p, f"{ql.get(q)}　{r.get('room_no')} {r.get('room_name') or ''}")
                 for q in ("room_photo", "ceil_photo", "ceil_photo2") for p in media_for(ds, r["globalid"], q)]
        fx_lines = []
        for f in fx_room.get(r["globalid"], []):
            x = enrich_fixture(f, fm, rules)
            fx_lines.append(f"・{x['type_label']}／{label('mounts', f.get('fx_mount'))}／{x['qty']}台／"
                            f"{label('lit_states', f.get('fx_lit'))}／{label('conditions', f.get('fx_cond'))}／"
                            f"高さ{f.get('fx_height_m') or '-'}m・{label('aerial', f.get('fx_aerial'))}／LED候補 {x['led']}")
            for q in ("fx_photo", "fx_plate"):
                for p in media_for(ds, f.get("globalid"), q):
                    items.append((p, f"{qlabel[q]}　{x['type_label']}（{x['qty']}台）"))
        for rp in ds["reports"]:
            if rp.get("parentglobalid") == r["globalid"]:
                for p in media_for(ds, rp.get("globalid"), "rp_photo"):
                    items.append((p, f"{label('rp_cats', rp.get('rp_cat'))}　{rp.get('rp_memo') or ''}"))
        dims = f"{_num(r.get('depth_m'), 0):g}m×{_num(r.get('width_m'), 0):g}m" if r.get("depth_m") else "寸法未入力"
        info = [f"用途：{label('uses', r.get('room_use'))}　寸法：{dims}　天井高：{r.get('ceil_h_m') or '-'}m　"
                f"点灯：{r.get('hours_day') or '-'}h×{r.get('days_year') or '-'}日　状態：{label('statuses', r.get('room_status'))}"] + fx_lines
        grid(items, f"{r.get('building')}　{r.get('floor')}　{r.get('room_no')}　{r.get('room_name') or ''}", info)

    # 現場写真報告・階の確認（部屋以外の記録）
    for m in ds["main"]:
        if m.get("rec_type") == "room":
            continue
        items = []
        for q in ("floor_photo", "floor_sketch"):
            items += [(p, qlabel[q]) for p in media_for(ds, m["globalid"], q)]
        lines = []
        if m.get("rec_type") == "floor":
            lines.append(f"扉の数：{m.get('door_count') or '-'}　申告部屋数：{m.get('declared_rooms') or '-'}　{m.get('floor_note') or ''}")
        for rp in ds["reports"]:
            if rp.get("parentglobalid") == m.get("globalid"):
                loc = f"（{rp['lat']:.6f}, {rp['lon']:.6f}）" if rp.get("lat") else ""
                lines.append(f"・[{label('rp_cats', rp.get('rp_cat'))}] {rp.get('rp_memo') or ''} {loc}")
                if rp.get("rp_comment"):
                    lines.append(f"　コメント：{rp['rp_comment']}")
                for p in media_for(ds, rp.get("globalid"), "rp_photo"):
                    items.append((p, f"{label('rp_cats', rp.get('rp_cat'))}　{rp.get('rp_memo') or ''}"))
        head = m.get("site_name") or f"{m.get('building')}　{m.get('floor')}　階の確認"
        grid(items, head, [ln for l0 in lines for ln in _wrap(l0, 60)][:12])
    c.save()
    return out_path


# ---------------------------------------------------------------- 簡易配置図（DXF + PDF）
_SYMBOL = {"line": (1.25, 0.18), "round": (0.15, 0.15), "high": (0.3, 0.3)}


def room_layout(r: dict, fxs: list[dict]) -> dict:
    """部屋の寸法と器具の配列から、器具の概略位置（m）を割り付ける。"""
    D = _num(r.get("depth_m")) or 0
    Wd = _num(r.get("width_m")) or 0
    assumed = not (D and Wd)
    if assumed:
        D = Wd = 6.0
    groups = [f for f in fxs if f["qty"] > 0]
    total = sum(f["qty"] for f in groups) or 1
    placed, y0 = [], 0.0
    for f in groups:
        band = D * f["qty"] / total
        if f.get("fx_method") == "array" and _num(f.get("fx_cols")) and _num(f.get("fx_rows")):
            cols, rws = int(_num(f["fx_cols"])), int(_num(f["fx_rows"]))
        else:
            cols = max(1, round(math.sqrt(f["qty"] * Wd / band)))
            cols = min(cols, f["qty"])
            rws = math.ceil(f["qty"] / cols)
        n = 0
        for j in range(rws):
            for k in range(cols):
                if n >= f["qty"]:
                    break
                placed.append({"x": Wd * (k + 0.5) / cols, "y": y0 + band * (j + 0.5) / rws,
                               "shape": f["shape"], "code": f.get("fx_type")})
                n += 1
        y0 += band
    return {"W": Wd, "D": D, "assumed": assumed, "items": placed}


def build_layouts(ds, out_dir: Path) -> list[Path]:
    import ezdxf
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    fm, rules = load_fixture_master(), load_led_rules()
    out_dir.mkdir(parents=True, exist_ok=True)
    fx_room = fixtures_by_room(ds)
    by_bldg = defaultdict(lambda: defaultdict(list))
    for r in rooms(ds):
        fxs = [enrich_fixture(f, fm, rules) for f in fx_room.get(r["globalid"], [])]
        by_bldg[r.get("building") or "建物"][r.get("floor") or "-"].append((r, fxs, room_layout(r, fxs)))
    paths = []
    font = _jp_font()
    for bldg, floors in by_bldg.items():
        safe = "".join(ch if ch not in '\\/:*?"<>|' else "_" for ch in bldg)
        # DXF（1建物1ファイル、階ごとに横一列、階は上へ積む）
        doc = ezdxf.new("R2010", setup=True)
        doc.styles.add("JP", font="msgothic.ttc")
        for name, color in (("ROOM", 7), ("FIXTURE", 1), ("TEXT", 3), ("NOTE", 8)):
            doc.layers.add(name, color=color)
        msp = doc.modelspace()
        y_base = 0.0
        for fl in sorted(floors, key=_floor_order):
            items = floors[fl]
            x = 0.0
            hmax = max(l["D"] for _, _, l in items)
            msp.add_text(f"{fl}", height=1.0, dxfattribs={"layer": "TEXT", "style": "JP"}).set_placement((-6, y_base + hmax / 2))
            for r, fxs, L in items:
                msp.add_lwpolyline([(x, y_base), (x + L["W"], y_base), (x + L["W"], y_base + L["D"]), (x, y_base + L["D"])],
                                   close=True, dxfattribs={"layer": "ROOM"})
                for it in L["items"]:
                    w, h = _SYMBOL.get(it["shape"], _SYMBOL["line"])
                    cx, cy = x + it["x"], y_base + it["y"]
                    if it["shape"] == "line":
                        msp.add_lwpolyline([(cx - w / 2, cy - h / 2), (cx + w / 2, cy - h / 2), (cx + w / 2, cy + h / 2),
                                            (cx - w / 2, cy + h / 2)], close=True, dxfattribs={"layer": "FIXTURE"})
                    else:
                        msp.add_circle((cx, cy), w, dxfattribs={"layer": "FIXTURE"})
                msp.add_text(f"{r.get('room_no')} {r.get('room_name') or ''}", height=0.35,
                             dxfattribs={"layer": "TEXT", "style": "JP"}).set_placement((x + 0.2, y_base + L["D"] + 0.3))
                ty = y_base - 0.6
                for f in fxs:
                    msp.add_text(f"{f['type_label']} {f['qty']}台", height=0.25,
                                 dxfattribs={"layer": "TEXT", "style": "JP"}).set_placement((x + 0.2, ty))
                    ty -= 0.4
                if L["assumed"]:
                    msp.add_text("寸法未入力（6m角で仮表示）", height=0.25,
                                 dxfattribs={"layer": "NOTE", "style": "JP"}).set_placement((x + 0.2, ty))
                x += L["W"] + 3.0
            y_base += hmax + 4.0 + 0.4 * max(len(f) for _, f, _ in items)
        p = out_dir / f"簡易配置図_{safe}.dxf"
        doc.saveas(p)
        paths.append(p)

        # PDF（階ごとに1ページ、A3横に縮尺を合わせて並べる）
        pp = out_dir / f"簡易配置図_{safe}.pdf"
        c = canvas.Canvas(str(pp), pagesize=landscape(A3))
        PW, PH = landscape(A3)
        for fl in sorted(floors, key=_floor_order):
            items = floors[fl]
            m = 15 * mm
            # 部屋を左から並べ、幅が足りなければ折り返す。全体が収まる縮尺を求める
            def pack(scale):
                pos, x, y, rowh = [], m, PH - 30 * mm, 0
                for r, fxs, L in items:
                    w, h = L["W"] * scale, L["D"] * scale
                    if x + w > PW - m and x > m:
                        x, y = m, y - rowh - 22 * mm
                        rowh = 0
                    pos.append((x, y - h))
                    x += w + 10 * mm
                    rowh = max(rowh, h + 4 * mm * len(fxs))
                return pos, y - rowh
            scale = 20 * mm
            while scale > 1 * mm:
                pos, bottom = pack(scale)
                if bottom > m:
                    break
                scale *= 0.85
            c.setFont(font, 14)
            c.drawString(m, PH - 15 * mm, f"簡易配置図　{bldg}　{fl}")
            c.setFont(font, 8)
            c.drawString(m, PH - 21 * mm, f"縮尺：1m = {scale / mm:.1f}mm（図面ではなく、寸法と器具の配列から作図した概略です）")
            for (r, fxs, L), (x0, y0) in zip(items, pos):
                c.setLineWidth(0.8)
                c.rect(x0, y0, L["W"] * scale, L["D"] * scale)
                c.setLineWidth(0.4)
                for it in L["items"]:
                    w, h = _SYMBOL.get(it["shape"], _SYMBOL["line"])
                    cx, cy = x0 + it["x"] * scale, y0 + (L["D"] - it["y"]) * scale
                    if it["shape"] == "line":
                        c.rect(cx - w * scale / 2, cy - max(h * scale, 1.2) / 2, w * scale, max(h * scale, 1.2), fill=0)
                    else:
                        c.circle(cx, cy, max(w * scale, 1.2))
                c.setFont(font, 8)
                c.drawString(x0, y0 + L["D"] * scale + 2 * mm, f"{r.get('room_no')} {r.get('room_name') or ''}"
                             + ("（寸法未入力）" if L["assumed"] else f"　横{L['W']:g}×縦{L['D']:g}m"))
                c.setFont(font, 7)
                for k, f in enumerate(fxs):
                    c.drawString(x0, y0 - (k + 1) * 3.6 * mm, f"{f['type_label']} {f['qty']}台")
            c.showPage()
        c.save()
        paths.append(pp)
    return paths


# ---------------------------------------------------------------- 地図用 GeoJSON
def build_geojson(ds, out_path: Path) -> Path:
    feats = []
    main_idx = {r.get("globalid"): r for r in ds["main"]}
    for r in ds["main"]:
        if r.get("lat") is not None and r.get("lon") is not None:
            feats.append({"type": "Feature", "geometry": {"type": "Point", "coordinates": [r["lon"], r["lat"]]},
                          "properties": {"種類": label("rec_types", r.get("rec_type")), "建物": r.get("building"),
                                         "階": r.get("floor"), "部屋": r.get("room_no") or r.get("site_name"),
                                         "報告者": r.get("reporter"), "日時": r.get("survey_dt")}})
    for rp in ds["reports"]:
        if rp.get("lat") is not None and rp.get("lon") is not None:
            p = main_idx.get(rp.get("parentglobalid"), {})
            feats.append({"type": "Feature", "geometry": {"type": "Point", "coordinates": [rp["lon"], rp["lat"]]},
                          "properties": {"種類": "報告事項", "区分": label("rp_cats", rp.get("rp_cat")),
                                         "場所": p.get("site_name") or p.get("room_no"), "メモ": rp.get("rp_memo"),
                                         "コメント": rp.get("rp_comment"),
                                         "写真": [Path(x).name for x in media_for(ds, rp.get("globalid"), "rp_photo")]}})
    out_path.write_text(json.dumps({"type": "FeatureCollection", "features": feats}, ensure_ascii=False, indent=1),
                        encoding="utf-8")
    return out_path


# ---------------------------------------------------------------- 写真整理（施設名・撮影箇所・撮影日）
_QNAME = {"room_photo": "入口・室名札", "ceil_photo": "天井", "ceil_photo2": "天井2", "fx_photo": "器具",
          "fx_plate": "銘板", "rp_photo": "報告", "rp_photo2": "報告", "rp_photo3": "報告", "rp_audio": "音声", "floor_photo": "案内図", "floor_sketch": "スケッチ"}


def export_photos(ds, out_dir: Path) -> Path:
    """写真を施設ごとのフォルダに「施設名_場所_撮影箇所_撮影日_連番」の名前でコピーし、写真一覧.xlsxを作る。"""
    import re
    import shutil

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    fm, rules = load_fixture_master(), load_led_rules()
    owner = {}  # 添付の親 globalid → (記録, 撮影箇所の補足)
    for r in ds["main"]:
        owner[r.get("globalid")] = (r, "")
    for f in ds["fixtures"]:
        p = next((r for r in ds["main"] if r.get("globalid") == f.get("parentglobalid")), {})
        owner[f.get("globalid")] = (p, enrich_fixture(f, fm, rules)["type_label"])
    for rp in ds["reports"]:
        p = next((r for r in ds["main"] if r.get("globalid") == rp.get("parentglobalid")), {})
        owner[rp.get("globalid")] = (p, label("rp_cats", rp.get("rp_cat")))
    safe = lambda t: re.sub(r'[\\/:*?"<>|\s]+', "", str(t or ""))[:40]
    wb = Workbook()
    ws = wb.active
    ws.title = "写真一覧"
    ws.append(["施設名", "階", "部屋・場所", "撮影箇所", "補足", "撮影日", "ファイル名", "元ファイル"])
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    count = Counter()
    for m in sorted(ds["media"], key=lambda m: m["path"]):
        rec, extra = owner.get(m["parent"], ({}, ""))
        fac = rec.get("building") or "施設不明"
        place = rec.get("room_no") or rec.get("site_name") or ("階の確認" if rec.get("rec_type") == "floor" else "")
        if rec.get("room_name"):
            place = f"{place}{rec['room_name']}"
        day = (rec.get("survey_dt") or "")[:10].replace("-", "") or "日付不明"
        kind = _QNAME.get(m["question"], m["question"])
        if rec.get("room_use") == "outdoor" or rec.get("floor") == "OUT":
            kind = {"room_photo": "全景", "ceil_photo": "照明全景", "ceil_photo2": "照明全景2"}.get(m["question"], kind)
        stem = "_".join(x for x in (safe(fac), safe(rec.get("floor")), safe(place), safe(kind + extra), day) if x)
        count[stem] += 1
        name = f"{stem}_{count[stem]:02d}{Path(m['path']).suffix.lower()}"
        dest = out_dir / safe(fac) / name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(m["path"], dest)
        ws.append([fac, rec.get("floor"), place, kind, extra, day, name, Path(m["path"]).name])
    for col, w in zip("ABCDEFGH", (28, 6, 18, 12, 22, 10, 70, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    p = out_dir / "写真一覧.xlsx"
    wb.save(p)
    return p
