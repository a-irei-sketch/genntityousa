# -*- coding: utf-8 -*-
"""動作確認用：Survey123 の Excel 書き出しを模した合成データと写真フォルダを作る（南大東村の調査を想定）。"""
import sys
import uuid
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw

out = Path(sys.argv[1] if len(sys.argv) > 1 else "sample")
media = out / "media"
media.mkdir(parents=True, exist_ok=True)


def gid():
    return str(uuid.uuid4())


def photo(parent, q, n, text, color):
    im = Image.new("RGB", (1200, 900), color)
    d = ImageDraw.Draw(im)
    d.rectangle([40, 40, 1160, 860], outline="white", width=6)
    d.text((80, 80), text, fill="white")
    for i in range(6):
        d.rectangle([150 + i * 160, 400, 270 + i * 160, 420], fill="white")
    im.save(media / f"{parent}_{q}_{n}.jpg", quality=80)


main, fx, rp = [], [], []
HOURS = {"meeting": (4, 245), "office": (10, 245), "corridor": (12, 245), "toilet": (10, 245),
         "storage": (1, 245), "outdoor": (3, 200), "classroom": (8, 200)}
W = {"FL40x2_G": 86, "FL40x1_G": 44, "FL20x1_G": 24, "FDL27x1": 30, "MH1000_FL": 1060, "HL500_FL": 500, "EMG_IL": 30}
# 施設, 階, 通し番号, 部屋名, 用途, 縦, 横, 天井高, 状態, 照明状況, 課題,
#   器具[(type, shape, lamp_kind, lamp_spec, mount, method, qty, cols, rows, height, aerial, removal, cond, defects)]
rooms = [
    ("F02", "1F", "1F-01", "集会室", "meeting", 12, 9, 3.2, "done", "bit_dark", "夜間の集会利用が多い",
     [("FL40x2_G", "inverted_fuji", "fl", "FL40SS・EX-N/37 ×2", "direct", "array", None, 4, 3, 3.2, "ladder",
       "fixture disposal", "poor", "yellow rust salt")]),
    ("F02", "1F", "1F-02", "事務室", "office", 5, 4, 2.6, "done", "enough", "",
     [("FL40x1_G", "trough", "fl", "FL40SS ×1", "direct", "each", 2, None, None, 2.6, "none", "fixture", "slight", "yellow")]),
    ("F02", "1F", "1F-03", "トイレ", "toilet", 3, 3, 2.4, "check", "dark", "1台不点。器具の型番が読めない",
     [("FDL27x1", "downlight", "compact", "FDL27EX-N", "recessed", "each", 2, None, None, 2.4, "none",
       "fixture ceiling_repair", "poor", "unlit")]),
    ("F11", "1F", "1F-01", "遊戯室", "classroom", 15, 10, 4.5, "done", "enough", "高天井。交換時はローリングタワーが必要",
     [("FL40x2_G", "louver", "fl", "FLR40S ×2 ラピッド", "direct", "array", None, 5, 4, 4.5, "tower",
       "fixture pcb disposal", "slight", "dirt"),
      ("EMG_IL", "emergency", "incandescent", "白熱30W", "direct", "each", 4, None, None, 4.5, "tower", "fixture", "good", "")]),
    ("F08", "OUT", "OUT-01", "コート照明", "outdoor", 36, 18, None, "done", "dark", "照度不足の声あり。ポール根元に塩害",
     [("MH1000_FL", "flood", "mh", "MH1000W", "pole", "each", 8, None, None, 10, "lift",
       "fixture pole disposal", "severe", "rust salt crack")]),
]
colors = ["#35577d", "#5d7d35", "#7d5a35", "#6b357d", "#357d74"]
for i, (fac, fl, no, name, use, d, w, h, st, light, issue, fxs) in enumerate(rooms):
    g = gid()
    hd, dy = HOURS[use]
    main.append({"ObjectID": i + 1, "GlobalID": "{" + g.upper() + "}", "報告者氏名": "伊禮", "案件名": "南大東村LED化現況調査",
                 "施設名": fac, "階": fl, "記録の種類": "部屋の調査", "調査日時": f"2026-12-1{4 + i // 2} 10:{10 + i}",
                 "通し番号": no, "部屋名（室名札）": name, "用途": use, "縦（m）": d, "横（m）": w,
                 "天井高（m）": h, "1日の点灯時間（h）": hd, "年間点灯日数": dy, "年間点灯時間（h）": hd * dy,
                 "非常灯・誘導灯の有無": "yes" if fac == "F11" else "no", "調査状態": st,
                 "照明状況（明るさの体感）": light, "施設利用上・更新上の課題": issue, "x": 131.2366 + i * 0.001, "y": 25.8467})
    photo(g.lower(), "room_photo", 1, f"ROOM {fac} {no}", colors[i])
    photo(g.lower(), "ceil_photo", 1, f"CEILING {fac} {no}", colors[i])
    for (t, shape, lk, spec, m, meth, q, c, r, hh, aer, rem, cond, dfc) in fxs:
        fg = gid()
        qty = c * r if meth == "array" else q
        fx.append({"GlobalID": fg, "ParentGlobalID": g, "fx_type": t, "fx_shape": shape, "fx_mount": m,
                   "fx_method": meth, "fx_cols": c, "fx_rows": r, "fx_qty_in": q, "fx_qty": qty, "fx_lit": "lit",
                   "fx_lamp_kind": lk, "fx_lamp_spec": spec, "fx_height_m": hh, "fx_aerial": aer, "fx_removal": rem,
                   "fx_cond": cond, "fx_defects": dfc, "fx_w_old": W[t]})
        photo(fg, "fx_photo", 1, f"FIXTURE {t}", colors[i])
        photo(fg, "fx_plate", 1, f"NAMEPLATE {t}", "#444444")
    if no == "1F-03":
        rg = gid()
        rp.append({"GlobalID": rg, "ParentGlobalID": g, "区分": "要対応", "現場メモ": "天井に雨漏り跡。器具交換前に補修が必要",
                   "コメント": "施設管理者へ連絡済み", "x": 131.2367, "y": 25.8468})
        photo(rg, "rp_photo", 1, "LEAK", "#7d3535")
# 階の確認（在所集会所1F：扉4に対し登録3室 → 漏れ警告）
for fac, fl, doors, dec in (("F02", "1F", 4, None), ("F11", "1F", 1, 1)):
    g = gid()
    main.append({"GlobalID": g, "報告者氏名": "伊禮", "案件名": "南大東村LED化現況調査", "施設名": fac, "階": fl,
                 "記録の種類": "階の確認", "調査日時": "2026-12-14 09:00",
                 "扉の数（この階・現地カウント）": doors, "施設側の申告部屋数": dec})
    photo(g, "floor_photo", 1, f"EVAC MAP {fac}", "#555555")
# 現場写真報告（屋外ステージ）
g = gid()
main.append({"GlobalID": g, "報告者氏名": "伊禮", "案件名": "南大東村LED化現況調査", "施設名": "F07", "階": "OUT",
             "記録の種類": "現場写真報告", "調査日時": "2026-12-15 15:00", "対象地名・場所": "ステージ上部照明",
             "x": 131.24, "y": 25.85})
rg = gid()
rp.append({"GlobalID": rg, "ParentGlobalID": g, "区分": "caution", "現場メモ": "投光器の取付金具が腐食。高所作業車の進入路あり",
           "コメント": "", "x": 131.2401, "y": 25.8501})
photo(rg, "rp_photo", 1, "STAGE", "#35357d")
photo(rg, "rp_photo2", 1, "STAGE 2", "#35357d")

with pd.ExcelWriter(out / "survey123_export.xlsx") as w:
    pd.DataFrame(main).to_excel(w, sheet_name="led_genkyo_chosa", index=False)
    pd.DataFrame(fx).to_excel(w, sheet_name="fixtures", index=False)
    pd.DataFrame(rp).to_excel(w, sheet_name="reports", index=False)
print("sample:", out, len(main), len(fx), len(rp))
