# -*- coding: utf-8 -*-
"""LED室内現況調査 Survey123用 XLSForm とマスタCSVを生成する。"""
import csv
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

出力先 = sys.argv[1] if len(sys.argv) > 1 else "out"
フォーム先 = os.path.join(出力先, "survey123_form")
メディア先 = os.path.join(フォーム先, "media")
os.makedirs(メディア先, exist_ok=True)

列 = ["type", "name", "label", "hint", "appearance", "required", "readonly", "default",
     "calculation", "constraint", "constraint_message", "relevant", "parameters",
     "body::esri:style", "bind::esri:parameters", "bind::esri:fieldType", "bind::esri:fieldLength"]

def q(type_, name, label="", **kw):
    行 = {c: "" for c in 列}
    行.update(type=type_, name=name, label=label)
    for k, v in kw.items():
        行[k.replace("__", "::").replace("_esri_", ":esri:")] = v
    return 行

写真 = "max-pixels=1920"
S = []
# ---------- 基本情報 ----------
S += [
    q("begin group", "hdr", "基本情報", appearance="field-list"),
    q("text", "reporter", "報告者氏名", required="yes", hint="Field App の「お気に入り回答」に登録すると毎回の入力が不要になります"),
    q("text", "project", "案件名", required="yes"),
    q("select_one_from_file facilities.csv", "building", "施設名", required="yes", appearance="autocomplete"),
    q("text", "building_other", "施設名（一覧にない場合）", required="yes", relevant="${building}='other'"),
    q("select_one floors", "floor", "階", required="yes", appearance="minimal"),
    q("select_one rec_types", "rec_type", "記録の種類", required="yes", appearance="horizontal"),
    q("dateTime", "survey_dt", "調査日時", default="now()"),
    q("geopoint", "location", "調査位置", hint="屋内でGPSが弱い場合は建物入口付近で取得してください"),
    q("end group", "hdr"),
]
# ---------- 部屋の調査 ----------
r = "${rec_type}='room'"
S += [
    q("begin group", "room", "部屋の調査", appearance="field-list", relevant=r),
    q("text", "room_no", "通し番号", required="yes",
      hint="例：2F-07。階段から時計回りなど、延べ順で付番します（部屋の一意な名前になります）"),
    q("image", "room_photo", "室名札・入口の写真", required="yes", parameters=写真,
      hint="どの部屋か第三者にわかるように撮影します（図面の代わり）。屋外は全景"),
    q("text", "room_name", "部屋名（室名札）", hint="iPhoneはキーボードの「テキストをスキャン」で室名札を読み取れます。室名札がなければ空欄"),
    q("select_one uses", "room_use", "用途", required="yes", appearance="minimal"),
    q("text", "room_use_other", "用途（その他）", relevant="${room_use}='other'"),
    q("decimal", "depth_m", "縦（m）", constraint=". > 0 and . < 300", constraint_message="0〜300mで入力"),
    q("decimal", "width_m", "横（m）", constraint=". > 0 and . < 300", constraint_message="0〜300mで入力"),
    q("decimal", "area_m2", "面積（m²）", readonly="yes",
      calculation="if(${depth_m} > 0 and ${width_m} > 0, round(${depth_m} * ${width_m}, 1), '')"),
    q("decimal", "ceil_h_m", "天井高（m）", constraint=". > 0 and . < 30", constraint_message="0〜30mで入力", relevant="${room_use} != 'outdoor'"),
    q("select_one ceil_types", "ceil_type", "天井材", appearance="minimal", relevant="${room_use} != 'outdoor'"),
    q("integer", "sw_circuits", "スイッチ回路数"),
    q("decimal", "hours_day", "1日の点灯時間（h）", required="yes",
      calculation="pulldata('usage_hours', 'hours_day', 'name', ${room_use})",
      bind__esri_parameters="calculationMode=whenEmpty", hint="用途から初期値が入ります。聞き取り結果で上書きしてください"),
    q("integer", "days_year", "年間点灯日数", required="yes",
      calculation="pulldata('usage_hours', 'days_year', 'name', ${room_use})",
      bind__esri_parameters="calculationMode=whenEmpty"),
    q("calculate", "annual_h", "年間点灯時間（h）", calculation="${hours_day} * ${days_year}",
      bind__esri_fieldType="esriFieldTypeDouble"),
    q("image", "ceil_photo", "天井見上げ写真（器具に印を付ける）", appearance="annotate", parameters=写真,
      hint="撮影後、器具1台ごとに印を付けます。印の数を下の「台数」に入力します。屋外は照明の全景"),
    q("image", "ceil_photo2", "天井見上げ写真（2枚目・広い部屋用）", appearance="annotate", parameters=写真),
    q("end group", "room"),
    # 器具（種類ごとに1件）
    q("begin repeat", "fixtures", "器具（種類・取付方式ごとに1件）", relevant=r),
    q("select_one_from_file fixture_types.csv", "fx_type", "器具タイプ", required="yes", appearance="autocomplete"),
    q("select_one shapes", "fx_shape", "形状", required="yes", appearance="minimal"),
    q("select_one mounts", "fx_mount", "取付方法", required="yes", appearance="horizontal"),
    q("select_one count_methods", "fx_method", "数え方", required="yes", appearance="horizontal", default="each"),
    q("integer", "fx_cols", "列数", required="yes", relevant="${fx_method}='array'"),
    q("integer", "fx_rows", "行数", required="yes", relevant="${fx_method}='array'"),
    q("integer", "fx_qty_in", "台数", required="yes", relevant="${fx_method}='each'"),
    q("calculate", "fx_qty", "台数（集計用）",
      calculation="if(${fx_method}='array', ${fx_cols} * ${fx_rows}, ${fx_qty_in})",
      bind__esri_fieldType="esriFieldTypeInteger"),
    q("select_one lit_states", "fx_lit", "点灯状況", appearance="horizontal", default="lit"),
    q("text", "fx_model", "メーカー・型番（銘板）", hint="読めない場合は銘板の写真だけで構いません（PC側で読み取り）"),
    q("image", "fx_photo", "器具の写真", parameters=写真),
    q("image", "fx_plate", "銘板の写真", parameters=写真),
    q("select_one lamp_kinds", "fx_lamp_kind", "ランプ区分", required="yes", appearance="minimal"),
    q("text", "fx_lamp_spec", "ランプ仕様（形式・色・灯数）", hint="例：FLR40S・W/M-36 ×2灯、ラピッド"),
    q("decimal", "fx_w_field", "銘板の消費電力（W/台）", hint="銘板で読めた場合に入力。空欄ならマスタの値を使います"),
    q("decimal", "fx_height_m", "設置高さ（m）", hint="床（地面）から器具までの高さ"),
    q("select_one aerial", "fx_aerial", "高所作業の要否", required="yes", appearance="horizontal"),
    q("select_multiple removal", "fx_removal", "既存設備の撤去条件"),
    q("text", "fx_removal_note", "撤去条件の補足"),
    q("select_one conditions", "fx_cond", "老朽度・劣化", required="yes", appearance="horizontal"),
    q("select_multiple defects", "fx_defects", "劣化の内容", relevant="${fx_cond} != 'good'"),
    q("calculate", "fx_w_master", "マスタ消費電力（W/台）",
      calculation="pulldata('fixture_types', 'watt', 'name', ${fx_type})", bind__esri_fieldType="esriFieldTypeDouble"),
    q("calculate", "fx_w_old", "既設消費電力（W/台）",
      calculation="if(${fx_w_field} != '' and ${fx_w_field} > 0, ${fx_w_field}, ${fx_w_master})",
      bind__esri_fieldType="esriFieldTypeDouble"),
    q("calculate", "fx_key", "置換キー", calculation="concat(${fx_type}, '|', ${fx_mount})"),
    q("calculate", "fx_led", "LED候補品番", calculation="pulldata('led_rules', 'led_code', 'key', ${fx_key})"),
    q("calculate", "fx_w_led", "LED消費電力（W/台）",
      calculation="pulldata('led_rules', 'led_watt', 'key', ${fx_key})", bind__esri_fieldType="esriFieldTypeDouble"),
    q("calculate", "fx_kwh", "年間削減電力量（kWh）",
      calculation="if(${fx_w_old} != '' and ${fx_w_led} != '', round((${fx_w_old} - ${fx_w_led}) * ${fx_qty} * ${annual_h} div 1000, 1), '')",
      bind__esri_fieldType="esriFieldTypeDouble"),
    q("note", "fx_result", "既設 ${fx_w_old}W → LED候補 ${fx_led}（${fx_w_led}W）／年間削減 約${fx_kwh}kWh",
      relevant="${fx_type} != ''"),
    q("text", "fx_note", "特記"),
    q("end repeat", "fixtures"),
    q("begin group", "room_end", "部屋のまとめ", appearance="field-list", relevant=r),
    q("calculate", "room_qty", "器具台数（部屋計）", calculation="sum(${fx_qty})", bind__esri_fieldType="esriFieldTypeInteger"),
    q("calculate", "room_kwh", "年間削減電力量（部屋計 kWh）", calculation="round(sum(${fx_kwh}), 1)",
      bind__esri_fieldType="esriFieldTypeDouble"),
    q("note", "room_sum", "この部屋：器具 ${room_qty} 台／年間削減 約 ${room_kwh} kWh"),
    q("select_one yn3", "emergency", "非常灯・誘導灯の有無", required="yes", appearance="horizontal"),
    q("decimal", "lux", "平均照度（lx）"),
    q("select_one light_levels", "room_light", "照明状況（明るさの体感）", appearance="horizontal"),
    q("text", "room_issue", "施設利用上・更新上の課題", appearance="multiline",
      hint="例：夜間利用が多い／高天井で交換に足場が必要／器具の一部が既にLED"),
    q("select_one statuses", "room_status", "調査状態", required="yes", appearance="horizontal"),
    q("end group", "room_end"),
]
# ---------- 階の確認（網羅チェック） ----------
f = "${rec_type}='floor'"
S += [
    q("begin group", "floor_chk", "階の確認（部屋の漏れチェック）", appearance="field-list", relevant=f),
    q("integer", "door_count", "扉の数（この階・現地カウント）", required="yes",
      hint="廊下から入れる扉を数えます。PC側で登録部屋数と突き合わせます"),
    q("integer", "declared_rooms", "施設側の申告部屋数"),
    q("image", "floor_photo", "避難案内図・フロア案内の写真", appearance="annotate", parameters=写真),
    q("image", "floor_sketch", "階の簡易スケッチ", appearance="draw", parameters=写真),
    q("text", "floor_note", "メモ", appearance="multiline"),
    q("end group", "floor_chk"),
]
# ---------- 現場写真報告（公園用フォーム相当） ----------
s = "${rec_type}='site'"
S += [
    q("begin group", "site", "現場写真報告", appearance="field-list", relevant=s),
    q("text", "site_name", "対象地名・場所", required="yes"),
    q("end group", "site"),
]
# ---------- 報告事項（全種類共通の繰り返し） ----------
S += [
    q("begin repeat", "reports", "報告事項（写真・メモ・音声・撮影場所）"),
    q("select_one rp_cats", "rp_cat", "区分", appearance="minimal"),
    q("image", "rp_photo", "該当写真", parameters=写真),
    q("image", "rp_photo2", "該当写真（2枚目）", parameters=写真),
    q("image", "rp_photo3", "該当写真（3枚目）", parameters=写真),
    q("text", "rp_memo", "現場メモ", appearance="multiline", constraint="string-length(.) <= 1000",
      constraint_message="1000字以内", bind__esri_fieldLength="1000"),
    q("audio", "rp_audio", "音声", bind__esri_parameters="max-duration=180"),
    q("geopoint", "rp_loc", "撮影場所"),
    q("text", "rp_comment", "コメント", bind__esri_fieldLength="1000"),
    q("end repeat", "reports"),
]

選択肢 = {
    "floors": [("B2", "B2"), ("B1", "B1")] + [(f"{i}F", f"{i}F") for i in range(1, 16)] + [("RF", "RF"), ("PH", "PH"), ("OUT", "屋外")],
    "rec_types": [("room", "部屋の調査"), ("floor", "階の確認"), ("site", "現場写真報告")],
    "uses": [("office", "事務室"), ("meeting", "会議室"), ("classroom", "教室"), ("corridor", "廊下"),
             ("stairs", "階段"), ("toilet", "トイレ"), ("storage", "倉庫"), ("machine", "機械室・電気室"),
             ("locker", "更衣室"), ("pantry", "給湯室"), ("entrance", "玄関・ホール"), ("gym", "体育館"),
             ("parking", "駐車場"), ("outdoor", "屋外"), ("other", "その他")],
    "ceil_types": [("board", "岩綿吸音板"), ("plaster", "石膏ボード"), ("system", "システム天井"),
                   ("exposed", "直天井（躯体現し）"), ("other", "その他")],
    "mounts": [("direct", "直付"), ("recessed", "埋込"), ("pendant", "吊下"), ("wall", "壁付"),
               ("pole", "ポール"), ("stand", "架台・スタンド")],
    "shapes": [("inverted_fuji", "ベースライト（逆富士）"), ("trough", "ベースライト（トラフ）"),
               ("open_bottom", "ベースライト（下面開放・埋込）"), ("louver", "ベースライト（ルーバー・カバー付）"),
               ("downlight", "ダウンライト"), ("ceiling", "シーリングライト"), ("bracket", "ブラケット"),
               ("pendant", "ペンダント"), ("highbay", "高天井用（ハイベイ）"), ("flood", "投光器"),
               ("street", "街路灯・防犯灯"), ("exit", "誘導灯"), ("emergency", "非常用照明"), ("other", "その他")],
    "lamp_kinds": [("fl", "蛍光灯（直管）"), ("fhf", "蛍光灯（Hf）"), ("compact", "蛍光灯（コンパクト・環形）"),
                   ("hg", "水銀灯"), ("mh", "メタルハライド"), ("hps", "高圧ナトリウム"), ("incandescent", "白熱灯"),
                   ("halogen", "ハロゲン"), ("led", "LED"), ("other", "その他")],
    "aerial": [("none", "不要（手が届く）"), ("ladder", "脚立"), ("tower", "ローリングタワー"),
               ("lift", "高所作業車"), ("scaffold", "足場")],
    "removal": [("fixture", "器具ごと交換"), ("lamp_only", "ランプ交換（直結工事）"), ("ceiling_repair", "天井補修が必要"),
                ("asbestos", "アスベスト含有の確認が必要"), ("pcb", "PCB含有安定器の確認が必要"),
                ("pole", "ポール・架台の補修が必要"), ("disposal", "産廃処理（水銀含有）"), ("other", "その他")],
    "conditions": [("good", "良好"), ("slight", "やや劣化"), ("poor", "劣化"), ("severe", "著しい劣化")],
    "defects": [("yellow", "カバー黄変"), ("rust", "錆・腐食"), ("crack", "破損・割れ"), ("flicker", "ちらつき"),
                ("unlit", "不点"), ("noise", "安定器の異音"), ("burn", "焦げ・変色"), ("dirt", "汚れ・虫の侵入"),
                ("salt", "塩害")],
    "light_levels": [("enough", "十分"), ("bit_dark", "やや暗い"), ("dark", "暗い"), ("too_bright", "明るすぎ")],
    "count_methods": [("each", "個別に数える"), ("array", "配列（列×行）")],
    "lit_states": [("lit", "点灯"), ("unlit", "不点"), ("missing", "ランプ欠品")],
    "yn3": [("yes", "有"), ("no", "無"), ("unknown", "未確認")],
    "statuses": [("done", "調査済"), ("check", "要確認")],
    "rp_cats": [("action", "要対応"), ("caution", "施工上の注意"), ("watch", "継続観察"),
                ("good", "良好"), ("other", "その他")],
}

wb = Workbook()
ws = wb.active
ws.title = "survey"
ws.append(列)
for 行 in S:
    ws.append([行[c] for c in 列])
wc = wb.create_sheet("choices")
wc.append(["list_name", "name", "label"])
for ln, items in 選択肢.items():
    for n, l in items:
        wc.append([ln, n, l])
    wc.append([])
wst = wb.create_sheet("settings")
wst.append(["form_title", "form_id", "instance_name", "style", "version"])
wst.append(["LED室内現況調査", "led_genkyo_chosa",
            "concat(${building}, ' ', ${floor}, ' ', ${room_no}, ' ', ${room_name}, ${site_name})",
            "pages", "1.2.0"])
for sh in (ws, wc, wst):
    for c in sh[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    for i, col in enumerate(sh.columns, 1):
        w = max(len(str(c.value or "")) for c in col)
        sh.column_dimensions[get_column_letter(i)].width = min(max(10, w * 1.2), 60)
    sh.freeze_panes = "A2"
    for row in sh.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top", wrap_text=False)
wb.save(os.path.join(フォーム先, "led_genkyo_chosa.xlsx"))

# ---------- マスタCSV（media フォルダ。pulldata と select_one_from_file が参照） ----------
# 消費電力は安定器損失込みの参考値（仮値）。自社で使っている値・メーカーカタログ値に置き換えること。
器具 = [
    # name, label, lamp, lamps, ballast, watt, shape
    ("FL20x1_G", "直管20形×1（グロー）", "FL20", 1, "グロー", 24, "line"),
    ("FL20x2_G", "直管20形×2（グロー）", "FL20", 2, "グロー", 46, "line"),
    ("FL40x1_G", "直管40形×1（グロー）", "FL40", 1, "グロー", 44, "line"),
    ("FL40x2_G", "直管40形×2（グロー）", "FL40", 2, "グロー", 86, "line"),
    ("FLR40x1_R", "直管40形×1（ラピッド）", "FLR40", 1, "ラピッド", 48, "line"),
    ("FLR40x2_R", "直管40形×2（ラピッド）", "FLR40", 2, "ラピッド", 88, "line"),
    ("FHF32x1_H", "Hf32形×1（高出力）", "FHF32", 1, "Hf", 50, "line"),
    ("FHF32x2_H", "Hf32形×2（高出力）", "FHF32", 2, "Hf", 90, "line"),
    ("FHF32x2_S", "Hf32形×2（定格出力）", "FHF32", 2, "Hf", 66, "line"),
    ("FDL27x1", "コンパクト形ダウンライト FDL27×1", "FDL27", 1, "-", 30, "round"),
    ("FDL27x2", "コンパクト形ダウンライト FDL27×2", "FDL27", 2, "-", 60, "round"),
    ("IL60_DL", "白熱ダウンライト 60W", "白熱60W", 1, "-", 60, "round"),
    ("FCL30_32", "環形 FCL30+32", "FCL", 2, "-", 66, "round"),
    ("HF250", "水銀灯 250W", "HF250", 1, "安定器", 265, "high"),
    ("HF400", "水銀灯 400W", "HF400", 1, "安定器", 420, "high"),
    ("EXIT_FL", "誘導灯（蛍光灯）", "FL20", 1, "-", 22, "line"),
    ("LED_EXIST", "既設LED（更新対象外）", "LED", 1, "-", 0, "line"),
    ("MH400_FL", "メタルハライド投光器 400W", "MH400", 1, "安定器", 430, "high"),
    ("MH1000_FL", "メタルハライド投光器 1kW", "MH1000", 1, "安定器", 1060, "high"),
    ("HL500_FL", "ハロゲン投光器 500W", "J500", 1, "-", 500, "high"),
    ("NH360", "高圧ナトリウム灯 360W", "NH360", 1, "安定器", 380, "high"),
    ("EMG_IL", "非常用照明（白熱）", "白熱", 1, "-", 30, "round"),
    ("OTHER", "その他（特記に記入）", "-", 1, "-", "", "line"),
]
with open(os.path.join(メディア先, "fixture_types.csv"), "w", newline="", encoding="utf-8-sig") as fp:
    w = csv.writer(fp)
    w.writerow(["name", "label", "lamp", "lamps", "ballast", "watt", "shape"])
    w.writerows(器具)

置換 = {  # 既設 → (LED候補品番[仮], LED消費電力W[仮])
    "FL20x1_G": ("LED-B20-1L", 10), "FL20x2_G": ("LED-B20-2L", 19),
    "FL40x1_G": ("LED-B40-2500", 18), "FL40x2_G": ("LED-B40-5200", 33),
    "FLR40x1_R": ("LED-B40-2500", 18), "FLR40x2_R": ("LED-B40-5200", 33),
    "FHF32x1_H": ("LED-B40-3200", 23), "FHF32x2_H": ("LED-B40-6900", 45),
    "FHF32x2_S": ("LED-B40-5200", 33), "FDL27x1": ("LED-DL100", 10),
    "FDL27x2": ("LED-DL150", 16), "IL60_DL": ("LED-DL60", 7),
    "FCL30_32": ("LED-CL8", 30), "HF250": ("LED-HB-250", 90),
    "HF400": ("LED-HB-400", 150), "MH400_FL": ("LED-FL-400", 150), "MH1000_FL": ("LED-FL-1000", 400), "HL500_FL": ("LED-FL-500", 100), "NH360": ("LED-ST-360", 100), "EMG_IL": ("LED-EMG", 3), "EXIT_FL": ("LED-EXIT-C", 3), "LED_EXIST": ("（対象外）", 0),
}
取付記号 = {"direct": "D", "recessed": "R", "pendant": "P", "wall": "W", "pole": "PL", "stand": "S"}
with open(os.path.join(メディア先, "led_rules.csv"), "w", newline="", encoding="utf-8-sig") as fp:
    w = csv.writer(fp)
    w.writerow(["key", "fx_type", "mount", "led_code", "led_watt"])
    for t, (code, watt) in 置換.items():
        for m, k in 取付記号.items():
            c = code if code.startswith("（") else f"仮_{code}-{k}"
            w.writerow([f"{t}|{m}", t, m, c, watt])

施設 = [  # 南大東村 15施設（依頼書の延床面積）
    ("F01", "旧東集会所", 100, "屋内"), ("F02", "在所集会所", 154, "屋内"), ("F03", "新東集会所", 110, "屋内"),
    ("F04", "池之沢集会所", 154, "屋内"), ("F05", "北集会所", 154, "屋内"), ("F06", "南集会所", 200, "屋内"),
    ("F07", "ふれあい広場（屋外ステージ）", 150, "屋外"), ("F08", "テニスコート", 200, "屋外"),
    ("F09", "星野洞（事務所・待合室）", 92, "屋内"), ("F10", "南大東村学習支援センター", 311, "屋内"),
    ("F11", "南大東村児童館", 451, "屋内"), ("F12", "南大東村高齢者生活福祉センター", 700, "屋内"),
    ("F13", "南大東村保健センター", 379, "屋内"), ("F14", "南大東村エコセンター", 320, "屋内"),
    ("F15", "旧空港施設（ターミナルビル）", 186, "屋内"), ("other", "その他（一覧にない施設）", "", ""),
]
with open(os.path.join(メディア先, "facilities.csv"), "w", newline="", encoding="utf-8-sig") as fp:
    w = csv.writer(fp)
    w.writerow(["name", "label", "floor_area_m2", "kind"])
    w.writerows(施設)

用途時間 = {  # 用途 → (1日の点灯時間h, 年間点灯日数)  仮値
    "office": (10, 245), "meeting": (4, 245), "classroom": (8, 200), "corridor": (12, 245),
    "stairs": (12, 245), "toilet": (10, 245), "storage": (1, 245), "machine": (1, 365),
    "locker": (3, 245), "pantry": (3, 245), "entrance": (12, 245), "gym": (6, 300),
    "parking": (24, 365), "outdoor": (12, 365), "other": (8, 245),
}
with open(os.path.join(メディア先, "usage_hours.csv"), "w", newline="", encoding="utf-8-sig") as fp:
    w = csv.writer(fp)
    w.writerow(["name", "label", "hours_day", "days_year"])
    ラベル = dict(選択肢["uses"])
    for k, (h, d) in 用途時間.items():
        w.writerow([k, ラベル[k], h, d])
import json
項目 = {r["name"]: {"type": r["type"], "label": r["label"]} for r in S if not r["type"].startswith(("begin", "end", "note"))}
with open(os.path.join(出力先, "fields.json"), "w", encoding="utf-8") as fp:
    json.dump({"fields": 項目, "choices": {k: dict(v) for k, v in 選択肢.items()}}, fp, ensure_ascii=False, indent=1)
print("ok", len(S), "rows")
