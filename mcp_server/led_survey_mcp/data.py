# -*- coding: utf-8 -*-
"""Survey123（LED室内現況調査）のデータ取得・読み込み・正規化。

取得元は2通り:
  1. ArcGIS API for Python で ArcGIS Online のフィーチャレイヤーを直接クエリ（写真・音声も取得）
  2. Survey123 Webサイトの「データ」→「エクスポート」→ Excel で書き出したファイル（属性のみ）

どちらもクエリ・ダウンロードのみで、ArcGIS上の解析・レポート・ジオコーディングは呼ばない
（＝クレジットを消費しない）。取得結果は作業フォルダに dataset.json としてキャッシュする。
"""
from __future__ import annotations

import datetime as _dt
import json
import math
import os
import re
from pathlib import Path

_HERE = Path(__file__).parent
_FIELDS = json.loads((_HERE / "fields.json").read_text(encoding="utf-8"))
FIELDS: dict = _FIELDS["fields"]
CHOICES: dict = _FIELDS["choices"]

# repeat（関連テーブル）を判定するための代表フィールド
_MAIN_KEY, _FX_KEY, _RP_KEY = "rec_type", "fx_type", "rp_memo"
_LABEL_TO_NAME = {v["label"]: k for k, v in FIELDS.items()}
_SYSTEM = {"globalid": "globalid", "parentglobalid": "parentglobalid", "objectid": "objectid",
           "x": "x", "y": "y", "creationdate": "created", "creator": "creator",
           "editdate": "edited", "editor": "editor"}


def master_dir() -> Path:
    return Path(os.environ.get("LED_SURVEY_MASTER_DIR", _HERE / "masters"))


def facility_master() -> dict:
    import csv
    p = master_dir() / "facilities.csv"
    if not p.exists():
        return {}
    with open(p, encoding="utf-8-sig") as fp:
        return {r["name"]: r for r in csv.DictReader(fp)}


def _apply_facility_labels(ds: dict):
    """施設コード（F01 など）を施設名に置き換える。コードは building_code に残す。"""
    fac = facility_master()
    for r in ds["main"]:
        code = r.get("building")
        if code is None or "building_code" in r:
            continue
        r["building_code"] = code
        if code == "other":
            r["building"] = r.get("building_other") or "その他"
        elif code in fac:
            r["building"] = fac[code]["label"]


def work_dir() -> Path:
    d = Path(os.environ.get("LED_SURVEY_WORKDIR", Path.home() / "LED現況調査"))
    d.mkdir(parents=True, exist_ok=True)
    return d


def project_dir(project: str) -> Path:
    safe = re.sub(r'[\\/:*?"<>|]', "_", project).strip() or "default"
    d = work_dir() / safe
    d.mkdir(parents=True, exist_ok=True)
    return d


# ---------------------------------------------------------------- 正規化
def _norm_col(col: str) -> str | None:
    c = str(col).strip()
    low = c.lower().replace(" ", "")
    if low in _SYSTEM:
        return _SYSTEM[low]
    if c in FIELDS:
        return c
    if low in FIELDS:
        return low
    if c in _LABEL_TO_NAME:
        return _LABEL_TO_NAME[c]
    return None


def _clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.isoformat()
    if hasattr(v, "isoformat") and not isinstance(v, str):  # pandas Timestamp
        try:
            return v.isoformat()
        except Exception:
            pass
    if isinstance(v, (int,)) and v > 10**11 and v < 10**14:  # epoch ms（API の日付）
        return _dt.datetime.fromtimestamp(v / 1000).isoformat(timespec="minutes")
    return v


def _to_code(field: str, v):
    """選択肢がラベルで書き出された場合にコードへ戻す。"""
    t = FIELDS.get(field, {}).get("type", "")
    if v is None or not t.startswith("select_one ") or "from_file" in t:
        return v
    lst = CHOICES.get(t.split()[1], {})
    if v in lst:
        return v
    for code, label in lst.items():
        if label == v:
            return code
    return v


def normalize_rows(rows: list[dict]) -> list[dict]:
    out = []
    for r in rows:
        n = {}
        for k, v in r.items():
            key = _norm_col(k)
            if key:
                n[key] = _to_code(key, _clean(v))
        for g in ("globalid", "parentglobalid"):
            if n.get(g):
                n[g] = str(n[g]).strip("{}").lower()
        out.append(n)
    return out


def _classify(rows: list[dict]) -> str | None:
    keys = set().union(*[r.keys() for r in rows]) if rows else set()
    if _FX_KEY in keys:
        return "fixtures"
    if _RP_KEY in keys or "rp_cat" in keys:
        return "reports"
    if _MAIN_KEY in keys or "room_no" in keys:
        return "main"
    return None


def empty_dataset(source: str) -> dict:
    return {"meta": {"source": source, "loaded_at": _dt.datetime.now().isoformat(timespec="seconds")},
            "main": [], "fixtures": [], "reports": [], "media": []}


# ---------------------------------------------------------------- Excel 書き出しから
def load_excel_export(path: str) -> dict:
    import pandas as pd
    ds = empty_dataset(f"excel:{path}")
    book = pd.read_excel(path, sheet_name=None)
    for name, df in book.items():
        rows = normalize_rows(df.to_dict("records"))
        kind = _classify(rows)
        if kind:
            ds[kind].extend(rows)
    for r in ds["main"]:
        if r.get("x") is not None and r.get("y") is not None:
            r["lon"], r["lat"] = r.pop("x"), r.pop("y")
    for r in ds["reports"]:
        if r.get("x") is not None and r.get("y") is not None:
            r["lon"], r["lat"] = r.pop("x"), r.pop("y")
    return ds


def attach_media_folder(ds: dict, media_dir: str) -> int:
    """{globalid}_{question}_{n}.jpg 形式で保存された写真フォルダを紐付ける（fetch で保存した形式と同じ）。"""
    n = 0
    for p in Path(media_dir).rglob("*"):
        m = re.match(r"([0-9a-fA-F-]{36})_([a-z0-9_]+?)_(\d+)\.\w+$", p.name)
        if p.is_file() and m:
            ds["media"].append({"parent": m.group(1).lower(), "question": m.group(2), "path": str(p)})
            n += 1
    return n


# ---------------------------------------------------------------- ArcGIS Online から直接
def connect_gis():
    from arcgis.gis import GIS  # 遅延 import（未インストールでも Excel 読込は動く）
    profile = os.environ.get("ARCGIS_PROFILE")
    url = os.environ.get("ARCGIS_URL", "https://www.arcgis.com")
    if profile:
        return GIS(profile=profile)
    user, pw = os.environ.get("ARCGIS_USERNAME"), os.environ.get("ARCGIS_PASSWORD")
    if user and pw:
        return GIS(url, user, pw)
    return GIS("pro") if os.environ.get("ARCGIS_USE_PRO") else GIS(url)


def _feature_service(gis, item_id: str):
    item = gis.content.get(item_id)
    if item is None:
        raise ValueError(f"アイテムが見つかりません: {item_id}")
    if item.type == "Form":  # Survey123 フォームのIDが渡された場合は関連フィーチャサービスへ
        rel = item.related_items("Survey2Service", "forward")
        if not rel:
            raise ValueError("フォームに紐づくフィーチャサービスが見つかりません")
        item = rel[0]
    return item


def fetch_from_arcgis(item_id: str, out_dir: Path, where: str = "1=1",
                      download_media: bool = True, since: str | None = None) -> dict:
    from arcgis.features import FeatureLayerCollection
    gis = connect_gis()
    item = _feature_service(gis, item_id)
    flc = FeatureLayerCollection.fromitem(item)
    ds = empty_dataset(f"arcgis:{item.id}")
    ds["meta"]["title"] = item.title
    media_dir = out_dir / "media"
    for lyr in list(flc.layers) + list(flc.tables):
        w = where
        if since:
            w = f"({where}) AND EditDate >= TIMESTAMP '{since} 00:00:00'"
        fs = lyr.query(where=w if lyr in flc.layers else "1=1", out_fields="*", return_geometry=True)
        raw = []
        for f in fs.features:
            a = dict(f.attributes)
            g = f.geometry or {}
            if "x" in g and "y" in g:
                a["x"], a["y"] = g["x"], g["y"]
            raw.append(a)
        rows = normalize_rows(raw)
        kind = _classify(rows)
        if not kind:
            continue
        for r in rows:
            if r.get("x") is not None:
                r["lon"], r["lat"] = r.pop("x"), r.pop("y")
        ds[kind].extend(rows)
        if download_media and lyr.properties.get("hasAttachments"):
            media_dir.mkdir(parents=True, exist_ok=True)
            for r in rows:
                oid, gid = r.get("objectid"), r.get("globalid")
                if oid is None or not gid:
                    continue
                for i, att in enumerate(lyr.attachments.get_list(oid=oid)):
                    q = (att.get("keywords") or "file").strip()
                    ext = Path(att.get("name", "a.jpg")).suffix or ".jpg"
                    dest = media_dir / f"{gid}_{q}_{i + 1}{ext}"
                    if not dest.exists():
                        tmp = lyr.attachments.download(oid=oid, attachment_id=att["id"], save_path=str(media_dir))
                        tmp = tmp[0] if isinstance(tmp, list) else tmp
                        os.replace(tmp, dest)
                    ds["media"].append({"parent": gid, "question": q, "path": str(dest)})
    # 位置が WebMercator の場合は緯度経度へ
    for r in ds["main"] + ds["reports"]:
        if r.get("lon") is not None and abs(r["lon"]) > 180:
            x, y = r["lon"], r["lat"]
            r["lon"] = x / 6378137.0 * 180 / math.pi
            r["lat"] = (2 * math.atan(math.exp(y / 6378137.0)) - math.pi / 2) * 180 / math.pi
    return ds


# ---------------------------------------------------------------- キャッシュ・補正
def save_dataset(ds: dict, project: str) -> Path:
    p = project_dir(project) / "dataset.json"
    p.write_text(json.dumps(ds, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return p


def load_dataset(project: str) -> dict:
    p = project_dir(project) / "dataset.json"
    if not p.exists():
        raise FileNotFoundError(f"案件「{project}」のデータがありません。先に survey_fetch か survey_load_export を実行してください")
    ds = json.loads(p.read_text(encoding="utf-8"))
    _apply_overrides(ds, project)
    _apply_facility_labels(ds)
    return ds


def _overrides_path(project: str) -> Path:
    return project_dir(project) / "corrections.json"


def add_override(project: str, globalid: str, field: str, value, note: str = "") -> int:
    p = _overrides_path(project)
    lst = json.loads(p.read_text(encoding="utf-8")) if p.exists() else []
    lst.append({"globalid": globalid.strip("{}").lower(), "field": field, "value": value, "note": note,
                "at": _dt.datetime.now().isoformat(timespec="seconds")})
    p.write_text(json.dumps(lst, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(lst)


def _apply_overrides(ds: dict, project: str):
    p = _overrides_path(project)
    if not p.exists():
        return
    idx = {r.get("globalid"): r for k in ("main", "fixtures", "reports") for r in ds[k]}
    for o in json.loads(p.read_text(encoding="utf-8")):
        r = idx.get(o["globalid"])
        if r is not None:
            r[o["field"]] = o["value"]


def media_for(ds: dict, globalid: str, question: str | None = None) -> list[str]:
    """question="rp_photo" なら rp_photo・rp_photo2・rp_photo3 をまとめて返す。"""
    return [m["path"] for m in ds["media"]
            if m["parent"] == globalid and (question is None or re.fullmatch(re.escape(question) + r"\d*", m["question"]))]


# ---------------------------------------------------------------- スマホアプリ（LED現調）の書き出しzipから
def load_app_zip(zip_path: str, out_dir: Path) -> dict:
    """スマホアプリが書き出した zip（data.json＋media/）を展開して dataset 形式にする。"""
    import zipfile
    zp = Path(zip_path)
    dest = out_dir / "app_import" / zp.stem
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zp) as z:
        for info in z.infolist():
            target = (dest / info.filename).resolve()
            if not str(target).startswith(str(dest.resolve())):  # zip の外へ書き出すパスは無視
                continue
            if info.is_dir():
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(z.read(info))
    data = json.loads((dest / "data.json").read_text(encoding="utf-8"))
    if data.get("format") != "led_gencho_app":
        raise ValueError("LED現調アプリの書き出しzipではありません")
    ds = empty_dataset(f"app:{zp.name}")
    ds["meta"]["exported_at"] = data.get("exported_at")
    for k in ("main", "fixtures", "reports"):
        ds[k] = [{kk: vv for kk, vv in r.items() if vv not in ("", None)} for r in data.get(k, [])]
    ds["media"] = [{"parent": m["parent"], "question": m["question"], "path": str(dest / m["file"]),
                    "taken_at": m.get("taken_at")} for m in data.get("media", [])]
    return ds


def merge_datasets(base: dict, new: dict) -> dict:
    """同じ記録（globalid）は更新日時が新しいほうを残し、写真は重複なく足し合わせる。"""
    for k in ("main", "fixtures", "reports"):
        idx = {r.get("globalid"): r for r in base.get(k, [])}
        for r in new.get(k, []):
            old = idx.get(r.get("globalid"))
            if old is None or (r.get("updated") or "") >= (old.get("updated") or ""):
                idx[r.get("globalid")] = r
        base[k] = list(idx.values())
    seen = {(m["parent"], m["question"], m.get("taken_at"), Path(m["path"]).name) for m in base.get("media", [])}
    for m in new.get("media", []):
        key = (m["parent"], m["question"], m.get("taken_at"), Path(m["path"]).name)
        if key not in seen:
            base["media"].append(m)
            seen.add(key)
    base["meta"]["source"] = f'{base["meta"].get("source")} + {new["meta"].get("source")}'
    return base
