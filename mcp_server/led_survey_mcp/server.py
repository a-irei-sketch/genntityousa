# -*- coding: utf-8 -*-
"""LED室内現況調査 Survey123 MCP サーバー。

Claude デスクトップアプリから、Survey123 の回答（写真・音声・位置を含む）を取得し、
器具台帳・写真帳・簡易配置図・網羅チェック・地図用データを作る。
ArcGIS へは読み取り（クエリ・添付ダウンロード）だけを行い、書き込みはしない。
"""
from __future__ import annotations

import datetime as _dt
import json
from pathlib import Path

from mcp.server.fastmcp import FastMCP, Image
from mcp.types import ToolAnnotations

from . import data, outputs

mcp = FastMCP("led-survey123")
RO = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
LOCAL = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=True, openWorldHint=False)


def _out(project: str, name: str) -> Path:
    d = data.project_dir(project) / "出力" / _dt.date.today().isoformat()
    d.mkdir(parents=True, exist_ok=True)
    return d / name


@mcp.tool(annotations=RO)
def survey_fetch(project: str, item_id: str, where: str = "1=1", download_media: bool = True,
                 since: str | None = None) -> str:
    """ArcGIS Online の Survey123（LED室内現況調査）から回答を取得し、作業フォルダに保存する。

    project: 保存先の案件名（作業フォルダ名になる）
    item_id: Survey123 フォーム、またはそのフィーチャサービスのアイテムID
    where: 絞り込み（例 "project='○○市役所'"、"building='本庁舎'"）
    download_media: 写真・音声もダウンロードする
    since: この日付（YYYY-MM-DD）以降に編集された回答だけ取得
    クエリと添付ダウンロードのみで、ArcGISのクレジットは消費しない。"""
    out = data.project_dir(project)
    ds = data.fetch_from_arcgis(item_id, out, where=where, download_media=download_media, since=since)
    p = data.save_dataset(ds, project)
    return json.dumps({"saved": str(p), "records": len(ds["main"]), "fixtures": len(ds["fixtures"]),
                       "reports": len(ds["reports"]), "media": len(ds["media"])}, ensure_ascii=False)


@mcp.tool(annotations=LOCAL)
def survey_load_export(project: str, excel_path: str, media_dir: str | None = None) -> str:
    """Survey123 の Web サイトで書き出した Excel（データ → エクスポート → Excel）を読み込む。
    ArcGIS API が使えない場合の代替。写真を使うときは media_dir に survey_fetch 形式の写真フォルダを指定。"""
    ds = data.load_excel_export(excel_path)
    n = data.attach_media_folder(ds, media_dir) if media_dir else 0
    p = data.save_dataset(ds, project)
    return json.dumps({"saved": str(p), "records": len(ds["main"]), "fixtures": len(ds["fixtures"]),
                       "reports": len(ds["reports"]), "media": n}, ensure_ascii=False)


@mcp.tool(annotations=LOCAL)
def app_load_export(project: str, zip_path: str, merge: bool = True) -> str:
    """スマホアプリ「LED現調」が書き出した zip を取り込む。
    merge=True なら、すでに取り込んだデータに足し合わせる（複数のスマホ・複数日の zip をまとめられる。
    同じ記録は更新日時が新しいほうを残す）。merge=False なら置き換える。"""
    out = data.project_dir(project)
    ds = data.load_app_zip(zip_path, out)
    p = out / "dataset.json"
    if merge and p.exists():
        base = json.loads(p.read_text(encoding="utf-8"))
        ds = data.merge_datasets(base, ds)
    p = data.save_dataset(ds, project)
    return json.dumps({"saved": str(p), "records": len(ds["main"]), "fixtures": len(ds["fixtures"]),
                       "reports": len(ds["reports"]), "media": len(ds["media"])}, ensure_ascii=False)


@mcp.tool(annotations=RO)
def survey_summary(project: str) -> str:
    """取得済みデータの概要：部屋数、器具台数（タイプ別）、年間削減電力量、要確認の部屋、報告事項、部屋の漏れ警告。"""
    return json.dumps(outputs.summary(data.load_dataset(project)), ensure_ascii=False, indent=1)


@mcp.tool(annotations=RO)
def coverage_check(project: str) -> str:
    """階ごとに「登録した部屋数」と「扉の数・申告部屋数」を突き合わせ、部屋の調査漏れを判定する（図面がない現場用）。"""
    return json.dumps(outputs.coverage(data.load_dataset(project)), ensure_ascii=False, indent=1)


@mcp.tool(annotations=RO)
def list_records(project: str, building: str | None = None, floor: str | None = None,
                 room_no: str | None = None, rec_type: str | None = None) -> str:
    """部屋・階の確認・現場写真報告の記録を一覧する。器具と報告事項、写真の保存先も返す。"""
    ds = data.load_dataset(project)
    out = []
    for r in ds["main"]:
        if building and r.get("building") != building or floor and r.get("floor") != floor \
                or room_no and r.get("room_no") != room_no or rec_type and r.get("rec_type") != rec_type:
            continue
        g = r.get("globalid")
        out.append({**{k: v for k, v in r.items() if v is not None},
                    "media": [m for m in ds["media"] if m["parent"] == g],
                    "fixtures": [{**f, "media": data.media_for(ds, f.get("globalid"))}
                                 for f in ds["fixtures"] if f.get("parentglobalid") == g],
                    "reports": [{**p, "media": data.media_for(ds, p.get("globalid"))}
                                for p in ds["reports"] if p.get("parentglobalid") == g]})
    return json.dumps(out[:200], ensure_ascii=False, indent=1, default=str)


@mcp.tool(annotations=RO)
def view_photo(path: str, max_px: int = 1280) -> Image:
    """写真を表示する。銘板の型番・室名札の読み取りや、器具タイプ・台数の確認に使う。"""
    from io import BytesIO

    from PIL import Image as PILImage, ImageOps
    im = ImageOps.exif_transpose(PILImage.open(path)).convert("RGB")
    im.thumbnail((max_px, max_px))
    buf = BytesIO()
    im.save(buf, format="JPEG", quality=85)
    return Image(data=buf.getvalue(), format="jpeg")


@mcp.tool(annotations=LOCAL)
def record_correction(project: str, globalid: str, field: str, value: str, note: str = "") -> str:
    """写真から読み取った型番・部屋名などの補正を記録する（PC側のみ。ArcGISのデータは変更しない）。
    以降の帳票作成で補正値が使われる。例: field='fx_model', value='FSA-41521-PS9'"""
    n = data.add_override(project, globalid, field, value, note)
    return f"補正を記録しました（累計{n}件）"


@mcp.tool(annotations=LOCAL)
def build_ledger(project: str, unit_price: float = 30.0, co2_factor: float = 0.434) -> str:
    """器具台帳Excelを作る：条件・部屋一覧・器具台帳・数量集計・省エネ算定・網羅チェック・報告事項。
    unit_price=電力量単価（円/kWh）、co2_factor=CO2排出係数（kg-CO2/kWh）。どちらも Excel の条件シートで後から変更可。"""
    p = outputs.build_ledger(data.load_dataset(project), _out(project, "器具台帳.xlsx"), unit_price, co2_factor)
    return str(p)


@mcp.tool(annotations=LOCAL)
def build_photo_album(project: str) -> str:
    """写真帳PDFを作る：部屋ごとに室名札・天井・器具・銘板の写真と器具一覧、続けて階の確認・現場写真報告。"""
    return str(outputs.build_photo_album(data.load_dataset(project), _out(project, "写真帳.pdf")))


@mcp.tool(annotations=LOCAL)
def build_layout(project: str) -> str:
    """簡易配置図を作る（建物ごとに DXF と PDF）。部屋の縦横寸法と器具の配列から概略位置を割り付ける。"""
    paths = outputs.build_layouts(data.load_dataset(project), _out(project, "x").parent)
    return "\n".join(str(p) for p in paths)


@mcp.tool(annotations=LOCAL)
def export_map_points(project: str) -> str:
    """調査位置と報告事項（撮影場所）をGeoJSONで書き出す。ArcGIS・QGIS・hazard_map_mcp などで地図に重ねられる。"""
    return str(outputs.build_geojson(data.load_dataset(project), _out(project, "報告地点.geojson")))


@mcp.tool(annotations=LOCAL)
def export_photos(project: str) -> str:
    """写真を施設ごとのフォルダに「施設名_階_場所_撮影箇所_撮影日_連番.jpg」で整理し、写真一覧.xlsxを作る（納品用）。"""
    return str(outputs.export_photos(data.load_dataset(project), _out(project, "x").parent / "写真整理"))


@mcp.tool(annotations=LOCAL)
def build_all(project: str, unit_price: float = 30.0, co2_factor: float = 0.434) -> str:
    """器具台帳・写真帳・簡易配置図・地図用データをまとめて作る。"""
    ds = data.load_dataset(project)
    ps = [outputs.build_ledger(ds, _out(project, "器具台帳.xlsx"), unit_price, co2_factor),
          outputs.build_photo_album(ds, _out(project, "写真帳.pdf")),
          *outputs.build_layouts(ds, _out(project, "x").parent),
          outputs.build_geojson(ds, _out(project, "報告地点.geojson")),
          outputs.export_photos(ds, _out(project, "x").parent / "写真整理")]
    return "\n".join(str(p) for p in ps)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
