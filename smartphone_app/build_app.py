# -*- coding: utf-8 -*-
"""スマホアプリ（PWA）を組み立てる。
マスタ（施設・器具・LED置換・点灯時間・選択肢）を survey123_form/media の CSV と MCP の fields.json から読み込み、
dist/ に公開用ファイル一式（index.html・sw.js・manifest・アイコン）を書き出す。

使い方:  python build_app.py
マスタCSVを直したら、もう一度実行して dist/ を公開し直す。
"""
import csv
import json
from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).parent
ROOT = HERE.parent
MEDIA = ROOT / "survey123_form" / "media"
FIELDS = ROOT / "mcp_server" / "led_survey_mcp" / "fields.json"
DIST = HERE / "dist"


def read_csv(name):
    with open(MEDIA / name, encoding="utf-8-sig") as fp:
        return list(csv.DictReader(fp))


def main():
    choices = json.loads(FIELDS.read_text(encoding="utf-8"))["choices"]
    masters = {
        "choices": choices,
        "facilities": [{"name": r["name"], "label": r["label"], "area": r["floor_area_m2"], "kind": r["kind"]}
                       for r in read_csv("facilities.csv")],
        "fixtures": [{"name": r["name"], "label": r["label"], "watt": r["watt"], "shape": r["shape"]}
                     for r in read_csv("fixture_types.csv")],
        "rules": {r["key"]: {"led_code": r["led_code"], "led_watt": r["led_watt"]} for r in read_csv("led_rules.csv")},
        "hours": {r["name"]: {"hours_day": r["hours_day"], "days_year": r["days_year"]} for r in read_csv("usage_hours.csv")},
    }
    tpl = (HERE / "src" / "index.template.html").read_text(encoding="utf-8")
    html = tpl.replace("/*__MASTERS__*/null", json.dumps(masters, ensure_ascii=False))
    version = html.split('APP_VERSION = "')[1].split('"')[0]
    DIST.mkdir(exist_ok=True)
    (DIST / "index.html").write_text(html, encoding="utf-8")

    (DIST / "manifest.webmanifest").write_text(json.dumps({
        "name": "LED現況調査", "short_name": "LED現調", "start_url": "./index.html", "scope": "./",
        "display": "standalone", "orientation": "portrait", "background_color": "#f3f4f6", "theme_color": "#1F3864",
        "icons": [{"src": "icon-192.png", "sizes": "192x192", "type": "image/png"},
                  {"src": "icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}],
    }, ensure_ascii=False, indent=1), encoding="utf-8")

    # オフラインでも起動できるよう、アプリ本体をキャッシュする（版が変わると入れ替わる）
    (DIST / "sw.js").write_text(f"""const CACHE = "led-gencho-{version}";
const FILES = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
self.addEventListener("install", e => {{ e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); }});
self.addEventListener("activate", e => {{ e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); }});
self.addEventListener("fetch", e => {{
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => {{ const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }})
    .catch(() => caches.match(e.request, {{ignoreSearch: true}}).then(r => r || caches.match("./index.html"))));
}});
""", encoding="utf-8")

    for size in (192, 512):
        im = Image.new("RGB", (size, size), "#1F3864")
        d = ImageDraw.Draw(im)
        s = size / 192
        d.ellipse([56 * s, 30 * s, 136 * s, 110 * s], fill="#FFD54F")           # 電球
        d.rectangle([76 * s, 104 * s, 116 * s, 128 * s], fill="#FFD54F")
        d.rectangle([78 * s, 132 * s, 114 * s, 140 * s], fill="#cfd8dc")
        d.rectangle([82 * s, 144 * s, 110 * s, 152 * s], fill="#cfd8dc")
        for i in range(3):                                                    # 台帳の線
            d.rectangle([40 * s, (160 + i * 0) * s, 152 * s, 166 * s], fill="#ffffff")
        im.save(DIST / f"icon-{size}.png")
    print("built", DIST, "version", version, f"{len(html) // 1024}KB")


if __name__ == "__main__":
    main()
