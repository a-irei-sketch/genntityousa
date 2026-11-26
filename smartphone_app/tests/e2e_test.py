# -*- coding: utf-8 -*-
"""スマホ画面サイズの Chrome（疑似カメラ・マイク・GPS）で、入力→撮影→書き出しまでを通しで確認する。
使い方: python tests/e2e_test.py <distフォルダ> <出力フォルダ>"""
import http.server
import socketserver
import sys
import threading
import time
from functools import partial
from pathlib import Path

from playwright.sync_api import sync_playwright

DIST, OUT = Path(sys.argv[1]), Path(sys.argv[2])
OUT.mkdir(parents=True, exist_ok=True)
handler = partial(http.server.SimpleHTTPRequestHandler, directory=str(DIST))
socketserver.TCPServer.allow_reuse_address = True
srv = socketserver.TCPServer(("127.0.0.1", 0), handler)
PORT = srv.server_address[1]
threading.Thread(target=srv.serve_forever, daemon=True).start()
shots = []


def shot(page, name):
    p = OUT / f"{len(shots) + 1:02d}_{name}.png"
    page.screenshot(path=str(p), full_page=True)
    shots.append(p)


with sync_playwright() as pw:
    br = pw.chromium.launch(executable_path="/usr/bin/google-chrome", args=[
        "--use-fake-device-for-media-stream", "--use-fake-ui-for-media-stream", "--no-sandbox"])
    ctx = br.new_context(viewport={"width": 390, "height": 844}, device_scale_factor=2, is_mobile=True, has_touch=True,
                         permissions=["camera", "microphone", "geolocation"],
                         geolocation={"latitude": 25.8467, "longitude": 131.2366, "accuracy": 8}, accept_downloads=True)
    page = ctx.new_page()
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.on("console", lambda m: m.type == "error" and errors.append(m.text))
    page.goto(f"http://127.0.0.1:{PORT}/index.html")
    page.wait_for_selector("text=施設を選ぶ")
    page.fill('[data-k="reporter"]', "伊禮")
    page.fill('[data-k="project"]', "南大東村LED化現況調査")
    time.sleep(.4)
    shot(page, "home")

    page.click("text=在所集会所")
    page.select_option('[data-k="newfloor"]', "1F")
    page.wait_for_selector("text=階の確認")
    page.fill('[data-k="door_count"]', "3")
    page.dispatch_event('[data-k="door_count"]', "change")
    time.sleep(.6)
    page.click("#addroom")
    page.wait_for_selector("text=器具（種類・取付方法ごと）")
    page.fill('[data-k="room_name"]', "集会室")

    def camera(tag_label=None, n=1, q=None):
        page.click(f'[data-cam="{q}"]')
        page.wait_for_selector("#cam:not([hidden])")
        page.wait_for_function("document.querySelector('#video').videoWidth > 0", timeout=10000)
        if tag_label:
            page.click(f'#camtags button:has-text("{tag_label}")')
        for _ in range(n):
            page.click("#shutter")
            time.sleep(.35)
        return page

    camera(n=2, q="room_photo")
    shot(page, "camera")
    page.click("#camdone")
    page.wait_for_selector("text=器具（種類・取付方法ごと）")
    page.select_option('[data-k="room_use"]', "meeting")
    time.sleep(.7)
    page.fill('[data-k="depth_m"]', "12")
    page.fill('[data-k="width_m"]', "9")
    page.fill('[data-k="ceil_h_m"]', "3.2")
    camera(q="ceil_photo", n=1)
    page.click("#camdone")
    page.wait_for_selector("#addfx")
    time.sleep(.4)
    shot(page, "room")

    # 器具
    page.click("#addfx")
    page.wait_for_selector("#q")
    page.fill("#q", "40形")
    page.click('#sl button:has-text("直管40形×2（グロー）")')
    page.click('[data-chips="fx_mount"] button:has-text("直付")')
    page.click('[data-chips="fx_method"] button:has-text("配列")')
    time.sleep(.7)
    for _ in range(4):
        page.click('[data-step="fx_cols"][data-d="1"]')
    for _ in range(3):
        page.click('[data-step="fx_rows"][data-d="1"]')
    page.select_option('[data-k="fx_shape"]', "inverted_fuji")
    page.select_option('[data-k="fx_lamp_kind"]', "fl")
    page.fill('[data-k="fx_height_m"]', "3.2")
    page.click('[data-chips="fx_aerial"] button:has-text("脚立")')
    page.click('[data-chips="fx_removal"] button:has-text("器具ごと交換")')
    page.click('[data-chips="fx_removal"] button:has-text("PCB")')
    page.click('[data-chips="fx_cond"] button:has-text("劣化")')
    page.click('[data-chips="fx_defects"] button:has-text("塩害")')
    camera(q="fx_photo", n=2)
    page.click('#camtags button:has-text("銘板")')
    page.click("#shutter")
    time.sleep(.4)
    page.click("#camdone")
    page.wait_for_selector("#res")
    time.sleep(.4)
    res = page.inner_text("#res")
    shot(page, "fixture")
    page.click("#donefx")
    page.wait_for_selector("#addrp")

    # 報告事項（メモ・GPS・音声）
    page.click("#addrp")
    page.wait_for_selector("#gps")
    page.click('[data-chips="rp_cat"] button:has-text("要対応")')
    page.fill('[data-k="rp_memo"]', "天井に雨漏り跡。器具交換前に補修が必要")
    page.click("#gps")
    page.wait_for_function("document.querySelector('#loc').textContent.includes('記録済み')", timeout=10000)
    page.click("#rec")
    time.sleep(1.5)
    page.click("#rec")
    page.wait_for_selector("audio", timeout=10000)
    camera(q="rp_photo", n=1)
    page.click("#camdone")
    page.wait_for_selector("#donerp")
    time.sleep(.4)
    shot(page, "report")
    page.click("#donerp")
    page.wait_for_selector("#addrp")
    page.click('[data-chips="emergency"] button:has-text("無")')
    page.click('[data-chips="room_status"] button:has-text("調査済")')
    time.sleep(.4)

    # 部屋を複製 → 2部屋目
    page.click("#copyroom")
    page.wait_for_selector("#addfx")
    page.fill('[data-k="room_name"]', "和室")
    time.sleep(.4)
    page.click("#next")
    page.wait_for_selector("#addroom")
    time.sleep(.4)
    shot(page, "floor")

    # 書き出し
    page.click("#menu")
    page.wait_for_selector("#go")
    page.click("#go")
    page.wait_for_selector("#out a.btn", timeout=20000)
    shot(page, "export")
    with page.expect_download() as d:
        page.click("#out a.btn")
    zp = OUT / d.value.suggested_filename
    d.value.save_as(zp)

    # 再読み込みしてもデータが残るか
    page.goto(f"http://127.0.0.1:{PORT}/index.html#/")
    page.wait_for_selector("text=施設を選ぶ")
    persisted = page.inner_text("main")
    shot(page, "home_after")
    br.close()
srv.shutdown()
print("RESULT", res)
print("ZIP", zp, zp.stat().st_size)
print("PERSIST", "在所集会所" in persisted and "2室" in persisted)
print("ERRORS", errors)
