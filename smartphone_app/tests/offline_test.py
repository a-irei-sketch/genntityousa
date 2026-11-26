# -*- coding: utf-8 -*-
"""一度開いたあと、通信を切っても起動でき、入力済みデータが残るかを確認する。"""
import http.server, socketserver, sys, threading, time
from functools import partial
from playwright.sync_api import sync_playwright
DIST = sys.argv[1]
srv = socketserver.TCPServer(("127.0.0.1", 0), partial(http.server.SimpleHTTPRequestHandler, directory=DIST))
PORT = srv.server_address[1]
threading.Thread(target=srv.serve_forever, daemon=True).start()
with sync_playwright() as pw:
    br = pw.chromium.launch(executable_path="/usr/bin/google-chrome", args=["--no-sandbox"])
    ctx = br.new_context(viewport={"width": 390, "height": 844}, is_mobile=True)
    page = ctx.new_page()
    page.goto(f"http://127.0.0.1:{PORT}/index.html")
    page.wait_for_selector("text=施設を選ぶ")
    page.fill('[data-k="reporter"]', "オフライン確認")
    page.wait_for_function("navigator.serviceWorker.controller !== null || (location.reload(), false)", timeout=15000)
    time.sleep(1)
    srv.shutdown(); srv.server_close()          # サーバーを止める＝圏外
    ctx.set_offline(True)
    page.goto(f"http://127.0.0.1:{PORT}/index.html#/")
    page.wait_for_selector("text=施設を選ぶ", timeout=10000)
    v = page.input_value('[data-k="reporter"]')
    print("OFFLINE_START OK, reporter =", v)
    br.close()
