#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""起一个本地静态站点，给「点一下就能下载」的交付页用。

    python tools/serve_downloads.py            # 默认 0.0.0.0:8000
    python tools/serve_downloads.py 8080

要点：
  · 绑定 0.0.0.0（预览代理才能访问到）
  · 访问 / 自动跳到 /site/index.html（下载页）
  · 不加任何 host / origin 白名单，允许被 iframe 嵌入
  · 静态文件直接从仓库读，名单/PPT 更新后刷新页面即可
"""
import os
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8000


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.path = "/site/index.html"
        return super().do_GET()

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, fmt, *args):
        sys.stdout.write("%s - %s\n" % (self.address_string(), fmt % args))
        sys.stdout.flush()


if __name__ == "__main__":
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"下载页: http://0.0.0.0:{PORT}/  (根目录 {ROOT})", flush=True)
    srv.serve_forever()
