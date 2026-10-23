# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

datas = [("app/web/index.html", "web"), ("app/fp_config.json", "."), ("assets/icon.ico", ".")]
if os.path.exists("app/fp_key.txt"):          # يُنشأ أثناء البناء من GitHub Secret (اختياري)
    datas.append(("app/fp_key.txt", "."))
binaries, hidden = [], ["zk", "pystray._win32"]
for pkg in ("webview", "pythonnet", "clr_loader"):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hidden += h

a = Analysis(["app/main.py"], pathex=["app"], binaries=binaries, datas=datas,
             hiddenimports=hidden, hookspath=[], runtime_hooks=[], excludes=["tkinter"])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
          name="FPSystem", console=False, icon="assets/icon.ico", upx=False)
