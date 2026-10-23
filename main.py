# -*- coding: utf-8 -*-
"""نظام البصمة - تطبيق ويندوز مستقل (exe): واجهة + وسيط الأجهزة + مزامنة تلقائية + أيقونة بجوار الساعة."""
import os
import sys
import threading
import time
import webbrowser

if sys.stdout is None:  # تطبيق نوافذ بلا كونسول
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import fp_bridge as fb

APP_NAME = "FPSystem"
PORT = int(fb.CFG["port"])
URL = "http://127.0.0.1:%d/app" % PORT
INTERVAL = max(1, int(fb.CFG.get("interval_minutes", 15)))
STATE = {"quitting": False, "window": None, "icon": None, "srv": None}


# ---------- التشغيل مع ويندوز ----------
def _run_key():
    import winreg
    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0,
                          winreg.KEY_ALL_ACCESS)


def autostart_enabled():
    try:
        import winreg
        with _run_key() as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except Exception:
        return False


def set_autostart(on):
    try:
        import winreg
        with _run_key() as k:
            if on:
                winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, '"%s" --minimized' % sys.executable)
            else:
                try:
                    winreg.DeleteValue(k, APP_NAME)
                except FileNotFoundError:
                    pass
    except Exception as e:
        fb.log.error("تعذر تعديل التشغيل مع ويندوز: %s", e)


# ---------- المزامنة الخلفية ----------
def sync_loop():
    fb.MODE = "app"
    time.sleep(10)
    while not STATE["quitting"]:
        try:
            if fb.has_key():
                fb.run_sync()
        except Exception as e:
            fb.log.error("خطأ في المزامنة: %s", e)
        for _ in range(INTERVAL * 60):
            if STATE["quitting"]:
                return
            time.sleep(1)


def sync_now():
    def _go():
        try:
            if fb.has_key():
                fb.run_sync(force=True)
        except Exception as e:
            fb.log.error("خطأ في المزامنة اليدوية: %s", e)
    threading.Thread(target=_go, daemon=True).start()


# ---------- الإيقاف ----------
def quit_app(*_):
    STATE["quitting"] = True
    try:
        if STATE["icon"]:
            STATE["icon"].stop()
    except Exception:
        pass
    try:
        if STATE["window"]:
            STATE["window"].destroy()
    except Exception:
        pass
    threading.Timer(2.0, lambda: os._exit(0)).start()


def show_window(*_):
    w = STATE["window"]
    if w:
        try:
            w.show()
            w.restore()
            return
        except Exception:
            pass
    webbrowser.open(URL)


# ---------- أيقونة الإشعارات ----------
def start_tray():
    try:
        import pystray
        from PIL import Image
        img = Image.open(os.path.join(fb.RES, "icon.ico")).convert("RGBA").resize((64, 64))
        menu = pystray.Menu(
            pystray.MenuItem("فتح النظام", show_window, default=True),
            pystray.MenuItem("سحب ومزامنة الآن", lambda *_: sync_now()),
            pystray.MenuItem("التشغيل مع ويندوز", lambda *_: set_autostart(not autostart_enabled()),
                             checked=lambda _i: autostart_enabled()),
            pystray.MenuItem("خروج", quit_app),
        )
        icon = pystray.Icon(APP_NAME, img, "نظام البصمة", menu)
        STATE["icon"] = icon
        icon.run_detached()
        return True
    except Exception as e:
        fb.log.warning("تعذر تشغيل أيقونة الإشعارات: %s", e)
        return False


def main():
    fb.setup_log()
    minimized = "--minimized" in sys.argv
    try:
        srv = fb.make_server()
    except OSError:  # نسخة أخرى تعمل: افتح واجهتها فقط
        if not minimized:
            webbrowser.open(URL)
        return 0
    STATE["srv"] = srv
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    threading.Thread(target=sync_loop, daemon=True).start()
    fb.log.info("نظام البصمة يعمل: %s | بيانات المستخدم: %s", URL, fb.HERE)

    tray = start_tray()
    try:
        import webview
        try:
            webview.settings["ALLOW_DOWNLOADS"] = True  # تصدير Excel/PDF
        except Exception:
            pass
        w = webview.create_window("نظام البصمة", URL, width=1360, height=860, min_size=(960, 640), hidden=minimized)
        STATE["window"] = w

        def on_closing():
            if STATE["quitting"] or not tray:
                return True
            w.hide()  # يبقى يعمل في الخلفية؛ الخروج من أيقونة الإشعارات
            return False

        w.events.closing += on_closing
        webview.start(storage_path=os.path.join(fb.HERE, "webview"), private_mode=False)
    except Exception as e:  # WebView2 غير متوفر: نستعمل المتصفح الافتراضي
        fb.log.warning("تعذر فتح النافذة المدمجة (%s). سيُفتح المتصفح الافتراضي", e)
        if not minimized:
            webbrowser.open(URL)
        while not STATE["quitting"]:
            time.sleep(1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
