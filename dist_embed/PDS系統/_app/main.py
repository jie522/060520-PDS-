"""
製令查詢系統 — GUI 版啟動器
使用 Edge --app 模式開啟，看起來像原生視窗（無網址列、無分頁）。
不依賴 pywebview/pythonnet，任何 Windows 10 電腦都能直接使用。
"""
import sys
import os
import threading
import time
import subprocess
import traceback

# ── 錯誤日誌 ──────────────────────────────────────────────
_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'error.log')

def _log_error(msg):
    try:
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass

# ── 路徑初始化（Embeddable / 一般執行）────────────────────
if getattr(sys, 'frozen', False):
    _APP_DIR = os.path.dirname(sys.executable)
    sys.path.insert(0, _APP_DIR)
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

_PARENT_DIR = os.path.dirname(_APP_DIR)
for _d in [_APP_DIR, _PARENT_DIR]:
    if _d not in sys.path:
        sys.path.insert(0, _d)

_py_exe_dir = os.path.dirname(os.path.abspath(sys.executable))
for _sp in [
    os.path.join(_py_exe_dir, 'Lib', 'site-packages'),
    os.path.join(_py_exe_dir, 'lib', 'site-packages'),
]:
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

# ── 讀取 config ──────────────────────────────────────────
try:
    import config
    PORT = getattr(config, 'FLASK_PORT', 5002)
    WIN_TITLE = getattr(config, 'WINDOW_TITLE', '詠基-加工部查詢系統PDIS')
except Exception:
    PORT = 5002
    WIN_TITLE = '詠基-加工部查詢系統PDIS'

URL = f"http://127.0.0.1:{PORT}/"


# ── 瀏覽器尋找與視窗開啟（搬到最前面，讓單例鎖檢查也能用）──────────
def _find_edge():
    """尋找 Microsoft Edge 執行檔路徑"""
    candidates = [
        os.path.join(os.environ.get('ProgramFiles(x86)', ''), 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
        os.path.join(os.environ.get('ProgramFiles', ''), 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def _find_chrome():
    """尋找 Google Chrome 執行檔路徑"""
    candidates = [
        os.path.join(os.environ.get('ProgramFiles', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('ProgramFiles(x86)', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def _is_flask_alive():
    """檢查 Flask 是否已在運行"""
    import urllib.request
    try:
        urllib.request.urlopen(f'http://127.0.0.1:{PORT}/', timeout=2)
        return True
    except Exception:
        return False


def _find_app_hwnd():
    """用 EnumWindows 找視窗標題「包含」關鍵字的視窗（標題會被 JS 改成「Maxclaw PDS系統 V...」，
    FindWindowW 只能精確比對會找不到）"""
    import ctypes as _c
    _u32 = _c.windll.user32
    _keywords = ('Maxclaw PDS', '製令查詢', '加工部查詢系統PDS', WIN_TITLE)
    _found = []

    @_c.WINFUNCTYPE(_c.c_bool, _c.c_void_p, _c.c_void_p)
    def _enum_proc(hwnd, lparam):
        if not _u32.IsWindowVisible(hwnd):
            return True
        length = _u32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = _c.create_unicode_buffer(length + 1)
        _u32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value
        if any(k in title for k in _keywords):
            _found.append(hwnd)
        return True

    try:
        _u32.EnumWindows(_enum_proc, 0)
    except Exception:
        pass
    return _found[0] if _found else 0


def _open_app_window():
    """用 Edge/Chrome --app 模式開啟（看起來像原生視窗）"""
    browser = _find_edge() or _find_chrome()
    if browser:
        import json

        # ── 使用固定的 profile 目錄（避免每次新目錄觸發 Edge 初始化視窗）──
        user_data = os.path.abspath(os.path.join(_APP_DIR, '..', '_edge_data'))
        default_dir = os.path.join(user_data, 'Default')
        os.makedirs(default_dir, exist_ok=True)

        # ── 刪除 session 還原相關檔案，避免重開時跳出舊分頁/視窗 ──────
        for _name in ('Sessions', 'Last Session', 'Last Tabs', 'Current Session', 'Current Tabs'):
            _p = os.path.join(default_dir, _name)
            try:
                if os.path.isdir(_p):
                    import shutil as _shutil
                    _shutil.rmtree(_p, ignore_errors=True)
                elif os.path.isfile(_p):
                    os.remove(_p)
            except Exception:
                pass

        # ── 清除 favicon 快取：app-icon.png 換新後，視窗/工作列圖示才會跟著更新 ──
        for _name in ('Favicons', 'Favicons-journal'):
            _p = os.path.join(default_dir, _name)
            try:
                if os.path.isfile(_p):
                    os.remove(_p)
            except Exception:
                pass
        try:
            import shutil as _shutil
            _shutil.rmtree(os.path.join(default_dir, 'Web Applications'), ignore_errors=True)
        except Exception:
            pass

        # ── 寫入 Preferences：強制禁止 session 還原，防止舊視窗重現 ──────
        prefs_path = os.path.join(default_dir, 'Preferences')
        try:
            prefs = {}
            if os.path.isfile(prefs_path):
                with open(prefs_path, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
            prefs.setdefault('session', {})
            prefs['session']['restore_on_startup'] = 1   # 1=開新分頁，不還原 session
            prefs['session']['startup_urls'] = []
            prefs.setdefault('profile', {})
            prefs['profile']['exited_cleanly'] = True
            prefs['profile']['exit_type'] = 'Normal'
            with open(prefs_path, 'w', encoding='utf-8') as f:
                json.dump(prefs, f)
        except Exception:
            pass

        cache_bust = int(time.time())
        subprocess.Popen([
            browser,
            f'--app={URL}?_t={cache_bust}',
            f'--window-size=1400,900',
            '--start-maximized',
            f'--user-data-dir={user_data}',
            '--no-first-run',
            '--no-default-browser-check',
            '--no-restore-last-session',
            '--disable-restore-session-state',
            '--disable-extensions',
            '--disable-background-networking',
            '--disable-background-mode',
            '--aggressive-cache-discard',
        ])
        return True
    return False


# ── 最早期單例鎖（在 Flask 匯入前，防止快速雙擊出現兩個實例）──────
# 把 Mutex 放在這裡：Flask import 需要 2-5 秒，若等到 main() 才建立，
# 兩次點擊都會通過檢查，造成雙視窗。
_MUTEX_HANDLE = None
try:
    import ctypes as _c
    _k32 = _c.WinDLL('kernel32', use_last_error=True)
    _MUTEX_HANDLE = _k32.CreateMutexW(None, False, f"PDS_製令查詢_{PORT}")
    if _c.get_last_error() == 183:          # ERROR_ALREADY_EXISTS
        # 另一個實例已在跑 → 找視窗帶前台，然後立刻退出（不繼續匯入 Flask）
        try:
            _u32 = _c.windll.user32
            # 輪詢找視窗：第一個實例可能還在匯入 Flask（2-5秒），
            # 視窗尚未開出來，先等一下再判斷，避免誤判「視窗已關閉」而多開一個
            _hwnd = 0
            for _i in range(8):
                _hwnd = _find_app_hwnd()
                if _hwnd:
                    break
                time.sleep(0.5)
            if _hwnd:
                _u32.ShowWindow(_hwnd, 9)       # SW_RESTORE
                _u32.SetForegroundWindow(_hwnd)
            elif _is_flask_alive():
                # 視窗已被關閉，但 Flask 仍在跑 → 開新視窗
                _open_app_window()
        except Exception:
            pass
        sys.exit(0)
except SystemExit:
    raise
except Exception:
    pass

# ── 匯入 Flask app ────────────────────────────────────────
# 強制確保 _APP_DIR 是 sys.path 的第一個（最高優先）
if sys.path[0] != _APP_DIR:
    while _APP_DIR in sys.path:
        sys.path.remove(_APP_DIR)
    sys.path.insert(0, _APP_DIR)

print(f"[DEBUG] sys.path 前3項:")
for i, p in enumerate(sys.path[:5]):
    print(f"  [{i}] {p}")
print(f"[DEBUG] _APP_DIR = {_APP_DIR}")
print(f"[DEBUG] app.py 存在? {os.path.isfile(os.path.join(_APP_DIR, 'app.py'))}")

try:
    import app as flask_app
    print(f"[DEBUG] 載入的 app.py 位置: {flask_app.__file__}")
    print(f"[DEBUG] APP_VERSION = {getattr(flask_app, 'APP_VERSION', '未定義!')}")
    # 印出 Flask 所有路由
    print("[DEBUG] Flask 已註冊路由:")
    for rule in flask_app.app.url_map.iter_rules():
        print(f"  {rule.rule} -> {rule.endpoint}")
except Exception as e:
    _log_error(f"匯入失敗:\n{traceback.format_exc()}")
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, f"程式載入失敗：\n{e}\n\n詳見 error.log", WIN_TITLE, 0x10)
    except Exception:
        pass
    sys.exit(1)


def _kill_old_process_on_port(port):
    """殺掉佔用指定 port 的舊進程"""
    try:
        result = subprocess.run(
            ['netstat', '-ano', '-p', 'TCP'],
            capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.splitlines():
            if f'127.0.0.1:{port}' in line and 'LISTENING' in line:
                parts = line.split()
                pid = parts[-1]
                if pid.isdigit() and int(pid) != os.getpid():
                    print(f"[!] 發現舊進程 PID={pid} 佔用 port {port}，正在關閉...")
                    subprocess.run(['taskkill', '/F', '/PID', pid],
                                   capture_output=True, timeout=5)
                    time.sleep(1)
                    print(f"[!] 已關閉舊進程")
                    return True
    except Exception as e:
        print(f"[!] 檢查 port 時發生錯誤: {e}")
    return False


def _run_flask():
    """在背景執行緒中啟動 Flask"""
    import logging
    logging.getLogger('werkzeug').setLevel(logging.ERROR)
    try:
        # 綁哪個網卡由 app._flask_host() 決定：預設 127.0.0.1（跟以前一樣只有本機連得到），
        # config.MOBILE_ACCESS 打開時才綁 0.0.0.0 讓手機掃 QR 進來（見 docs/equipment-maintenance.md）
        flask_app.app.run(
            host=flask_app._flask_host(), port=PORT,
            debug=False, use_reloader=False, threaded=True,
        )
    except Exception as e:
        print(f"[!] Flask 啟動失敗: {e}")
        _log_error(f"Flask 啟動失敗: {e}\n{traceback.format_exc()}")


def _wait_for_flask():
    """等待 Flask 就緒"""
    import urllib.request
    for _ in range(30):
        try:
            urllib.request.urlopen(URL, timeout=1)
            return True
        except Exception:
            time.sleep(0.5)
    return False



def main():
    try:
        # ── 注意：單例 Mutex 已在 module level 建立（Flask import 之前），
        # 若本次到達這裡，代表本進程是第一個實例，直接往下執行。

        # ── 若 Flask 已在運行，直接開新視窗即可 ───────────
        if _is_flask_alive():
            print(f"[1] 偵測到 Flask 已在運行（port={PORT}），直接開啟視窗")
            if _open_app_window():
                print("[2] 視窗已開啟")
            else:
                import webbrowser
                webbrowser.open(URL)
        else:
            # 先殺掉佔用 port 的舊進程
            _kill_old_process_on_port(PORT)

            print(f"[1] 啟動 Flask (port={PORT})...")
            flask_thread = threading.Thread(target=_run_flask, daemon=True)
            flask_thread.start()

            threading.Thread(target=flask_app._preload_cache, daemon=True).start()

            print(f"[2] 等待 Flask 就緒 ({URL})...")
            if not _wait_for_flask():
                msg = "Flask 伺服器啟動逾時"
                print(f"[!] {msg}")
                _log_error(msg)
                try:
                    import ctypes
                    ctypes.windll.user32.MessageBoxW(0, msg, WIN_TITLE, 0x10)
                except Exception:
                    pass
                sys.exit(1)

            print("[3] Flask 就緒，開啟視窗...")
            if _open_app_window():
                print("[4] 已用 Edge/Chrome --app 模式開啟")
            else:
                import webbrowser
                webbrowser.open(URL)
                print("[4] 已用預設瀏覽器開啟")

        # 保持程式運行（Flask 在背景執行緒）
        print(f"[5] 系統運行中... 按 Ctrl+C 結束")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n結束程式")

    except Exception as e:
        _log_error(f"主程式錯誤:\n{traceback.format_exc()}")
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0, f"啟動失敗：\n{e}\n\n詳見 error.log", WIN_TITLE, 0x10)
        except Exception:
            pass
        sys.exit(1)


if __name__ == '__main__':
    main()
