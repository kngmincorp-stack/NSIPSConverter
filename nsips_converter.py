# -*- coding: utf-8 -*-
"""
NSIPS TXT → CSV 転記ツール
------------------------------------
レセコン(NSIPS)が出力するTXTファイルを監視し、新しいTXTが出来たら
内容を抽出して、元TXTと同名の .csv として出力フォルダへ保存する常駐GUIツール。

・処理済みTXTはファイル名で記録し、監視再開・再起動後も再処理しない
・Windowsスタートアップ登録チェックボックス
・パッチ更新システム(GitHub Releasesベース自動更新) / 更新履歴表示
依存: 標準ライブラリのみ (openpyxl不要)
"""

import os
import sys
import csv
import json
import ssl
import threading
import traceback
import datetime
import tempfile
import subprocess
import urllib.request

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import winreg  # Windows のみ
except ImportError:
    winreg = None

# ---------------------------------------------------------------------------
# パス解決 (PyInstaller onefile / スクリプト 両対応)
# ---------------------------------------------------------------------------
def base_dir():
    """設定を置く永続ディレクトリ(= exe/スクリプトのある場所)。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(name):
    """同梱リソース(VERSION/CHANGELOG)のパス。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        p = os.path.join(sys._MEIPASS, name)
        if os.path.exists(p):
            return p
    return os.path.join(base_dir(), name)


BASE_DIR      = base_dir()
CONFIG_PATH   = os.path.join(BASE_DIR, "config.json")
POLL_INTERVAL = 3.0                    # 監視ポーリング間隔(秒)
CSV_ENCODING  = "cp932"                # 出力CSVの文字コード (Shift-JIS)

# Windows スタートアップ登録
STARTUP_KEY   = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_NAME  = "NSIPSConverter"

# GitHub Releases 自動更新
GITHUB_REPO   = "kngmincorp-stack/NSIPSConverter"
API_URL       = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO
EXE_NAME      = "NSIPSConverter.exe"   # リリース資産 / 実行ファイル名
UPDATE_TIMEOUT = 15

try:
    import certifi
    _SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except Exception:
    try:
        _SSL_CTX = ssl.create_default_context()
    except Exception:
        _SSL_CTX = ssl._create_unverified_context()


def read_version():
    for p in (resource_path("VERSION"), os.path.join(BASE_DIR, "VERSION")):
        try:
            with open(p, encoding="utf-8") as f:
                t = f.read().strip()
            if t:
                return t
        except OSError:
            continue
    return "0.0.0"


APP_VERSION = read_version()


def read_changelog():
    for p in (resource_path("CHANGELOG.md"), os.path.join(BASE_DIR, "CHANGELOG.md")):
        try:
            with open(p, encoding="utf-8") as f:
                t = f.read().strip()
            if t:
                return t
        except OSError:
            continue
    return "更新履歴ファイル (CHANGELOG.md) が見つかりませんでした。"


def compare_versions(a, b):
    def parts(v):
        out = []
        for x in v.split("."):
            try:
                out.append(int(x))
            except ValueError:
                out.append(0)
        return out
    pa, pb = parts(a), parts(b)
    for x, y in zip(pa, pb):
        if x != y:
            return x - y
    return len(pa) - len(pb)


# ---------------------------------------------------------------------------
# TXT 解析
# ---------------------------------------------------------------------------
def _fmt_date(s):
    """'20260630' → '2026/06/30'。8桁でなければそのまま返す。"""
    s = (s or "").strip()
    if len(s) == 8 and s.isdigit():
        return "%s/%s/%s" % (s[0:4], s[4:6], s[6:8])
    return s


def parse_txt(path):
    """
    NSIPS TXT を解析して転記用 dict を返す。解析不可なら ValueError。

    「1,」行 = 患者情報 (固定位置)
        [1]患者番号 [2]カナ氏名 [3]漢字氏名 [4]性別(1男/2女) [5]生年月日8桁
        [6]郵便番号 [7]住所 → 出力しない(飛ばす)
        [14]保険者番号 [16]記号 [17]番号 [18]本人家族(1本人/2家族) [-1]枝番(行末)
    「2,」行 = 調剤情報
        [4]調剤(来局)日8桁。複数あれば最新を最終来局日とする。
    """
    with open(path, encoding="cp932") as f:
        text = f.read()

    row1 = None
    visit_dates = []
    for line in text.splitlines():
        cols = line.split(",")
        if not cols:
            continue
        rt = cols[0]
        if rt == "1" and row1 is None:
            row1 = cols
        elif rt == "2":
            if len(cols) > 4:
                d = cols[4].strip()
                if len(d) == 8 and d.isdigit():
                    visit_dates.append(d)

    if row1 is None or len(row1) < 19:
        raise ValueError("患者情報(『1,』行)が見つからないか項目が不足しています")

    def g(i):
        return row1[i].strip() if i < len(row1) else ""

    sex = {"1": "男", "2": "女"}.get(g(4), g(4))
    honke = {"1": "本人", "2": "家族"}.get(g(18), g(18))

    edaban = ""
    for v in reversed(row1):
        if v.strip() != "":
            edaban = v.strip()
            break

    last_visit = max(visit_dates) if visit_dates else ""

    return {
        "患者番号":   g(1),
        "カナ氏名":   g(2),
        "漢字氏名":   g(3),
        "生年月日":   _fmt_date(g(5)),
        "性別":       sex,
        "保険者番号": g(14),
        "記号":       g(16),
        "番号":       g(17),
        "枝番":       edaban,
        "本/家":      honke,
        "最終来局日": _fmt_date(last_visit),
    }


# 出力CSVの列順(=見本Excelの列順)
CSV_COLUMNS = [
    "患者番号", "カナ氏名", "漢字氏名", "生年月日", "性別", "保険者番号",
    "記号", "番号", "枝番", "本/家", "最終来局日",
]


def write_csv(data, out_path):
    """1患者分を見出し付きCSV(Shift-JIS)で出力。"""
    with open(out_path, "w", encoding=CSV_ENCODING, newline="", errors="replace") as f:
        w = csv.writer(f)
        w.writerow(CSV_COLUMNS)
        w.writerow([data.get(k, "") for k in CSV_COLUMNS])


# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "watch_folder": "",
    "output_folder": "",
    "processed": [],           # 処理済みTXTのファイル名リスト
    "last_seen_version": "",
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Windows スタートアップ登録
# ---------------------------------------------------------------------------
def startup_enabled():
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as k:
            val, _ = winreg.QueryValueEx(k, STARTUP_NAME)
            return bool(val)
    except OSError:
        return False


def set_startup(enable):
    """スタートアップ登録/解除。frozen(exe)時のみ意味を持つ。"""
    if winreg is None:
        raise OSError("Windows 以外では利用できません")
    exe = sys.executable
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if enable:
            winreg.SetValueEx(k, STARTUP_NAME, 0, winreg.REG_SZ, '"%s"' % exe)
        else:
            try:
                winreg.DeleteValue(k, STARTUP_NAME)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# 自動更新 (GitHub Releases / 単一exe置換)
# ---------------------------------------------------------------------------
class Updater:
    def __init__(self, app):
        self.app = app
        self._latest = None
        self._url = None
        self._busy = False

    def check_async(self, silent=True):
        threading.Thread(target=self._check, args=(silent,), daemon=True).start()

    def _check(self, silent):
        try:
            req = urllib.request.Request(
                API_URL, headers={"Accept": "application/vnd.github.v3+json"})
            with urllib.request.urlopen(req, timeout=UPDATE_TIMEOUT, context=_SSL_CTX) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            remote = str(data.get("tag_name", "")).lstrip("v")
            url = None
            for a in data.get("assets", []):
                n = a.get("name", "")
                if isinstance(n, str) and n.lower().endswith(".exe"):
                    url = a.get("browser_download_url")
                    break
            self.app.after(0, lambda: self._done(remote, url, silent))
        except Exception as e:
            self.app.after(0, lambda: self._fail(e, silent))

    def _fail(self, e, silent):
        if not silent:
            self.app.log("更新チェック失敗: %s" % e)

    def _done(self, remote, url, silent):
        if not remote:
            if not silent:
                self.app.log("リリース情報が取得できませんでした")
            return
        if compare_versions(remote, APP_VERSION) > 0 and url:
            self._latest, self._url = remote, url
            self.app.log("新しいバージョン v%s が利用可能です（現在 v%s）" % (remote, APP_VERSION))
            self._ask()
        elif not silent:
            self.app.log("最新版です (v%s)" % APP_VERSION)

    def _ask(self):
        if not (self._url and self._latest):
            return
        if not getattr(sys, "frozen", False):
            self.app.log("※開発モードのため自動更新はスキップします（exe版で有効）")
            return
        ok = messagebox.askyesno(
            "更新",
            "v%s に更新しますか？\n現在: v%s\n\n更新するとアプリが自動的に再起動されます。"
            % (self._latest, APP_VERSION))
        if not ok:
            self.app.log("更新をキャンセルしました")
            return
        self.app.log("ダウンロード中…")
        self._busy = True
        threading.Thread(target=self._apply, daemon=True).start()

    def _apply(self):
        try:
            tmp_dir = tempfile.mkdtemp(prefix="nsips_update_")
            new_exe = os.path.join(tmp_dir, EXE_NAME)
            with urllib.request.urlopen(self._url, context=_SSL_CTX) as resp, open(new_exe, "wb") as f:
                while True:
                    chunk = resp.read(1024 * 256)
                    if not chunk:
                        break
                    f.write(chunk)
            self.app.after(0, lambda: self.app.log("ダウンロード完了。更新を適用します…"))
            self.app.after(0, lambda: self._swap(new_exe, tmp_dir))
        except Exception as e:
            self._busy = False
            self.app.after(0, lambda: self.app.log("更新エラー: %s" % e))

    def _swap(self, new_exe, tmp_dir):
        cur_exe = sys.executable
        log_path = os.path.join(BASE_DIR, "update_error.log")
        lines = [
            "@echo off",
            "chcp 65001 > nul",
            "echo NSIPS Converter を更新中...",
            "timeout /t 3 /nobreak > nul",
            'taskkill /F /IM "%s" > nul 2>&1' % EXE_NAME,
            "timeout /t 2 /nobreak > nul",
            'copy /y "%s" "%s" > nul' % (new_exe, cur_exe),
            "if errorlevel 1 (",
            '  echo [ERROR] copy failed > "%s"' % log_path,
            "  exit /b 1",
            ")",
            'rmdir /s /q "%s" 2> nul' % tmp_dir,
            'start "" "%s"' % cur_exe,
            'del "%~f0"',
        ]
        bat = os.path.join(tmp_dir, "apply_update.bat")
        try:
            with open(bat, "w", encoding="utf-8") as f:
                f.write("\r\n".join(lines) + "\r\n")
            subprocess.Popen(["cmd", "/c", bat],
                             creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0))
        except OSError as e:
            self._busy = False
            self.app.log("更新の起動に失敗: %s" % e)
            return
        self.app.root.destroy()


# ---------------------------------------------------------------------------
# GUI アプリ
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.processed = set(self.cfg.get("processed", []))
        self.monitoring = False
        self._stop_evt = threading.Event()
        self.updater = Updater(self)

        root.title("NSIPS TXT → CSV 転記ツール  v%s" % APP_VERSION)
        root.geometry("790x580")
        root.minsize(700, 500)
        pad = {"padx": 6, "pady": 4}

        # 監視元フォルダ
        frm_w = ttk.LabelFrame(root, text="監視元フォルダ (NSIPS TXT出力先)")
        frm_w.pack(fill="x", **pad)
        self.var_watch = tk.StringVar(value=self.cfg.get("watch_folder", ""))
        ttk.Entry(frm_w, textvariable=self.var_watch).pack(side="left", fill="x", expand=True, padx=6, pady=6)
        ttk.Button(frm_w, text="参照…", command=self.pick_watch).pack(side="left", padx=3)
        ttk.Button(frm_w, text="開く", command=lambda: self.open_folder(self.var_watch.get())).pack(side="left", padx=3, pady=6)

        # 出力先フォルダ
        frm_o = ttk.LabelFrame(root, text="出力先フォルダ (CSV保存先)")
        frm_o.pack(fill="x", **pad)
        self.var_out = tk.StringVar(value=self.cfg.get("output_folder", ""))
        ttk.Entry(frm_o, textvariable=self.var_out).pack(side="left", fill="x", expand=True, padx=6, pady=6)
        ttk.Button(frm_o, text="参照…", command=self.pick_out).pack(side="left", padx=3)
        ttk.Button(frm_o, text="開く", command=lambda: self.open_folder(self.var_out.get())).pack(side="left", padx=3, pady=6)

        # 操作ボタン
        frm_b = ttk.Frame(root)
        frm_b.pack(fill="x", **pad)
        self.btn_toggle = ttk.Button(frm_b, text="監視開始", command=self.toggle_monitor)
        self.btn_toggle.pack(side="left", padx=4)
        ttk.Button(frm_b, text="今すぐ全処理", command=self.process_all_now).pack(side="left", padx=4)
        ttk.Button(frm_b, text="処理済み履歴クリア", command=self.clear_processed).pack(side="left", padx=4)
        self.lbl_status = ttk.Label(frm_b, text="停止中", foreground="#a00")
        self.lbl_status.pack(side="right", padx=8)

        # オプション行(スタートアップ登録)
        frm_s = ttk.Frame(root)
        frm_s.pack(fill="x", **pad)
        self.var_startup = tk.BooleanVar(value=startup_enabled())
        self.cb_startup = ttk.Checkbutton(
            frm_s, text="Windowsスタートアップに登録（PC起動時に自動で立ち上げる）",
            variable=self.var_startup, command=self.toggle_startup)
        self.cb_startup.pack(side="left", padx=4)
        if not getattr(sys, "frozen", False):
            self.cb_startup.configure(state="disabled")
            ttk.Label(frm_s, text="(exe版で有効)", foreground="#999").pack(side="left")

        # 更新・履歴ボタン
        frm_u = ttk.Frame(root)
        frm_u.pack(fill="x", **pad)
        ttk.Button(frm_u, text="更新履歴", command=self.show_changelog).pack(side="left", padx=4)
        ttk.Button(frm_u, text="更新を確認", command=lambda: self.updater.check_async(silent=False)).pack(side="left", padx=4)
        ttk.Label(frm_u, text="バージョン v%s" % APP_VERSION, foreground="#666").pack(side="right", padx=8)

        # ログ
        frm_l = ttk.LabelFrame(root, text="ログ")
        frm_l.pack(fill="both", expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(frm_l, height=12, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.log("ツールを起動しました。 (v%s)" % APP_VERSION)

        self.maybe_show_changelog_on_update()
        self.updater.check_async(silent=True)
        # フォルダが設定済みなら自動で監視開始(スタートアップ起動時も無操作で稼働)
        self.root.after(500, self.auto_start_if_ready)

    def auto_start_if_ready(self):
        if (self.var_watch.get() and os.path.isdir(self.var_watch.get())
                and self.var_out.get() and os.path.isdir(self.var_out.get())):
            self.log("設定済みフォルダを検出。自動で監視を開始します。")
            self.start_monitor()

    # after() ラッパ
    def after(self, ms, fn):
        try:
            self.root.after(ms, fn)
        except Exception:
            pass

    # ログ
    def log(self, msg):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = "[%s] %s\n" % (ts, msg)
        def _append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", line)
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _append)

    # スタートアップ
    def toggle_startup(self):
        want = self.var_startup.get()
        try:
            set_startup(want)
            self.log("Windowsスタートアップ登録: %s" % ("ON" if want else "OFF"))
        except OSError as e:
            self.log("スタートアップ設定に失敗: %s" % e)
            self.var_startup.set(startup_enabled())

    # 更新履歴
    def show_changelog(self):
        self._changelog_window(read_changelog(), "更新履歴")

    def maybe_show_changelog_on_update(self):
        last = self.cfg.get("last_seen_version", "")
        if last != APP_VERSION:
            if last:
                self._changelog_window(read_changelog(), "更新されました (v%s)" % APP_VERSION)
            self.cfg["last_seen_version"] = APP_VERSION
            self.persist()

    def _changelog_window(self, text, title):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("560x460")
        box = scrolledtext.ScrolledText(win, wrap="word")
        box.pack(fill="both", expand=True, padx=8, pady=8)
        box.insert("1.0", text)
        box.configure(state="disabled")
        ttk.Button(win, text="閉じる", command=win.destroy).pack(pady=6)

    # フォルダ操作
    def pick_watch(self):
        d = filedialog.askdirectory(title="監視元フォルダを選択", initialdir=self.var_watch.get() or "D:/")
        if d:
            self.var_watch.set(os.path.normpath(d))
            self.persist()

    def pick_out(self):
        d = filedialog.askdirectory(title="出力先フォルダを選択", initialdir=self.var_out.get() or "D:/")
        if d:
            self.var_out.set(os.path.normpath(d))
            self.persist()

    def open_folder(self, path):
        if path and os.path.isdir(path):
            try:
                os.startfile(path)
            except Exception as e:
                self.log("フォルダを開けません: %s" % e)
        else:
            messagebox.showwarning("フォルダ", "有効なフォルダが指定されていません。")

    def persist(self):
        self.cfg["watch_folder"] = self.var_watch.get()
        self.cfg["output_folder"] = self.var_out.get()
        self.cfg["processed"] = sorted(self.processed)
        save_config(self.cfg)

    # 処理 (ファイル名で処理済み判定 → 再起動・監視再開後も再処理しない)
    def process_file(self, path):
        out_dir = self.var_out.get()
        if not out_dir or not os.path.isdir(out_dir):
            self.log("⚠ 出力先フォルダが未設定です。")
            return False
        name = os.path.basename(path)
        try:
            data = parse_txt(path)
        except Exception as e:
            self.log("✖ 解析失敗 %s : %s" % (name, e))
            return False
        out_path = os.path.join(out_dir, os.path.splitext(name)[0] + ".csv")
        try:
            write_csv(data, out_path)
        except Exception as e:
            self.log("✖ CSV出力失敗 %s : %s" % (name, e))
            return False
        self.log("✔ %s → %s (%s %s %s)" % (
            name, os.path.basename(out_path),
            data["患者番号"], data["漢字氏名"], data["最終来局日"]))
        return True

    def scan_once(self, force=False):
        wdir = self.var_watch.get()
        if not wdir or not os.path.isdir(wdir):
            if force:
                self.log("⚠ 監視元フォルダが未設定です。")
            return
        try:
            files = [f for f in os.listdir(wdir) if f.lower().endswith(".txt")]
        except Exception as e:
            self.log("フォルダ読み取りエラー: %s" % e)
            return
        # 未処理(ファイル名ベース)のみ抽出
        todo = [f for f in sorted(files) if f not in self.processed]
        if not todo:
            if force:
                self.log("新規に処理するTXTはありませんでした。")
            return
        if len(todo) > 5:
            self.log("未処理 %d 件を処理します…" % len(todo))
        done = 0
        for f in todo:
            if self.process_file(os.path.join(wdir, f)):
                self.processed.add(f)
                done += 1
            else:
                # 失敗した場合も再試行を避けるなら add するが、
                # ここでは成功したものだけ記録し、失敗は次回再試行する
                pass
        if done:
            self.persist()
            self.log("%d 件を処理しました。" % done)

    def process_all_now(self):
        threading.Thread(target=lambda: self.scan_once(force=True), daemon=True).start()

    def clear_processed(self):
        if messagebox.askyesno("確認", "処理済み履歴をクリアします。\n次回スキャン時に全TXTが再処理されます。よろしいですか？"):
            self.processed.clear()
            self.persist()
            self.log("処理済み履歴をクリアしました。")

    # 監視スレッド
    def toggle_monitor(self):
        if self.monitoring:
            self.stop_monitor()
        else:
            self.start_monitor()

    def start_monitor(self):
        if self.monitoring:
            return
        if not self.var_watch.get() or not os.path.isdir(self.var_watch.get()):
            messagebox.showwarning("監視", "有効な監視元フォルダを指定してください。")
            return
        if not self.var_out.get() or not os.path.isdir(self.var_out.get()):
            messagebox.showwarning("監視", "有効な出力先フォルダを指定してください。")
            return
        self.persist()
        self.monitoring = True
        self._stop_evt.clear()
        threading.Thread(target=self._monitor_loop, daemon=True).start()
        self.btn_toggle.configure(text="監視停止")
        self.lbl_status.configure(text="監視中", foreground="#080")
        self.log("監視を開始しました（%.0f秒間隔）: %s" % (POLL_INTERVAL, self.var_watch.get()))

    def stop_monitor(self):
        self.monitoring = False
        self._stop_evt.set()
        self.btn_toggle.configure(text="監視開始")
        self.lbl_status.configure(text="停止中", foreground="#a00")
        self.log("監視を停止しました。")

    def _monitor_loop(self):
        while not self._stop_evt.is_set():
            try:
                self.scan_once(force=False)
            except Exception:
                self.log("監視エラー:\n" + traceback.format_exc())
            self._stop_evt.wait(POLL_INTERVAL)

    # 終了
    def on_close(self):
        self.stop_monitor()
        self.persist()
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
