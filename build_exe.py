# -*- coding: utf-8 -*-
"""NSIPS Converter - EXE build script (PyInstaller onefile)。

template.xlsx / VERSION / CHANGELOG.md を同梱した単一 exe を dist に生成する。
生成物: dist/NSIPSConverter.exe （別PCにこれ1つコピーすれば動作）
"""
import os
import shutil
import subprocess
import sys

APP = "NSIPSConverter"
ENTRY = "nsips_converter.py"


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print("=" * 48)
    print("  NSIPS Converter - EXE Build")
    print("=" * 48)

    # [1/4] 依存導入
    if os.environ.get("SKIP_PIP", "").strip() not in ("", "0", "false", "False"):
        print("[1/4] skip pip (SKIP_PIP set)")
    else:
        print("[1/4] install dependencies...")
        subprocess.run([sys.executable, "-m", "pip", "install",
                        "openpyxl", "pyinstaller", "certifi",
                        "--quiet", "--disable-pip-version-check"])

    # [2/4] clean
    print("[2/4] clean old build...")
    for d in ("build", "dist"):
        if os.path.isdir(d):
            shutil.rmtree(d)
    for f in (APP + ".spec",):
        if os.path.isfile(f):
            os.remove(f)

    # [3/4] build (onefile, windowed)
    print("[3/4] PyInstaller build...")
    sep = ";"   # Windows の --add-data 区切り
    cmd = [
        sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm",
        "--onefile", "--windowed", "--name", APP,
        "--add-data", "VERSION%s." % sep,
        "--add-data", "CHANGELOG.md%s." % sep,
        "--hidden-import", "certifi",
        ENTRY,
    ]
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print("[ERROR] PyInstaller failed:", r.returncode)
        return

    exe = os.path.join("dist", APP + ".exe")
    if not os.path.isfile(exe):
        print("[ERROR] EXE not found:", os.path.abspath(exe))
        return

    # [4/4] 分かりやすい場所(プロジェクト直下の 配布/ )へコピー
    if os.path.isdir("build"):
        shutil.rmtree("build")
    if os.path.isfile(APP + ".spec"):
        os.remove(APP + ".spec")

    dist_folder = "配布"
    os.makedirs(dist_folder, exist_ok=True)
    final = os.path.join(dist_folder, APP + ".exe")
    shutil.copy2(exe, final)

    mb = os.path.getsize(exe) / (1024 * 1024)
    print("=" * 48)
    print("  BUILD OK!")
    print("  EXE(配布用): %s (%.1f MB)" % (os.path.abspath(final), mb))
    print("  ※このexe1つを別PCにコピーすれば動作します")
    print("=" * 48)


if __name__ == "__main__":
    main()
