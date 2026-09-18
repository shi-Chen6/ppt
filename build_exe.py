# -*- coding: utf-8 -*-
"""
一键打包脚本：把 bootstrap.py 打成 onefile exe（内嵌嵌入式 Python + get-pip + 主程序）
用法：
  python build_exe.py            # 完整打包，产物 dist/PPT-Tool.exe
  python build_exe.py --stage-only   # 仅下载缓存 + 准备 build/stage（调试 bootstrap.py 用）
"""
import os
import shutil
import subprocess
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
BUILD = os.path.join(HERE, "build")
CACHE = os.path.join(BUILD, "cache")
STAGE = os.path.join(BUILD, "stage")
WORK = os.path.join(BUILD, "work")
DIST = os.path.join(HERE, "dist")
VENV = os.path.join(BUILD, "venv")

EMBED_URL = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip"
GETPIP_URL = "https://bootstrap.pypa.io/get-pip.py"
EXE_NAME = "PPT-Tool"


def download(url, dest):
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    if os.path.isfile(dest) and os.path.getsize(dest) > 0:
        print("[缓存] 已存在 %s" % dest)
        return
    print("[下载] %s" % url)

    def hook(count, block, total):
        if total > 0:
            pct = min(100, count * block * 100 // total)
            sys.stdout.write("\r  进度 %d%% (%.1f MB)" % (pct, count * block / 1e6))
            sys.stdout.flush()

    tmp = dest + ".part"
    urllib.request.urlretrieve(url, tmp, reporthook=hook)
    os.replace(tmp, dest)
    print("\n[下载] 完成 → %s" % dest)


def _rmtree(path):
    """手动递归删除（不依赖 shutil.rmtree，兼容受限环境）。"""
    if not os.path.isdir(path):
        return
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            try:
                os.remove(os.path.join(root, name))
            except OSError:
                pass
        for name in dirs:
            try:
                os.rmdir(os.path.join(root, name))
            except OSError:
                pass
    try:
        os.rmdir(path)
    except OSError:
        pass


def stage_resources():
    os.makedirs(STAGE, exist_ok=True)
    embed = os.path.join(CACHE, "python_embed.zip")
    getpip = os.path.join(CACHE, "get-pip.py")
    download(EMBED_URL, embed)
    download(GETPIP_URL, getpip)
    shutil.copyfile(embed, os.path.join(STAGE, "python_embed.zip"))
    shutil.copyfile(getpip, os.path.join(STAGE, "get-pip.py"))
    for name in ("app.py", "chart_renderer.py", "requirements.txt", "bootstrap.py"):
        shutil.copyfile(os.path.join(HERE, name), os.path.join(STAGE, name))
    # 覆盖式复制，不删除旧目录（避免受限环境下删除失败；PyInstaller 会整体重新收集）
    shutil.copytree(os.path.join(HERE, "static"), os.path.join(STAGE, "static"),
                    dirs_exist_ok=True)
    print("[staging] 资源已就绪 → %s" % STAGE)


def ensure_build_env():
    py = os.path.join(VENV, "Scripts", "python.exe")
    if not os.path.isfile(py):
        print("[venv] 创建构建虚拟环境…")
        subprocess.run([sys.executable, "-m", "venv", VENV], check=True)
    print("[venv] 安装 PyInstaller…")
    subprocess.run([py, "-m", "pip", "install", "-q", "--disable-pip-version-check",
                    "pyinstaller"], check=True)
    return py


def run_pyinstaller(py):
    os.makedirs(DIST, exist_ok=True)
    sep = ";" if os.name == "nt" else ":"
    cmd = [
        py, "-m", "PyInstaller",
        "--onefile", "--noconsole", "--clean", "--noconfirm",
        "--name", EXE_NAME,
        "--distpath", DIST,
        "--workpath", WORK,
        "--specpath", BUILD,
        "--add-data", os.path.join(STAGE, "python_embed.zip") + sep + ".",
        "--add-data", os.path.join(STAGE, "get-pip.py") + sep + ".",
        "--add-data", os.path.join(STAGE, "app.py") + sep + ".",
        "--add-data", os.path.join(STAGE, "chart_renderer.py") + sep + ".",
        "--add-data", os.path.join(STAGE, "requirements.txt") + sep + ".",
        "--add-data", os.path.join(STAGE, "static") + sep + "static",
        os.path.join(HERE, "bootstrap.py"),
    ]
    print("[PyInstaller] %s" % " ".join(cmd))
    subprocess.run(cmd, check=True)
    exe = os.path.join(DIST, EXE_NAME + (".exe" if os.name == "nt" else ""))
    size_mb = os.path.getsize(exe) / 1e6
    print("=" * 52)
    print("打包完成: %s  (%.1f MB)" % (exe, size_mb))
    print("=" * 52)
    return exe


def main(argv):
    stage_resources()
    if "--stage-only" in argv:
        print("[stage-only] 完成")
        return 0
    py = ensure_build_env()
    run_pyinstaller(py)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
