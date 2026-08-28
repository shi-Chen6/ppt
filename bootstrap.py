# -*- coding: utf-8 -*-
"""
PPT 智能生成工具 —— 首次运行引导器 (bootstrapper)
=================================================

打包形态：PyInstaller onefile（--noconsole），内嵌资源：
  python_embed.zip   嵌入式 Python 运行时（python-3.12.x-embed-amd64.zip）
  get-pip.py         pip 引导脚本
  app.py / requirements.txt / static/   主程序

首次运行流程：
  1) 若 127.0.0.1:5000 已在服务 → 直接开浏览器退出
  2) 释放嵌入式 Python 到 %LOCALAPPDATA%\\PPTTool\\runtime（含版本指纹，变更自动重建）
  3) 释放主程序文件到 %LOCALAPPDATA%\\PPTTool\\app（output/ 保留）
  4) 哨兵文件 .deps_ok 命中 → 直接启动；否则逐包 find_spec + 版本约束检测
  5) 有缺失 → tkinter 进度窗口：网络/磁盘预检 → get-pip → 逐包 pip install
     （进度条分段：预检 0-10 / pip 引导 10-20 / 逐包安装 20-95 / 复检哨兵 95-100）
  6) 安装成功复检 → 写哨兵 → pythonw 拉起 app.py → 等端口就绪开浏览器
  7) 失败 → 错误分类（网络/磁盘/版本/权限），自动换镜像重试，日志落盘

调试模式（源码运行）：
  python bootstrap.py --headless --no-launch [--mirror https://pypi.tuna.tsinghua.edu.cn/simple]
"""
import glob
import hashlib
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
import webbrowser
import zipfile

APP_NAME = "PPTTool"
APP_TITLE = "PPT 智能生成工具"
PORT = 5000
HOME_URL = "http://127.0.0.1:%d/" % PORT
DISK_MIN_BYTES = 800 * 1024 * 1024  # 800MB 余量

MIRRORS = [
    ("清华源", "https://pypi.tuna.tsinghua.edu.cn/simple"),
    ("阿里源", "https://mirrors.aliyun.com/pypi/simple"),
    ("官方源", "https://pypi.org/simple"),
]

NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW，防止子进程闪黑框

DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), APP_NAME)
RUNTIME_DIR = os.path.join(DATA_DIR, "runtime")
APP_DIR = os.path.join(DATA_DIR, "app")
LOG_DIR = os.path.join(DATA_DIR, "logs")
SENTINEL = os.path.join(DATA_DIR, ".deps_ok")
RUNTIME_VER_FILE = os.path.join(DATA_DIR, ".runtime_ver")
CHECK_SCRIPT = os.path.join(DATA_DIR, "check_deps.py")

_LOG_PATH = None


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


# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------
def log(msg):
    global _LOG_PATH
    if _LOG_PATH is None:
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            _LOG_PATH = os.path.join(LOG_DIR, "install_%s.log" % time.strftime("%Y%m%d_%H%M%S"))
        except OSError:
            return
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(time.strftime("[%H:%M:%S] ") + str(msg) + "\n")
    except OSError:
        pass


# ---------------------------------------------------------------------------
# 资源定位：frozen → _MEIPASS；源码运行 → build/stage（打包脚本准备）
# ---------------------------------------------------------------------------
def resource_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    stage = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build", "stage")
    if os.path.isdir(stage):
        return stage
    return os.path.dirname(os.path.abspath(__file__))


def embed_zip_path():
    return os.path.join(resource_dir(), "python_embed.zip")


def get_pip_path():
    return os.path.join(resource_dir(), "get-pip.py")


def embed_hash():
    h = hashlib.sha256()
    with open(embed_zip_path(), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


# ---------------------------------------------------------------------------
# 运行时 / 主程序文件释放
# ---------------------------------------------------------------------------
def ensure_runtime(force=False):
    """释放嵌入式 Python 到 RUNTIME_DIR；embed zip 变化时重建。返回 embed 指纹。"""
    eh = embed_hash()
    cur = None
    if os.path.isfile(RUNTIME_VER_FILE):
        try:
            with open(RUNTIME_VER_FILE, encoding="utf-8") as f:
                cur = f.read().strip()
        except OSError:
            cur = None
    ok = (not force and cur == eh and os.path.isfile(os.path.join(RUNTIME_DIR, "python.exe")))
    if not ok:
        log("释放嵌入式 Python 运行时 (指纹 %s)" % eh)
        if os.path.isdir(RUNTIME_DIR):
            _rmtree(RUNTIME_DIR)
        os.makedirs(RUNTIME_DIR, exist_ok=True)
        with zipfile.ZipFile(embed_zip_path()) as zf:
            zf.extractall(RUNTIME_DIR)
        # 改写 pythonXXX._pth：启用 site 机制 + site-packages
        for pth in glob.glob(os.path.join(RUNTIME_DIR, "python*._pth")):
            try:
                with open(pth, encoding="utf-8") as f:
                    lines = [ln.strip() for ln in f.read().splitlines() if ln.strip() and not ln.strip().startswith("#")]
            except OSError:
                lines = []
            keep = [ln for ln in lines if ln.endswith(".zip") or ln == "."]
            with open(pth, "w", encoding="utf-8") as f:
                f.write("\n".join(keep + ["Lib\\site-packages", "import site"]) + "\n")
        with open(RUNTIME_VER_FILE, "w", encoding="utf-8") as f:
            f.write(eh)
    return eh


def ensure_app_files():
    """释放主程序文件到 APP_DIR（覆盖式更新，output/ 不受影响）。"""
    src = resource_dir()
    os.makedirs(APP_DIR, exist_ok=True)
    for name in ("app.py", "requirements.txt"):
        shutil.copyfile(os.path.join(src, name), os.path.join(APP_DIR, name))
    sdir = os.path.join(src, "static")
    ddir = os.path.join(APP_DIR, "static")
    if os.path.isdir(sdir):
        if os.path.isdir(ddir):
            _rmtree(ddir)
        shutil.copytree(sdir, ddir)


def py_exe(windowed=False):
    return os.path.join(RUNTIME_DIR, "pythonw.exe" if windowed else "python.exe")


def child_env():
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


# ---------------------------------------------------------------------------
# 依赖检测
# ---------------------------------------------------------------------------
CHECK_CODE = r'''
# -*- coding: utf-8 -*-
"""在目标运行时里执行：逐包 find_spec + 版本约束比对，输出缺失清单 JSON。"""
import importlib.metadata
import importlib.util
import json
import re
import sys

MODULE_MAP = {
    "flask": "flask", "requests": "requests", "python-pptx": "pptx",
    "python-docx": "docx", "matplotlib": "matplotlib", "Pillow": "PIL",
}


def ver_tuple(s):
    nums = []
    for part in s.split("."):
        m = re.match(r"^(\d+)", part)
        if not m:
            break
        nums.append(int(m.group(1)))
    return tuple(nums) if nums else (0,)


missing = []
for raw in open(sys.argv[1], encoding="utf-8-sig"):
    line = raw.strip()
    if not line or line.startswith("#"):
        continue
    m = re.match(r"^([A-Za-z0-9_][A-Za-z0-9._\-]*)\s*(>=|==|~=|<=|>|<)?\s*([0-9][0-9A-Za-z.\*]*)?", line)
    if not m:
        continue
    name, op, want = m.group(1), m.group(2), m.group(3)
    module = MODULE_MAP.get(name, name.replace("-", "_"))
    if importlib.util.find_spec(module) is None:
        missing.append(line)
        continue
    if op and want:
        try:
            installed = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            missing.append(line)
            continue
        iv, wv = ver_tuple(installed), ver_tuple(want)
        ok = iv >= wv if op in (">=", "~=", ">") else iv <= wv if op in ("<=", "<") else iv == wv
        if not ok:
            missing.append(line)
print(json.dumps({"missing": missing}))
'''


def run_check():
    """用 runtime 的 python 执行依赖检测。返回缺失 requirement 行列表；None=检测环境损坏。"""
    try:
        with open(CHECK_SCRIPT, "w", encoding="utf-8") as f:
            f.write(CHECK_CODE)
        r = subprocess.run(
            [py_exe(), CHECK_SCRIPT, os.path.join(APP_DIR, "requirements.txt")],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=120, env=child_env(), creationflags=NO_WINDOW,
        )
        if r.returncode != 0:
            log("依赖检测脚本报错: %s" % r.stderr[-500:])
            return None
        data = json.loads(r.stdout.strip().splitlines()[-1])
        return data.get("missing", [])
    except Exception as e:  # noqa: BLE001
        log("依赖检测异常: %s" % e)
        return None


def all_requirement_lines():
    out = []
    with open(os.path.join(APP_DIR, "requirements.txt"), encoding="utf-8-sig") as f:
        for raw in f:
            line = raw.strip()
            if line and not line.startswith("#"):
                out.append(line)
    return out


def fingerprint(embed_hash_value):
    with open(os.path.join(APP_DIR, "requirements.txt"), "rb") as f:
        return hashlib.sha256(f.read() + embed_hash_value.encode()).hexdigest()[:20]


def write_sentinel(value):
    with open(SENTINEL, "w", encoding="utf-8") as f:
        f.write(value)


def read_sentinel():
    try:
        with open(SENTINEL, encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return None


# ---------------------------------------------------------------------------
# 预检
# ---------------------------------------------------------------------------
def port_open():
    s = socket.socket()
    s.settimeout(0.6)
    try:
        s.connect(("127.0.0.1", PORT))
        return True
    except OSError:
        return False
    finally:
        s.close()


def check_network(mirror):
    try:
        urllib.request.urlopen(mirror, timeout=8)
        return True, ""
    except Exception as e:  # noqa: BLE001
        return False, str(e)


def check_disk():
    try:
        return shutil.disk_usage(DATA_DIR).free >= DISK_MIN_BYTES, shutil.disk_usage(DATA_DIR).free
    except OSError:
        return True, 0


# ---------------------------------------------------------------------------
# 安装器（逻辑核心，GUI / headless 共用）
# ---------------------------------------------------------------------------
def classify_error(lines, returncode):
    """把 pip 输出尾部 + 退出码分类成人话。"""
    text = "\n".join(lines[-60:])
    if "No space left" in text or "WinError 1128" in text:
        return {"kind": "disk", "title": "磁盘空间不足",
                "hint": "请释放 C 盘（或应用所在盘）空间后点击重试。"}
    if re.search(r"ReadTimeoutError|ConnectionError|MaxRetry|retries|ProxyError|"
                 r"Failed to establish|Could not fetch|Connection reset|Connection aborted|"
                 r"Temporary failure|URLError|Network is unreachable", text, re.I):
        return {"kind": "network", "title": "网络连接失败",
                "hint": "已自动切换镜像源重试。若持续失败，请检查网络/代理后手动换源重试。"}
    if "Could not find a version" in text or "ResolutionImpossible" in text or "No matching distribution" in text:
        return {"kind": "version", "title": "依赖版本无法解析",
                "hint": "当前 Python 运行时找不到满足 requirements 的包版本，请重新下载最新安装包。"}
    if "PermissionError" in text or "Access is denied" in text or "WinError 5" in text:
        return {"kind": "permission", "title": "写入被拒绝",
                "hint": "可能是杀毒软件拦截。请将 %s 加入信任区后重试。" % DATA_DIR}
    if " externally-managed-environment " in text:
        return {"kind": "env", "title": "运行环境异常",
                "hint": "运行环境标记为外部管理，正在自动重建。"}
    return {"kind": "unknown", "title": "安装失败（退出码 %s）" % returncode,
            "hint": "请点击「查看日志」获取详情，或重试 / 更换镜像源。"}


class Installer:
    """执行完整安装流程。回调：on_stage/on_progress/on_pkg_start/on_pkg_done/on_line。"""

    def __init__(self, reqs, mirror, callbacks):
        self.reqs = reqs
        self.mirror = mirror
        self.cb = callbacks
        self.cancelled = False

    def cancel(self):
        self.cancelled = True

    # ---- 子进程输出逐行读取（可中断）----
    def _pump(self, cmd, tag):
        """运行子进程，逐行回调。返回 (returncode, all_lines)。"""
        lines = []
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
            env=child_env(), creationflags=NO_WINDOW,
        )
        try:
            for line in proc.stdout:
                line = line.rstrip()
                if not line:
                    continue
                lines.append(line)
                log("[%s] %s" % (tag, line))
                if re.match(r"^(Collecting|Downloading|Using cached|Installing collected|"
                            r"Successfully installed|Requirement already)", line):
                    self.cb["on_line"]("%s｜%s" % (tag, line[:110]))
                if self.cancelled:
                    proc.terminate()
                    break
            proc.wait(timeout=30)
        except Exception:
            proc.kill()
        return proc.returncode or 0, lines

    def _pip_ok(self):
        r = subprocess.run(
            [py_exe(), "-m", "pip", "--version"], capture_output=True,
            env=child_env(), creationflags=NO_WINDOW,
        )
        return r.returncode == 0

    def run(self):
        """返回 (ok, error_dict, cancelled)。"""
        cb = self.cb

        # ---- 阶段一：环境预检 0-10% ----
        cb["on_stage"]("正在检查网络与磁盘环境…")
        ok, err = check_network(self.mirror)
        if not ok:
            return False, {"kind": "network", "title": "无法访问 pip 镜像源",
                           "hint": "请检查网络连接，或在下方切换镜像源后重试。"
                                   "\n详细信息：" + err[:200]}, False
        ok, free = check_disk()
        if not ok:
            return False, {"kind": "disk", "title": "磁盘空间不足",
                           "hint": "安装约需 800MB 可用空间，当前剩余 %.1f GB。" % (free / 1e9)}, False
        cb["on_progress"](10)

        # ---- 阶段二：pip 引导 10-20% ----
        if not self._pip_ok():
            cb["on_stage"]("正在初始化 pip 包管理器…")
            code, lines = self._pump(
                [py_exe(), get_pip_path(), "--no-warn-script-location", "-i", self.mirror],
                "pip-init",
            )
            if self.cancelled:
                return False, None, True
            if code != 0 or not self._pip_ok():
                return False, classify_error(lines, code), False
        cb["on_progress"](20)

        # ---- 阶段三：逐包安装 20-95% ----
        total = len(self.reqs)
        for i, req in enumerate(self.reqs):
            if self.cancelled:
                return False, None, True
            name = re.match(r"^([A-Za-z0-9._\-]+)", req).group(1)
            cb["on_pkg_start"](i, name, total)
            cmd = [py_exe(), "-m", "pip", "install", req,
                   "-i", self.mirror, "--no-warn-script-location",
                   "--progress-bar", "off", "--disable-pip-version-check"]
            code, lines = self._pump(cmd, name)
            if self.cancelled:
                return False, None, True
            if code != 0:
                return False, classify_error(lines, code), False
            cb["on_pkg_done"](i, name, total)
        cb["on_progress"](95)

        # ---- 阶段四：复检 + 哨兵 95-100% ----
        cb["on_stage"]("正在校验依赖完整性…")
        missing = run_check()
        if missing is None or missing:
            detail = "复检仍缺失: %s" % (missing or "检测脚本异常")
            log(detail)
            return False, {"kind": "verify", "title": "依赖校验未通过",
                           "hint": "安装后复检发现问题，请重试；若持续失败请查看日志。"}, False
        cb["on_progress"](100)
        return True, None, False


def run_with_retry(reqs, mirror_idx, make_installer, sleep_between=True):
    """自动重试：网络类错误自动换镜像（最多 3 个源）；其余错误原地重试 1 次。"""
    errors = []
    idx = mirror_idx
    for attempt in range(3):
        if attempt and sleep_between:
            time.sleep(3 * attempt)
        name, url = MIRRORS[idx % len(MIRRORS)]
        log("安装尝试 %d/3，镜像: %s (%s)" % (attempt + 1, name, url))
        inst = make_installer(url)
        ok, err, cancelled = inst.run()
        if ok or cancelled:
            return ok, err, cancelled, idx % len(MIRRORS)
        errors.append(err or {})
        if err.get("kind") in ("network",):
            idx += 1  # 换下一个镜像
        elif err.get("kind") in ("disk", "version", "permission"):
            break     # 重试无意义
    return False, errors[-1], False, idx % len(MIRRORS)


# ---------------------------------------------------------------------------
# 启动主程序
# ---------------------------------------------------------------------------
def launch_app():
    os.makedirs(LOG_DIR, exist_ok=True)
    app_log_path = os.path.join(LOG_DIR, "app.log")
    log("启动主程序: %s" % os.path.join(APP_DIR, "app.py"))
    try:
        app_log = open(app_log_path, "ab")
        subprocess.Popen(
            [py_exe(windowed=True), os.path.join(APP_DIR, "app.py")],
            cwd=APP_DIR, stdout=app_log, stderr=app_log, stdin=subprocess.DEVNULL,
            env=child_env(), creationflags=NO_WINDOW, close_fds=True,
        )
    except Exception as e:  # noqa: BLE001
        log("启动失败: %s\n%s" % (e, traceback.format_exc()))
        return False
    return True


def wait_and_open_browser(timeout=40):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if port_open():
            break
        time.sleep(0.4)
    try:
        webbrowser.open(HOME_URL)
    except Exception:  # noqa: BLE001
        pass
    log("浏览器已打开 %s" % HOME_URL)


# ---------------------------------------------------------------------------
# GUI 安装窗口
# ---------------------------------------------------------------------------
def install_gui(reqs, mirror_idx, fp):
    import tkinter as tk
    from tkinter import ttk, messagebox

    q = queue.Queue()
    root = tk.Tk()
    root.title(APP_TITLE + " - 首次运行环境准备")
    root.geometry("560x380")
    root.resizable(False, False)

    ttk.Label(root, text="首次运行准备", font=("Microsoft YaHei UI", 13, "bold")).pack(pady=(14, 2))
    ttk.Label(root, text="正在为本机准备 Python 运行环境（仅需一次，之后秒速启动）",
              foreground="#666").pack()

    status_var = tk.StringVar(value="准备中…")
    ttk.Label(root, textvariable=status_var, font=("Microsoft YaHei UI", 10, "bold")).pack(pady=(8, 4))

    bar = ttk.Progressbar(root, maximum=100, length=500, mode="determinate")
    bar.pack(pady=2)

    out = tk.Text(root, height=7, width=66, font=("Consolas", 9), state="disabled",
                  background="#fafafa", relief="flat")
    out.pack(padx=14, pady=6, fill="both")

    bottom = ttk.Frame(root)
    bottom.pack(side="bottom", fill="x", padx=14, pady=8)

    mirror_var = tk.StringVar(value=MIRRORS[mirror_idx][0])
    mirror_combo = ttk.Combobox(bottom, textvariable=mirror_var, state="readonly",
                                values=[n for n, _ in MIRRORS], width=10)
    btn_retry = ttk.Button(bottom, text="重试")
    btn_log = ttk.Button(bottom, text="查看日志")
    btn_copy = ttk.Button(bottom, text="复制错误信息")
    btn_cancel = ttk.Button(bottom, text="取消")

    state = {"worker": None, "done": False, "last_error": None, "fp": fp, "inst": None, "ok": False}

    def set_buttons(mode):
        for b in (btn_retry, btn_log, btn_copy, btn_cancel, mirror_combo):
            b.pack_forget()
        if mode == "installing":
            btn_cancel.pack(side="right")
            mirror_combo.pack(side="left")
            mirror_combo.state(["disabled"])
        else:  # error
            mirror_combo.state(["!disabled"])
            mirror_combo.pack(side="left")
            btn_cancel.pack(side="right")
            btn_retry.pack(side="right", padx=6)
            btn_copy.pack(side="right", padx=6)
            btn_log.pack(side="right", padx=6)

    def out_write(text):
        out.config(state="normal")
        out.insert("end", text + "\n")
        out.see("end")
        line_count = int(out.index("end-1c").split(".")[0])
        if line_count > 200:
            out.delete("1.0", "%d.0" % (line_count - 200))
        out.config(state="disabled")

    # ---- worker 线程 ----
    def start_worker(idx):
        def make_installer(mirror_url):
            def safe(name, *a, **k):
                q.put((name, a, k))
            inst = Installer(reqs, mirror_url, {
                "on_stage": lambda t: safe("stage", t),
                "on_progress": lambda p: safe("progress", p),
                "on_pkg_start": lambda i, n, t: safe("pkg_start", i, n, t),
                "on_pkg_done": lambda i, n, t: safe("pkg_done", i, n, t),
                "on_line": lambda s: safe("line", s),
            })
            state["inst"] = inst
            return inst
        def worker():
            ok, err, cancelled, used_idx = run_with_retry(reqs, idx, make_installer)
            q.put(("finished", (ok, err, cancelled, used_idx), {}))
        t = threading.Thread(target=worker, daemon=True)
        state["worker"] = t
        t.start()

    # ---- 队列轮询 ----
    def poll():
        try:
            while True:
                kind, args, kwargs = q.get_nowait()
                if kind == "stage":
                    status_var.set(args[0])
                elif kind == "progress":
                    bar.config(mode="determinate")
                    bar["value"] = args[0]
                elif kind == "pkg_start":
                    i, name, total = args
                    status_var.set("正在安装依赖 (%d/%d)：%s" % (i + 1, total, name))
                    bar.config(mode="indeterminate")
                    bar.start(25)
                elif kind == "pkg_done":
                    i, _, total = args
                    bar.stop()
                    bar.config(mode="determinate")
                    bar["value"] = 20 + 75.0 * (i + 1) / total
                elif kind == "line":
                    out_write(args[0])
                elif kind == "finished":
                    ok, err, cancelled, used_idx = args
                    state["done"] = True
                    if ok:
                        status_var.set("环境准备完成，正在启动主程序…")
                        bar.config(mode="determinate")
                        bar["value"] = 100
                        root.after(400, finish_success)
                    elif cancelled:
                        status_var.set("已取消安装。")
                        set_buttons("error")
                    else:
                        state["last_error"] = err
                        bar.stop()
                        bar.config(mode="determinate")
                        show_error(err)
        except queue.Empty:
            pass
        if not state["done"]:
            root.after(100, poll)

    def show_error(err):
        status_var.set("安装失败：%s" % err.get("title", "未知错误"))
        out_write("")
        out_write("【原因】%s" % err.get("title", ""))
        out_write("【建议】%s" % err.get("hint", ""))
        set_buttons("error")

    def finish_success():
        write_sentinel(state["fp"])
        state["ok"] = True
        root.destroy()
        launch_app()
        threading.Thread(target=wait_and_open_browser, daemon=True).start()

    # ---- 按钮行为 ----
    def on_cancel():
        if state["worker"] and state["worker"].is_alive():
            if not messagebox.askokcancel(APP_TITLE, "确定取消安装吗？下次启动将重新检测并安装。"):
                return
            state["done"] = True
            if state["inst"]:
                state["inst"].cancel()
            root.destroy()
        else:
            root.destroy()

    def on_retry():
        for i, (n, _) in enumerate(MIRRORS):
            if n == mirror_var.get():
                idx = i
                break
        else:
            idx = 0
        out_write("")
        out_write("── 使用 %s 重新安装 ──" % MIRRORS[idx][0])
        state["done"] = False
        set_buttons("installing")
        status_var.set("正在重试…")
        root.after(100, poll)
        start_worker(idx)

    def on_log():
        path = _LOG_PATH or LOG_DIR
        try:
            os.startfile(path)  # noqa: S606
        except OSError:
            messagebox.showerror(APP_TITLE, "日志目录不存在：%s" % path)

    def on_copy():
        err = state["last_error"] or {}
        text = "%s\n%s\n\n日志：%s" % (err.get("title", ""), err.get("hint", ""), _LOG_PATH or LOG_DIR)
        root.clipboard_clear()
        root.clipboard_append(text)
        messagebox.showinfo(APP_TITLE, "错误信息已复制到剪贴板。")

    btn_retry.config(command=on_retry)
    btn_log.config(command=on_log)
    btn_copy.config(command=on_copy)
    btn_cancel.config(command=on_cancel)
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    set_buttons("installing")
    start_worker(mirror_idx)
    root.after(100, poll)
    root.mainloop()
    return state["ok"]


# ---------------------------------------------------------------------------
# headless 模式（调试用）
# ---------------------------------------------------------------------------
def install_headless(reqs, mirror_idx):
    def cb():
        def on_stage(t):
            print("[阶段] %s" % t, flush=True)
        def on_progress(p):
            print("[进度] %d%%" % p, flush=True)
        def on_pkg_start(i, n, t):
            print("[安装] (%d/%d) %s" % (i + 1, t, n), flush=True)
        def on_pkg_done(i, n, t):
            print("[完成] %s" % n, flush=True)
        def on_line(s):
            print("    %s" % s, flush=True)
        return {"on_stage": on_stage, "on_progress": on_progress,
                "on_pkg_start": on_pkg_start, "on_pkg_done": on_pkg_done, "on_line": on_line}

    callbacks = cb()

    def make_installer(mirror_url):
        return Installer(reqs, mirror_url, callbacks)

    ok, err, cancelled, used = run_with_retry(reqs, mirror_idx, make_installer, sleep_between=False)
    if cancelled:
        print("[结果] 已取消", flush=True)
        return False
    if not ok:
        print("[结果] 失败：%s\n%s" % (err.get("title"), err.get("hint")), flush=True)
        return False
    print("[结果] 全部依赖安装成功", flush=True)
    return True


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main(argv):
    headless = "--headless" in argv
    no_launch = "--no-launch" in argv
    mirror_idx = 0
    for i, a in enumerate(argv):
        if a == "--mirror" and i + 1 < len(argv):
            for j, (n, url) in enumerate(MIRRORS):
                if url == argv[i + 1] or n == argv[i + 1]:
                    mirror_idx = j
    log("引导器启动 frozen=%s argv=%s" % (getattr(sys, "frozen", False), argv))

    # 0) 服务已在运行 → 直接开浏览器
    if port_open():
        log("端口 %d 已在服务，直接打开浏览器" % PORT)
        webbrowser.open(HOME_URL)
        return 0

    # 1) 释放运行时与主程序文件
    try:
        eh = ensure_runtime()
        ensure_app_files()
    except Exception as e:  # noqa: BLE001
        log("资源释放失败: %s\n%s" % (e, traceback.format_exc()))
        if headless:
            print("[结果] 资源释放失败: %s" % e, flush=True)
        else:
            _fatal_dialog("初始化失败", "无法准备运行环境：%s\n日志：%s" % (e, _LOG_PATH or LOG_DIR))
        return 1

    fp = fingerprint(eh)

    # 2) 哨兵 / 依赖检测
    missing = None
    if read_sentinel() == fp:
        missing = []
        log("哨兵命中，跳过依赖检测")
    else:
        missing = run_check()
        if missing is None:
            # 检测环境损坏 → 重建运行时，全量安装（自愈路径）
            log("依赖检测环境损坏，重建运行时")
            eh = ensure_runtime(force=True)
            fp = fingerprint(eh)
            missing = all_requirement_lines()

    # 3) 环境完整 → 写哨兵直接启动
    if not missing:
        log("依赖完整，直接启动")
        write_sentinel(fp)
        if no_launch:
            print("[结果] 环境就绪（--no-launch，跳过启动）", flush=True)
            return 0
        launch_app()
        wait_and_open_browser()
        return 0

    log("缺失依赖 %d 项: %s" % (len(missing), ", ".join(missing)))

    # 4) 安装流程
    if headless:
        ok = install_headless(missing, mirror_idx)
    else:
        ok = install_gui(missing, mirror_idx, fp)
    if not ok:
        return 1
    write_sentinel(fp)  # GUI 成功时已写，这里幂等兜底

    if no_launch:
        print("[结果] 安装完成（--no-launch，跳过启动）", flush=True)
        return 0
    launch_app()
    wait_and_open_browser()
    return 0


def _fatal_dialog(title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:  # noqa: BLE001
        pass


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
