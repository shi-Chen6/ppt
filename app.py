# -*- coding: utf-8 -*-
"""
PPT 智能生成工具 —— 后端服务

两条链路：
  A. 手动逐页：/api/optimize（优化提示词）→ /api/generate（生图）→ /api/build_ppt（拼 PPT）
  B. 文档一键：/api/parse_doc（解析文档）→ /api/outline（生成大纲）→ /api/split_content（拆文案）
     → /api/optimize（生成 prompt）→ /api/generate（生图，可选锁长相）→ /api/build_ppt

生图模型：中转站 gpt-image-2
锁长相：走 /images/edits 端点传参考图
依赖：flask / requests / python-pptx / python-docx
运行：python app.py  然后浏览器打开 http://127.0.0.1:5000
"""
import base64
import io
import json
import os
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from flask import Flask, jsonify, request, send_from_directory
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
os.makedirs(IMAGES_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", static_url_path="")

# ---------------------------------------------------------------------------
# System Prompt 集
# ---------------------------------------------------------------------------
OPTIMIZE_SYSTEM = (
    "你是一位专业的演示文稿视觉设计师与 AI 生图提示词专家。"
    "用户会给出一页 PPT 的内容描述（可能是中文），请把它改写成一段用于 gpt-image-2 模型"
    "生成 16:9 演示页面的高质量提示词。\n\n"
    "要求：\n"
    "1. 输出一段完整连贯的英文提示词（画面与风格描述用英文，便于生图模型理解构图、光影与质感）；"
    "但页面中需要出现的文字必须保持用户原文的语种与措辞——例如标题是中文，就原样保留中文，"
    "并用双引号精确标出每段文字。\n"
    "2. 明确写出：16:9 横向构图、整体视觉风格（商务/科技/简约/手绘等，可从描述推断，"
    "不确定就选\"现代简约商务风\"）、主色调与配色、排版布局（标题/正文/图形的摆放与层级）。\n"
    "3. 需要出现文字时，把每段要渲染的文字用双引号精确写出，并强调\"文字必须清晰、无错别字、"
    "对齐正确、不要截断\"。\n"
    "4. 若页面需要示意图表（流程图、思维导图、气泡图、时间轴、韦恩图等），"
    "请在提示词中明确写出图表类型、整体结构、节点/层级、连接关系（箭头/连线）以及每个节点要标注的文字，"
    "文字用中文并用双引号精确标出，并强调\"图表布局清晰、节点对齐、连线方向正确、标签不重叠\"。\n"
    "5. 补充画面质感、光线、留白、细节等能让画面更精致的描述。\n"
    "6. 只输出最终提示词本身，用一段连续文字输出，不要任何解释、不要分条、不要加标题或前后缀。"
)

STYLE_MASTER_SYSTEM = (
    "你是一位演示文稿视觉设计师与 AI 生图提示词专家。"
    "用户会给出一个“风格意图”（可能是中文风格描述，也可能是若干页课件的内容描述），"
    "请据此输出一段用于 gpt-image-2 生成 16:9 演示文稿“风格母版样张”的高质量英文提示词。\n\n"
    "要求：\n"
    "1. 这是一张纯“视觉风格样张”，不含任何具体文字、标题或正文，只体现整体设计系统："
    "主色调与配色方案、背景质感、字体风格、装饰元素、版式骨架、图标/插图风格、光影与留白。\n"
    "2. 若用户给的是页面内容，请先从内容推断适合的教学课件风格再输出。例如信息技术/计算机/网络类"
    "课程宜偏科技感：深蓝或靛蓝主色、几何线条、电路/代码/网络节点等装饰、扁平图标、清晰层级与留白。\n"
    "3. 明确写出：16:9 横向构图、风格关键词、主色调与配色、版式骨架与装饰元素。\n"
    "4. 强调“画面不含可读文字、仅展示统一视觉风格”。\n"
    "5. 只输出一段连续英文提示词，不要解释、不要分条、不要标题或前后缀。"
)

OUTLINE_SYSTEM = (
    "你是一位资深课件设计师，擅长把资料整理成适合学生学习的 PPT。\n"
    "请根据用户提供的文档内容，生成一份 {page_count} 页的 PPT 大纲。\n"
    "目标受众：{grade}学生。\n"
    "要求：\n"
    "1. 严格输出 JSON 数组，每页一个对象，字段为：title（本页标题）、"
    "points（本页要点的字符串数组，数量按内容多少灵活决定，通常 2~8 个、宁精勿滥）、"
    "visual（配图建议，一句话描述画面应包含的元素；若该页用图表表达更清晰，请描述图表类型与结构）、"
    "chart（可选字段，本页最适合的图表类型，取值仅限 flowchart/mindmap/bubble/timeline/venn/radar/bar/pie，"
    "普通配图页省略该字段）。\n"
    "2. 标题简洁、要点精炼，语言贴合 {grade} 学生的理解水平。\n"
    "3. 覆盖文档的核心内容，逻辑由浅入深、层层递进。\n"
    "4. 只输出 JSON 本身，不要任何解释、不要 markdown 代码块标记。"
)

SPLIT_SYSTEM = (
    "你是一位资深课件文案专家，擅长把内容改写成适合学生理解的 PPT 页面文案。\n"
    "目标受众：{grade}学生。\n"
    "用户会给出整份 PPT 的大纲（JSON）和原始文档内容，请为每一页生成最终页面文案。\n"
    "要求：\n"
    "1. 严格输出 JSON 数组，每页一个对象，字段为：title（标题）、"
    "content（本页正文，用换行符分隔的简练条目，条目数量按内容灵活决定）、"
    "visual（配图建议，描述画面应包含的具体元素与场景；若用图表，描述图表类型与结构）、"
    "chart（可选字段，本页图表类型，取值仅限 flowchart/mindmap/bubble/timeline/venn/radar/bar/pie，普通页面省略）、"
    "chart_data（可选字段，仅当 chart 为 radar/bar/pie 且原文有明确数据时输出，结构为 title 图表标题、"
    "labels 分类名数组、values 数值数组；数据必须来自原文、不得编造，无明确数据则省略该字段）。\n"
    "2. 文案认知适配：{grade_hint}。\n"
    "3. 每页信息量适中，避免堆砌，标题与正文呼应。\n"
    "4. 只输出 JSON 本身，不要任何解释、不要 markdown 代码块标记。"
)

GRADE_HINTS = {
    "小学": "语言通俗、多用生活化比喻和拟人化表达、句子短、画面感强",
    "初中": "概念用生活例子解释、由具体到抽象、可设置小问题引发思考",
    "高中": "逻辑清晰、结构化表达、可包含概念关系与推导",
    "大学": "术语准确、体系完整、专业严谨",
}


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _ok(data):
    return jsonify({"ok": True, **data})


def _err(msg, status=400):
    return jsonify({"ok": False, "error": msg}), status


def _validate_filename(name):
    """只允许纯文件名，防止路径穿越。"""
    return os.path.basename(name)


def _download_image(url, headers=None, skip_proxy=False):
    """中转站可能返回图片 URL 而不是 base64，这里负责下载。"""
    proxies = _proxy_cfg(skip_proxy)
    r = requests.get(url, headers=headers or {}, timeout=120, proxies=proxies)
    r.raise_for_status()
    return r.content


def _proxy_cfg(skip_proxy):
    """skip_proxy=True 时让 requests 忽略系统/环境变量里的代理，直接连接。"""
    if skip_proxy:
        return {"http": None, "https": None}
    return None


def _save_image(img_bytes, index, prefix="page"):
    """按真实图片格式确定扩展名并落盘，兼容 png/jpeg/webp 等。

    prefix 用于区分页面图（page_）与风格母版候选图（master_）。
    """
    ext = "png"
    try:
        im = Image.open(io.BytesIO(img_bytes))
        fmt = (im.format or "PNG").lower()
        ext = "jpg" if fmt == "jpeg" else (fmt if fmt in ("png", "webp") else "png")
    except Exception:
        ext = "png"
    fname = f"{prefix}_{int(index):02d}_{int(time.time() * 1000)}.{ext}"
    with open(os.path.join(IMAGES_DIR, fname), "wb") as f:
        f.write(img_bytes)
    return fname


def _candidate_urls(base_url):
    """生成候选请求地址：兼容「带 /v1」和「不带 /v1」两种中转站写法。"""
    base = base_url.rstrip("/")
    if base.endswith("/v1"):
        return [base + "/images/generations"]
    return [base + "/images/generations", base + "/v1/images/generations"]


def _candidate_urls_edits(base_url):
    base = base_url.rstrip("/")
    if base.endswith("/v1"):
        return [base + "/images/edits"]
    return [base + "/images/edits", base + "/v1/images/edits"]


def _call_deepseek(key, base_url, model, system, user, temperature=0.7, max_tokens=1400, timeout=120, skip_proxy=False):
    """统一的 DeepSeek chat 调用。"""
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }
    if "reasoner" in model:
        body["max_tokens"] = max_tokens  # reasoner 不支持 temperature
    else:
        body["temperature"] = temperature
        body["max_tokens"] = max_tokens
    proxies = _proxy_cfg(skip_proxy)
    try:
        r = requests.post(
            base_url + "/chat/completions",
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
            json=body,
            timeout=timeout,
            proxies=proxies,
        )
    except requests.exceptions.ProxyError as e:
        raise RuntimeError(
            "代理连接失败（ProxyError）：当前系统/网络设置了代理但无法连通目标服务器。"
            "请在「① API 设置」勾选「跳过系统代理（直连）」后重试。详情：" + str(e)
        )
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}：{r.text[:300]}")
    payload = r.json()
    return payload["choices"][0]["message"]["content"].strip()


def _parse_json(content):
    """从模型返回文本里稳健提取 JSON（容忍 markdown 代码块包裹）。"""
    content = content.strip()
    if content.startswith("```"):
        lines = content.split("\n")
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        content = "\n".join(lines).strip()
    try:
        return json.loads(content)
    except Exception:
        start = content.find("[")
        end = content.rfind("]")
        if start != -1 and end != -1 and end > start:
            return json.loads(content[start:end + 1])
        raise


def _read_docx(content):
    """从 docx 字节流提取纯文本（段落 + 表格）。"""
    from docx import Document
    doc = Document(io.BytesIO(content))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _extract_text(content, ext):
    """按扩展名解析文本；txt/md 做多编码兜底。"""
    ext = (ext or "").lower()
    if ext in (".docx", ".doc"):
        return _read_docx(content)
    raw = content
    if isinstance(raw, bytes):
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                return raw.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return raw.decode("utf-8", errors="ignore")
    return raw


def _text2img_once(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy=False, prefix="page"):
    """generations 端点：纯文本生图（纯逻辑，返回 dict，不依赖 Flask 上下文）。

    返回 {"ok": True, "filename":..., "url":...} 或 {"ok": False, "error":...}。
    prefix 决定落盘文件名前缀：页面图用 page_，风格母版候选图用 master_。
    可在工作线程中安全调用（不触碰 jsonify / current_app）。
    """
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    proxies = _proxy_cfg(skip_proxy)

    def make_body(with_extra):
        body = {"model": model, "prompt": prompt, "n": 1, "size": size}
        if with_extra:
            body["quality"] = quality
            body["format"] = fmt
        return body

    last_err = None
    proxy_err = False
    for url in _candidate_urls(base_url):
        for with_extra in (True, False):
            try:
                r = requests.post(url, headers=headers, json=make_body(with_extra), timeout=300, proxies=proxies)
                if r.status_code in (401, 403):
                    return {"ok": False, "error": f"鉴权失败（HTTP {r.status_code}）：{r.text[:300]}。请检查 Key 或 API 地址"}
                if r.status_code == 404:
                    last_err = f"HTTP 404：{r.text[:200]}（路径 {url} 不存在）"
                    break
                if r.status_code >= 400:
                    last_err = f"HTTP {r.status_code}：{r.text[:300]}"
                    continue
                payload = r.json()
                item = (payload.get("data") or [{}])[0]
                img_bytes = None
                if item.get("b64_json"):
                    img_bytes = base64.b64decode(item["b64_json"])
                elif item.get("url"):
                    img_bytes = _download_image(item["url"], headers, skip_proxy)
                if not img_bytes:
                    last_err = "接口返回里既没有 b64_json 也没有 url，可能是非标准中转站格式"
                    continue
                fname = _save_image(img_bytes, index, prefix)
                return {"ok": True, "filename": fname, "url": f"/output/images/{fname}"}
            except requests.exceptions.ProxyError as e:
                proxy_err = True
                last_err = ("代理连接失败（ProxyError）：当前系统/网络设置了代理，但代理无法连通目标服务器。"
                            "请在「① API 设置」勾选「跳过系统代理（直连）」后重试。详情：" + str(e))
                break
            except requests.exceptions.RequestException as e:
                last_err = f"网络请求失败：{e}"
                continue
            except Exception as e:  # noqa: BLE001
                last_err = f"图片生成失败：{e}"
                continue
        if proxy_err:
            break
    return {"ok": False, "error": f"图片生成失败：{last_err}"}


def _gen_text2img(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy=False, prefix="page"):
    """generations 端点视图包装：调纯逻辑函数并构造 Flask 响应（仅在请求线程调用）。"""
    r = _text2img_once(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy, prefix)
    if r.get("ok"):
        return _ok({"filename": r["filename"], "url": r["url"]})
    return _err(r.get("error") or "图片生成失败")


# 锁长相 / 锁风格指令：edits 端点把参考图作为强条件，约束后续出图
_LOCK_FACE = "Keep every character's face, hairstyle and clothing identical to the reference image."
_LOCK_STYLE = (
    "Strictly preserve the overall visual style of the reference image in this new image: "
    "keep the exact same color palette, background texture and tone, typography and font style, "
    "decorative motifs, layout skeleton, and overall aesthetic as the reference. "
    "If any character appears, also keep their face, hairstyle and clothing identical. "
    "Only change the content and text as described in the prompt."
)


def _gen_edit(key, base_url, model, prompt, size, index, ref_path, skip_proxy=False, lock_scope="style"):
    """edits 端点：传参考图锁长相 / 锁整体风格。

    lock_scope:
      "style" —— 锁定整页视觉风格（配色/背景/字体/装饰/版式）+ 人物（默认，用于统一整批课件风格）
      "face"  —— 仅锁定人物长相、发型、服装
    """
    headers = {"Authorization": f"Bearer {key}"}
    proxies = _proxy_cfg(skip_proxy)
    with open(ref_path, "rb") as f:
        ref_bytes = f.read()
    ext = os.path.splitext(ref_path)[1].lower().lstrip(".")
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "webp": "image/webp"}.get(ext, "image/png")
    files = {"image": (os.path.basename(ref_path), ref_bytes, mime)}
    lock = _LOCK_STYLE if lock_scope != "face" else _LOCK_FACE
    data = {"model": model, "prompt": prompt + " " + lock, "size": size}

    last_err = None
    for url in _candidate_urls_edits(base_url):
        try:
            r = requests.post(url, headers=headers, files=files, data=data, timeout=300, proxies=proxies)
            if r.status_code in (401, 403):
                return _err(f"鉴权失败（HTTP {r.status_code}）：{r.text[:300]}。请检查 Key 或 API 地址")
            if r.status_code == 404:
                last_err = f"HTTP 404：{r.text[:200]}（edits 路径不存在，该中转站可能不支持参考图）"
                break
            if r.status_code >= 400:
                last_err = f"HTTP {r.status_code}：{r.text[:300]}"
                continue
            payload = r.json()
            item = (payload.get("data") or [{}])[0]
            img_bytes = None
            if item.get("b64_json"):
                img_bytes = base64.b64decode(item["b64_json"])
            elif item.get("url"):
                img_bytes = _download_image(item["url"], headers, skip_proxy)
            if not img_bytes:
                last_err = "接口返回里既没有 b64_json 也没有 url"
                continue
            fname = _save_image(img_bytes, index)
            return _ok({"filename": fname, "url": f"/output/images/{fname}"})
        except requests.exceptions.ProxyError as e:
            return _err("代理连接失败（ProxyError）：当前系统/网络设置了代理，但代理无法连通目标服务器。"
                        "请在「① API 设置」勾选「跳过系统代理（直连）」后重试。详情：" + str(e))
        except requests.exceptions.RequestException as e:
            last_err = f"网络请求失败：{e}"
            continue
        except Exception as e:  # noqa: BLE001
            last_err = f"图片生成失败：{e}"
            continue
    return _err(f"图片生成失败：{last_err}")


# ---------------------------------------------------------------------------
# 页面
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/output/<path:filename>")
def serve_output(filename):
    return send_from_directory(OUTPUT_DIR, filename)


# ---------------------------------------------------------------------------
# B 链：文档解析 / 大纲 / 文案
# ---------------------------------------------------------------------------
@app.route("/api/parse_doc", methods=["POST"])
def parse_doc():
    # 支持文件上传（docx/txt/md）或 JSON 里直接粘贴文本
    if request.content_type and "multipart" in request.content_type:
        f = request.files.get("file")
        if not f:
            return _err("未收到文件")
        ext = os.path.splitext(f.filename or "")[1].lower()
        text = _extract_text(f.read(), ext)
    else:
        data = request.get_json(force=True, silent=True) or {}
        text = (data.get("text") or "").strip()
        if not text:
            return _err("请上传文件或直接粘贴文本")
    text = (text or "").strip()
    if not text:
        return _err("未能从文档中提取到文字，请检查文件格式")
    return _ok({"text": text, "length": len(text)})


@app.route("/api/outline", methods=["POST"])
def outline():
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("deepseek_key") or "").strip()
    base_url = (data.get("deepseek_base_url") or "https://api.deepseek.com").strip().rstrip("/")
    model = (data.get("deepseek_model") or "deepseek-chat").strip()
    text = (data.get("text") or "").strip()
    page_count = data.get("page_count") or 8
    grade = (data.get("grade") or "通用").strip()
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 DeepSeek API Key")
    if not text:
        return _err("请先上传或粘贴文档内容")

    try:
        page_count = max(1, min(int(page_count), 40))
    except (TypeError, ValueError):
        page_count = 8

    system = OUTLINE_SYSTEM.format(page_count=page_count, grade=grade)
    # 文档过长则截断（保留开头 + 结尾），防止超上下文
    user = text if len(text) <= 12000 else text[:9000] + "\n\n……(中间省略)……\n\n" + text[-3000:]

    try:
        content = _call_deepseek(key, base_url, model, system, user, max_tokens=3000, skip_proxy=skip_proxy)
        outline_json = _parse_json(content)
        if not isinstance(outline_json, list):
            return _err("大纲格式异常，请重试")
        return _ok({"outline": outline_json})
    except Exception as e:  # noqa: BLE001
        return _err(f"大纲生成失败：{e}", 500)


@app.route("/api/split_content", methods=["POST"])
def split_content():
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("deepseek_key") or "").strip()
    base_url = (data.get("deepseek_base_url") or "https://api.deepseek.com").strip().rstrip("/")
    model = (data.get("deepseek_model") or "deepseek-chat").strip()
    text = (data.get("text") or "").strip()
    outline = data.get("outline") or []
    grade = (data.get("grade") or "通用").strip()
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 DeepSeek API Key")
    if not outline:
        return _err("大纲为空，请先生成大纲")

    grade_hint = GRADE_HINTS.get(grade, "通俗易懂、条理清晰、重点突出")
    system = SPLIT_SYSTEM.format(grade=grade, grade_hint=grade_hint)
    src = text if len(text) <= 10000 else text[:8000] + "\n\n……(中间省略)……\n\n" + text[-2000:]
    user = "PPT 大纲（JSON）：\n" + json.dumps(outline, ensure_ascii=False) + "\n\n原始文档内容：\n" + src

    try:
        content = _call_deepseek(key, base_url, model, system, user, max_tokens=4000, skip_proxy=skip_proxy)
        pages = _parse_json(content)
        if not isinstance(pages, list):
            return _err("文案格式异常，请重试")
        return _ok({"pages": pages})
    except Exception as e:  # noqa: BLE001
        return _err(f"文案拆解失败：{e}", 500)


# ---------------------------------------------------------------------------
# 设定图上传
# ---------------------------------------------------------------------------
@app.route("/api/upload_ref", methods=["POST"])
def upload_ref():
    f = request.files.get("file")
    if not f:
        return _err("未收到文件")
    ext = os.path.splitext(f.filename or "")[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".webp"):
        return _err("请上传 png / jpg / webp 图片")
    data = f.read()
    if not data:
        return _err("文件为空")
    e = ext.lstrip(".")
    if e == "jpeg":
        e = "jpg"
    fname = f"ref_{int(time.time() * 1000)}.{e}"
    with open(os.path.join(IMAGES_DIR, fname), "wb") as out:
        out.write(data)
    return _ok({"filename": fname, "url": f"/output/images/{fname}"})


# ---------------------------------------------------------------------------
# AI 生成风格母版（可选）：按风格意图生成 1~3 张候选样张
# ---------------------------------------------------------------------------
@app.route("/api/gen_style_master", methods=["POST"])
def gen_style_master():
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("openai_key") or "").strip()
    base_url = (data.get("openai_base_url") or "").strip().rstrip("/")
    model = (data.get("openai_model") or data.get("model") or "gpt-image-2").strip()
    size = data.get("size") or "1920x1080"
    quality = data.get("quality") or "high"
    fmt = data.get("format") or "png"
    skip_proxy = data.get("skip_proxy") or False

    style_desc = (data.get("style_desc") or "").strip()
    pages = data.get("pages") or []
    try:
        count = int(data.get("count") or 2)
    except Exception:
        count = 2
    count = max(1, min(3, count))

    ds_key = (data.get("deepseek_key") or "").strip()
    ds_base = (data.get("deepseek_base_url") or "https://api.deepseek.com").strip().rstrip("/")
    ds_model = (data.get("deepseek_model") or "deepseek-chat").strip()

    if not key:
        return _err("请先填写 OpenAI API Key")
    if not base_url:
        return _err("请先填写 OpenAI API 地址")
    if not style_desc and not pages:
        return _err("请填写风格描述，或先添加页面内容以便自动归纳风格")

    # 由 DeepSeek 把风格意图 → 英文母版生图 prompt（无 Key 则套固定模板）
    prompt = None
    if ds_key:
        try:
            if style_desc:
                user_msg = f"风格意图：{style_desc}"
            else:
                joined = "\n".join([p for p in pages if p][:20])
                user_msg = ("以下是若干页课件的内容描述，请据此归纳适合的整体视觉风格"
                            "并输出母版提示词：\n\n" + joined)
            prompt = _call_deepseek(ds_key, ds_base, ds_model, STYLE_MASTER_SYSTEM, user_msg,
                                    temperature=0.7, max_tokens=900, skip_proxy=skip_proxy)
            prompt = (prompt or "").strip() or None
        except Exception as e:  # noqa: BLE001
            return _err(f"风格归纳失败：{e}")

    if not prompt:
        intent = style_desc or "modern minimal teaching slide style"
        prompt = (
            "A 16:9 presentation slide master design sheet, NO readable text content. "
            "Visual style: " + intent + ". "
            "Consistent color palette, background texture, typography, decorative motifs, "
            "iconography, layout grid, lighting and whitespace. "
            "Clean, professional, suitable for classroom teaching slides. "
            "No letters, no words, no captions — purely a visual style reference."
        )

    # 并发生成 count 张候选样张（master_ 前缀，避免与页面图混淆）
    results = [None] * count

    def _one(i):
        return i, _text2img_once(key, base_url, model, prompt, size, quality, fmt, i,
                                 skip_proxy=skip_proxy, prefix="master")

    with ThreadPoolExecutor(max_workers=count) as ex:
        futs = [ex.submit(_one, i) for i in range(count)]
        for fu in as_completed(futs):
            i, payload = fu.result()
            results[i] = payload

    candidates = []
    last_err = None
    for payload in results:
        if not payload:
            continue
        if payload.get("ok"):
            candidates.append({"filename": payload["filename"], "url": payload["url"]})
        else:
            last_err = payload.get("error")
    if not candidates:
        return _err(last_err or "候选图生成失败")
    return _ok({"candidates": candidates})


# ---------------------------------------------------------------------------
# 精确图表绘制（matplotlib：雷达图 / 柱状图 / 饼图）
# ---------------------------------------------------------------------------
@app.route("/api/render_chart", methods=["POST"])
def render_chart():
    data = request.get_json(force=True, silent=True) or {}
    chart_type = (data.get("chart") or "").strip()
    chart_data = data.get("chart_data") or {}
    index = data.get("index") or 0

    if not chart_type:
        return _err("缺少图表类型")

    try:
        from chart_renderer import render as render_chart_img
        png_bytes = render_chart_img(chart_type, chart_data)
    except Exception as e:  # noqa: BLE001
        return _err(f"图表绘制失败：{e}", 500)

    fname = f"chart_{int(index):02d}_{int(time.time() * 1000)}.png"
    with open(os.path.join(IMAGES_DIR, fname), "wb") as f:
        f.write(png_bytes)
    return _ok({"filename": fname, "url": f"/output/images/{fname}"})


# ---------------------------------------------------------------------------
# 优化提示词（DeepSeek）
# ---------------------------------------------------------------------------
@app.route("/api/optimize", methods=["POST"])
def optimize():
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("deepseek_key") or "").strip()
    description = (data.get("description") or "").strip()
    base_url = (data.get("deepseek_base_url") or "https://api.deepseek.com").strip().rstrip("/")
    model = (data.get("deepseek_model") or "deepseek-chat").strip()
    style = (data.get("style") or "").strip()
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 DeepSeek API Key")
    if not description:
        return _err("页面描述不能为空")

    user_msg = description
    if style:
        user_msg = f"统一风格要求：{style}\n\n页面内容描述：{description}"

    try:
        content = _call_deepseek(key, base_url, model, OPTIMIZE_SYSTEM, user_msg, skip_proxy=skip_proxy)
        return _ok({"prompt": content})
    except Exception as e:  # noqa: BLE001
        return _err(f"DeepSeek 调用失败：{e}", 500)


# ---------------------------------------------------------------------------
# 生成图片（gpt-image-2，可选锁长相）
# ---------------------------------------------------------------------------
@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("openai_key") or "").strip()
    prompt = (data.get("prompt") or "").strip()
    base_url = (data.get("openai_base_url") or "").strip().rstrip("/")
    model = (data.get("model") or "gpt-image-2").strip()
    size = data.get("size") or "1280x720"
    quality = data.get("quality") or "high"
    fmt = data.get("format") or "png"
    index = data.get("index") or 0
    ref_image = (data.get("ref_image") or "").strip()
    lock_scope = (data.get("lock_scope") or "style").strip()
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 OpenAI API Key")
    if not base_url:
        return _err("请先填写 OpenAI API 地址")
    if not prompt:
        return _err("提示词不能为空")

    if ref_image:
        safe = _validate_filename(ref_image)
        ref_path = os.path.join(IMAGES_DIR, safe)
        if not os.path.isfile(ref_path):
            return _err("风格母版不存在，请重新上传")
        return _gen_edit(key, base_url, model, prompt, size, index, ref_path, skip_proxy, lock_scope)

    return _gen_text2img(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy)


# ---------------------------------------------------------------------------
# 拼装 PPTX
# ---------------------------------------------------------------------------
@app.route("/api/build_ppt", methods=["POST"])
def build_ppt():
    data = request.get_json(force=True, silent=True) or {}
    images = data.get("images") or []
    title = (data.get("title") or "演示文稿").strip() or "演示文稿"

    if not images:
        return _err("还没有生成任何图片，请先生成")

    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9 宽屏
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for fn in images:
        safe = _validate_filename(fn)
        fpath = os.path.join(IMAGES_DIR, safe)
        if not os.path.isfile(fpath):
            return _err(f"找不到图片：{safe}")
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(fpath, 0, 0, width=prs.slide_width, height=prs.slide_height)

    deck_name = f"{title}_{int(time.time() * 1000)}.pptx"
    deck_path = os.path.join(OUTPUT_DIR, deck_name)
    prs.save(deck_path)
    return _ok({"filename": deck_name, "url": f"/output/{deck_name}"})


# ---------------------------------------------------------------------------
# 打包全部图片为 ZIP
# ---------------------------------------------------------------------------
@app.route("/api/zip_images", methods=["POST"])
def zip_images():
    data = request.get_json(force=True, silent=True) or {}
    images = data.get("images") or []
    if not images:
        return _err("还没有生成任何图片")

    zname = f"images_{int(time.time() * 1000)}.zip"
    zpath = os.path.join(OUTPUT_DIR, zname)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn in images:
            safe = _validate_filename(fn)
            fpath = os.path.join(IMAGES_DIR, safe)
            if os.path.isfile(fpath):
                zf.write(fpath, arcname=safe)
    return _ok({"filename": zname, "url": f"/output/{zname}"})


if __name__ == "__main__":
    print("=" * 52)
    print("  PPT 智能生成工具已启动")
    print("  请在浏览器打开： http://127.0.0.1:5000")
    print("  按 Ctrl+C 停止服务")
    print("=" * 52)
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
