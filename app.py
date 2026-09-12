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
import hashlib
import io
import json
import os
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote

import requests
from flask import Flask, jsonify, request, send_from_directory
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
MASTERS_DIR = os.path.join(OUTPUT_DIR, "masters")  # 母版参考图单独存放
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(MASTERS_DIR, exist_ok=True)

# 自定义图片目录：由前端「① API 设置 → 图片保存目录」随请求传入；
# None 表示使用默认目录（output/images + output/masters）。
CUSTOM_IMAGES_DIR = None
_KNOWN_CUSTOM_DIRS = []   # 本会话出现过的自定义目录（新→旧），用于跨目录查找历史图片

# 当前项目子文件夹名（如 "20260831_计算机网络基础"）：
# 页面图 / 图表落到 <images_dir>/<project_name>/，母版 / 参考图仍统一存 <images_dir>/masters/。
CURRENT_PROJECT_NAME = None
_KNOWN_PROJECT_DIRS = []  # 历史项目子文件夹名（新→旧），用于跨项目查找历史图片

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


def _remember_images_dir(images_dir):
    """记录当前生效的自定义图片目录；空值回到默认目录。"""
    global CUSTOM_IMAGES_DIR
    d = (images_dir or "").strip()
    if d:
        d = os.path.abspath(d)
        if d not in _KNOWN_CUSTOM_DIRS:
            _KNOWN_CUSTOM_DIRS.insert(0, d)
        CUSTOM_IMAGES_DIR = d
    else:
        CUSTOM_IMAGES_DIR = None


def _sanitize_project_name(name):
    """清洗项目名：去掉 Windows 文件名非法字符，空格→下划线，折叠连续下划线。"""
    if not name:
        return ""
    safe = name.strip()
    for ch in '\\/:*?"<>|':
        safe = safe.replace(ch, "")
    safe = safe.replace(" ", "_")
    while "__" in safe:
        safe = safe.replace("__", "_")
    return safe.strip("_")


def _build_project_name(raw_name):
    """组合成 YYYYMMDD_<项目名>；raw_name 为空时仅用 YYYYMMDD。"""
    date_part = time.strftime("%Y%m%d")
    safe = _sanitize_project_name(raw_name)
    return f"{date_part}_{safe}" if safe else date_part


def _make_project_name_unique(parent_dir, base_name):
    """若 parent_dir 下已存在同名子文件夹，追加 _2 / _3 ... 直到不重名。"""
    if not base_name:
        base_name = time.strftime("%Y%m%d")
    candidate = base_name
    n = 2
    while os.path.isdir(os.path.join(parent_dir, candidate)):
        candidate = f"{base_name}_{n}"
        n += 1
    return candidate


def _remember_project(project_name):
    """记录当前活跃项目子文件夹；空值清除。同时加入历史列表用于跨项目查找。"""
    global CURRENT_PROJECT_NAME
    pn = (project_name or "").strip()
    if pn:
        if pn not in _KNOWN_PROJECT_DIRS:
            _KNOWN_PROJECT_DIRS.insert(0, pn)
        CURRENT_PROJECT_NAME = pn
    else:
        CURRENT_PROJECT_NAME = None


def _save_dirs(images_dir=None, project_name=None):
    """返回 (页面图保存目录, 母版保存目录)，目录不存在则自动创建。

    页面图 / 图表存 <images_dir>/<project_name>/，母版 / 参考图存 <images_dir>/masters/。
    project_name 为空时退回到 <images_dir> 根目录（兼容旧逻辑）。
    images_dir 为空时用默认 IMAGES_DIR / MASTERS_DIR。
    """
    d = (images_dir or "").strip()
    pn = (project_name or "").strip()
    if d:
        root = os.path.abspath(d)
        mas = os.path.join(root, "masters")
        os.makedirs(mas, exist_ok=True)
        if pn:
            img = os.path.join(root, pn)
            os.makedirs(img, exist_ok=True)
        else:
            img = root
            os.makedirs(img, exist_ok=True)
        return img, mas
    os.makedirs(IMAGES_DIR, exist_ok=True)
    os.makedirs(MASTERS_DIR, exist_ok=True)
    if pn:
        img = os.path.join(IMAGES_DIR, pn)
        os.makedirs(img, exist_ok=True)
        return img, MASTERS_DIR
    return IMAGES_DIR, MASTERS_DIR


def _image_roots():
    """按优先级列出所有可能存放图片的目录。

    优先级：当前自定义目录的当前项目子文件夹 → 当前自定义目录 → 当前目录的 masters →
    当前自定义目录下的历史项目子文件夹 → 历史自定义目录及其项目子文件夹/masters → 默认目录。
    """
    roots, seen = [], set()

    def add(p):
        p = os.path.abspath(p)
        if p not in seen:
            seen.add(p)
            roots.append(p)

    if CUSTOM_IMAGES_DIR:
        # 当前活跃项目子文件夹优先级最高，便于重新生成时跨目录定位最新图
        if CURRENT_PROJECT_NAME:
            add(os.path.join(CUSTOM_IMAGES_DIR, CURRENT_PROJECT_NAME))
        add(CUSTOM_IMAGES_DIR)
        add(os.path.join(CUSTOM_IMAGES_DIR, "masters"))
        # 当前自定义目录下所有历史项目子文件夹
        for pn in _KNOWN_PROJECT_DIRS:
            add(os.path.join(CUSTOM_IMAGES_DIR, pn))
    for d in _KNOWN_CUSTOM_DIRS:
        if d != CUSTOM_IMAGES_DIR:
            add(d)
            for pn in _KNOWN_PROJECT_DIRS:
                add(os.path.join(d, pn))
            add(os.path.join(d, "masters"))
    add(IMAGES_DIR)
    add(MASTERS_DIR)
    return roots


def _find_image(fname):
    """跨目录按文件名定位图片，返回绝对路径；找不到返回 None。"""
    safe = _validate_filename(fname)
    for d in _image_roots():
        p = os.path.join(d, safe)
        if os.path.isfile(p):
            return p
    return None


def _image_url(fname):
    """统一的图片访问地址：按文件名跨目录定位，与图片实际存放目录无关。"""
    return f"/api/image/{fname}"


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


def _save_image(img_bytes, index, prefix="page", target_dir=None):
    """按真实图片格式确定扩展名并落盘，兼容 png/jpeg/webp 等。

    prefix 用于区分页面图（page_）与风格母版候选图（master_）。
    target_dir 指定落盘目录，默认 IMAGES_DIR；母版图传 MASTERS_DIR。
    """
    ext = "png"
    try:
        im = Image.open(io.BytesIO(img_bytes))
        fmt = (im.format or "PNG").lower()
        ext = "jpg" if fmt == "jpeg" else (fmt if fmt in ("png", "webp") else "png")
    except Exception:
        ext = "png"
    fname = f"{prefix}_{int(index):02d}_{int(time.time() * 1000)}.{ext}"
    save_dir = target_dir or IMAGES_DIR
    with open(os.path.join(save_dir, fname), "wb") as f:
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


def _text2img_once(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy=False, prefix="page", images_dir=None, project_name=None):
    """generations 端点：纯文本生图（纯逻辑，返回 dict，不依赖 Flask 上下文）。

    返回 {"ok": True, "filename":..., "url":...} 或 {"ok": False, "error":...}。
    prefix 决定落盘文件名前缀：页面图用 page_，风格母版候选图用 master_。
    images_dir 指定自定义图片保存根目录；project_name 指定项目子文件夹名（页面图落到其下，
    母版图存 masters 子文件夹）。可在工作线程中安全调用（不触碰 jsonify / current_app）。
    """
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    proxies = _proxy_cfg(skip_proxy)
    img_dir, mas_dir = _save_dirs(images_dir, project_name)

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
                is_master = prefix == "master"
                target_dir = mas_dir if is_master else img_dir
                fname = _save_image(img_bytes, index, prefix, target_dir=target_dir)
                return {"ok": True, "filename": fname, "url": _image_url(fname)}
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


def _gen_text2img(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy=False, prefix="page", images_dir=None, project_name=None):
    """generations 端点视图包装：调纯逻辑函数并构造 Flask 响应（仅在请求线程调用）。"""
    r = _text2img_once(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy, prefix, images_dir, project_name)
    if r.get("ok"):
        return _ok({"filename": r["filename"], "url": r["url"]})
    return _err(r.get("error") or "图片生成失败")


# ---------------------------------------------------------------------------
# 多参考图：角色清单 / 拼贴降级 / 能力探测
# ---------------------------------------------------------------------------
# GPT Image 系列在 /images/edits 上最多接受 16 张输入图；
# 前端默认只建议 6 张，张数越多指令互相打架的概率越高。
HARD_REF_LIMIT = 16
MAX_REF_IMAGES = 6
REF_INDEX_NAME = "refs_index.json"

# 角色 → 写入 prompt 清单块的英文标签
REF_ROLE_LABEL = {
    "style": "STYLE MASTER",
    "character": "CHARACTER",
    "layout": "LAYOUT",
    "element": "ELEMENT",
    "custom": "EXTRA INSTRUCTION",
}
# 冲突消解优先级：风格 > 人物 > 版式 > 素材 > 自定义
REF_ROLE_ORDER = {"style": 0, "character": 1, "layout": 2, "element": 3, "custom": 4}

_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS


def _role_instruction(role, who, note=""):
    """按角色生成该张参考图对应的英文指令（who 形如 "image 1" 或 "panel 1"）。"""
    if role == "character":
        return ("keep the face, hairstyle, clothing and body proportions of the character "
                f"in {who} identical")
    if role == "layout":
        return (f"follow only the composition skeleton of {who} - the zones for title, body "
                "and visuals plus their relative proportions; do NOT copy its colours or text")
    if role == "element":
        return f"reuse the object, icon or illustration shown in {who} as a visual asset"
    if role == "custom":
        n = (note or "").strip()
        if n:
            return f"apply this extra instruction tied to {who}: {n}"
        return f"use {who} as an additional visual reference"
    return (f"reproduce the overall visual style of {who}: the exact colour palette, "
            "background texture and tone, typography feel, decorative motifs, layout grid, "
            "lighting and whitespace")


def build_ref_manifest(items, mode="multi"):
    """把有序参考图集编译成注入 prompt 的英文清单块。

    items: [{"role":..., "note":...}, ...]，顺序即模型眼中的编号。
    mode="multi"     —— 按 Image 1 / 2 / 3 引用（原生多图）。
    mode="composite" —— 按 contact sheet 的 Panel 1 / 2 / 3 引用（拼贴单图降级）。
    """
    n = len(items)
    if mode == "composite":
        head = (f"A single reference image is attached. It is a contact sheet made of {n} numbered "
                "panels separated by thin grey dividers, each panel carrying its number in the "
                "bottom-right corner. The dividers and the corner numbers are layout guides only - "
                "do NOT render them in the output. Treat each panel as a separate reference:")
    elif n == 1:
        head = "One reference image is attached. Follow it according to this role:"
    else:
        head = f"{n} reference images are attached, in order. Follow each one according to its role:"
    lines = [head]
    for k, it in enumerate(items, start=1):
        role = (it.get("role") or "style").strip()
        if role not in REF_ROLE_LABEL:
            role = "style"
        who = f"panel {k}" if mode == "composite" else f"image {k}"
        lines.append(f"{k}. {REF_ROLE_LABEL[role]} - {who.capitalize()}: "
                     f"{_role_instruction(role, who, it.get('note'))}.")
    lines.append("Do NOT copy any text, logo or watermark from the reference images.")
    return "\n".join(lines)


def _compose_contact_sheet(paths, cell_max=1024, pad=10):
    """把多张参考图拼成一张带编号角标的 contact sheet（L2 降级用），返回 PNG 字节。"""
    imgs = []
    for p in paths:
        try:
            im = Image.open(p)
            im.load()
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[-1])
                im = bg
            elif im.mode != "RGB":
                im = im.convert("RGB")
            imgs.append(im.copy())
        except Exception:  # noqa: BLE001
            continue
    if not imgs:
        raise RuntimeError("没有可用的参考图")
    n = len(imgs)
    cols = 1 if n == 1 else (2 if n <= 4 else 3)
    rows = (n + cols - 1) // cols
    cell_w = min(max(im.width for im in imgs), cell_max)
    cell_h = min(max(im.height for im in imgs), cell_max)
    sheet = Image.new("RGB", (cols * cell_w + (cols + 1) * pad, rows * cell_h + (rows + 1) * pad),
                      (255, 255, 255))
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.load_default()
    except Exception:  # noqa: BLE001
        font = None
    for i, im in enumerate(imgs):
        r, c = divmod(i, cols)
        sc = min(cell_w / max(1, im.width), cell_h / max(1, im.height))
        w, h = max(1, int(im.width * sc)), max(1, int(im.height * sc))
        thumb = im.resize((w, h), _RESAMPLE)
        x = pad + c * (cell_w + pad) + (cell_w - w) // 2
        y = pad + r * (cell_h + pad) + (cell_h - h) // 2
        sheet.paste(thumb, (x, y))
        bx, by = x + w - 26, y + h - 26
        draw.rectangle([bx, by, bx + 20, by + 20], fill=(255, 255, 255), outline=(120, 120, 120))
        if font:
            draw.text((bx + 6, by + 4), str(i + 1), fill=(40, 40, 40), font=font)
    buf = io.BytesIO()
    sheet.save(buf, format="PNG")
    return buf.getvalue()


# ---- 参考图元数据索引（masters/refs_index.json）----
def _refs_index_path(mas_dir):
    return os.path.join(mas_dir, REF_INDEX_NAME)


def _load_refs_index(mas_dir):
    try:
        with open(_refs_index_path(mas_dir), "r", encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict) and isinstance(d.get("items"), dict):
            return d
    except Exception:  # noqa: BLE001
        pass
    return {"schemaVersion": 1, "items": {}}


def _save_refs_index(mas_dir, data):
    try:
        with open(_refs_index_path(mas_dir), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:  # noqa: BLE001
        pass


# ---- 中转站多图能力探测（结果按 base_url 缓存）----
_CAPABILITY_CACHE = {}
_CAPABILITY_FILE = os.path.join(OUTPUT_DIR, ".edit_capability.json")


def _load_capability():
    global _CAPABILITY_CACHE
    if _CAPABILITY_CACHE:
        return _CAPABILITY_CACHE
    try:
        with open(_CAPABILITY_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict):
            _CAPABILITY_CACHE = d
    except Exception:  # noqa: BLE001
        _CAPABILITY_CACHE = {}
    return _CAPABILITY_CACHE


def _save_capability():
    try:
        with open(_CAPABILITY_FILE, "w", encoding="utf-8") as f:
            json.dump(_CAPABILITY_CACHE, f, ensure_ascii=False, indent=2)
    except Exception:  # noqa: BLE001
        pass


def _tiny_png(rgb, size=96):
    buf = io.BytesIO()
    Image.new("RGB", (size, size), rgb).save(buf, format="PNG")
    return buf.getvalue()


def _probe_edit_capability(key, base_url, model, skip_proxy=False):
    """真实发一次极小的 edits 请求，判断中转站支持哪种多图传参方式。

    命中顺序：image[] 重复 → image 重复 → 均不通则视为仅支持单图。
    注意：部分中转站会「静默丢弃」第二张图，探测无法覆盖这种情况。
    """
    headers = {"Authorization": f"Bearer {key}"}
    proxies = _proxy_cfg(skip_proxy)
    a, b = _tiny_png((220, 60, 60)), _tiny_png((60, 90, 220))
    reachable = False
    last_err = ""
    for kind, field in (("native_multi_array", "image[]"), ("native_multi_repeat", "image")):
        files = [(field, ("p1.png", a, "image/png")), (field, ("p2.png", b, "image/png"))]
        data = {"model": model, "prompt": ".", "size": "1024x1024", "quality": "low", "n": 1}
        for url in _candidate_urls_edits(base_url):
            try:
                r = requests.post(url, headers=headers, files=files, data=data,
                                  timeout=180, proxies=proxies)
            except Exception as e:  # noqa: BLE001
                last_err = str(e)
                continue
            if r.status_code in (401, 403):
                return {"ok": False, "error": f"鉴权失败（HTTP {r.status_code}）：{r.text[:200]}"}
            if r.status_code == 404:
                continue
            reachable = True
            if r.status_code < 400:
                return {"ok": True, "mode": kind, "max_images": HARD_REF_LIMIT,
                        "probed_at": int(time.time())}
            last_err = f"HTTP {r.status_code}：{r.text[:200]}"
    return {"ok": True, "mode": ("single_only" if reachable else "unknown"), "max_images": 1,
            "probed_at": int(time.time()), "note": last_err}


# 锁长相 / 锁风格指令：edits 端点把参考图作为强条件，约束后续出图
_LOCK_FACE = "Keep every character's face, hairstyle and clothing identical to the reference image."
_LOCK_STYLE = (
    "Strictly preserve the overall visual style of the reference image in this new image: "
    "keep the exact same color palette, background texture and tone, typography and font style, "
    "decorative motifs, layout skeleton, and overall aesthetic as the reference. "
    "If any character appears, also keep their face, hairstyle and clothing identical. "
    "Only change the content and text as described in the prompt."
)


def _post_edit_once(headers, proxies, base_url, files, data, skip_proxy=False):
    """发一次 edits 请求，返回 (kind, value)。

    kind: ok（value=图片字节）/ auth / proxy / network / notfound / err（value=错误文案）。
    网络类失败不再重试其它传参方案，避免把一次超时放大成多次超时。
    """
    last = "未知错误"
    for url in _candidate_urls_edits(base_url):
        try:
            r = requests.post(url, headers=headers, files=files, data=data,
                              timeout=300, proxies=proxies)
        except requests.exceptions.ProxyError as e:
            return "proxy", ("代理连接失败（ProxyError）：当前系统/网络设置了代理，但代理无法连通目标服务器。"
                             "请在「① API 设置」勾选「跳过系统代理（直连）」后重试。详情：" + str(e))
        except requests.exceptions.RequestException as e:
            return "network", f"网络请求失败：{e}"
        if r.status_code in (401, 403):
            return "auth", f"鉴权失败（HTTP {r.status_code}）：{r.text[:300]}。请检查 Key 或 API 地址"
        if r.status_code == 404:
            last = f"HTTP 404：{r.text[:200]}（edits 路径不存在，该中转站可能不支持参考图）"
            continue
        if r.status_code >= 400:
            # 400/422 通常意味着传参方式不被接受，交给外层换一种方案重试
            return "err", f"HTTP {r.status_code}：{r.text[:300]}"
        try:
            payload = r.json()
        except Exception:  # noqa: BLE001
            last = "接口返回非 JSON，无法解析图片"
            continue
        item = (payload.get("data") or [{}])[0]
        img_bytes = None
        if item.get("b64_json"):
            img_bytes = base64.b64decode(item["b64_json"])
        elif item.get("url"):
            try:
                img_bytes = _download_image(item["url"], headers, skip_proxy)
            except Exception as e:  # noqa: BLE001
                last = f"下载结果图失败：{e}"
                continue
        if not img_bytes:
            last = "接口返回里既没有 b64_json 也没有 url"
            continue
        return "ok", img_bytes
    return "err", last


def _gen_edit(key, base_url, model, prompt, size, index, ref_items, skip_proxy=False,
              lock_scope="style", images_dir=None, project_name=None, input_fidelity="auto"):
    """edits 端点：多参考图 + 角色绑定 + 自动降级。

    降级顺序（每级失败自动落到下一级，回传 mode/degraded 供前端提示）：
      L1 native_multi_array   —— 重复 image[] 字段（GPT Image 原生多图）
      L1 native_multi_repeat  —— 重复 image 字段（兼容只认单名字段的中转站）
      L2 composite_single     —— 多图拼成 contact sheet 当单图发
      L3 single_primary       —— 只发优先级最高的一张（风格 > 人物 > 版式 > 素材）
    input_fidelity: auto / high / low；auto 时仅当存在风格或人物参考图才用 high。
    """
    headers = {"Authorization": f"Bearer {key}"}
    proxies = _proxy_cfg(skip_proxy)

    resolved = []
    for it in (ref_items or []):
        path = _find_image((it or {}).get("filename") or "")
        if not path:
            continue
        try:
            w = int((it or {}).get("weight") or 3)
        except (TypeError, ValueError):
            w = 3
        role = ((it or {}).get("role") or "style").strip() or "style"
        if role not in REF_ROLE_LABEL:
            role = "style"
        resolved.append({"path": path, "role": role, "note": (it or {}).get("note") or "",
                         "weight": max(1, min(5, w))})
    if not resolved:
        return _err("参考图不存在，请重新上传或移除后再生成")

    resolved = resolved[:HARD_REF_LIMIT]
    # 编号稳定化：按「角色优先级 + 权重」排序后固定，禁用项已在入参侧剔除
    resolved.sort(key=lambda r: (REF_ROLE_ORDER.get(r["role"], 9), -r["weight"]))

    lock = _LOCK_FACE if lock_scope == "face" else _LOCK_STYLE
    if input_fidelity == "low":
        want_fidelity = False
    elif input_fidelity == "high":
        want_fidelity = True
    else:
        want_fidelity = any(r["role"] in ("character", "style") for r in resolved)

    def part_of(path):
        with open(path, "rb") as f:
            b = f.read()
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                "webp": "image/webp"}.get(ext, "image/png")
        return os.path.basename(path), mime, b

    attempts = []
    if len(resolved) > 1:
        multi = [part_of(r["path"]) for r in resolved]
        manifest_multi = build_ref_manifest(resolved, "multi")
        attempts.append(("native_multi_array", "image[]", multi, manifest_multi))
        attempts.append(("native_multi_repeat", "image", multi, manifest_multi))
        try:
            sheet = _compose_contact_sheet([r["path"] for r in resolved])
            attempts.append(("composite_single", "image",
                             [("reference_contact_sheet.png", "image/png", sheet)],
                             build_ref_manifest(resolved, "composite")))
        except Exception:  # noqa: BLE001
            pass
    attempts.append(("single_primary", "image", [part_of(resolved[0]["path"])],
                     build_ref_manifest(resolved[:1], "multi")))

    last_err = "未知错误"
    for mode_label, field, parts, manifest in attempts:
        files = [(field, (fn, b, mime)) for (fn, mime, b) in parts]
        for use_fidelity in ([True, False] if want_fidelity else [False]):
            data = {"model": model,
                    "prompt": manifest + "\n\n" + prompt + " " + lock,
                    "size": size}
            if use_fidelity:
                data["input_fidelity"] = "high"
            kind, value = _post_edit_once(headers, proxies, base_url, files, data, skip_proxy)
            if kind == "ok":
                img_dir, _ = _save_dirs(images_dir, project_name)
                fname = _save_image(value, index, target_dir=img_dir)
                used = len(resolved) if mode_label != "composite_single" else len(resolved)
                return _ok({
                    "filename": fname,
                    "url": _image_url(fname),
                    "mode": mode_label,
                    "ref_used": used,
                    "ref_total": len(resolved),
                    "degraded": mode_label not in ("native_multi_array", "native_multi_repeat"),
                    "input_fidelity": "high" if use_fidelity else "low",
                })
            if kind in ("auth", "proxy", "network"):
                return _err(value)
            last_err = value
            if "input_fidelity" in (value or ""):
                continue  # 换用不带 input_fidelity 的重试
            break         # 该方案整体不可用，落到下一个方案
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


@app.route("/api/image/<path:filename>")
def serve_image(filename):
    """按文件名跨目录返回图片（页面图/图表/母版/参考图，兼容自定义目录与默认目录）。"""
    path = _find_image(filename)
    if not path:
        return _err("图片不存在", 404)
    return send_from_directory(os.path.dirname(path), os.path.basename(path))


@app.route("/api/set_images_dir", methods=["POST"])
def set_images_dir():
    """轻量接口：仅登记自定义图片目录，供前端刷新页面后恢复预览定位。"""
    data = request.get_json(force=True, silent=True) or {}
    _remember_images_dir(data.get("images_dir"))
    return _ok({})


@app.route("/api/new_project", methods=["POST"])
def new_project():
    """新建一个项目子文件夹：组合 YYYYMMDD_<项目名>，重名追加 _2/_3。

    入参：images_dir（可空，空则用默认 output/images）、project_name（raw，可空）。
    出参：project_name（最终子文件夹名）、project_dir（绝对路径）。
    页面图 / 图表将落到该子文件夹；母版 / 参考图仍存 images_dir/masters。
    """
    data = request.get_json(force=True, silent=True) or {}
    images_dir = (data.get("images_dir") or "").strip() or None
    raw_name = (data.get("project_name") or "").strip()
    _remember_images_dir(images_dir)

    parent = os.path.abspath(images_dir) if images_dir else IMAGES_DIR
    os.makedirs(parent, exist_ok=True)
    # 同时确保 masters 子文件夹存在（母版/参考图统一存放位置，与项目无关）
    os.makedirs(os.path.join(parent, "masters"), exist_ok=True)
    base = _build_project_name(raw_name)            # YYYYMMDD_<safe>
    final = _make_project_name_unique(parent, base)  # 重名追加 _2/_3
    project_dir = os.path.join(parent, final)
    os.makedirs(project_dir, exist_ok=True)
    _remember_project(final)
    return _ok({"project_name": final, "project_dir": os.path.abspath(project_dir)})


# ---------------------------------------------------------------------------
# B 链：文档解析 / 大纲 / 文案
# ---------------------------------------------------------------------------
@app.route("/api/parse_doc", methods=["POST"])
def parse_doc():
    # 支持文件上传（docx/txt/md）或 JSON 里直接粘贴文本
    images_dir = None
    project_name = None
    file_basename = None
    if request.content_type and "multipart" in request.content_type:
        f = request.files.get("file")
        if not f:
            return _err("未收到文件")
        ext = os.path.splitext(f.filename or "")[1].lower()
        content = f.read()
        file_basename = os.path.splitext(os.path.basename(f.filename or ""))[0]
        images_dir = (request.form.get("images_dir") or "").strip() or None
        project_name = (request.form.get("project_name") or "").strip() or None
        text = _extract_text(content, ext)
    else:
        data = request.get_json(force=True, silent=True) or {}
        text = (data.get("text") or "").strip()
        if not text:
            return _err("请上传文件或直接粘贴文本")
        images_dir = (data.get("images_dir") or "").strip() or None
        project_name = (data.get("project_name") or "").strip() or None
    text = (text or "").strip()
    if not text:
        return _err("未能从文档中提取到文字，请检查文件格式")

    _remember_images_dir(images_dir)
    # 前端已传活跃项目就直接记忆；否则后端按文档名兜底新建一个并回传最终名
    if project_name:
        _remember_project(project_name)
    elif not CURRENT_PROJECT_NAME:
        raw = file_basename or time.strftime("%Y%m%d")
        parent = os.path.abspath(images_dir) if images_dir else IMAGES_DIR
        os.makedirs(parent, exist_ok=True)
        base = _build_project_name(raw)
        final = _make_project_name_unique(parent, base)
        os.makedirs(os.path.join(parent, final), exist_ok=True)
        _remember_project(final)
        project_name = final
    else:
        project_name = CURRENT_PROJECT_NAME

    return _ok({"text": text, "length": len(text), "project_name": project_name})


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
    images_dir = (data.get("images_dir") or "").strip() or None
    project_name = (data.get("project_name") or "").strip() or None

    if not key:
        return _err("请先填写 DeepSeek API Key")
    if not text:
        return _err("请先上传或粘贴文档内容")

    _remember_images_dir(images_dir)
    # 跳过 parse_doc 直接走 outline 时兜底建项目（用文本前若干字作候选名）
    if project_name:
        _remember_project(project_name)
    elif not CURRENT_PROJECT_NAME:
        parent = os.path.abspath(images_dir) if images_dir else IMAGES_DIR
        os.makedirs(parent, exist_ok=True)
        raw = (text[:20].replace("\n", " ").replace("\r", " ").strip()) or time.strftime("%Y%m%d")
        base = _build_project_name(raw)
        final = _make_project_name_unique(parent, base)
        os.makedirs(os.path.join(parent, final), exist_ok=True)
        _remember_project(final)
        project_name = final
    else:
        project_name = CURRENT_PROJECT_NAME

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
        return _ok({"outline": outline_json, "project_name": project_name})
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
# 参考图上传 / 图库 / 删除 / 能力探测
# ---------------------------------------------------------------------------
@app.route("/api/upload_ref", methods=["POST"])
def upload_ref():
    """上传 1~N 张参考图（风格母版 / 人物 / 版式 / 素材）。

    兼容两种字段：files（可重复，多张）与 file（单张，旧调用）。
    落盘到 <images_dir>/masters/，并把角色/权重元数据写入 refs_index.json。
    """
    files = [f for f in (request.files.getlist("files") or []) if f and f.filename]
    if not files:
        single = request.files.get("file")
        if single and single.filename:
            files = [single]
    if not files:
        return _err("未收到文件")

    images_dir = (request.form.get("images_dir") or "").strip() or None
    _remember_images_dir(images_dir)
    _, mas_dir = _save_dirs(images_dir)
    index = _load_refs_index(mas_dir)
    index.setdefault("items", {})

    items, skipped = [], []
    idx = 1
    for f in files:
        raw_name = os.path.basename(f.filename or "")
        ext = os.path.splitext(raw_name)[1].lower()
        if ext not in (".png", ".jpg", ".jpeg", ".webp"):
            skipped.append({"name": raw_name, "reason": "仅支持 png / jpg / webp"})
            continue
        data = f.read()
        if not data:
            skipped.append({"name": raw_name, "reason": "文件为空"})
            continue
        if len(data) > 10 * 1024 * 1024:
            skipped.append({"name": raw_name, "reason": "超过 10MB 限制"})
            continue
        digest = hashlib.sha1(data).hexdigest()
        dup = next((k for k, v in index["items"].items()
                    if isinstance(v, dict) and v.get("sha1") == digest
                    and os.path.isfile(os.path.join(mas_dir, k))), None)
        if dup:
            items.append({"filename": dup, "url": _image_url(dup), "size": len(data),
                          "sha1": digest, "duplicated": True})
            continue
        e = "jpg" if ext.lstrip(".") == "jpeg" else ext.lstrip(".")
        base_fname = f"ref_{int(time.time() * 1000)}_{idx}"
        fname = f"{base_fname}.{e}"
        n = 2
        while os.path.exists(os.path.join(mas_dir, fname)):
            fname = f"{base_fname}_{n}.{e}"
            n += 1
        with open(os.path.join(mas_dir, fname), "wb") as out:
            out.write(data)
        index["items"][fname] = {"role": "style", "weight": 5, "note": "",
                                 "sha1": digest, "addedAt": int(time.time() * 1000)}
        items.append({"filename": fname, "url": _image_url(fname), "size": len(data),
                      "sha1": digest, "duplicated": False})
        idx += 1
    _save_refs_index(mas_dir, index)

    if not items:
        return _err(skipped[0]["reason"] if skipped else "没有可用的图片")
    # ref_image / ref_url 为旧前端的兼容字段，指向本次第一张
    return _ok({"items": items, "skipped": skipped,
                "ref_image": items[0]["filename"], "ref_url": items[0]["url"]})


@app.route("/api/list_refs", methods=["GET"])
def list_refs():
    """列出参考图库：masters 目录磁盘扫描 ∪ refs_index.json 元数据。"""
    images_dir = (request.args.get("images_dir") or "").strip()
    _remember_images_dir(images_dir)
    _, mas_dir = _save_dirs(images_dir)
    index = _load_refs_index(mas_dir)
    meta = index.get("items") if isinstance(index.get("items"), dict) else {}
    items = []
    if os.path.isdir(mas_dir):
        for name in sorted(os.listdir(mas_dir)):
            if not name.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                continue
            full = os.path.join(mas_dir, name)
            if not os.path.isfile(full):
                continue
            m = meta.get(name) or {}
            try:
                st = os.stat(full)
                size, mtime = st.st_size, int(st.st_mtime * 1000)
            except OSError:
                size, mtime = 0, 0
            items.append({"filename": name, "url": _image_url(name), "size": size, "mtime": mtime,
                          "role": m.get("role") or "style", "weight": m.get("weight") or 5,
                          "note": m.get("note") or ""})
    items.sort(key=lambda x: x.get("mtime") or 0, reverse=True)
    return _ok({"items": items, "max_images": MAX_REF_IMAGES, "hard_limit": HARD_REF_LIMIT})


@app.route("/api/delete_ref", methods=["POST"])
def delete_ref():
    """删除 masters 目录内的一张参考图（含索引条目），不触碰页面图。"""
    data = request.get_json(force=True, silent=True) or {}
    name = _validate_filename(data.get("filename") or "")
    if not name:
        return _err("缺少文件名")
    images_dir = (data.get("images_dir") or "").strip() or None
    _remember_images_dir(images_dir)
    _, mas_dir = _save_dirs(images_dir)
    target = os.path.abspath(os.path.join(mas_dir, name))
    if os.path.dirname(target) != os.path.abspath(mas_dir):
        return _err("拒绝访问：只能删除 masters 目录内的参考图", 403)
    if not os.path.isfile(target):
        return _err("文件不存在", 404)
    try:
        os.remove(target)
    except OSError as e:
        return _err(f"删除失败：{e}")
    index = _load_refs_index(mas_dir)
    if isinstance(index.get("items"), dict) and name in index["items"]:
        index["items"].pop(name, None)
        _save_refs_index(mas_dir, index)
    return _ok({"filename": name})


@app.route("/api/probe_edit_capability", methods=["POST"])
def probe_edit_capability_route():
    """探测中转站的 edits 端点支持哪种多图传参方式（结果按 base_url 缓存）。"""
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("openai_key") or "").strip()
    base_url = (data.get("openai_base_url") or "").strip().rstrip("/")
    model = (data.get("openai_model") or data.get("model") or "gpt-image-2").strip()
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 OpenAI API Key")
    if not base_url:
        return _err("请先填写 OpenAI API 地址")

    cache = _load_capability()
    cached = cache.get(base_url)
    if cached and not data.get("force"):
        return _ok({"mode": cached.get("mode"), "max_images": cached.get("max_images"),
                    "probed_at": cached.get("probed_at"), "cached": True})

    res = _probe_edit_capability(key, base_url, model, skip_proxy)
    if not res.get("ok"):
        return _err(res.get("error") or "探测失败")
    cache[base_url] = {"mode": res.get("mode"), "max_images": res.get("max_images"),
                       "probed_at": res.get("probed_at")}
    _save_capability()
    res["cached"] = False
    return _ok(res)



# ---------------------------------------------------------------------------
# AI 生成风格母版（可选）：按风格意图生成 1~3 张候选样张
# ---------------------------------------------------------------------------
@app.route("/api/optimize_style_prompt", methods=["POST"])
def optimize_style_prompt():
    """仅调用 DeepSeek 把风格意图 → 英文母版生图 prompt，便于前端预览与编辑后再生图。

    与 /api/optimize（页面级）对应，但本接口使用 STYLE_MASTER_SYSTEM，
    输出的是不含具体文字、纯视觉风格的母版提示词。
    """
    data = request.get_json(force=True, silent=True) or {}
    key = (data.get("deepseek_key") or "").strip()
    base_url = (data.get("deepseek_base_url") or "https://api.deepseek.com").strip().rstrip("/")
    model = (data.get("deepseek_model") or "deepseek-chat").strip()
    style_desc = (data.get("style_desc") or "").strip()
    pages = data.get("pages") or []
    skip_proxy = data.get("skip_proxy") or False

    if not key:
        return _err("请先填写 DeepSeek API Key")
    if not style_desc and not pages:
        return _err("请填写风格描述，或先添加页面内容以便自动归纳风格")

    try:
        if style_desc:
            user_msg = f"风格意图：{style_desc}"
        else:
            joined = "\n".join([p for p in pages if p][:20])
            user_msg = ("以下是若干页课件的内容描述，请据此归纳适合的整体视觉风格"
                        "并输出母版提示词：\n\n" + joined)
        prompt = _call_deepseek(key, base_url, model, STYLE_MASTER_SYSTEM, user_msg,
                                temperature=0.7, max_tokens=900, skip_proxy=skip_proxy)
        prompt = (prompt or "").strip()
        if not prompt:
            return _err("DeepSeek 返回空提示词，请重试")
        return _ok({"prompt": prompt})
    except Exception as e:  # noqa: BLE001
        return _err(f"风格提示词优化失败：{e}", 500)


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
    images_dir = (data.get("images_dir") or "").strip() or None
    _remember_images_dir(images_dir)

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

    # 优先使用前端传来的「已优化提示词」（用户可见并可编辑）；否则现场由 DeepSeek 优化
    prompt = (data.get("prompt") or "").strip()
    if not prompt and ds_key:
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
                                 skip_proxy=skip_proxy, prefix="master", images_dir=images_dir)

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
    images_dir = (data.get("images_dir") or "").strip() or None
    project_name = (data.get("project_name") or "").strip() or None
    _remember_images_dir(images_dir)
    if project_name:
        _remember_project(project_name)

    if not chart_type:
        return _err("缺少图表类型")

    try:
        from chart_renderer import render as render_chart_img
        png_bytes = render_chart_img(chart_type, chart_data)
    except Exception as e:  # noqa: BLE001
        return _err(f"图表绘制失败：{e}", 500)

    img_dir, _ = _save_dirs(images_dir, project_name)
    fname = f"chart_{int(index):02d}_{int(time.time() * 1000)}.png"
    with open(os.path.join(img_dir, fname), "wb") as f:
        f.write(png_bytes)
    return _ok({"filename": fname, "url": _image_url(fname)})


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
    ref_images = data.get("ref_images") or []
    lock_scope = (data.get("lock_scope") or "style").strip()
    input_fidelity = (data.get("input_fidelity") or "auto").strip()
    skip_proxy = data.get("skip_proxy") or False
    images_dir = (data.get("images_dir") or "").strip() or None
    project_name = (data.get("project_name") or "").strip() or None
    _remember_images_dir(images_dir)
    if project_name:
        _remember_project(project_name)

    if not key:
        return _err("请先填写 OpenAI API Key")
    if not base_url:
        return _err("请先填写 OpenAI API 地址")
    if not prompt:
        return _err("提示词不能为空")

    # 组装有序参考图集：新字段 ref_images 优先；旧字段 ref_image 兜底包装成一张
    ref_items = []
    if isinstance(ref_images, list):
        for it in ref_images:
            if isinstance(it, str):
                if it.strip():
                    ref_items.append({"filename": it.strip(), "role": "style", "weight": 5})
            elif isinstance(it, dict) and (it.get("filename") or "").strip():
                ref_items.append({"filename": it["filename"].strip(),
                                  "role": (it.get("role") or "style").strip() or "style",
                                  "weight": it.get("weight") or 5,
                                  "note": it.get("note") or ""})
    if not ref_items and ref_image:
        ref_items.append({"filename": ref_image,
                          "role": "character" if lock_scope == "face" else "style",
                          "weight": 5})

    if ref_items:
        return _gen_edit(key, base_url, model, prompt, size, index, ref_items, skip_proxy,
                         lock_scope, images_dir, project_name, input_fidelity)

    return _gen_text2img(key, base_url, model, prompt, size, quality, fmt, index, skip_proxy, images_dir=images_dir, project_name=project_name)


# ---------------------------------------------------------------------------
# 拼装 PPTX
# ---------------------------------------------------------------------------
@app.route("/api/build_ppt", methods=["POST"])
def build_ppt():
    data = request.get_json(force=True, silent=True) or {}
    images = data.get("images") or []
    title = (data.get("title") or "演示文稿").strip() or "演示文稿"
    output_dir = (data.get("output_dir") or "").strip()
    images_dir = (data.get("images_dir") or "").strip() or None
    project_name = (data.get("project_name") or "").strip() or None
    _remember_images_dir(images_dir)
    if project_name:
        _remember_project(project_name)

    if not images:
        return _err("还没有生成任何图片，请先生成")

    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9 宽屏
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for fn in images:
        fpath = _find_image(fn)  # 跨目录定位图片（自定义目录 > 默认目录）
        if not fpath:
            return _err(f"找不到图片：{_validate_filename(fn)}")
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(fpath, 0, 0, width=prs.slide_width, height=prs.slide_height)

    deck_name = f"{title}_{int(time.time() * 1000)}.pptx"
    save_dir = output_dir if output_dir else OUTPUT_DIR
    try:
        os.makedirs(save_dir, exist_ok=True)
    except Exception as e:
        return _err(f"输出目录创建失败：{e}")
    deck_path = os.path.join(save_dir, deck_name)
    prs.save(deck_path)
    if os.path.abspath(save_dir) == os.path.abspath(OUTPUT_DIR):
        url = f"/output/{deck_name}"
    else:
        url = f"/api/download?path={quote(deck_path)}"
    return _ok({"filename": deck_name, "url": url, "abs_path": os.path.abspath(deck_path)})


# ---------------------------------------------------------------------------
# 打包全部图片为 ZIP
# ---------------------------------------------------------------------------
@app.route("/api/zip_images", methods=["POST"])
def zip_images():
    data = request.get_json(force=True, silent=True) or {}
    images = data.get("images") or []
    output_dir = (data.get("output_dir") or "").strip()
    images_dir = (data.get("images_dir") or "").strip() or None
    project_name = (data.get("project_name") or "").strip() or None
    _remember_images_dir(images_dir)
    if project_name:
        _remember_project(project_name)
    if not images:
        return _err("还没有生成任何图片")

    zname = f"images_{int(time.time() * 1000)}.zip"
    save_dir = output_dir if output_dir else OUTPUT_DIR
    try:
        os.makedirs(save_dir, exist_ok=True)
    except Exception as e:
        return _err(f"输出目录创建失败：{e}")
    zpath = os.path.join(save_dir, zname)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn in images:
            safe = _validate_filename(fn)
            fpath = _find_image(safe)  # 跨目录定位图片
            if fpath:
                zf.write(fpath, arcname=safe)
    if os.path.abspath(save_dir) == os.path.abspath(OUTPUT_DIR):
        url = f"/output/{zname}"
    else:
        url = f"/api/download?path={quote(zpath)}"
    return _ok({"filename": zname, "url": url, "abs_path": os.path.abspath(zpath)})


# ---------------------------------------------------------------------------
# 通用文件下载（用于自定义输出目录的文件）
# ---------------------------------------------------------------------------
@app.route("/api/download")
def download_file():
    path = request.args.get("path", "")
    if not path:
        return _err("缺少文件路径")
    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        return _err("文件不存在")
    dirname = os.path.dirname(abs_path)
    basename = os.path.basename(abs_path)
    return send_from_directory(dirname, basename, as_attachment=True)


if __name__ == "__main__":
    print("=" * 52)
    print("  PPT 智能生成工具已启动")
    print("  请在浏览器打开： http://127.0.0.1:5000")
    print("  按 Ctrl+C 停止服务")
    print("=" * 52)
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
