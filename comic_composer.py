# -*- coding: utf-8 -*-
"""
漫画页拼格与气泡绘制（纯 Pillow，无外部服务依赖）。

输入：每页若干格画面（PIL Image 或文件路径）+ 每格的对白 / 旁白
输出：竖版 A4（2480×3508，300dpi）漫画页 PNG；多页可合并导出 PDF。

设计原则（与 chart_renderer.py 一致）：
  精确内容（格子边框、对白文字、页码）一律由程序绘制，
  不让生图模型碰文字，保证中文对白零错字、可随时改。

独立自测：
  python comic_composer.py
  会在 output/ 下生成示例格图、两页示例漫画与 comic_demo.pdf。
"""
import os

from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# 画布与版式常量（竖版 A4 / 300dpi）
# ---------------------------------------------------------------------------
A4_SIZE = (2480, 3508)
MARGIN = 120                     # 页边距
GUTTER = 40                      # 格子间隙
BORDER = 6                       # 格子黑框宽度
FOOTER_H = 120                   # 页脚（页码）高度
INK = (15, 15, 15)               # 边框 / 文字墨色
PAPER = (255, 255, 255)          # 纸面
NARRATION_BG = (255, 250, 235)   # 旁白框底色
PLACEHOLDER_BG = (238, 238, 240)  # 待生成格子底色
BUBBLE_MAX_W = 0.62              # 气泡最大宽度占格宽比例
TAIL_LEN = 46                    # 气泡尾巴长度

_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS

# 版式 → 格数（1=单格全页、2v=上下 2 格、2x2=四格、3v=三格竖排、1+2=上宽下二、2+1=上二下宽）
LAYOUT_COUNTS = {"1": 1, "2v": 2, "2x2": 4, "3v": 3, "1+2": 3, "2+1": 3}


def layout_for_count(n):
    """按格数挑一个默认版式。"""
    return {1: "1", 2: "2v", 3: "1+2"}.get(n, "2x2")


# ---------------------------------------------------------------------------
# 字体（中文兜底，思路同 chart_renderer.py）
# ---------------------------------------------------------------------------
_FONT_FILES = [
    "msyh.ttc",                  # 微软雅黑
    "msyhbd.ttc",                # 微软雅黑 Bold
    "simhei.ttf",                # 黑体
    "simsun.ttc",                # 宋体
    "PingFangSC.ttf",
    "NotoSansCJKsc-Regular.otf",
    "wqy-microhei.ttc",
]


def _load_font(size):
    """按候选列表加载中文字体；全失败退回 PIL 默认字体（仅 ASCII 可读）。"""
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for name in _FONT_FILES:
        for path in (os.path.join(windir, "Fonts", name), name):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()


# ---------------------------------------------------------------------------
# 版式切分
# ---------------------------------------------------------------------------
def _layout_rects(layout, rect):
    """把内容区按版式切成格子矩形 [(x, y, w, h), ...]，阅读顺序从左到右、从上到下。"""
    x, y, w, h = rect
    layout = (layout or "").strip()
    if layout == "1":
        return [(x, y, w, h)]
    if layout == "2v":
        ph = (h - GUTTER) // 2
        return [(x, y, w, ph), (x, y + ph + GUTTER, w, h - ph - GUTTER)]
    if layout == "3v":
        ph = (h - 2 * GUTTER) // 3
        return [(x, y, w, ph), (x, y + ph + GUTTER, w, ph),
                (x, y + 2 * (ph + GUTTER), w, h - 2 * (ph + GUTTER))]
    if layout in ("1+2", "2+1"):
        ph = (h - GUTTER) // 2
        pw = (w - GUTTER) // 2
        wide_top = (x, y, w, ph)
        pair_top = [(x, y, pw, ph), (x + pw + GUTTER, y, w - pw - GUTTER, ph)]
        wide_bottom = (x, y + ph + GUTTER, w, h - ph - GUTTER)
        pair_bottom = [(x, y + ph + GUTTER, pw, h - ph - GUTTER),
                       (x + pw + GUTTER, y + ph + GUTTER, w - pw - GUTTER, h - ph - GUTTER)]
        return [wide_top] + pair_bottom if layout == "1+2" else pair_top + [wide_bottom]
    # 默认 2x2
    pw = (w - GUTTER) // 2
    ph = (h - GUTTER) // 2
    return [
        (x, y, pw, ph),
        (x + pw + GUTTER, y, w - pw - GUTTER, ph),
        (x, y + ph + GUTTER, pw, h - ph - GUTTER),
        (x + pw + GUTTER, y + ph + GUTTER, w - pw - GUTTER, h - ph - GUTTER),
    ]


# ---------------------------------------------------------------------------
# 绘制原语
# ---------------------------------------------------------------------------
def _cover(im, w, h):
    """等比放大铺满目标区域后居中裁剪（cover 模式），返回 RGB 图。"""
    if im.mode != "RGB":
        im = im.convert("RGB")
    iw, ih = im.size
    s = max(w / iw, h / ih)
    nw, nh = max(w, int(iw * s + 0.5)), max(h, int(ih * s + 0.5))
    im = im.resize((nw, nh), _RESAMPLE)
    ox = (nw - w) // 2
    oy = (nh - h) // 2
    return im.crop((ox, oy, ox + w, oy + h))


def _wrap_text(draw, text, font, max_width):
    """按像素宽度对中英混排文本换行。"""
    lines = []
    for raw in str(text).split("\n"):
        line = ""
        for ch in raw:
            if draw.textlength(line + ch, font=font) <= max_width:
                line += ch
            elif line:
                lines.append(line)
                line = ch
            else:                       # 单字即超宽（极端窄气泡），硬塞
                lines.append(ch)
                line = ""
        lines.append(line)
    return lines or [""]


def _draw_narration(draw, rect, text, font, inset=28):
    """旁白框：格内左上角直角浅底框。返回占用的纵向高度（含下间距）。"""
    x, y, w, _ = rect
    pad = 20
    lines = _wrap_text(draw, text, font, int(w * 0.7) - pad * 2)
    line_h = font.size + 10
    tw = max((draw.textlength(l, font=font) for l in lines), default=0)
    bw, bh = int(tw) + pad * 2, line_h * len(lines) + pad * 2
    x0, y0 = x + inset, y + inset
    draw.rectangle([x0, y0, x0 + bw, y0 + bh], fill=NARRATION_BG, outline=INK, width=4)
    for i, l in enumerate(lines):
        draw.text((x0 + pad, y0 + pad + i * line_h), l, font=font, fill=INK)
    return bh + 16


def _draw_bubble(draw, rect, text, font, anchor="tl", offset=28):
    """对白气泡：圆角白底框 + 指向格内的小尾巴。

    anchor: tl/tr（格内上方，尾巴朝下）或 bl/br（格内下方，尾巴朝上）；
    offset: 距格边（顶或底）的距离，同侧多气泡由调用方累加实现堆叠。
    返回占用的纵向高度（气泡 + 尾巴 + 间距）。
    """
    x, y, w, h = rect
    pad = 24
    lines = _wrap_text(draw, text, font, int(w * BUBBLE_MAX_W) - pad * 2)
    line_h = font.size + 14
    tw = max((draw.textlength(l, font=font) for l in lines), default=0)
    bw, bh = int(tw) + pad * 2, line_h * len(lines) + pad * 2
    inset_x = 30
    x0 = x + inset_x if anchor.endswith("l") else x + w - bw - inset_x
    top_side = anchor.startswith("t")
    y0 = y + offset if top_side else y + h - bh - offset

    radius = min(28, bh // 2)
    draw.rounded_rectangle([x0, y0, x0 + bw, y0 + bh], radius=radius,
                           fill=PAPER, outline=INK, width=5)

    # 尾巴：根部两个挂点 + 偏向格中心的尖端
    cx = x + w / 2
    base_cx = x0 + bw * (0.35 if cx < x0 + bw / 2 else 0.65)
    base_cx = max(x0 + radius + 20, min(base_cx, x0 + bw - radius - 20))
    half = 24
    tip_x = base_cx + (18 if cx > base_cx else -18)
    if top_side:
        y_base, tip_y = y0 + bh - 3, y0 + bh + TAIL_LEN
    else:
        y_base, tip_y = y0 + 3, y0 - TAIL_LEN
    p_left, p_right, p_tip = (base_cx - half, y_base), (base_cx + half, y_base), (tip_x, tip_y)
    draw.polygon([p_left, p_right, p_tip], fill=PAPER)
    draw.line([p_left, p_right], fill=PAPER, width=8)   # 抹掉根部气泡边框，让尾巴连通
    draw.line([p_left, p_tip], fill=INK, width=5)
    draw.line([p_right, p_tip], fill=INK, width=5)

    for i, l in enumerate(lines):
        lw = draw.textlength(l, font=font)
        draw.text((x0 + (bw - lw) / 2, y0 + pad + i * line_h), l, font=font, fill=INK)
    return bh + TAIL_LEN + 16


def _draw_footer(draw, size, page_no):
    """页码：- N -，居中画在页脚区。"""
    W, H = size
    f = _load_font(44)
    t = f"- {page_no} -"
    tw = draw.textlength(t, font=f)
    draw.text(((W - tw) / 2, H - MARGIN - FOOTER_H + 36), t, font=f, fill=(110, 110, 115))


# ---------------------------------------------------------------------------
# 页面合成
# ---------------------------------------------------------------------------
def compose_page(panels, layout="2x2", page_no=None, size=A4_SIZE):
    """把一页的若干格拼成一张竖版 A4 漫画页，返回 PIL Image。

    panels: [{"image": PIL.Image | None, "narration": str,
              "dialogue": [{"speaker": str, "text": str, "pos": "tl/tr/bl/br"}]}]
    image 为 None 的格画「待生成」占位，便于部分生成时也能预览拼页。
    """
    W, H = size
    canvas = Image.new("RGB", size, PAPER)
    draw = ImageDraw.Draw(canvas)
    content = (MARGIN, MARGIN, W - 2 * MARGIN, H - 2 * MARGIN - FOOTER_H)
    rects = _layout_rects(layout, content)
    dfont = _load_font(56)
    nfont = _load_font(50)

    for i, rect in enumerate(rects):
        panel = panels[i] if i < len(panels) else {}
        x, y, w, h = (int(v) for v in rect)
        im = panel.get("image")
        if im is not None:
            canvas.paste(_cover(im, w, h), (x, y))
        else:
            draw.rectangle([x, y, x + w, y + h], fill=PLACEHOLDER_BG)
            ph_font = _load_font(72)
            t = f"第 {i + 1} 格 · 待生成"
            tw = draw.textlength(t, font=ph_font)
            draw.text((x + (w - tw) / 2, y + h / 2 - 40), t, font=ph_font, fill=(150, 150, 155))
        draw.rectangle([x, y, x + w, y + h], outline=INK, width=BORDER)

        inset_top = 28
        narration = (panel.get("narration") or "").strip()
        if narration:
            inset_top += _draw_narration(draw, (x, y, w, h), narration, nfont)
        cursors = {"tl": inset_top, "tr": inset_top, "bl": 28, "br": 28}
        for j, d in enumerate(panel.get("dialogue") or []):
            speaker = (d.get("speaker") or "").strip()
            text = (d.get("text") or "").strip()
            if not text:
                continue
            label = f"{speaker}：{text}" if speaker and speaker != "旁白" else text
            anchor = (d.get("pos") or "").strip().lower()
            if anchor not in ("tl", "tr", "bl", "br"):
                anchor = "tl" if j % 2 == 0 else "tr"
            cursors[anchor] += _draw_bubble(draw, (x, y, w, h), label, dfont,
                                            anchor=anchor, offset=cursors[anchor])

    if page_no is not None:
        _draw_footer(draw, size, page_no)
    return canvas


def frame_fullpage(im, page_no=None, size=A4_SIZE):
    """整页模式：把一张竖版整页图贴进 A4 画布（cover 铺满内容区 + 黑框 + 页码）。"""
    W, H = size
    canvas = Image.new("RGB", size, PAPER)
    draw = ImageDraw.Draw(canvas)
    x, y, w, h = MARGIN, MARGIN, W - 2 * MARGIN, H - 2 * MARGIN - FOOTER_H
    canvas.paste(_cover(im, w, h), (x, y))
    draw.rectangle([x, y, x + w, y + h], outline=INK, width=BORDER)
    if page_no is not None:
        _draw_footer(draw, size, page_no)
    return canvas


def save_pdf(images, path):
    """多页漫画合并导出 PDF（Pillow 原生支持，无新依赖）。"""
    ims = [im if im.mode == "RGB" else im.convert("RGB") for im in images]
    if not ims:
        raise ValueError("没有可导出的页面")
    ims[0].save(path, "PDF", save_all=True, append_images=ims[1:], resolution=300.0)


# ---------------------------------------------------------------------------
# 自测：生成两页示例漫画（占位色块代替真实格图，验证拼格 / 气泡 / 页码 / PDF）
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    base = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base, "output")
    img_dir = os.path.join(out_dir, "images")
    os.makedirs(img_dir, exist_ok=True)

    def placeholder(w, h, color, label):
        im = Image.new("RGB", (w, h), color)
        d = ImageDraw.Draw(im)
        f = _load_font(90)
        tw = d.textlength(label, font=f)
        d.text(((w - tw) / 2, h / 2 - 50), label, font=f, fill=(255, 255, 255))
        return im

    demo_panels = [
        placeholder(1024, 1024, (91, 141, 239), "教室"),
        placeholder(1024, 1024, (42, 157, 143), "电脑屏幕"),
        placeholder(1024, 1024, (244, 162, 97), "路由器"),
        placeholder(1024, 1024, (231, 111, 81), "网络拓扑"),
    ]
    demo_files = []
    for i, im in enumerate(demo_panels, start=1):
        fn = os.path.join(img_dir, f"demo_panel_{i}.png")
        im.save(fn)
        demo_files.append(fn)

    page1 = compose_page([
        {"image": demo_panels[0], "narration": "周五下午的信息技术课",
         "dialogue": [{"speaker": "小明", "text": "老师，为什么网页打不开了？", "pos": "tr"}]},
        {"image": demo_panels[1],
         "dialogue": [{"speaker": "老师", "text": "先看看是网络问题还是电脑问题。", "pos": "tl"},
                      {"speaker": "小明", "text": "咦，右下角有个黄色感叹号！", "pos": "br"}]},
        {"image": demo_panels[2],
         "dialogue": [{"speaker": "老师", "text": "检查路由器指示灯，红色说明外网断了。", "pos": "tl"}]},
        {"image": demo_panels[3], "narration": "原来网络就像一条路，一环断了都到不了终点",
         "dialogue": [{"speaker": "小明", "text": "我懂了，要一段一段排查！", "pos": "bl"}]},
    ], layout="2x2", page_no=1)

    page2 = compose_page([
        {"image": demo_panels[3], "narration": "第二天，小明当起了网络小医生",
         "dialogue": [{"speaker": "小明", "text": "先查电脑，再查路由器，最后打电话问运营商。", "pos": "tr"}]},
        {"image": demo_panels[1],
         "dialogue": [{"speaker": "同桌", "text": "我家 Wi-Fi 很慢怎么办？", "pos": "tl"}]},
        {"image": None,
         "dialogue": [{"speaker": "小明", "text": "离路由器近一点，或者换个人少的信道！", "pos": "bl"}]},
    ], layout="1+2", page_no=2)

    p1_path = os.path.join(out_dir, "comic_demo_p1.png")
    p2_path = os.path.join(out_dir, "comic_demo_p2.png")
    pdf_path = os.path.join(out_dir, "comic_demo.pdf")
    page1.save(p1_path)
    page2.save(p2_path)
    save_pdf([page1, page2], pdf_path)

    print("示例格图：", *demo_files, sep="\n  ")
    print("示例漫画页：\n  {}\n  {}".format(p1_path, p2_path))
    print("示例 PDF：\n  {}".format(pdf_path))
