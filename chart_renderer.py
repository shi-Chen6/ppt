# -*- coding: utf-8 -*-
"""
用 matplotlib 精确绘制图表（雷达图 / 柱状图 / 饼图），输出高清 PNG。

chart_data 结构（dict）：
  title  : 图表标题（可选）
  labels : 分类名列表（list[str] 或逗号分隔字符串）
  values : 数值列表（list[float]）
"""
import io
import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# 中文字体兜底
_CJK = ["Microsoft YaHei", "SimHei", "SimSun", "PingFang SC", "Noto Sans CJK SC", "WenQuanYi Micro Hei"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.sans-serif"] = _CJK + list(plt.rcParams.get("font.sans-serif", []))


def _get(data, *keys):
    if isinstance(data, dict):
        for k in keys:
            if k in data:
                return data[k]
    return None


def _labels(data):
    v = _get(data, "labels")
    if isinstance(v, list):
        return [str(x) for x in v]
    if isinstance(v, str):
        return [x.strip() for x in v.replace("，", ",").split(",") if x.strip()]
    return []


def _values(data):
    v = _get(data, "values")
    if isinstance(v, list):
        out = []
        for x in v:
            try:
                out.append(float(x))
            except (TypeError, ValueError):
                out.append(0.0)
        return out
    return []


def _title(data):
    t = _get(data, "title")
    return str(t) if t else ""


def _save(fig, dpi):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render(chart_type, data, dpi=200):
    chart_type = (chart_type or "").strip().lower()
    if chart_type == "radar":
        return _radar(data, dpi)
    if chart_type == "bar":
        return _bar(data, dpi)
    if chart_type == "pie":
        return _pie(data, dpi)
    raise ValueError(f"不支持的图表类型：{chart_type}")


def _radar(data, dpi):
    labels = _labels(data)
    values = _values(data)
    if not labels or not values:
        raise ValueError("雷达图需要 labels 和 values")
    if len(labels) != len(values):
        values = (values + [0.0] * len(labels))[:len(labels)]

    angles = [2 * math.pi * i / len(labels) for i in range(len(labels))]
    vals = values + values[:1]
    angs = angles + angles[:1]

    fig = plt.figure(figsize=(8, 8))
    ax = fig.add_subplot(111, polar=True)
    ax.plot(angs, vals, "o-", linewidth=2, color="#378ADD")
    ax.fill(angs, vals, alpha=0.25, color="#378ADD")
    ax.set_xticks(angles)
    ax.set_xticklabels(labels)
    if _title(data):
        ax.set_title(_title(data), pad=22)
    return _save(fig, dpi)


def _bar(data, dpi):
    labels = _labels(data)
    values = _values(data)
    if not labels or not values:
        raise ValueError("柱状图需要 labels 和 values")
    if len(labels) != len(values):
        values = (values + [0.0] * len(labels))[:len(labels)]

    fig = plt.figure(figsize=(8, 5))
    ax = fig.add_subplot(111)
    ax.bar(labels, values, color="#378ADD", width=0.6)
    if _title(data):
        ax.set_title(_title(data))
    if len(labels) > 6:
        plt.xticks(rotation=45, ha="right")
    return _save(fig, dpi)


def _pie(data, dpi):
    labels = _labels(data)
    values = _values(data)
    if not labels or not values:
        raise ValueError("饼图需要 labels 和 values")
    if len(labels) != len(values):
        values = (values + [0.0] * len(labels))[:len(labels)]

    fig = plt.figure(figsize=(8, 6))
    ax = fig.add_subplot(111)
    ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
    ax.axis("equal")
    if _title(data):
        ax.set_title(_title(data))
    return _save(fig, dpi)
