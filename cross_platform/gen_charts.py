#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""为各高校 WPS/data.json 生成缺失的图表 PNG（数据全部取自 data.json 自身）。

做了三件事：
  1) 数据可还原的图表（柱/饼）→ 用 matplotlib 重绘，中文用 PingFang SC；
  2) 校园照片（images/*.jpg，原包未提供真实照片）→ 生成主题色占位图；
  3) 无数据支撑的图表（如招生计划分布）→ 生成「示意图」占位图。

输出直接写入各高校目录（与 data.json 中 file 字段一致，含 images/ 子目录），
随后 build_office.py 运行时会自动找到这些图片。
"""
import os
import re
import json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["PingFang SC", "Arial Unicode MS",
                                   "Heiti SC", "Songti SC", "Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

HERE = os.path.dirname(os.path.abspath(__file__))
WPS = os.path.normpath(os.path.join(HERE, "..", "WPS"))


def num(s):
    if s is None:
        return None
    s = str(s).strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def get_tables(slide):
    return [e for e in slide.get("elements", []) if e.get("type") == "table"]


def find_image(spec, filename):
    for sd in spec.get("slides", []):
        for el in sd.get("elements", []):
            f = el.get("file") or el.get("src") or el.get("path") or el.get("image")
            if f and os.path.basename(str(f)) == os.path.basename(filename):
                return el
    return None


def series_from_tables(spec, slide_idx, table_idxs, label_col, value_col,
                       topn=None, exclude=None):
    exclude = set(exclude or [])
    pairs = []
    slide = spec["slides"][slide_idx]
    tables = get_tables(slide)
    for ti in table_idxs:
        t = tables[ti]
        rows = t.get("data")
        if not rows:
            continue
        for r in rows[1:]:
            if not isinstance(r, list):
                continue
            lab = r[label_col] if label_col < len(r) else ""
            if lab in exclude:
                continue
            v = num(r[value_col]) if value_col < len(r) else None
            if v is None:
                continue
            pairs.append((str(lab), v))
    pairs.sort(key=lambda x: -x[1])
    if topn:
        pairs = pairs[:topn]
    return pairs


def make_bar(out, pairs, title, color, el_w, el_h, topn):
    labels = [p[0] for p in pairs]
    vals = [p[1] for p in pairs]
    n = max(len(pairs), 1)
    h = max(el_h / 100.0, n * 0.34)
    w = max(el_w / 100.0, 4.0)
    fig, ax = plt.subplots(figsize=(w, h))
    y = range(n)
    bars = ax.barh(y, vals, color=color, height=0.7)
    ax.set_yticks(list(y))
    ax.set_yticklabels(labels, fontsize=12)
    ax.invert_yaxis()
    for i, v in enumerate(vals):
        ax.text(v, i, f"  {int(v)}", va="center", fontsize=10, color="#333333")
    ax.set_title(title, fontsize=15, fontweight="bold", color=color, pad=10)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(left=False)
    ax.set_xlim(0, max(vals) * 1.12)
    plt.tight_layout()
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)


def make_pie(out, labels, values, title, color, el_w, el_h):
    fig, ax = plt.subplots(figsize=(max(el_w / 100.0, 3.4), max(el_h / 100.0, 3.2)))
    cmap = [color, "#2E5090", "#D2691E", "#3A6B2C", "#8B1A6B", "#B8860B", "#007F6E"]
    colors = [cmap[i % len(cmap)] for i in range(len(values))]
    wedges, _ = ax.pie(values, colors=colors, startangle=90,
                       wedgeprops=dict(width=0.42, edgecolor="white"))
    ax.legend(wedges, [f"{l}  {int(v)}" for l, v in zip(labels, values)],
              loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=11, frameon=False)
    ax.set_title(title, fontsize=15, fontweight="bold", color=color)
    plt.tight_layout()
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)


def make_placeholder(out, label, color, el_w, el_h, ext):
    w = max(el_w / 100.0, 3.0)
    h = max(el_h / 100.0, 2.2)
    fig, ax = plt.subplots(figsize=(w, h))
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.03, 0.03), 0.94, 0.94,
                                boxstyle="round,pad=0.01,rounding_size=0.04",
                                linewidth=1.5, edgecolor=color, facecolor="#F4F6F8"))
    ax.text(0.5, 0.58, label, ha="center", va="center", fontsize=15,
            fontweight="bold", color="#444444")
    ax.text(0.5, 0.40, "示意图 · 原图未随包提供", ha="center", va="center",
            fontsize=11, color="#9AA0A6")
    plt.tight_layout()
    fig.savefig(out, dpi=150, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)


# 每个高校的图表配方：file -> ("pie"|"bar", {...})
CONFIG = {
    "北大": dict(
        json="pku_data.json", color="#8B0012",
        charts={
            "chart_channels.png": ("pie", dict(
                labels=["普通类统招", "强基计划", "国家专项", "竞赛保送", "医学部"],
                values=[1411, 900, 250, 200, 956], title="招生渠道分布")),
            "chart_scores.png": ("bar", dict(
                slide=3, tables=[0], label_col=0, value_col=1, topn=8,
                title="各省物理类投档线")),
            "chart_admissions.png": ("bar", dict(
                slide=4, tables=[0], label_col=1, value_col=2, topn=10,
                title="各省录取人数 TOP10")),
        },
        photos=[],
    ),
    "华中科大": dict(
        json="hust_data.json", color="#004098",
        charts={
            "images/chart_types.png": ("pie", dict(
                labels=["湖北", "其他省份"], values=[1768, 7305 - 1768],
                title="招生地域分布")),
            "images/chart_provinces.png": ("bar", dict(
                slide=3, tables=[0], label_col=1, value_col=2, topn=8,
                title="各省录取人数 TOP12")),
            "images/chart_scores.png": ("bar", dict(
                slide=4, tables=[0], label_col=0, value_col=1, topn=10,
                title="六省物理类投档线")),
            "images/chart_majors.png": ("bar", dict(
                slide=5, tables=[0], label_col=0, value_col=1, topn=8,
                title="热门专业分数线（山东）")),
        },
        photos=[
            ("images/campus_1.jpg", "校园风光 ①"),
            ("images/campus_2.jpg", "校园风光 ②"),
            ("images/campus_3.jpg", "校园风光 ③"),
            ("images/dormitory.jpg", "学生宿舍"),
            ("images/lab.jpg", "实验室"),
        ],
    ),
    "南科大": dict(
        json="sustech_data.json", color="#006B3F",
        charts={
            "chart_mode.png": ("pie", dict(
                labels=["高考成绩 60%", "校测 30%", "学业 10%"],
                values=[60, 30, 10], title="「631」综合评价模式")),
            "chart_top12.png": ("bar", dict(
                slide=4, tables=[0], label_col=1, value_col=2, topn=7,
                title="各省录取人数 TOP12")),
            "chart_avgscores.png": ("bar", dict(
                slide=3, tables=[0, 1], label_col=0, value_col=2, topn=12,
                exclude=["合计"], title="各省高考平均分")),
            "chart_gd_compare.png": ("placeholder", dict(label="广东三校投档线对比")),
        },
        photos=[],
    ),
    "浙大城市学院": dict(
        json="zucc_data.json", color="#005A9C",
        charts={
            "chart_plan.png": ("placeholder", dict(label="招生计划分布")),
            "chart_majors.png": ("bar", dict(
                slide=4, tables=[0], label_col=0, value_col=1, topn=9,
                title="浙江省热门专业分数线")),
            "chart_scores_prov.png": ("bar", dict(
                slide=3, tables=[0, 1], label_col=0, value_col=1, topn=14,
                exclude=["分数区间", "物理类", "历史类", "招生范围", "16省招生", "34专业"],
                title="各省物理类录取分数线")),
        },
        photos=[],
    ),
}


def generate(school, cfg):
    sdir = os.path.join(WPS, school)
    json_path = os.path.join(sdir, cfg["json"])
    with open(json_path, "r", encoding="utf-8") as f:
        spec = json.load(f)
    color = cfg["color"]
    made = []

    for filename, (kind, kw) in cfg["charts"].items():
        out_path = os.path.join(sdir, filename)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        if os.path.exists(out_path):
            continue
        el = find_image(spec, filename) or {}
        el_w = el.get("w", 460)
        el_h = el.get("h", 330)
        if kind == "pie":
            make_pie(out_path, kw["labels"], kw["values"], kw["title"], color, el_w, el_h)
        elif kind == "bar":
            pairs = series_from_tables(spec, kw["slide"], kw["tables"],
                                       kw["label_col"], kw["value_col"],
                                       kw.get("topn"), kw.get("exclude"))
            if not pairs:
                make_placeholder(out_path, kw.get("title", filename), color, el_w, el_h,
                                 os.path.splitext(filename)[1])
            else:
                make_bar(out_path, pairs, kw["title"], color, el_w, el_h, kw.get("topn", len(pairs)))
        else:  # placeholder
            make_placeholder(out_path, kw["label"], color, el_w, el_h,
                             os.path.splitext(filename)[1])
        made.append(filename)

    for filename, label in cfg.get("photos", []):
        out_path = os.path.join(sdir, filename)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        if os.path.exists(out_path):
            continue
        el = find_image(spec, filename) or {}
        make_placeholder(out_path, label, color, el.get("w", 295), el.get("h", 200),
                         os.path.splitext(filename)[1])
        made.append(filename)

    return made


def main():
    total = 0
    for school, cfg in CONFIG.items():
        made = generate(school, cfg)
        print(f"[{school}] 生成 {len(made)} 个图片: {made}")
        total += len(made)
    print(f"\n全部完成，共生成 {total} 个图片文件。")


if __name__ == "__main__":
    main()
