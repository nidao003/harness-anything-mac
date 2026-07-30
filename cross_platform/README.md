# macOS 跨平台 Office 生成引擎

本目录提供一套**不依赖任何 Windows COM / WPS 自动化**的 Office 文档生成方案，
让 `harness-anything-mac` 在 macOS（以及 Linux / Windows）上也能直接产出
真实的 `pptx / docx / xlsx`，并一键导出 PDF。

> 背景：上游的 `cli-anything/office` 在 Mac 上的 `export render` 仍硬编码 WPS COM，
> 无法在 macOS 运行。本引擎用 `python-pptx / python-docx / openpyxl` 直接生成文档模型，
> 再用本机 LibreOffice (`soffice`) 导出 PDF，与平台无关、无需安装 WPS。

## 能力

| 格式 | 命令 | 说明 |
|------|------|------|
| PPTX + PDF | `python build_office.py ppt --json spec.json --out deck.pptx --pdf` | 960×540 逻辑画布，元素类型齐全 |
| DOCX + PDF | `python build_office.py docx --json spec.json --out report.docx --pdf` | 标题/段落/列表/表格/图片/分页 |
| XLSX + PDF | `python build_office.py xlsx --json spec.json --out data.xlsx --pdf` | 多 Sheet、表头高亮、冻结首行 |

已在 macOS 实测：三件套 + PDF 导出全部通过，中文渲染正常（macOS 自动把
`SimHei/微软雅黑` 映射为 `PingFang SC`，`SimSun/宋体` 映射为 `Songti SC`）。

## 依赖

```bash
pip install python-pptx python-docx openpyxl matplotlib
# PDF 导出需要 LibreOffice
brew install --cask libreoffice      # macOS
```

## 快速开始

```bash
VENV="$HOME/.workbuddy/binaries/python/envs/default/bin/python"   # 或任意含依赖的 python
ENGINE="cross_platform/build_office.py"

# PPTX + PDF
"$VENV" "$ENGINE" ppt  --json spec.json --out deck.pptx  [--pdf] [--bg template.png] [--bg-color #006B3F]

# DOCX + PDF
"$VENV" "$ENGINE" docx --json spec.json --out report.docx [--pdf]

# XLSX + PDF
"$VENV" "$ENGINE" xlsx --json spec.json --out data.xlsx   [--pdf]
```

- `--bg template.png`：每页铺满的背景图（如各校 `WPS/<校>/template_bg.png`）。
- `--bg-color #RRGGBB`：纯色背景兜底。
- 图片 `file` / `src` 路径相对于**当前工作目录**解析，生成含图片的 PPT 时建议 `cd` 到数据所在目录再运行。

## PPT JSON Schema（960×540 逻辑画布）

```json
{
  "theme": {"primary": "#006B3F"},
  "slides": [
    {
      "title": "页面标题（自动置顶）",
      "elements": [
        {"type":"text",        "x":60,"y":150,"w":840,"h":120,"text":"标题","fs":44,"bold":true,"color":"#006B3F","align":"center"},
        {"type":"bullets",     "x":60,"y":300,"w":840,"h":200,"items":["要点一","要点二"],"ordered":false},
        {"type":"table",       "x":60,"y":140,"w":840,"h":300,"header":true,"header_color":"#006B3F","rows":[["A","B"],["v1","v2"]]},
        {"type":"image",       "x":..,"y":..,"w":..,"h":..,"src":"/path/to/img.png"},
        {"type":"tagline_bar", "x":160,"y":400,"w":640,"h":50,"text":"标语","color":"#006B3F","fs":18},
        {"type":"cards_2x3",   "x":60,"y":130,"w":840,"h":360,"items":[{"title":"卡标题","body":"内容"}]},
        {"type":"cards_1x4_info","x":..,"y":..,"w":..,"h":..,"items":[{"num":"4490","label":"计划"}]},
        {"type":"card_list_wide","x":..,"y":..,"w":..,"h":..,"items":[{"num":"01","title":"标题","sub":"副标题"}]},
        {"type":"cards_1x3_big","x":..,"y":..,"w":..,"h":..,"items":[{"title":"标题","desc":"说明"}]},
        {"type":"cards_2x2_four","x":..,"y":..,"w":..,"h":..,"items":[{"title":"标题","desc":"说明"}]},
        {"type":"card_row_5",  "x":..,"y":..,"w":..,"h":..,"items":[{"icon":"强基","title":"标题","desc":"说明"}]},
        {"type":"line",        "x":..,"y":..,"w":..,"h":..,"color":"#006B3F"}
      ]
    }
  ]
}
```

`table` 兼容两种 schema：旧版 `rows` 为二维数据数组；新版 `rows`/`cols` 为 int、
数据放在 `data` 字段（含 `th_fs`/`td_fs`/`header_color`）。`align` 支持数字
`1/2/3` 或 `"left"/"center"/"right"`。

## WPS 五校招生模板 + 图表生成器

`WPS/` 下含 5 所高校招生数据模板（清华 / 北大 / 华中科大 / 南科大 / 浙大城市学院），
每校有 Windows 版 `build_*.py`（Mac 不可用）+ `*_data.json`（纯数据，跨平台可用）
+ `template_bg.png` 背景。

模板引用的图表 / 校园照片原始资源未随包分发，因此提供 `gen_charts.py` 从
`*_data.json` **自身数据**重绘：

```bash
"$VENV" cross_platform/gen_charts.py
```

它会为每校生成：可还原的图表（录取人数柱状图、分数线对比、招生渠道 / 631 模式饼图等）
以 matplotlib 绘制（macOS 上中文用 PingFang SC），校园照片与无数据图表则生成主题色
「示意图」占位框。生成的图片写入 `WPS/<校>/`（含 `images/` 子目录），与 `data.json`
中的 `file` 字段一一对应。

随后逐校生成完整 PPTX + PDF（需 `cd` 到各校目录，图片相对路径才能解析）：

```bash
WPS="$PWD/WPS"
for s in 清华 北大 华中科大 南科大 浙大城市学院; do
  ( cd "$WPS/$s" && "$VENV" "$ENGINE" ppt --json ${s}_data.json \
      --out /tmp/$s.pptx --bg template_bg.png --pdf )
done
```

成品 PPTX + PDF 见各校目录或 `WPS/_demos/`（若已生成）。

## 已知限制

- 本引擎与上游 `cli-anything/office` 是**两套并存的方案**：Mac / 跨平台用本引擎；
  上游的 WPS COM 自动化仍仅限 Windows。
- 校园实景照片为真实图片，无法凭空生成，统一用「示意图」占位框处理；放入真实照片
  （按 `data.json` 的 `file` 路径）后引擎会自动采用。
- `gen_charts.py` 的图表数据取自 `data.json` 中已结构化的表格 / 文字，若模板后续调整
  字段，需同步更新 `gen_charts.py` 中的配方（`CONFIG`）。
