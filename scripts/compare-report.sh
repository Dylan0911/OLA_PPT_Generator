#!/bin/bash
# 用法: ./scripts/compare-report.sh <生成版.pptx> <人工版.pptx> <输出目录名>
# 依赖: LibreOffice (soffice) + poppler (pdftoppm)。按段落对齐逐页左右对比,标注字号/加粗/对齐差异。
set -e
SCRIPTDIR="$(cd "$(dirname "$0")" && pwd)"
GEN="$1"; HUM="$2"; OUT="$HOME/Downloads/$3"
W=$(mktemp -d)
cd "$W"
mkdir gen_png hum_png gen_x hum_x
# 结构解析用【原始】pptx（保真）；渲染截图用【粗体注入副本】（让 LibreOffice 显示楷体加粗）。
(cd gen_x && unzip -q "$GEN") & (cd hum_x && unzip -q "$HUM") & wait
python3 "$SCRIPTDIR/prep_render_copy.py" "$GEN" "$W/gen.pptx"
python3 "$SCRIPTDIR/prep_render_copy.py" "$HUM" "$W/hum.pptx"
soffice --headless --convert-to pdf gen.pptx --outdir . >/dev/null 2>&1
soffice --headless --convert-to pdf hum.pptx --outdir . >/dev/null 2>&1
pdftoppm -png -r 100 gen.pdf gen_png/s & pdftoppm -png -r 100 hum.pdf hum_png/s & wait
export OUT
python3 "$SCRIPTDIR/compare_report.py" "$W" "$OUT"
