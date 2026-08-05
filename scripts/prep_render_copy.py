#!/usr/bin/env python3
# 为报告渲染准备 PPTX 副本：LibreOffice 在 Mac 上对「楷体」不合成粗体（无粗字重），
# 导致加粗看不出。这里【只给 b="1" 的 rPr 注入有粗体字重的字体(苹方 PingFang SC)】，
# 正文楷体不动 → 渲染截图里加粗词明显变粗，肉眼可辨。不影响交付的 PPT。
# 用法: prep_render_copy.py <in.pptx> <out.pptx>
import sys, os, re, zipfile, shutil, tempfile

IN, OUT = sys.argv[1], sys.argv[2]
# 用黑体渲染加粗词（楷体本身没粗字重、加粗几乎看不出；黑体笔画重，报告里一眼可辨）。
INJ = '<a:latin typeface="Heiti SC"/><a:ea typeface="Heiti SC"/>'

def fix_rpr(tag):
    if 'b="1"' not in tag:
        return tag
    if tag.endswith("/>"):  # 自闭合 rPr → 展开后注入字体
        return tag[:-2] + ">" + INJ + "</a:rPr>"
    # 带子元素 rPr：去掉已有 latin/ea，在结尾注入
    t = re.sub(r"<a:latin\b[^>]*/>", "", tag)
    t = re.sub(r"<a:ea\b[^>]*/>", "", t)
    return t.replace("</a:rPr>", INJ + "</a:rPr>")

def fix(xml):
    # 只匹配 <a:rPr …>（不含 endParaRPr）；两种形态：自闭合 / 带子元素。
    return re.sub(r"<a:rPr\b[^>]*/>|<a:rPr\b[^>]*>[\s\S]*?</a:rPr>", lambda m: fix_rpr(m.group(0)), xml)

tmp = tempfile.mkdtemp()
with zipfile.ZipFile(IN) as z:
    z.extractall(tmp)
sd = os.path.join(tmp, "ppt", "slides")
for f in os.listdir(sd):
    if re.match(r"slide\d+\.xml$", f):
        p = os.path.join(sd, f)
        with open(p, encoding="utf8") as fh:
            data = fh.read()
        with open(p, "w", encoding="utf8") as fh:
            fh.write(fix(data))

if os.path.exists(OUT):
    os.remove(OUT)
base = os.getcwd()
os.chdir(tmp)
os.system(f'zip -qr "{OUT}" .')
os.chdir(base)
shutil.rmtree(tmp)
print("prepared", OUT)
