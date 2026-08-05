#!/usr/bin/env python3
# 把对比报告文件夹打成【单个自包含 HTML】：图片压小 + base64 内嵌，一个文件即可手机转发查看。
# 用法: embed_report.py <报告文件夹> <输出单文件.html>
import re, os, sys, base64, io
from PIL import Image

SRC, OUT = sys.argv[1], sys.argv[2]
html = open(os.path.join(SRC, 'index.html'), encoding='utf8').read()
cache = {}

def data_uri(rel):
    p = os.path.join(SRC, rel)
    if rel in cache: return cache[rel]
    if not os.path.exists(p):
        cache[rel] = rel; return rel
    im = Image.open(p).convert('RGB')
    w = 460
    if im.width > w: im = im.resize((w, round(im.height * w / im.width)))
    buf = io.BytesIO(); im.save(buf, 'JPEG', quality=58, optimize=True)
    uri = 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode()
    cache[rel] = uri
    return uri

html = re.sub(r"src='((?:gen|hum)/\d+\.png)'", lambda m: "src='" + data_uri(m.group(1)) + "'", html)
open(OUT, 'w', encoding='utf8').write(html)
print(f"{OUT}  {os.path.getsize(OUT)/1024/1024:.1f} MB  ({len(cache)} imgs)")
