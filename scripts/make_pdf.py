#!/usr/bin/env python3
# 把对比报告做成 PDF（iPhone/微信直接可看）。每页 = 一对「生成|人工」截图 + 下方差异标注。
# 用法: make_pdf.py <report目录(含 gen/ hum/ 的png)> <gen.pptx> <hum.pptx> <out.pdf> [标题]
import re, glob, os, sys
from PIL import Image, ImageDraw, ImageFont

REPORT, GEN, HUM, OUT = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
TITLE = sys.argv[5] if len(sys.argv) > 5 else "对比报告"

# 中文字体（标注用）
def font(sz):
    for p in ["/System/Library/Fonts/PingFang.ttc", os.path.expanduser("~/Library/Fonts/汉仪楷体简.ttf")]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, sz)
            except: pass
    return ImageFont.load_default()
F_H = font(34); F_N = font(26); F_CAP = font(21)

import zipfile
def slides_xml(pptx):
    z = zipfile.ZipFile(pptx)
    files = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)], key=lambda n: int(re.search(r'slide(\d+)', n).group(1)))
    return z, files
def shp(x, i):
    m = re.search(r'<p:sp>(?:(?!</p:sp>)[\s\S])*?<p:cNvPr id="%s"(?:(?!</p:sp>)[\s\S])*?</p:sp>' % i, x)
    return m.group(0) if m else ''
def txt(s): return ''.join(re.findall(r'<a:t>([\s\S]*?)</a:t>', s)).strip()
_cache = {}
def _read(z, p):
    key = (z.filename, p)  # 生成版/人工版 slide 同名，必须按 zip 区分，否则缓存串味
    if key not in _cache:
        try: _cache[key] = z.read(p).decode('utf8')
        except: _cache[key] = ''
    return _cache[key]
def inh_sz(z, sp):
    phidx = re.search(r'<p:ph[^>]*idx="(\d+)"', sp); phtype = re.search(r'<p:ph[^>]*type="(\w+)"', sp)
    for lay in [n for n in z.namelist() if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)]:
        for lsp in re.findall(r'<p:sp>[\s\S]*?</p:sp>', _read(z, lay)):
            lidx = re.search(r'<p:ph[^>]*idx="(\d+)"', lsp); ltype = re.search(r'<p:ph[^>]*type="(\w+)"', lsp)
            if (phidx and lidx and lidx.group(1) == phidx.group(1)) or (not phidx and phtype and ltype and ltype.group(1) == phtype.group(1)):
                d = re.search(r'<a:defRPr[^>]*sz="(\d+)"', lsp)
                if d: return int(d.group(1))
    return None
def box(z, sp):
    if not sp or not txt(sp): return None
    szs = sorted(set(int(s) for s in re.findall(r'sz="(\d+)"', sp)))
    fs = re.search(r'fontScale="(\d+)"', sp)
    algn = sorted(set(re.findall(r'algn="(\w+)"', sp)))
    bold = any('b="1"' in p and txt(p) for p in re.findall(r'<a:p>[\s\S]*?</a:p>', sp))
    eff = szs if szs else ([inh_sz(z, sp)] if inh_sz(z, sp) else [])
    return {"sz": eff, "scale": round(int(fs.group(1)) / 1000) if fs else 100, "algn": algn, "bold": bold}
SECS = ["进堂咏","致候礼","忏悔礼","光荣颂","集祷经","读经一","答唱咏","读经二","福音前欢呼","福音","信经","信友祷文","奉献咏","献礼经","感恩经","常年期主日颂谢词","圣、圣、圣","Memorial","颂谢词","Doxology","天主经","平安礼","羔羊颂","领圣体礼","领主咏","领圣体后经","堂区报告","降福","圣弥额尔","求圣母","礼成咏"]
def meta(pptx):
    z, files = slides_xml(pptx); out = []
    for f in files:
        x = _read(z, f); n = int(re.search(r'slide(\d+)', f).group(1))
        t = txt(shp(x, '2')) or txt(shp(x, '10'))
        sec = next((s for s in ["福音前欢呼"] + SECS if s in t), "其他")
        if sec == "福音" and "福音前欢呼" in t: sec = "福音前欢呼"
        out.append({"n": n, "sec": sec, "title": t[:18], "cn": box(z, shp(x, '3')), "en": box(z, shp(x, '4'))})
    return out
def runs(ms):
    r = []
    for m in ms:
        if r and r[-1][0] == m["sec"]: r[-1][1].append(m)
        else: r.append([m["sec"], [m]])
    return r

def ptstr(b):
    if not b or not b['sz']: return '—'
    base = '/'.join(str(s // 100) for s in b['sz'])
    return f"{base}pt" + (f"·缩{b['scale']}%" if b['scale'] != 100 else "")
def alstr(b):
    m = {'just': '两端', 'l': '左', 'ctr': '居中', 'r': '右'}
    return '/'.join(m.get(a, a) for a in b['algn']) if b['algn'] else '默认'
def notelines(g, h):
    if g is None: return ["· 仅人工版有此页"]
    if h is None: return ["· 仅生成版有此页"]
    out = []
    for lab, k in (('中文', 'cn'), ('英文', 'en')):
        gb, hb = g.get(k), h.get(k)
        if not gb and not hb: continue
        if bool(gb) != bool(hb): out.append(f"· {lab}：一侧无内容框"); continue
        seg = [f"字号 生成{ptstr(gb)} {'＝' if ptstr(gb)==ptstr(hb) else 'vs'} 人工{ptstr(hb)}"]
        if alstr(gb) != alstr(hb): seg.append(f"对齐 生成{alstr(gb)} vs 人工{alstr(hb)}")
        if gb['bold'] != hb['bold']: seg.append(f"加粗 生成{'有' if gb['bold'] else '无'} vs 人工{'有' if hb['bold'] else '无'}")
        out.append(f"· {lab}：" + "；".join(seg))
    return out or ["· 版式一致"]

gr = runs(meta(GEN)); hr = runs(meta(HUM))
W, IMG_W = 1200, 585
pages = []
def load(sub, n):
    p = os.path.join(REPORT, sub, f"{n}.png")
    if not os.path.exists(p): return None
    im = Image.open(p).convert('RGB')
    return im.resize((IMG_W, round(im.height * IMG_W / im.width)))

# 封面
cov = Image.new('RGB', (W, 500), 'white'); d = ImageDraw.Draw(cov)
d.text((40, 200), TITLE, fill=(40, 70, 60), font=font(46))
d.text((40, 280), "左=自动生成　　右=人工制作", fill=(90, 90, 90), font=F_N)
pages.append(cov)

for (gs, gp), (hs, hp) in zip(gr, hr):
    sec = gs if gs != '其他' else (gp[0]['title'] or '过渡页')
    for i in range(max(len(gp), len(hp))):
        g = gp[i] if i < len(gp) else None
        h = hp[i] if i < len(hp) else None
        gi = load('gen', g['n']) if g else None
        hi = load('hum', h['n']) if h else None
        ih = max((im.height for im in (gi, hi) if im), default=360)
        notes = notelines(g, h)
        page_h = 60 + ih + 24 + len(notes) * 30 + 30
        pg = Image.new('RGB', (W, page_h), 'white'); d = ImageDraw.Draw(pg)
        d.text((20, 16), sec, fill=(45, 70, 60), font=F_H)
        d.line((20, 54, W - 20, 54), fill=(85, 114, 93), width=2)
        y = 66
        for im, x, cap in ((gi, 14, f"生成 P{g['n']}" if g else "生成(无)"), (hi, 14 + IMG_W + 8, f"人工 P{h['n']}" if h else "人工(无)")):
            if im:
                pg.paste(im, (x, y)); d.rectangle((x, y, x + IMG_W, y + im.height), outline=(200, 200, 200))
                d.text((x + 4, y + im.height + 3), cap, fill=(120, 120, 120), font=F_CAP)
            else:
                d.rectangle((x, y, x + IMG_W, y + ih), fill=(240, 240, 240)); d.text((x + IMG_W // 2 - 30, y + ih // 2), "无对应页", fill=(170, 170, 170), font=F_CAP)
        ty = y + ih + 26
        for ln in notes:
            col = (200, 0, 0) if ('vs' in ln or '仅' in ln) else (60, 60, 60)
            d.text((22, ty), ln, fill=col, font=F_N); ty += 30
        pages.append(pg)

pages[0].save(OUT, save_all=True, append_images=pages[1:], resolution=110)
print(f"{OUT}  {os.path.getsize(OUT)/1024/1024:.1f} MB  {len(pages)} 页")
