#!/usr/bin/env python3
# 用法: compare_report.py <workdir(含 gen_x/hum_x/gen_png/hum_png)> <输出目录> [标题]
# 手机友好的左右对比报告：响应式(窄屏上下堆叠)、无总览表、无渲染说明、解析继承字号。
import re, glob, html, os, shutil, sys

W, OUT = sys.argv[1], sys.argv[2]
TITLE = sys.argv[3] if len(sys.argv) > 3 else "生成 vs 人工 对比"

def slides(d):
    return sorted(glob.glob(f'{d}/ppt/slides/slide*.xml'), key=lambda f: int(re.search(r'slide(\d+)', f).group(1)))
def shp(x, i):
    m = re.search(r'<p:sp>(?:(?!</p:sp>)[\s\S])*?<p:cNvPr id="%s"(?:(?!</p:sp>)[\s\S])*?</p:sp>' % i, x)
    return m.group(0) if m else ''
def txt(s):
    return ''.join(re.findall(r'<a:t>([\s\S]*?)</a:t>', s)).strip()

# ── 继承字号解析：无显式 sz 时查 ph idx/type 在 layout / master 的 defRPr sz ──
_cache = {}
def _read(path):
    if path not in _cache:
        try: _cache[path] = open(path, encoding='utf8').read()
        except: _cache[path] = ''
    return _cache[path]
def inherited_sz(root, slidefile, sp):
    phidx = re.search(r'<p:ph[^>]*idx="(\d+)"', sp)
    phtype = re.search(r'<p:ph[^>]*type="(\w+)"', sp)
    rels = _read(slidefile.replace('slides/', 'slides/_rels/') + '.rels')
    lay = re.search(r'slideLayout\d+\.xml', rels)
    if lay:
        lx = _read(f'{root}/ppt/slideLayouts/' + lay.group(0))
        for lsp in re.findall(r'<p:sp>[\s\S]*?</p:sp>', lx):
            lidx = re.search(r'<p:ph[^>]*idx="(\d+)"', lsp); ltype = re.search(r'<p:ph[^>]*type="(\w+)"', lsp)
            if (phidx and lidx and lidx.group(1) == phidx.group(1)) or (not phidx and phtype and ltype and ltype.group(1) == phtype.group(1)):
                d = re.search(r'<a:defRPr[^>]*sz="(\d+)"', lsp)
                if d: return int(d.group(1))
    for mf in glob.glob(f'{root}/ppt/slideMasters/slideMaster*.xml'):
        d = re.search(r'<p:bodyStyle>[\s\S]*?<a:lvl1pPr[^>]*>[\s\S]*?<a:defRPr[^>]*sz="(\d+)"', _read(mf))
        if d: return int(d.group(1))
    return None

def boxinfo(root, slidefile, sp):
    if not sp or not txt(sp): return None
    szs = sorted(set(int(s) for s in re.findall(r'sz="(\d+)"', sp)))
    fs = re.search(r'fontScale="(\d+)"', sp)
    algn = sorted(set(re.findall(r'algn="(\w+)"', sp)))
    bold = [txt(p)[:12] for p in re.findall(r'<a:p>[\s\S]*?</a:p>', sp) if txt(p) and 'b="1"' in p]
    if szs:
        eff = szs
    else:
        inh = inherited_sz(root, slidefile, sp)
        eff = [inh] if inh else []
    return {"sz": eff, "scale": round(int(fs.group(1)) / 1000) if fs else 100, "algn": algn, "bold": bold}

SECS = ["进堂咏","致候礼","忏悔礼","光荣颂","集祷经","读经一","答唱咏","读经二","福音前欢呼","福音","信经","信友祷文","奉献咏","献礼经","感恩经","常年期主日颂谢词","圣、圣、圣","Memorial","颂谢词","Doxology","天主经","平安礼","羔羊颂","领圣体礼","领主咏","领圣体后经","堂区报告","降福","圣弥额尔","求圣母","礼成咏"]
def classify(t):
    for s in ["福音前欢呼"] + SECS:
        if s in t: return s
    return "其他"
def meta(root):
    out = []
    for f in slides(root):
        x = open(f, encoding='utf8').read()
        n = int(re.search(r'slide(\d+)', f).group(1))
        t = ''
        for i in ('2', '10'):
            tt = txt(shp(x, i))
            if tt: t = tt; break
        sec = classify(t)
        if sec == "福音" and "福音前欢呼" in t: sec = "福音前欢呼"
        out.append({"n": n, "sec": sec, "title": t[:20], "cn": boxinfo(root, f, shp(x, '3')), "en": boxinfo(root, f, shp(x, '4'))})
    return out
def runs(ms):
    r = []
    for m in ms:
        if r and r[-1][0] == m["sec"]: r[-1][1].append(m)
        else: r.append([m["sec"], [m]])
    return r

gen = meta(W + '/gen_x'); hum = meta(W + '/hum_x')
gr = runs(gen); hr = runs(hum)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT + '/gen'); os.makedirs(OUT + '/hum')
for src, dst in ((W + '/gen_png', OUT + '/gen'), (W + '/hum_png', OUT + '/hum')):
    for f in glob.glob(src + '/s-*.png'):
        n = int(re.search(r's-0*(\d+)\.png', f).group(1))
        shutil.copy(f, f'{dst}/{n}.png')

def pt(b):
    if not b or not b['sz']: return '—'
    base = '/'.join(str(s // 100) for s in b['sz'])
    return f"{base}pt" + (f"·缩{b['scale']}%" if b['scale'] != 100 else "")
def al(b):
    m = {'just': '两端', 'l': '左', 'ctr': '居中', 'r': '右'}
    return '/'.join(m.get(a, a) for a in b['algn']) if b['algn'] else '默认'
def notes(g, h):
    if g is None: return '<span class=add>仅生成版有此页</span>'
    if h is None: return '<span class=miss>仅人工版有此页</span>'
    rows = []
    for lang, key in (('中', 'cn'), ('英', 'en')):
        gb, hb = g.get(key), h.get(key)
        if not gb and not hb: continue
        if bool(gb) != bool(hb):
            rows.append(f"<b>{lang}文</b> 一侧无内容框"); continue
        seg = []
        gp, hp = pt(gb), pt(hb)
        seg.append(f"字号 生成{gp} <span class='{'ok' if gp==hp else 'ne'}'>{'一致' if gp==hp else 'vs'}</span> 人工{hp}")
        ga, ha = al(gb), al(hb)
        if ga != ha: seg.append(f"对齐 生成{ga} <span class=ne>vs</span> 人工{ha}")
        if (gb['bold'] == []) != (hb['bold'] == []):
            seg.append(f"加粗 生成{'有' if gb['bold'] else '无'} <span class=ne>vs</span> 人工{'有' if hb['bold'] else '无'}")
        rows.append(f"<b>{lang}文</b> " + "；".join(seg))
    return "<br>".join(rows) or "<span class=ok>版式一致</span>"

blocks, counts = [], []
for (gs, gp), (hs, hp) in zip(gr, hr):
    sec = gs if gs != '其他' else (gp[0]['title'] or '过渡页')
    counts.append(sec)
    tag = f" <em>生成{len(gp)}页 / 人工{len(hp)}页</em>" if len(gp) != len(hp) else ""
    blocks.append(f"<h2 id='s{len(counts)}'>{html.escape(sec)}{tag}</h2>")
    for i in range(max(len(gp), len(hp))):
        g = gp[i] if i < len(gp) else None
        h = hp[i] if i < len(hp) else None
        gi = f"<figure><img src='gen/{g['n']}.png' loading=lazy><figcaption>生成 P{g['n']}</figcaption></figure>" if g else "<figure class=empty><div>—</div><figcaption>生成(无)</figcaption></figure>"
        hi = f"<figure><img src='hum/{h['n']}.png' loading=lazy><figcaption>人工 P{h['n']}</figcaption></figure>" if h else "<figure class=empty><div>—</div><figcaption>人工(无)</figcaption></figure>"
        blocks.append(f"<div class=pair><div class=imgs>{gi}{hi}</div><div class=notes>{notes(g,h)}</div></div>")

toc = "".join(f"<a href='#s{i+1}'>{html.escape(s)}</a>" for i, s in enumerate(counts))
open(OUT + '/index.html', 'w', encoding='utf8').write(f"""<!DOCTYPE html><html lang=zh><head>
<meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>{html.escape(TITLE)}</title>
<style>
:root{{--g:#55725D}}*{{box-sizing:border-box}}
body{{font-family:-apple-system,'PingFang SC',sans-serif;margin:0;background:#f4f3ef;color:#222;font-size:15px}}
header{{position:sticky;top:0;background:var(--g);color:#fff;padding:11px 14px;z-index:9}}
header h1{{font-size:16px;margin:0}}
.toc{{display:flex;flex-wrap:wrap;gap:6px;padding:9px 12px;background:#fff;border-bottom:1px solid #ddd}}
.toc a{{font-size:12px;color:var(--g);text-decoration:none;background:#eef2ef;padding:3px 9px;border-radius:12px;white-space:nowrap}}
main{{padding:10px}}
h2{{font-size:15px;margin:22px 4px 8px;color:var(--g);border-bottom:2px solid var(--g);padding-bottom:3px}}
h2 em{{font-size:12px;color:#b40;font-style:normal}}
.pair{{background:#fff;border:1px solid #e0e0e0;border-radius:10px;padding:8px;margin:10px 0}}
.imgs{{display:flex;gap:8px}}
figure{{flex:1;margin:0;text-align:center;min-width:0}}
img{{width:100%;border:1px solid #ccc;border-radius:4px;display:block}}
figcaption{{font-size:11px;color:#777;margin-top:3px}}
.empty div{{display:flex;align-items:center;justify-content:center;aspect-ratio:4/3;background:#f0f0f0;color:#bbb;border-radius:4px}}
.notes{{margin-top:7px;font-size:13px;line-height:1.7;background:#faf9f6;border-left:3px solid var(--g);padding:6px 10px;border-radius:0 6px 6px 0}}
.ok{{color:#2a7;font-weight:bold}}.ne{{color:#c00;font-weight:bold;padding:0 2px}}.miss{{color:#c00}}.add{{color:#06c}}
@media(max-width:640px){{.imgs{{flex-direction:column}}}}
</style></head><body>
<header><h1>{html.escape(TITLE)}</h1></header>
<div class=toc>{toc}</div>
<main>{''.join(blocks)}</main>
</body></html>""")
print("OK", OUT)
