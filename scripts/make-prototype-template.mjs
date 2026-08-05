// 把绿色常年期成品 deck 改造成【原型占位符模板】，供程序自动填充：
//   · 每个可变段（集祷经/读经一/读经二/答唱咏/福音前欢呼/福音/信友祷文/献礼经/领圣体后经）
//     折叠成【一张原型页】，放命名占位符 {{读经一_CN}} / {{读经一_EN}}；该段其余页删除。
//     （填充程序按当周内容克隆这张原型页、自动分页。）
//   · 封面 → {{主日标题_CN}} / {{主日标题_EN}}
//   · 歌咏(进堂咏/奉献咏/领主咏/礼成咏) + 堂区报告 → 红字「请手动填入」
//   · 固定经文（致候礼/忏悔礼/光荣颂/信经/颂谢词/感恩经/天主经/羔羊颂…）原文保留，
//     只修排版：英文两端对齐→左对齐、文字溢出→烤入 fontScale。
// 形状约定：id2=标题、id3=中文、id4=英文；封面标题在 id4。
//
// 用法: node make-prototype-template.mjs <source.pptx> <out.pptx>
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

const sourcePptx = process.argv[2];
const outPptx = process.argv[3];
if (!sourcePptx || !outPptx) {
  console.error("用法: node make-prototype-template.mjs <source.pptx> <out.pptx>");
  process.exit(1);
}

// 可变段：按此顺序识别（福音前欢呼 必须排在 福音 之前，否则 "福音前欢呼" 标题会被 "福音" 抢先命中）。
const VAR_ORDER = ["集祷经", "读经一", "读经二", "答唱咏", "福音前欢呼", "福音", "信友祷文", "献礼经", "领圣体后经"];
const RED_SECTIONS = ["进堂咏", "奉献咏", "领主咏", "礼成咏", "堂区报告"];
const RED_LABEL = {
  进堂咏: "进堂咏 · 请手动填入本周歌词",
  奉献咏: "奉献咏 · 请手动填入本周歌词",
  领主咏: "领主咏 / 领圣体歌 · 请手动填入",
  礼成咏: "礼成咏 · 请手动填入本周歌词",
  堂区报告: "本周堂区报告 · 请在此手动填入",
};
const POSTURE = /^\s*(请?(站立|坐下|跪下|起立)|全体起立|Stand|Sit|Kneel|Please stand|Please sit|Please kneel)\s*$/;
const CHINESE_FONT = "楷体";

const escapeXml = (v) => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const hasChinese = (t) => /[㐀-鿿]/.test(t);
const hasLatin = (t) => /[A-Za-z]/.test(t);

function shapeBlock(xml, id) {
  return xml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${id}"(?:(?!</p:sp>)[\\s\\S])*?</p:sp>`))?.[0] || null;
}
const shapeText = (sp) => [...(sp || "").matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join("").trim();
// 按段落取形状文字（<a:p> 之间、<a:br/> 处 = 换行）—— 估行数用；否则多段文字被当一整行、行数被低估。
function shapeParaText(sp) {
  return [...(sp || "").matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
    .map((p) =>
      [...p[0].replace(/<a:br\b[^>]*\/>/g, "<a:t>\n</a:t>").matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join("")
    )
    .join("\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}
function deleteShape(xml, id) {
  const sp = shapeBlock(xml, id);
  return sp ? xml.replace(sp, "") : xml;
}
const deleteAllPics = (xml) => xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, "");
function textShapeIds(xml) {
  const ids = [];
  for (const sp of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const id = sp[0].match(/<p:cNvPr id="(\d+)"/)?.[1];
    if (id && /<a:t>/.test(sp[0])) ids.push(id);
  }
  return ids;
}

// 命名占位符（统一格式：中文楷体 5400、英文 3200、左对齐、normAutofit）。
// 全卷统一字号：中文 4800 / 英文 2800（= 对话页家族现有规格；填充程序按自然容量分页、不缩字，
// 固定页也由填充程序归一化到同一标准 → 全卷字号恒定）。
const CANON_RPR = {
  cn: '<a:rPr lang="zh-CN" altLang="en-US" sz="4800" dirty="0"/>',
  en: '<a:rPr lang="en-HK" altLang="zh-CN" sz="2800" dirty="0"/>',
};
const CANON_ALGN = { cn: "just", en: "l" };
function putToken(xml, shapeId, token, lang) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp || !/<p:txBody>/.test(sp)) return xml;
  const newTx = `<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="${CANON_ALGN[lang]}"><a:buNone/></a:pPr><a:r>${CANON_RPR[lang]}<a:t>${escapeXml(token)}</a:t></a:r></a:p></p:txBody>`;
  return xml.replace(sp, sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx));
}
// 红字「请手动填入」——直接用目标内容的排版（粘贴内容后只需把颜色改黑，无需调字体字号）：
//   hymn（歌词）= 人工版歌词标准：楷体 6600 粗体、居中、顶端对齐；
//   announce（堂区报告）= 楷体 3200 粗体、左对齐、顶端对齐。
const RED_STYLE = {
  hymn: { sz: 6600, algn: "ctr", anchor: "t" },
  announce: { sz: 3200, algn: "l", anchor: "t" },
};
function putRedNotice(xml, shapeId, label, kind = "hymn") {
  const sp = shapeBlock(xml, shapeId);
  if (!sp) return xml;
  const st = RED_STYLE[kind] || RED_STYLE.hymn;
  const bodyPr = `<a:bodyPr rtlCol="0" anchor="${st.anchor}"><a:normAutofit/></a:bodyPr>`;
  const newTx = `<p:txBody>${bodyPr}<a:lstStyle/><a:p><a:pPr algn="${st.algn}"/><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="${st.sz}" b="1" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="${CHINESE_FONT}"/><a:ea typeface="${CHINESE_FONT}"/></a:rPr><a:t>${escapeXml(label)}</a:t></a:r></a:p></p:txBody>`;
  return xml.replace(sp, sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx));
}
// 封面标题：整体重建 txBody —— 中文 4800 / 英文 4400 粗体白字（人工版标准），
// 英文写死 EB Garamond 与正文一致；标题框加高到人工版的 2143314 EMU
// （原生成框矮 14%，中文一行+英文折两行放不下 → 顶部被裁掉）。
const COVER_BOX_CY = 2143314;
function coverTitle(xml) {
  const sp = shapeBlock(xml, "4");
  if (!sp) return xml;
  let newSp = sp.replace(/(<a:ext cx="\d+" cy=")(\d+)("\/>)/, (m, a, cy, b) => `${a}${Math.max(Number(cy), COVER_BOX_CY)}${b}`);
  const fillW = '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>';
  const newTx =
    `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPts val="5200"/></a:lnSpc></a:pPr>` +
    `<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="4800" b="1" dirty="0">${fillW}<a:ea typeface="${CHINESE_FONT}"/></a:rPr><a:t>{{主日标题_CN}}</a:t></a:r>` +
    `<a:br><a:rPr lang="en-HK" altLang="zh-TW" sz="4400" b="1" dirty="0">${fillW}</a:rPr></a:br>` +
    `<a:r><a:rPr lang="en-US" altLang="zh-CN" sz="4400" b="1" dirty="0">${fillW}<a:latin typeface="EB Garamond"/></a:rPr><a:t>{{主日标题_EN}}</a:t></a:r>` +
    `<a:endParaRPr lang="en-HK" sz="4400" b="1" dirty="0">${fillW}</a:endParaRPr></a:p></p:txBody>`;
  newSp = newSp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx);
  return xml.replace(sp, newSp);
}

// ── 固定页排版修复（同 make-placeholder-template）：英文左对齐 + 溢出烤缩放 ──
const BOX_FALLBACK = { "3": { cx: 4863060, cy: 5254561, sz: 4800, lang: "cn" }, "4": { cx: 3138026, cy: 5254560, sz: 2800, lang: "en" } };
// 字宽（em）：中日韩 1、其它 0.55 —— 按人工 deck 实测校准（166 字符英文@2800 恰好 11 行放下）。
const emUnits = (t) => [...String(t)].reduce((u, ch) => u + (/[　-鿿＀-￯]/.test(ch) ? 1 : 0.55), 0);
const rowCount = (t, cap) => String(t).split("\n").reduce((r, l) => r + Math.max(1, Math.ceil(emUnits(l) / cap)), 0);
function resolveBoxSz(sp, rels, shapeId) {
  const fb = BOX_FALLBACK[shapeId] || { cx: 8000000, cy: 5300000, sz: 3200, lang: "en" };
  const szs = [...sp.matchAll(/sz="(\d+)"/g)].map((m) => +m[1]);
  let sz = szs.length ? Math.max(...szs) : null;
  const own = sp.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
  let cx = own ? +own[1] : null;
  let cy = own ? +own[2] : null;
  const phIdx = sp.match(/<p:ph[^>]*idx="(\d+)"/)?.[1];
  const phType = sp.match(/<p:ph[^>]*type="(\w+)"/)?.[1];
  const layoutName = (rels || "").match(/slideLayout\d+\.xml/)?.[0];
  if ((!cx || !sz) && layoutName) {
    try {
      const lx = readFileSync(join(work, "ppt/slideLayouts", layoutName), "utf8");
      for (const lsp of lx.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
        const lIdx = lsp[0].match(/<p:ph[^>]*idx="(\d+)"/)?.[1];
        const lType = lsp[0].match(/<p:ph[^>]*type="(\w+)"/)?.[1];
        if ((phIdx && lIdx === phIdx) || (!phIdx && phType && lType === phType)) {
          const ext = lsp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"/);
          if (ext) { cx = cx || +ext[1]; cy = cy || +ext[2]; }
          const dsz = lsp[0].match(/<a:defRPr[^>]*sz="(\d+)"/)?.[1];
          if (dsz && !sz) sz = +dsz;
        }
      }
    } catch {}
  }
  return { cx: cx || fb.cx, cy: cy || fb.cy, sz: sz || fb.sz, lang: fb.lang };
}
function setAutofit(bodyPr, scalePct) {
  const fit = scalePct >= 100 ? "<a:normAutofit/>" : `<a:normAutofit fontScale="${Math.round(scalePct * 1000)}" lnSpcReduction="${scalePct < 80 ? 10000 : 0}"/>`;
  if (/<a:(normAutofit|noAutofit|spAutoFit)\b/.test(bodyPr)) return bodyPr.replace(/<a:(normAutofit|noAutofit|spAutoFit)\b[^>]*\/>/, fit);
  if (/<a:bodyPr\b[^>]*\/>/.test(bodyPr)) return bodyPr.replace(/(<a:bodyPr\b[^>]*?)\s*\/>/, `$1>${fit}</a:bodyPr>`);
  return bodyPr.replace(/(<a:bodyPr\b[^>]*>)/, `$1${fit}`);
}
function fixFixedBox(xml, rels, shapeId) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp || !/<p:txBody>/.test(sp)) return xml;
  const t = shapeParaText(sp);
  if (!t || t.includes("{{")) return xml;
  const { cx, cy, sz } = resolveBoxSz(sp, rels, shapeId);
  // 语言按内容判断（英文框里也可能混中文短语）：中日韩占比 >30% 按中文行高。
  const cjk = [...t].filter((ch) => /[　-鿿＀-￯]/.test(ch)).length;
  const lang = cjk > t.length * 0.3 ? "cn" : "en";
  let nsp = sp;
  if (lang === "en") nsp = nsp.replace(/algn="just"/g, 'algn="l"');
  const fontEmu = (sz / 100) * 12700;
  const lineHFactor = lang === "en" ? 1.15 : 1.18;
  // 迭代求缩放（与填充程序同一模型）：从 100% 逐步降到估算能放下为止。
  let scale = 100;
  while (scale > 45) {
    const f = fontEmu * (scale / 100);
    const cap = Math.max(1, (cx * 0.94) / f);
    const maxl = Math.max(1, Math.floor((cy * 0.96) / (f * lineHFactor)));
    if (rowCount(t, cap) <= maxl) break;
    scale -= 3;
  }
  const existing = Number(nsp.match(/normAutofit fontScale="(\d+)"/)?.[1] || 0) / 1000 || 100;
  const finalScale = Math.min(existing, scale);
  if (finalScale < 100) {
    const bodyPr = nsp.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/)?.[0] || "<a:bodyPr/>";
    nsp = nsp.replace(bodyPr, setAutofit(bodyPr, finalScale));
  }
  return xml.replace(sp, nsp);
}

// ───────────────────────── 解包 ─────────────────────────
const work = "/tmp/make-prototype-template-work";
rmSync(work, { recursive: true, force: true });
mkdirSync(work, { recursive: true });
execFileSync("unzip", ["-q", sourcePptx, "-d", work]);

const slideDir = join(work, "ppt/slides");
const relsDir = join(slideDir, "_rels");
const slideCount = Math.max(...readdirSync(slideDir).map((f) => f.match(/^slide(\d+)\.xml$/)?.[1]).filter(Boolean).map(Number));

// ───────────────────────── 逐页处理 + 折叠 ─────────────────────────
const seen = new Set(); // 已保留原型的可变段
const kept = []; // {xml, rels}
const report = { prototypes: [], deleted: 0, red: [], fixed: 0 };

for (let i = 1; i <= slideCount; i += 1) {
  let xml = readFileSync(join(slideDir, `slide${i}.xml`), "utf8");
  const relsPath = join(relsDir, `slide${i}.xml.rels`);
  const rels = existsSync(relsPath) ? readFileSync(relsPath, "utf8") : "";

  // 封面
  if (i === 1) {
    kept.push({ xml: coverTitle(xml), rels });
    continue;
  }

  // 标题栏优先 id2；歌咏页（进堂咏/奉献咏/礼成咏）标题在 id10、id2 为空。
  const title = shapeText(shapeBlock(xml, "2")) || shapeText(shapeBlock(xml, "10"));

  // 歌咏 / 堂区报告 → 红字
  const redKey = RED_SECTIONS.find((k) => title.includes(k));
  if (redKey) {
    const contentIds = textShapeIds(xml).filter((id) => id !== "2" && id !== "10" && shapeText(shapeBlock(xml, id)) && !POSTURE.test(shapeText(shapeBlock(xml, id))));
    let noticeId = contentIds.find((id) => shapeText(shapeBlock(xml, id)).includes("请手动填入"));
    if (!noticeId && contentIds.length) noticeId = contentIds.slice().sort((a, b) => shapeText(shapeBlock(xml, b)).length - shapeText(shapeBlock(xml, a)).length)[0];
    if (redKey === "堂区报告" && /手动填入/.test(xml) && /<p:pic>/.test(xml)) {
      // 已是「原版设计 + 红字」的移植页（含图标）→ 原样保留，不再清理。
      kept.push({ xml, rels });
      report.red.push(`${redKey}(src${i},保留设计)`);
      continue;
    }
    if (noticeId) {
      xml = putRedNotice(xml, noticeId, RED_LABEL[redKey], redKey === "堂区报告" ? "announce" : "hymn");
      if (redKey === "堂区报告") {
        for (const id of contentIds) if (id !== noticeId) xml = deleteShape(xml, id);
        xml = deleteAllPics(xml);
      } else {
        for (const id of contentIds) if (id !== noticeId) xml = deleteShape(xml, id);
      }
      report.red.push(`${redKey}(src${i})`);
    }
    kept.push({ xml, rels });
    continue;
  }

  // 可变段 → 折叠成原型页（首次保留 + 命名占位符；其余删除）
  const varKey = VAR_ORDER.find((k) => title.includes(k));
  if (varKey) {
    if (seen.has(varKey)) {
      report.deleted += 1; // 该段的后续页 → 删除
      continue;
    }
    seen.add(varKey);
    // 只要该框存在就放占位符（首页英文框可能当前为空，但结构上要有 {{_EN}} 槽位供程序填）。
    xml = putToken(xml, "3", `{{${varKey}_CN}}`, "cn");
    xml = putToken(xml, "4", `{{${varKey}_EN}}`, "en");
    kept.push({ xml, rels });
    report.prototypes.push(varKey);
    continue;
  }

  // 固定经文：保留 + 排版修复
  xml = fixFixedBox(fixFixedBox(xml, rels, "3"), rels, "4");
  kept.push({ xml, rels });
  report.fixed += 1;
}

// ───────────────────────── 重新编号 + 重打包 ─────────────────────────
for (const f of readdirSync(slideDir)) if (/^slide\d+\.xml$/.test(f)) rmSync(join(slideDir, f));
for (const f of readdirSync(relsDir)) if (/^slide\d+\.xml\.rels$/.test(f)) rmSync(join(relsDir, f));
kept.forEach((s, idx) => {
  const n = idx + 1;
  writeFileSync(join(slideDir, `slide${n}.xml`), s.xml);
  writeFileSync(join(relsDir, `slide${n}.xml.rels`), s.rels || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
});

let pres = readFileSync(join(work, "ppt/presentation.xml"), "utf8");
const sldIds = kept.map((_s, i) => `<p:sldId id="${2147483648 - kept.length + i}" r:id="rIdSlide${i + 1}"/>`).join("");
pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIds}</p:sldIdLst>`);
writeFileSync(join(work, "ppt/presentation.xml"), pres);

let prels = readFileSync(join(work, "ppt/_rels/presentation.xml.rels"), "utf8");
const nonSlide = [...prels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).filter((r) => !r.includes('/relationships/slide"'));
const slideRels = kept.map((_s, i) => `<Relationship Id="rIdSlide${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`);
prels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nonSlide.join("")}${slideRels.join("")}</Relationships>`;
writeFileSync(join(work, "ppt/_rels/presentation.xml.rels"), prels);

let ct = readFileSync(join(work, "[Content_Types].xml"), "utf8");
ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "");
const overrides = kept.map((_s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
ct = ct.replace("</Types>", `${overrides}</Types>`);
writeFileSync(join(work, "[Content_Types].xml"), ct);

const appPath = join(work, "docProps/app.xml");
if (existsSync(appPath)) {
  let app = readFileSync(appPath, "utf8");
  app = app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${kept.length}</Slides>`);
  writeFileSync(appPath, app);
}

mkdirSync(dirname(outPptx), { recursive: true });
if (existsSync(outPptx)) rmSync(outPptx);
execFileSync("zip", ["-qr", outPptx, "."], { cwd: work });

console.log(JSON.stringify({ out: outPptx, from: slideCount, slides: kept.length, prototypes: report.prototypes, deletedVarPages: report.deleted, red: report.red, fixedKept: report.fixed }, null, 2));
