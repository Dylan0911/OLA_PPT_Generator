// 把「常年期第十二主日」成品 deck 转成【占位符模板】：
// 固定内容原样保留；每个可变段折叠成一张「原型页」，其文字框换成 {段落_CN}/{段落_EN} 占位符，
// 并保留该形状原本的字体/字号/颜色/对齐。以后生成只需把占位符替换成文档内容即可。
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

const sourcePptx = process.argv[2];
const outPptx = process.argv[3];
if (!sourcePptx || !outPptx) {
  console.error("用法: node make-template.mjs <source.pptx> <out.pptx>");
  process.exit(1);
}

// 可变段配置：keep=原型页；其后到 del 的页删除；ph=形状ID→占位符。
const SECTIONS = [
  { keep: 2, del: 7, ph: { 8: "{进堂咏_CN}" } },
  { keep: 21, del: 22, ph: { 3: "{集祷经_CN}", 4: "{集祷经_EN}" } },
  { keep: 25, del: 32, ph: { 3: "{读经一_CN}", 4: "{读经一_EN}" } },
  { keep: 33, del: 42, ph: { 3: "{答唱咏_CN}", 4: "{答唱咏_EN}" } },
  { keep: 43, del: 47, ph: { 3: "{读经二_CN}", 4: "{读经二_EN}" } },
  { keep: 48, del: 48, ph: { 3: "{福音前欢呼_CN}", 4: "{福音前欢呼_EN}" } },
  { keep: 49, del: 57, ph: { 3: "{福音_CN}", 4: "{福音_EN}" } },
  { keep: 67, del: 78, ph: { 3: "{信友祷文_CN}" } },
  { keep: 81, del: 83, ph: { 8: "{奉献咏_CN}" } },
  { keep: 84, del: 86, ph: { 3: "{献礼经_CN}", 4: "{献礼经_EN}" } },
  { keep: 240, del: 244, ph: { 3: "{领主咏_CN}", 4: "{领主咏_EN}" } },
  { keep: 245, del: 246, ph: { 3: "{领圣体后经_CN}", 4: "{领圣体后经_EN}" } },
  { keep: 248, del: 252, ph: { 4: "{堂区报告}" } }, // 247 分隔页保留
  { keep: 260, del: 264, ph: { 8: "{礼成咏_CN}" } },
];

const work = "/tmp/make-template-work";
rmSync(work, { recursive: true, force: true });
mkdirSync(work, { recursive: true });
execFileSync("unzip", ["-q", sourcePptx, "-d", work]);

const slideCount = Math.max(
  ...readdirSync(join(work, "ppt/slides"))
    .map((f) => f.match(/^slide(\d+)\.xml$/)?.[1])
    .filter(Boolean)
    .map(Number)
);

const escapeXml = (v) => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// 在指定形状里放占位符，保留 bodyPr + 首个 run 的 rPr + 首段 pPr（即原本的字体/字号/颜色/对齐）。
function putPlaceholder(xml, shapeId, token) {
  const shapeRe = new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${shapeId}"(?:(?!</p:sp>)[\\s\\S])*?</p:sp>`);
  if (!shapeRe.test(xml)) return xml;
  return xml.replace(shapeRe, (sp) => {
    const tx = sp.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)?.[0];
    if (!tx) return sp;
    const bodyPr = tx.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/)?.[0] || "<a:bodyPr/>";
    const rpr = tx.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || "";
    const pPr = tx.match(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] || "";
    const newTx = `<p:txBody>${bodyPr}<a:lstStyle/><a:p>${pPr}<a:r>${rpr}<a:t>${escapeXml(token)}</a:t></a:r></a:p></p:txBody>`;
    return sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx);
  });
}

// 封面标题（shape 4）：中英两行占位符，沿用居中 + 4800/4400 + 白色格式。
function coverPlaceholder(xml) {
  const shapeRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<p:cNvPr id="4"(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/;
  return xml.replace(shapeRe, (sp) => {
    const newTx = `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPts val="5200"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="4800" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>{主日标题_CN}</a:t></a:r><a:br><a:rPr lang="en-HK" altLang="zh-TW" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr></a:br><a:r><a:rPr lang="en-US" altLang="zh-CN" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>{主日标题_EN}</a:t></a:r><a:endParaRPr lang="en-HK" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:endParaRPr></a:p></p:txBody>`;
    return sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx);
  });
}

// 建删除集 + 原型页→占位符映射
const deleteSet = new Set();
const phByStart = new Map();
for (const s of SECTIONS) {
  phByStart.set(s.keep, s.ph);
  for (let i = s.keep + 1; i <= s.del; i += 1) deleteSet.add(i);
}

// 组装新页序
const kept = [];
for (let i = 1; i <= slideCount; i += 1) {
  if (deleteSet.has(i)) continue;
  let xml = readFileSync(join(work, `ppt/slides/slide${i}.xml`), "utf8");
  if (i === 1) xml = coverPlaceholder(xml);
  if (phByStart.has(i)) {
    for (const [shapeId, token] of Object.entries(phByStart.get(i))) xml = putPlaceholder(xml, shapeId, token);
  }
  kept.push({ src: i, xml });
}

// 写新页（重新编号）+ 搬运 rels
const relsDir = join(work, "ppt/slides/_rels");
const oldRels = new Map();
for (let i = 1; i <= slideCount; i += 1) {
  const p = join(relsDir, `slide${i}.xml.rels`);
  if (existsSync(p)) oldRels.set(i, readFileSync(p, "utf8"));
}
for (const f of readdirSync(join(work, "ppt/slides"))) if (/^slide\d+\.xml$/.test(f)) rmSync(join(work, "ppt/slides", f));
for (const f of readdirSync(relsDir)) if (/^slide\d+\.xml\.rels$/.test(f)) rmSync(join(relsDir, f));
kept.forEach((s, idx) => {
  const n = idx + 1;
  writeFileSync(join(work, `ppt/slides/slide${n}.xml`), s.xml);
  writeFileSync(join(relsDir, `slide${n}.xml.rels`), oldRels.get(s.src) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
});

// presentation.xml sldIdLst
let pres = readFileSync(join(work, "ppt/presentation.xml"), "utf8");
const sldIds = kept.map((_s, i) => `<p:sldId id="${2147483648 - kept.length + i}" r:id="rIdSlide${i + 1}"/>`).join("");
pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIds}</p:sldIdLst>`);
writeFileSync(join(work, "ppt/presentation.xml"), pres);

// presentation.xml.rels：保留非slide关系，重建slide关系
let prels = readFileSync(join(work, "ppt/_rels/presentation.xml.rels"), "utf8");
const nonSlide = [...prels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).filter((r) => !r.includes("/relationships/slide\""));
const slideRels = kept.map((_s, i) => `<Relationship Id="rIdSlide${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`);
prels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nonSlide.join("")}${slideRels.join("")}</Relationships>`;
writeFileSync(join(work, "ppt/_rels/presentation.xml.rels"), prels);

// [Content_Types].xml slide overrides
let ct = readFileSync(join(work, "[Content_Types].xml"), "utf8");
ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "");
const overrides = kept.map((_s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
ct = ct.replace("</Types>", `${overrides}</Types>`);
writeFileSync(join(work, "[Content_Types].xml"), ct);

// docProps/app.xml Slides 数
const appPath = join(work, "docProps/app.xml");
if (existsSync(appPath)) {
  let app = readFileSync(appPath, "utf8");
  app = app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${kept.length}</Slides>`);
  writeFileSync(appPath, app);
}

mkdirSync(dirname(outPptx), { recursive: true });
if (existsSync(outPptx)) rmSync(outPptx);
execFileSync("zip", ["-qr", outPptx, "."], { cwd: work });
console.log(JSON.stringify({ out: outPptx, slides: kept.length, from: slideCount }, null, 2));
