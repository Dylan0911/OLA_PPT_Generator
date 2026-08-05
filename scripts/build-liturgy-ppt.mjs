import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const sourcePptx = process.argv[2];
const workspace = process.argv[3];
const finalPptx = process.argv[4];

if (!sourcePptx || !workspace || !finalPptx) {
  console.error("Usage: node build-lent5-full-template.mjs <source.pptx> <workspace> <final.pptx>");
  process.exit(1);
}

const here = dirname(fileURLToPath(import.meta.url));
const sectionsPath = process.argv[5] || "";
const inputPayload = sectionsPath ? JSON.parse(readFileSync(sectionsPath, "utf8")) : null;
let sections;
if (inputPayload?.sections) {
  sections = inputPayload.sections;
} else {
  const fallbackSectionsCode = readFileSync(join(here, "build-lent5-deck.mjs"), "utf8").match(
    /const sections = ([\s\S]*?);\n\nfunction escapeXml/
  )[1];
  sections = vm.runInNewContext(fallbackSectionsCode);
}
const meta = {
  titleCn: inputPayload?.meta?.titleCn || "常年期主日",
  titleEn: inputPayload?.meta?.titleEn || "",
  includeGloria: inputPayload?.meta?.includeGloria !== false,
};

// 绿色「常年期」模板原有配色：主绿色（横幅/装饰形状）+ 浅绿点缀。
// 用户在界面选择主题色后，统一把这两个颜色替换成所选色及其浅色衍生，
// 正文深色 282120 等保持不变。（旧棕金模板用 BF913B / F2E1AC。）
const DEFAULT_PRIMARY = "55725D";
const DEFAULT_LIGHT = "BCCCBE";

function normalizeHex(value) {
  if (!value) return "";
  let hex = String(value).trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{3}$/.test(hex)) hex = hex.split("").map((c) => c + c).join("");
  return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : "";
}

function lightTint(hex, amount = 0.8) {
  const channel = (i) => parseInt(hex.slice(i, i + 2), 16);
  const mix = (c) => Math.round(c + (255 - c) * amount);
  return [channel(0), channel(2), channel(4)]
    .map((c) => mix(c).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

const themePrimary = normalizeHex(inputPayload?.meta?.themeColor);
const themeRecolor =
  themePrimary && themePrimary !== DEFAULT_PRIMARY
    ? new Map([
        [DEFAULT_PRIMARY, themePrimary],
        [DEFAULT_LIGHT, lightTint(themePrimary)],
      ])
    : null;

function applyThemeRecolor(xml) {
  if (!themeRecolor) return xml;
  return xml.replace(/val="([0-9A-Fa-f]{6})"/g, (match, hex) => {
    const repl = themeRecolor.get(hex.toUpperCase());
    return repl ? `val="${repl}"` : match;
  });
}

// 字体策略：模板若把 DFKai-SB 以「子集」方式嵌入（旧棕金模板），新字会缺字乱码，
// 这时改用系统「楷体」；若模板没有嵌入字体（绿色常年期模板），则沿用模板原字体
// DFKai-SB，保持与保留页完全一致。CHINESE_FONT 在解包后根据是否嵌入字体决定（见下）。
const LEGACY_CHINESE_FONT = "DFKai-SB";

// 移除指定字体的失效嵌入（嵌入条目 + 关系 + 字体文件），避免缺字回退并减小体积。
function removeEmbeddedFont(typeface) {
  const presPath = join(packageDir, "ppt/presentation.xml");
  if (!existsSync(presPath)) return;
  let pres = readFileSync(presPath, "utf8");
  const blockRe = new RegExp(
    `<p:embeddedFont>(?:(?!</p:embeddedFont>)[\\s\\S])*?${typeface}(?:(?!</p:embeddedFont>)[\\s\\S])*?</p:embeddedFont>`
  );
  const block = pres.match(blockRe);
  if (!block) return;
  const rid = block[0].match(/r:id="(rId\d+)"/)?.[1];
  pres = pres.replace(blockRe, "");
  writeFileSync(presPath, pres);
  if (!rid) return;
  const relsPath = join(packageDir, "ppt/_rels/presentation.xml.rels");
  if (!existsSync(relsPath)) return;
  let rels = readFileSync(relsPath, "utf8");
  const relRe = new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*/>`);
  const target = rels.match(relRe)?.[0].match(/Target="([^"]+)"/)?.[1];
  rels = rels.replace(relRe, "");
  writeFileSync(relsPath, rels);
  if (target) {
    const fontFile = join(packageDir, "ppt", target.replace(/^\.?\/?/, ""));
    if (existsSync(fontFile)) unlinkSync(fontFile);
  }
}

const byTitle = new Map(sections.map((s) => [s.titleCn, { ...s }]));

const packageDir = join(workspace, "package-lent5-full-template");
if (existsSync(packageDir)) renameSync(packageDir, `${packageDir}.previous-${Date.now()}`);
mkdirSync(packageDir, { recursive: true });
execFileSync("unzip", ["-q", sourcePptx, "-d", packageDir], { stdio: "inherit" });

// 统一中文字体为「楷体」（= 模板主题的东亚字体）。模板里显式写死的 DFKai-SB 在没装该字体的
// 环境（如 iPad/Keynote）会回退成难看的黑体，所以一律改名为楷体，与继承楷体的页面保持一致。
const templateEmbedsFonts =
  existsSync(join(packageDir, "ppt/fonts")) &&
  readdirSync(join(packageDir, "ppt/fonts")).some((f) => f.endsWith(".fntdata"));
const CHINESE_FONT = "楷体";

const sourceSlideCount = Math.max(
  ...readdirSync(join(packageDir, "ppt/slides"))
    .map((file) => file.match(/^slide(\d+)\.xml$/)?.[1])
    .filter(Boolean)
    .map(Number)
);

const sourceSlides = new Map();
for (let i = 1; i <= sourceSlideCount; i += 1) {
  sourceSlides.set(i, {
    xml: readFileSync(join(packageDir, `ppt/slides/slide${i}.xml`), "utf8"),
    rels: existsSync(join(packageDir, `ppt/slides/_rels/slide${i}.xml.rels`))
      ? readFileSync(join(packageDir, `ppt/slides/_rels/slide${i}.xml.rels`), "utf8")
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
  });
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function units(text, lang) {
  let total = 0;
  for (const ch of String(text || "")) {
    if (ch === "\n") total += lang === "cn" ? 8 : 16;
    else if (lang === "cn") total += /[\u0000-\u007f]/.test(ch) ? 0.48 : 1;
    else total += 1;
  }
  return total;
}

function splitLong(text, maxUnits, lang) {
  const chunks = [];
  let rest = String(text || "").trim();
  while (units(rest, lang) > maxUnits) {
    let best = -1;
    const re = lang === "cn" ? /[。！？；：，、]/g : /[.!?;:,]\s/g;
    for (const m of rest.matchAll(re)) {
      if (units(rest.slice(0, m.index + 1), lang) <= maxUnits) best = m.index + 1;
    }
    if (best < Math.max(8, maxUnits * 0.45)) {
      let acc = 0;
      let lastSpace = 0;
      best = 0;
      for (const ch of rest) {
        acc += units(ch, lang);
        if (acc > maxUnits) break;
        best += ch.length;
        if (lang === "en" && /\s/.test(ch)) lastSpace = best;
      }
      if (lang === "en" && lastSpace > Math.max(8, best * 0.55)) best = lastSpace;
    }
    best = adjustBreakPosition(rest, best, maxUnits, lang);
    chunks.push(rest.slice(0, best).trim());
    rest = rest.slice(best).trim();
  }
  if (rest) chunks.push(rest);
  return chunks;
}

function adjustBreakPosition(text, index, maxUnits, lang) {
  let best = Math.max(0, index);
  const clean = String(text || "");
  const closing = "”’」』）)]}》,，.。;；:：!！?？";
  while (best < clean.length && closing.includes(clean[best]) && units(clean.slice(0, best + 1), lang) <= maxUnits * 1.18) {
    best += 1;
  }
  const left = clean.slice(0, best);
  const right = clean.slice(best);
  const lastOpen = Math.max(left.lastIndexOf("“"), left.lastIndexOf("‘"), left.lastIndexOf("「"), left.lastIndexOf("『"));
  const lastClose = Math.max(left.lastIndexOf("”"), left.lastIndexOf("’"), left.lastIndexOf("」"), left.lastIndexOf("』"));
  if (lastOpen > lastClose) {
    const closeInRight = right.search(/[”’」』]/u);
    if (closeInRight >= 0) {
      const candidate = best + closeInRight + 1;
      if (units(clean.slice(0, candidate), lang) <= maxUnits * 1.25) best = candidate;
    }
  }
  const opening = "“‘「『([（";
  if (best > 0 && opening.includes(clean[best - 1])) {
    const nextStop = right.search(/[。！？；，,.;!?]/u);
    const candidate = nextStop >= 0 ? best + nextStop + 1 : Math.min(clean.length, best + 8);
    if (units(clean.slice(0, candidate), lang) <= maxUnits * 1.2) best = candidate;
  }
  return best || index;
}

const chunkConfig = {
  singleCn: { max: 96, sentence: 28, maxLines: 5 },
  noGloria: { max: 42, sentence: 42, maxLines: 2 },
  dialogueCn: { max: 92, sentence: 52, maxLines: 5 },
  faithfulCn: { max: 92, sentence: 92, maxLines: 4 },
  prayerCn: { max: 58, sentence: 50, maxLines: 2 },
  twoCn: { max: 70, sentence: 64, maxLines: 3 },
  twoEn: { max: 235, sentence: 215, maxLines: 5 },
  gospelCn: { max: 78, sentence: 70, maxLines: 3 },
  gospelEn: { max: 255, sentence: 230, maxLines: 5 },
};

function splitParts(text, lang, sentenceMax) {
  return String(text || "")
    .replace(/\r/g, "")
    .split(/\n+/)
    .flatMap((p) => {
      const t = p.trim();
      if (!t) return [""];
      return units(t, lang) <= sentenceMax ? [t] : splitLong(t, sentenceMax, lang);
    });
}

function chunkText(text, lang, mode) {
  const clean = String(text || "").trim();
  if (!clean) return [];
  const config = chunkConfig[mode];
  const out = [];
  let cur = [];
  for (const part of splitParts(clean, lang, config.sentence)) {
    if (part === "") {
      if (cur.length) out.push(cur.join("\n"));
      cur = [];
      continue;
    }
    const trial = cur.length ? `${cur.join("\n")}\n${part}` : part;
    if (cur.length && (units(trial, lang) > config.max || cur.length >= config.maxLines)) {
      out.push(cur.join("\n"));
      cur = [part];
    } else {
      cur.push(part);
    }
  }
  if (cur.length) out.push(cur.join("\n"));
  return out.flatMap((chunk) => (units(chunk, lang) > config.max * 1.12 ? splitLong(chunk, config.max, lang) : [chunk]));
}

function paragraphsOf(text) {
  return String(text || "")
    .replace(/\r/g, "")
    .split(/\n+/)
    .map((p) => p.trim())
    .filter(Boolean);
}

function splitOnce(text, lang) {
  const clean = String(text || "").trim();
  if (units(clean, lang) < 10) return [clean];
  const target = units(clean, lang) / 2;
  let best = -1;
  const re = lang === "cn" ? /[。！？；：，、]/g : /[.!?;:,]\s/g;
  for (const m of clean.matchAll(re)) {
    const pos = m.index + 1;
    const leftUnits = units(clean.slice(0, pos), lang);
    if (leftUnits <= target && leftUnits >= target * 0.55) best = pos;
  }
  if (best < 0) {
    let acc = 0;
    let lastSpace = 0;
    best = 0;
    for (const ch of clean) {
      acc += units(ch, lang);
      if (acc >= target) break;
      best += ch.length;
      if (lang === "en" && /\s/.test(ch)) lastSpace = best;
    }
    if (lang === "en" && lastSpace > Math.max(8, best * 0.55)) best = lastSpace;
  }
  best = adjustBreakPosition(clean, best, units(clean, lang) * 0.62, lang);
  const left = clean.slice(0, best).trim();
  const right = clean.slice(best).trim();
  return left && right ? [left, right] : [clean];
}

function growChunksTo(text, lang, mode, target) {
  const chunks = chunkText(text, lang, mode);
  while (chunks.length < target) {
    let largest = -1;
    let largestUnits = 0;
    for (let i = 0; i < chunks.length; i += 1) {
      const u = units(chunks[i], lang);
      if (u > largestUnits) {
        largest = i;
        largestUnits = u;
      }
    }
    if (largest < 0 || largestUnits < 12) break;
    const split = splitOnce(chunks[largest], lang);
    if (split.length < 2) break;
    chunks.splice(largest, 1, ...split);
  }
  return chunks;
}

// 把一段文字拆成句子（按换行 + 中英句末标点），用于比例对齐。
function sentencesOf(text, lang) {
  const clean = String(text || "").replace(/\r/g, "").trim();
  if (!clean) return [];
  const out = [];
  for (const line of clean.split(/\n+/)) {
    const t = line.trim();
    if (!t) continue;
    const re = lang === "cn" ? /[^。！？；]*[。！？；]?/g : /[^.!?]*[.!?]?/g;
    const parts = (t.match(re) || []).map((s) => s.trim()).filter(Boolean);
    if (parts.length) out.push(...parts);
    else out.push(t);
  }
  return out;
}

// 句子数不够分成 minCount 组时，把最长的句子按次级标点（逗号/顿号/分号）再切，
// 直到片段够数或无法再分 —— 避免比例切分时某一侧出现空页。
function sentencePieces(text, lang, minCount) {
  const pieces = sentencesOf(text, lang);
  const secondary = lang === "cn" ? /[，、；：]/g : /[,;:]/g;
  while (pieces.length < minCount) {
    let idx = -1;
    let best = 0;
    for (let i = 0; i < pieces.length; i += 1) {
      const u = units(pieces[i], lang);
      if (u > best && secondary.test(pieces[i])) {
        best = u;
        idx = i;
      }
      secondary.lastIndex = 0;
    }
    if (idx < 0) break;
    const piece = pieces[idx];
    const marks = [...piece.matchAll(secondary)];
    if (!marks.length) break;
    const cut = marks[Math.floor((marks.length - 1) / 2)].index + 1;
    const left = piece.slice(0, cut).trim();
    const right = piece.slice(cut).trim();
    if (!left || !right) break;
    pieces.splice(idx, 1, left, right);
  }
  return pieces;
}

// 把文字按「累计长度比例」切成 parts 组：第 i 组覆盖整段的 [i/parts, (i+1)/parts) 区间。
// 中英用同一个 parts 调用，就能保证第 i 页的中文与英文落在原文的同一相对位置 —— 互相对应。
function splitByProportion(text, lang, parts) {
  if (parts <= 1 || !String(text || "").trim()) {
    const t = String(text || "").trim();
    return t ? [t] : [];
  }
  const sents = sentencePieces(text, lang, parts);
  if (sents.length <= 1) return [String(text).trim()];
  const groups = Array.from({ length: parts }, () => []);
  const total = sents.reduce((sum, s) => sum + units(s, lang), 0) || 1;
  let acc = 0;
  let gi = 0;
  for (let k = 0; k < sents.length; k += 1) {
    groups[gi].push(sents[k]);
    acc += units(sents[k], lang);
    const remaining = sents.length - k - 1;
    const slotsLeft = parts - 1 - gi;
    // 必须给后面每组都留至少一句 → 余量刚好时强制进位；否则按比例进位。
    if (slotsLeft > 0 && (remaining <= slotsLeft || acc >= ((gi + 1) / parts) * total)) {
      gi += 1;
    }
  }
  const joiner = lang === "cn" ? "" : " ";
  return groups.map((g) => g.join(joiner).trim());
}

function alignPair(cnText, enText, cnMode, enMode) {
  const cnChunks = chunkText(cnText, "cn", cnMode);
  const enChunks = chunkText(enText, "en", enMode);
  const count = Math.max(cnChunks.length, enChunks.length, 1);
  // 中英在相同比例位置切分，保证逐页对应（替代各自独立切片再硬凑数量）。
  const cn = count > 1 && String(cnText || "").trim() ? splitByProportion(cnText, "cn", count) : cnChunks;
  const en = count > 1 && String(enText || "").trim() ? splitByProportion(enText, "en", count) : enChunks;
  return Array.from({ length: count }, (_v, i) => ({
    cn: cn[i] || "",
    en: en[i] || "",
  }));
}

function chunkBilingual(cn, en, cnMode = "twoCn", enMode = "twoEn", pairMode = "auto") {
  const cnParas = paragraphsOf(cn);
  const enParas = paragraphsOf(en);
  if (pairMode === "global") return alignPair(cn, en, cnMode, enMode);
  if (pairMode === "auto") {
    if (cnParas.length <= 1 || enParas.length <= 1) return alignPair(cn, en, cnMode, enMode);
    const ratio = Math.max(cnParas.length, enParas.length) / Math.max(1, Math.min(cnParas.length, enParas.length));
    if (ratio > 1.65) return alignPair(cn, en, cnMode, enMode);
  }
  const max = Math.max(cnParas.length, enParas.length, 1);
  const slides = [];
  for (let i = 0; i < max; i += 1) {
    slides.push(...alignPair(cnParas[i] || "", enParas[i] || "", cnMode, enMode));
  }
  return slides;
}

function chunkBilingualGlobal(cn, en, cnMode = "twoCn", enMode = "twoEn") {
  return alignPair(cn, en, cnMode, enMode);
}

function chunkBilingualProportional(cn, en, cnMode = "twoCn", enMode = "twoEn") {
  const cnParas = paragraphsOf(cn);
  const enParas = paragraphsOf(en);
  if (cnParas.length <= 1 || enParas.length <= 1) return alignPair(cn, en, cnMode, enMode);

  const cnTotal = cnParas.reduce((sum, p) => sum + units(p, "cn"), 0) || 1;
  const enTotal = enParas.reduce((sum, p) => sum + units(p, "en"), 0) || 1;
  let cnSeen = 0;
  let enSeen = 0;
  let enIndex = 0;
  const slides = [];

  for (let i = 0; i < cnParas.length; i += 1) {
    cnSeen += units(cnParas[i], "cn");
    const remainingCn = cnParas.length - i - 1;
    const targetEn = (cnSeen / cnTotal) * enTotal;
    const group = [];
    while (enIndex < enParas.length - remainingCn) {
      const nextUnits = units(enParas[enIndex], "en");
      const mustTake = group.length === 0;
      const currentGap = Math.abs(targetEn - enSeen);
      const nextGap = Math.abs(targetEn - (enSeen + nextUnits));
      if (!mustTake && nextGap > currentGap) break;
      group.push(enParas[enIndex]);
      enSeen += nextUnits;
      enIndex += 1;
    }
    slides.push(...alignPair(cnParas[i], group.join("\n"), cnMode, enMode));
  }

  if (enIndex < enParas.length) {
    const tail = enParas.slice(enIndex).join("\n");
    const last = slides[slides.length - 1];
    if (last && !last.en) last.en = tail;
    else slides.push(...alignPair("", tail, cnMode, enMode));
  }
  return slides;
}

function semanticParts(text, lang, mode) {
  return splitParts(text, lang, chunkConfig[mode].sentence)
    .map((part) => part.trim())
    .filter(Boolean);
}

function chunkBilingualProgress(cn, en, cnMode = "twoCn", enMode = "twoEn") {
  const cnParts = semanticParts(cn, "cn", cnMode);
  const enParts = semanticParts(en, "en", enMode);
  if (!cnParts.length || !enParts.length) return alignPair(cn, en, cnMode, enMode);

  const cnTotal = cnParts.reduce((sum, p) => sum + units(p, "cn"), 0) || 1;
  const enTotal = enParts.reduce((sum, p) => sum + units(p, "en"), 0) || 1;
  let cnSeen = 0;
  let enSeen = 0;
  let enIndex = 0;
  const slides = [];

  for (let i = 0; i < cnParts.length; i += 1) {
    cnSeen += units(cnParts[i], "cn");
    const remainingCn = cnParts.length - i - 1;
    const targetEn = (cnSeen / cnTotal) * enTotal;
    const group = [];
    while (enIndex < enParts.length - remainingCn) {
      const nextUnits = units(enParts[enIndex], "en");
      const mustTake = group.length === 0;
      const currentGap = Math.abs(targetEn - enSeen);
      const nextGap = Math.abs(targetEn - (enSeen + nextUnits));
      if (!mustTake && nextGap > currentGap) break;
      group.push(enParts[enIndex]);
      enSeen += nextUnits;
      enIndex += 1;
    }
    slides.push(...alignPair(cnParts[i], group.join("\n"), cnMode, enMode));
  }

  if (enIndex < enParts.length) {
    const tail = enParts.slice(enIndex).join("\n");
    const last = slides[slides.length - 1];
    if (last && canAppendChunk(last.en, tail, "en", enMode, 1.12)) last.en = appendText(last.en, tail);
    else slides.push(...alignPair("", tail, cnMode, enMode));
  }
  return slides;
}

function splitPoemLine(line, maxUnits = 15) {
  const clean = String(line || "").trim();
  if (!clean) return [];
  const fragments = [];
  const re = /[，,。.;；！？!?]/g;
  let last = 0;
  for (const m of clean.matchAll(re)) {
    const end = m.index + 1;
    fragments.push(clean.slice(last, end).trim());
    last = end;
  }
  if (last < clean.length) fragments.push(clean.slice(last).trim());
  const source = fragments.filter(Boolean).length ? fragments.filter(Boolean) : [clean];
  return source.flatMap((item) => (units(item, "cn") > maxUnits * 1.25 ? splitLong(item, maxUnits, "cn") : [item]));
}

function formatPoemText(text) {
  return paragraphsOf(text)
    .flatMap((line) => splitPoemLine(line))
    .join("\n");
}

function dropLines(text, regexes) {
  return String(text || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter((line) => line && !regexes.some((re) => re.test(line)))
    .join("\n");
}

function scriptureRefRegex() {
  return /^(若|路|玛|谷|宗|咏|罗|格|伯|雅|默|Jn|John|Lk|Luke|Mt|Matthew|Mk|Mark|Acts|Romans|Psalm|Ps)\s*[\d０-９]/i;
}

function splitInlineResponses(text) {
  return String(text || "")
    .replace(/([^\n])\s*(——上主的圣言。)/g, "$1\n$2")
    .replace(/([^\n])\s*(——基督的福音。)/g, "$1\n$2")
    .replace(/([^\n])\s+(The word of the Lord\.)/gi, "$1\n$2")
    .replace(/([^\n])\s+(The Gospel of the Lord\.)/gi, "$1\n$2")
    .replace(/[ \t]+(答[:：])/g, "\n$1")
    .replace(/[ \t]+(众[:：])/g, "\n$1")
    .replace(/(众[:：])\s*\n\s*(阿们。?)/g, "$1$2")
    .replace(/[ \t]+(R\.\s+)/g, "\n$1")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function isCnResponse(line) {
  return /^(答|众)[:：]/.test(String(line || "").trim());
}

function isEnResponse(line) {
  return /^(R\.|Amen\.?$)/i.test(String(line || "").trim());
}

function richifyResponses(text, lang = "cn", forceBold = false) {
  return splitInlineResponses(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const bold = forceBold || (lang === "cn" ? isCnResponse(line) : isEnResponse(line));
      return bold ? { text: line, bold: true } : line;
    });
}

function richifyFaithful(text) {
  return splitInlineResponses(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      if (/^答[:：]\s*(求主俯听我们。?|阿们。?)$/u.test(line)) return { text: line, bold: true };
      const role = line.match(/^(主祭|领)[:：]/u)?.[0];
      if (role) {
        return {
          runs: [
            { text: role, bold: true },
            { text: line.slice(role.length) },
          ],
        };
      }
      return line;
    });
}

function stripReadingCitation(line) {
  return String(line || "")
    .replace(/\s+[0-9０-９][0-9０-９:：;；,，.\-–—a-zA-Z\s]*$/u, "")
    .replace(/\s+/g, " ")
    .trim();
}

function section(title) {
  return byTitle.get(title) || { titleCn: title, titleEn: "", cn: "", en: "" };
}

function entranceHymnText() {
  const hymn = section("进堂咏唱").cn || "";
  const antiphon = section("进堂咏").cn || "";
  return dropLines(hymn || antiphon, [/^咏/, /^不念/, /^Ps\b/i, /^Psalm\b/i]);
}

function gospelCitationFromCnIntro(intro) {
  const match = String(intro || "").match(/(?:恭读)?圣?(玛窦|马尔谷|马可|路加|若望)福音\s*([0-9０-９][0-9０-９:：,，.\-–—a-zA-Z\s]*)/u);
  if (!match) return "";
  const books = {
    玛窦: "Matthew",
    马尔谷: "Mark",
    马可: "Mark",
    路加: "Luke",
    若望: "John",
  };
  const citation = match[2]
    .replace(/[０-９]/g, (ch) => String(ch.charCodeAt(0) - 0xff10))
    .replace(/：/g, ":")
    .replace(/，/g, ",")
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, "")
    .trim();
  return books[match[1]] && citation ? `${books[match[1]]} ${citation}` : "";
}

function trimPrayerText(text, title) {
  const cutPatterns =
    title === "献礼经"
      ? [/^复活期.*颂谢词/, /^颂谢词/, /^Preface\b/i, /^主祭[:：]\s*愿主与你们同在/, /^The Lord be with you\b/i]
      : [];
  const lines = paragraphsOf(text);
  if (title === "献礼经") {
    const endAt = lines.findIndex((line) => /以上所求是靠我们的主基督|Through Christ our Lord/i.test(line));
    if (endAt >= 0) return lines.slice(0, endAt + 1).join("\n");
  }
  const cutAt = lines.findIndex((line) => cutPatterns.some((re) => re.test(line)));
  return (cutAt >= 0 ? lines.slice(0, cutAt) : lines).join("\n");
}

function normalizePrayer(title, options = {}) {
  let cn = trimPrayerText(section(title).cn, title);
  let en = trimPrayerText(section(title).en, title);
  cn = cn.replace(/\s*阿们。?\s*$/u, "").trim();
  en = en.replace(/\s*(Amen\.?|R\.\s*Amen\.?)\s*$/i, "").trim();

  if (options.letUsPray) {
    cn = cn ? `请大家祈祷：${cn}` : "请大家祈祷：";
    en = en ? `Let us pray. ${en}` : "Let us pray.";
  }
  if (options.responseAmen !== false) {
    cn = `${cn}\n答：阿们。`;
    en = `${en}\nR. Amen.`;
  }
  return { cn: splitInlineResponses(cn), en: splitInlineResponses(en) };
}

function runPr(lang, style = "inherited", part = {}) {
  const bold = part.bold ? ' b="1"' : "";
  const fill = part.color ? `<a:solidFill><a:srgbClr val="${part.color}"/></a:solidFill>` : "";
  if (lang === "cn") {
    const size = style === "hymnCn" ? "5800" : "4800";
    return `<a:rPr lang="zh-CN" altLang="en-US" sz="${size}"${bold} dirty="0">${fill}<a:latin typeface="${CHINESE_FONT}"/><a:ea typeface="${CHINESE_FONT}"/></a:rPr>`;
  }
  return `<a:rPr lang="en-HK" sz="2800"${bold} dirty="0">${fill}<a:latin typeface="EB Garamond"/><a:ea typeface="EB Garamond"/></a:rPr>`;
}

function lineParts(line) {
  if (line && typeof line === "object" && Array.isArray(line.runs)) return line.runs;
  if (line && typeof line === "object") return [{ text: line.text || "", bold: Boolean(line.bold), color: line.color }];
  return [{ text: String(line ?? "") }];
}

function makeParagraphs(text, lang, style = "inherited") {
  const lines = Array.isArray(text) ? text : String(text || "").split("\n");
  if (!lines.length) lines.push("");
  return lines
    .map((line) => {
      const runs = lineParts(line)
        .map((part) => `<a:r>${runPr(lang, style, part)}<a:t>${escapeXml(part.text || "")}</a:t></a:r>`)
        .join("");
      const align = style === "faithful" ? ' algn="l"' : "";
      const lineSpacing = lang === "en" ? '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>' : "";
      return `<a:p><a:pPr marL="0" indent="0"${align}>${lineSpacing}<a:buNone/></a:pPr>${runs}<a:endParaRPr dirty="0"/></a:p>`;
    })
    .join("");
}

function txBody(text, lang, style) {
  const fit = style === "singleCn" || style === "faithful" || style === "hymnCn" ? "<a:noAutofit/>" : '<a:normAutofit fontScale="90000" lnSpcReduction="12000"/>';
  return `<p:txBody><a:bodyPr>${fit}</a:bodyPr><a:lstStyle/>${makeParagraphs(text, lang, style)}</p:txBody>`;
}

// ── 按模板保真填充 ──
// 从模板形状抽取「参考格式」（bodyPr、首个 run 的字号/字体/颜色、首段对齐），
// 用它来渲染我们填入的文字 → 字体/字号/颜色与模板一致，溢出由 autofit 自动缩放。
function extractTemplateStyle(slideXml, shapeId) {
  const sp = slideXml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${shapeId}"[\\s\\S]*?</p:sp>`));
  if (!sp) return {};
  return {
    bodyPr: sp[0].match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/)?.[0] || "<a:bodyPr/>",
    rpr: sp[0].match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || "",
    pPr: sp[0].match(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] || "",
  };
}

function refRprParts(rpr) {
  return {
    sz: rpr.match(/\bsz="(\d+)"/)?.[1] || "",
    typeface: rpr.match(/typeface="([^"]+)"/)?.[1] || "",
    fill: rpr.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0] || "",
  };
}

function buildRunRpr(ref, lang, part) {
  const langAttr = lang === "cn" ? 'lang="zh-CN" altLang="en-US"' : 'lang="en-HK" altLang="zh-CN"';
  const sz = ref.sz ? ` sz="${ref.sz}"` : "";
  const bold = part.bold ? ' b="1"' : "";
  const fill = part.color ? `<a:solidFill><a:srgbClr val="${part.color}"/></a:solidFill>` : ref.fill;
  // 模板该形状显式指定了字体就沿用；否则【不写字体】，让它继承占位符/主题的东亚字体（楷体）。
  // 之前强写 DFKai-SB 会盖掉继承的楷体，导致在没装 DFKai-SB 的环境里回退成难看的黑体。
  const tf = ref.typeface ? `<a:latin typeface="${ref.typeface}"/><a:ea typeface="${ref.typeface}"/>` : "";
  return `<a:rPr ${langAttr}${sz}${bold} dirty="0">${fill}${tf}</a:rPr>`;
}

// 把 bodyPr 的 autofit 设为 normAutofit，并烤入一个 fontScale（百分数*1000）。
// Keynote 不会自动应用 bare normAutofit，所以必须把缩放比例预先写进去，文字才不会溢出。
function setAutofit(bodyPr, scalePct) {
  const fit = scalePct >= 100 ? "<a:normAutofit/>" : `<a:normAutofit fontScale="${Math.round(scalePct * 1000)}" lnSpcReduction="${scalePct < 80 ? 10000 : 0}"/>`;
  if (/<a:(normAutofit|noAutofit|spAutoFit)\b/.test(bodyPr)) {
    return bodyPr.replace(/<a:(normAutofit|noAutofit|spAutoFit)\b[^>]*\/>/, fit);
  }
  if (/<a:bodyPr\b[^>]*\/>/.test(bodyPr)) {
    return bodyPr.replace(/(<a:bodyPr\b[^>]*?)\s*\/>/, `$1>${fit}</a:bodyPr>`);
  }
  return bodyPr.replace(/(<a:bodyPr\b[^>]*>)/, `$1${fit}`);
}

function lineText(line) {
  if (Array.isArray(line)) return line.map(lineText).join("");
  if (line && typeof line === "object") return Array.isArray(line.runs) ? line.runs.map((r) => r.text || "").join("") : line.text || "";
  return String(line ?? "");
}

// 估算文字「行数」：中文按字宽 1、英文/数字按 0.5；按每行可容字数换算所占行数。
function estimateRows(text, lang, charsPerLine) {
  const lines = Array.isArray(text) ? text.map(lineText) : String(text || "").split("\n");
  let rows = 0;
  for (const line of lines) {
    let len = 0;
    for (const ch of String(line)) len += /[　-〿㐀-鿿＀-￯]/.test(ch) ? 1 : 0.5;
    rows += Math.max(1, Math.ceil(len / charsPerLine));
  }
  return rows;
}

// 解析某源页某形状的文本框尺寸（EMU）：优先用形状自身 xfrm，否则查版式占位符，再否则用默认。
function resolveShapeBox(sourceNum, shapeId) {
  const slide = sourceSlides.get(sourceNum);
  const slideXml = slide?.xml || "";
  const sp = slideXml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?id="${shapeId}"[\\s\\S]*?</p:sp>`)) || [];
  const own = (sp[0] || "").match(/<a:ext cx="(\d+)" cy="(\d+)"/);
  if (own) return { cx: +own[1], cy: +own[2] };
  const phIdx = (sp[0] || "").match(/<p:ph[^>]*idx="(\d+)"/)?.[1];
  const phType = (sp[0] || "").match(/<p:ph[^>]*type="(\w+)"/)?.[1];
  const layoutFile = (slide?.rels || "").match(/slideLayout\d+\.xml/)?.[0];
  if (layoutFile) {
    try {
      const lx = readFileSync(join(packageDir, "ppt/slideLayouts", layoutFile), "utf8");
      for (const lsp of lx.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
        const lIdx = lsp[0].match(/<p:ph[^>]*idx="(\d+)"/)?.[1];
        const lType = lsp[0].match(/<p:ph[^>]*type="(\w+)"/)?.[1];
        const matchIdx = phIdx && lIdx === phIdx;
        const matchType = !phIdx && phType && lType === phType;
        const matchBody = !phIdx && !phType && lIdx && !lType;
        if (matchIdx || matchType || matchBody) {
          const ext = lsp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"/);
          if (ext) return { cx: +ext[1], cy: +ext[2] };
        }
      }
    } catch {}
  }
  return { cx: 8000000, cy: 5300000 }; // 兜底：约占大半页
}

function fillBody(sourceNum, shapeId, text, lang) {
  const sourceXml = sourceSlides.get(sourceNum)?.xml || "";
  const tpl = extractTemplateStyle(sourceXml, shapeId);
  const ref = refRprParts(tpl.rpr || "");
  const algn = (tpl.pPr || "").match(/algn="(\w+)"/)?.[1];

  // 计算 fontScale，保证不溢出（Keynote 不会自动缩放）。
  const sz = Number(ref.sz) || (lang === "cn" ? 5400 : 3200);
  const box = resolveShapeBox(sourceNum, shapeId);
  const fontEmu = (sz / 100) * 12700;
  const charW = lang === "cn" ? fontEmu * 1.0 : fontEmu * 0.55;
  const lineH = fontEmu * 1.22;
  const charsPerLine = Math.max(1, Math.floor((box.cx * 0.95) / charW));
  const maxLines = Math.max(1, Math.floor((box.cy * 0.96) / lineH));
  const rows = estimateRows(text, lang, charsPerLine);
  const scalePct = rows <= maxLines ? 100 : Math.max(45, Math.round(Math.sqrt(maxLines / rows) * 100));
  const bodyPr = setAutofit(tpl.bodyPr || "<a:bodyPr/>", scalePct);

  const lines = Array.isArray(text) ? text : String(text || "").split("\n");
  if (!lines.length) lines.push("");
  const paras = lines
    .map((line) => {
      const runs = lineParts(line)
        .map((part) => `<a:r>${buildRunRpr(ref, lang, part)}<a:t>${escapeXml(part.text || "")}</a:t></a:r>`)
        .join("");
      const a = algn ? ` algn="${algn}"` : "";
      const lineSpacing = lang === "en" ? '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>' : "";
      return `<a:p><a:pPr marL="0" indent="0"${a}>${lineSpacing}<a:buNone/></a:pPr>${runs}<a:endParaRPr dirty="0"/></a:p>`;
    })
    .join("");
  return `<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody>`;
}

function titleBody(cn, en, order = "cn-en") {
  const first = order === "en-cn" ? [en, "en-HK"] : [cn, "zh-TW"];
  const second = order === "en-cn" ? [cn, "zh-TW"] : [en, "en-US"];
  return `<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="${first[1]}" altLang="en-US" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>${escapeXml(first[0])}</a:t></a:r><a:r><a:rPr lang="zh-TW" altLang="en-US" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>　</a:t></a:r><a:r><a:rPr lang="${second[1]}" altLang="zh-TW" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>${escapeXml(second[0])}</a:t></a:r><a:endParaRPr dirty="0"/></a:p></p:txBody>`;
}

function coverTitleBody() {
  return `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPts val="5200"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="4800" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>${escapeXml(meta.titleCn)}</a:t></a:r><a:br><a:rPr lang="en-HK" altLang="zh-TW" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr></a:br><a:r><a:rPr lang="en-US" altLang="zh-CN" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>${escapeXml(meta.titleEn)}</a:t></a:r><a:endParaRPr lang="en-HK" sz="4400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:endParaRPr></a:p></p:txBody>`;
}

function replaceShapeTxBody(slideXml, shapeId, replacement) {
  const re = new RegExp(`(<p:sp>[\\s\\S]*?<p:cNvPr id="${shapeId}"[\\s\\S]*?<p:txBody>)[\\s\\S]*?(<\\/p:txBody>[\\s\\S]*?<\\/p:sp>)`);
  if (!re.test(slideXml)) throw new Error(`Could not find shape ${shapeId}`);
  return slideXml.replace(re, (_m, start, end) => {
    const open = start.replace(/<p:txBody>$/, "");
    return `${open}${replacement}${end.replace(/^<\/p:txBody>/, "")}`;
  });
}

const POSTURE_INSTRUCTION = /^\s*(请?(站立|坐下|跪下|起立)|全体起立|Stand|Sit|Kneel|Please stand|Please sit|Please kneel)\s*$/;

function clearExtraTextShapes(slideXml, keepIds) {
  return slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (block) => {
    const id = block.match(/<p:cNvPr id="([^"]+)"/)?.[1];
    if (!id || keepIds.has(id) || !block.includes("<p:txBody>") || !block.includes("<a:t>")) return block;
    // 保留「站立／坐下／跪下」等姿势提示框，不清空。
    const text = [...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join("").trim();
    if (POSTURE_INSTRUCTION.test(text)) return block;
    return block.replace(
      /<p:txBody>[\s\S]*?<\/p:txBody>/,
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr dirty="0"/></a:p></p:txBody>'
    );
  });
}

function normalizeRPrXml(match, attrs = "", inner = "", lang = "cn", selfClosing = false) {
  const cleanAttrs = attrs
    .replace(/\s(?:lang|altLang|sz)="[^"]*"/g, "")
    .trim();
  const baseAttrs =
    lang === "cn"
      ? `lang="zh-CN" altLang="en-US" sz="4800"`
      : `lang="en-HK" altLang="zh-CN" sz="2800"`;
  const typeface = lang === "cn" ? CHINESE_FONT : "EB Garamond";
  const mergedAttrs = cleanAttrs ? `${baseAttrs} ${cleanAttrs}` : baseAttrs;
  const content = selfClosing ? "" : inner.replace(/<a:(?:latin|ea|cs|sym)\b[^>]*\/>/g, "");
  return `<a:rPr ${mergedAttrs}>${content}<a:latin typeface="${typeface}"/><a:ea typeface="${typeface}"/></a:rPr>`;
}

function normalizeParagraphSpacing(block, lang) {
  if (lang !== "en") return block;
  return block
    .replace(/<a:pPr\b([^>]*)\/>/g, (_m, attrs) => `<a:pPr${attrs}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buNone/></a:pPr>`)
    .replace(/<a:pPr\b([^>]*)>([\s\S]*?)<\/a:pPr>/g, (_m, attrs, inner) => {
      const cleanInner = inner.replace(/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/g, "");
      return `<a:pPr${attrs}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>${cleanInner}</a:pPr>`;
    });
}

function normalizeBodyTextStyles(slideXml) {
  return slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (block) => {
    const id = block.match(/<p:cNvPr id="([^"]+)"/)?.[1];
    const lang = id === "3" ? "cn" : id === "4" ? "en" : "";
    if (!lang || !block.includes("<p:txBody>")) return block;
    let out = block.replace(/<a:rPr\b([^>]*)\/>/g, (match, attrs) => normalizeRPrXml(match, attrs, "", lang, true));
    out = out.replace(/<a:rPr\b([^>]*)>([\s\S]*?)<\/a:rPr>/g, (match, attrs, inner) => normalizeRPrXml(match, attrs, inner, lang, false));
    return normalizeParagraphSpacing(out, lang);
  });
}

function renderSpec(spec) {
  const source = sourceSlides.get(spec.source);
  if (!source) throw new Error(`Missing source slide ${spec.source}`);
  let xml = source.xml;
  // 保留页：原样输出，不改字体/字号/间距，确保与模板完全一致。
  if (spec.kind === "preserve") return { xml: source.xml, rels: source.rels };
  if (spec.kind === "cover") {
    xml = replaceShapeTxBody(xml, "4", coverTitleBody());
  } else if (spec.kind === "single-cn") {
    xml = replaceShapeTxBody(xml, "2", titleBody(spec.titleCn, spec.titleEn));
    xml = replaceShapeTxBody(xml, "3", fillBody(spec.source,"3", spec.cnRich || spec.cn, "cn"));
  } else if (spec.kind === "two-col") {
    xml = replaceShapeTxBody(xml, "2", titleBody(spec.titleCn, spec.titleEn));
    xml = replaceShapeTxBody(xml, "3", fillBody(spec.source,"3", spec.cnRich || richifyResponses(spec.cn || "", "cn"), "cn"));
    xml = replaceShapeTxBody(xml, "4", fillBody(spec.source,"4", spec.enRich || richifyResponses(spec.en || "", "en"), "en"));
    xml = clearExtraTextShapes(xml, new Set(["2", "3", "4", ...(spec.keepIds || [])]));
  } else if (spec.kind === "faithful") {
    xml = replaceShapeTxBody(xml, "2", titleBody("信友祷文", "Prayers of the Faithful", "en-cn"));
    xml = replaceShapeTxBody(xml, "3", fillBody(spec.source,"3", spec.cnRich || richifyResponses(spec.cn || "", "cn"), "cn"));
    xml = clearExtraTextShapes(xml, new Set(["2", "3", ...(spec.keepIds || [])]));
  } else if (spec.kind === "cn-keep-en") {
    xml = replaceShapeTxBody(xml, "2", titleBody(spec.titleCn, spec.titleEn));
    xml = replaceShapeTxBody(xml, "3", fillBody(spec.source,"3", spec.cn || "", "cn"));
  } else if (spec.kind === "blank-notice") {
    // 文档里没有的内容（堂区报告、每周圣歌等）：清空旧内容，放红色「请手动填入」提示。
    // 保留标题框（shape 2 或 10）与提示框，其余清空；「站立/坐下」由 clearExtraTextShapes 保留。
    const shape = spec.noticeShape || "3";
    const notice = `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="3600" b="1" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="${CHINESE_FONT}"/><a:ea typeface="${CHINESE_FONT}"/></a:rPr><a:t>${escapeXml(spec.notice)}</a:t></a:r></a:p></p:txBody>`;
    xml = replaceShapeTxBody(xml, shape, notice);
    xml = clearExtraTextShapes(xml, new Set(["2", "10", shape, ...(spec.keepShapes || [])]));
  }
  return { xml, rels: source.rels };
}

function usePattern(pattern, index) {
  return pattern[Math.min(index, pattern.length - 1)];
}

function singleCnSpecs(title, pattern, titleEn, text, mode = "singleCn") {
  return chunkText(formatPoemText(text), "cn", mode).map((cn, i) => ({
    kind: "single-cn",
    source: usePattern(pattern, i),
    titleCn: title,
    titleEn,
    cn,
  }));
}

function manualLine(text) {
  return [{ text, bold: true, color: "C00000" }];
}

function manualSingleCnSpecs(title, pattern, titleEn, labels) {
  return labels.map((label, i) => ({
    kind: "single-cn",
    source: usePattern(pattern, i),
    titleCn: title,
    titleEn,
    cnRich: manualLine(label),
    cnStyle: "hymnCn",
  }));
}

function manualTwoColSpec(title, source, titleEn, cnLabel, enLabel) {
  return {
    kind: "two-col",
    source,
    titleCn: title,
    titleEn,
    cnRich: manualLine(cnLabel),
    enRich: manualLine(enLabel),
  };
}

function twoColSpecs(title, pattern, titleEn, cn, en, cnMode = "twoCn", enMode = "twoEn", pairMode = "auto") {
  const pairs = pairMode === "global" ? chunkBilingualGlobal(cn, en, cnMode, enMode) : chunkBilingual(cn, en, cnMode, enMode, pairMode);
  return pairs.map((pair, i) => ({
    kind: "two-col",
    source: usePattern(pattern, i),
    titleCn: title,
    titleEn,
    cn: pair.cn,
    en: pair.en,
  }));
}

function twoColSpecsMin(title, pattern, titleEn, cn, en, minCount, cnMode = "twoCn", enMode = "twoEn", pairMode = "global") {
  const pairs = pairMode === "global" ? chunkBilingualGlobal(cn, en, cnMode, enMode) : chunkBilingual(cn, en, cnMode, enMode, pairMode);
  while (pairs.length < minCount) {
    let largest = -1;
    let largestUnits = 0;
    for (let i = 0; i < pairs.length; i += 1) {
      const size = units(pairs[i].cn, "cn") + units(pairs[i].en, "en") / 3;
      if (size > largestUnits) {
        largest = i;
        largestUnits = size;
      }
    }
    if (largest < 0 || largestUnits < 18) break;
    const item = pairs[largest];
    const cnParts = units(item.cn, "cn") > 16 ? splitOnce(item.cn, "cn") : [item.cn, ""];
    const enParts = units(item.en, "en") > 28 ? splitOnce(item.en, "en") : [item.en, ""];
    if (cnParts.length < 2 && enParts.length < 2) break;
    pairs.splice(
      largest,
      1,
      { cn: cnParts[0] || "", en: enParts[0] || "" },
      { cn: cnParts[1] || "", en: enParts[1] || "" }
    );
  }
  return pairs.map((pair, i) => ({
    kind: "two-col",
    source: usePattern(pattern, i),
    titleCn: title,
      titleEn,
      cn: pair.cn,
      en: pair.en,
  }));
}

function lineCount(text) {
  return String(text || "")
    .split("\n")
    .filter((line) => line.trim()).length;
}

function appendText(current, extra) {
  const left = String(current || "").trim();
  const right = String(extra || "").trim();
  if (!left) return right;
  if (!right) return left;
  return `${left}\n${right}`;
}

function isShortResponse(text, lang) {
  const clean = String(text || "").trim();
  if (!clean) return false;
  if (lang === "cn") return /^(答|众)[:：][^\n]{0,18}$/u.test(clean);
  return /^(R\.\s*[^\n]{0,84}|Amen\.?)$/i.test(clean);
}

function canAppendChunk(current, extra, lang, mode, slack = 1.18) {
  if (!String(extra || "").trim()) return true;
  const merged = appendText(current, extra);
  const config = chunkConfig[mode];
  return units(merged, lang) <= config.max * slack && lineCount(merged) <= config.maxLines + 1;
}

function mergeShortResponsePairs(pairs, cnMode = "twoCn", enMode = "twoEn") {
  const out = [];
  for (const pair of pairs) {
    const cn = String(pair.cn || "").trim();
    const en = String(pair.en || "").trim();
    const tail = isShortResponse(cn, "cn") || isShortResponse(en, "en");
    const prev = out[out.length - 1];
    if (
      tail &&
      prev &&
      canAppendChunk(prev.cn, cn, "cn", cnMode, 1.24) &&
      canAppendChunk(prev.en, en, "en", enMode, 1.18)
    ) {
      prev.cn = appendText(prev.cn, cn);
      prev.en = appendText(prev.en, en);
    } else {
      out.push({ cn, en });
    }
  }
  return out;
}

function isLiturgicalEnding(line, lang, kind) {
  const clean = String(line || "").trim();
  if (lang === "cn") {
    return kind === "gospel" ? /^——基督的福音。$/u.test(clean) : /^——上主的圣言。$/u.test(clean);
  }
  return kind === "gospel" ? /^The Gospel of the Lord\.$/i.test(clean) : /^The word of the Lord\.$/i.test(clean);
}

function splitLiturgicalTail(text, lang, kind) {
  const lines = splitInlineResponses(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const tail = [];
  while (lines.length) {
    const last = lines[lines.length - 1];
    const isTail = lang === "cn" ? isCnResponse(last) || isLiturgicalEnding(last, lang, kind) : isEnResponse(last) || isLiturgicalEnding(last, lang, kind);
    if (!isTail) break;
    tail.unshift(lines.pop());
  }
  return { body: lines.join("\n"), tail: tail.join("\n") };
}

function prepareReadingAlignment(text, lang) {
  const clean = splitInlineResponses(text);
  if (lang !== "cn") return clean;
  return clean.replace(/。(?=因为)/g, "。\n");
}

function appendTailPair(pairs, cnTail, enTail, cnMode, enMode) {
  const cn = String(cnTail || "").trim();
  const en = String(enTail || "").trim();
  if (!cn && !en) return mergeShortResponsePairs(pairs, cnMode, enMode);
  const prev = pairs[pairs.length - 1];
  if (
    prev &&
    canAppendChunk(prev.cn, cn, "cn", cnMode, 1.42) &&
    canAppendChunk(prev.en, en, "en", enMode, 1.16)
  ) {
    prev.cn = appendText(prev.cn, cn);
    prev.en = appendText(prev.en, en);
  } else {
    pairs.push({ cn, en });
  }
  return mergeShortResponsePairs(pairs, cnMode, enMode);
}

function pairsToTwoColSpecs(title, pattern, titleEn, pairs, cnMode = "twoCn", enMode = "twoEn") {
  return pairs
    .flatMap((pair) => alignPair(pair.cn || "", pair.en || "", cnMode, enMode))
    .map((pair, i) => ({
      kind: "two-col",
      source: usePattern(pattern, i),
      titleCn: title,
      titleEn,
      cn: pair.cn,
      en: pair.en,
    }));
}

function prayerSpecs(title, pattern, titleEn, prayer, minCount = 0, keepIdsBySource = {}) {
  let pairs = chunkBilingualProportional(prayer.cn, prayer.en, "twoCn", "twoEn");
  if (minCount) pairs = growPairsTo(pairs, minCount);
  pairs = mergeShortResponsePairs(pairs, "twoCn", "twoEn");
  return pairs.map((pair, i) => {
    const source = pattern.length > 1 && i === pairs.length - 1 ? pattern[pattern.length - 1] : usePattern(pattern, i);
    return {
      kind: "two-col",
      source,
      titleCn: title,
      titleEn,
      cn: pair.cn,
      en: pair.en,
      keepIds: keepIdsBySource[source] || [],
    };
  });
}

function offeringSpecs() {
  const fixed = [
    {
      cn: "各位兄弟姊妹，请你们祈祷，望全能的天主圣父，收纳我和你们共同奉献的圣祭。",
      en: "Pray, brethren (brothers and sisters), that my sacrifice and yours may be acceptable to God, the almighty Father.",
    },
    {
      cn: "望上主从你的手中，收纳这个圣祭，为赞美并光荣祂的圣名，也为我们和祂整个圣教会的益处。",
      en: "May the Lord accept the sacrifice at your hands for the praise and glory of his name, for our good and the good of all his holy Church.",
    },
  ].map((pair, i) => ({
    kind: "two-col",
    source: usePattern([84, 85], i),
    titleCn: "献礼经",
    titleEn: "Prayer over the Offerings",
    cn: pair.cn,
    en: pair.en,
  }));
  const prayerOverOfferings = normalizePrayer("献礼经");
  return [
    ...fixed,
    ...prayerSpecs("献礼经", [86], "Prayer over the Offerings", prayerOverOfferings, 1),
  ];
}

function growPairsTo(pairs, minCount) {
  while (pairs.length < minCount) {
    let largest = -1;
    let largestUnits = 0;
    for (let i = 0; i < pairs.length; i += 1) {
      const size = units(pairs[i].cn, "cn") + units(pairs[i].en, "en") / 3;
      if (size > largestUnits) {
        largest = i;
        largestUnits = size;
      }
    }
    if (largest < 0 || largestUnits < 18) break;
    const item = pairs[largest];
    const cnParts = units(item.cn, "cn") > 16 ? splitOnce(item.cn, "cn") : [item.cn, ""];
    const enParts = units(item.en, "en") > 28 ? splitOnce(item.en, "en") : [item.en, ""];
    pairs.splice(
      largest,
      1,
      { cn: cnParts[0] || "", en: enParts[0] || "" },
      { cn: cnParts[1] || "", en: enParts[1] || "" }
    );
  }
  return pairs;
}

function solemnBlessingSpecs() {
  const cnLines = paragraphsOf(section("隆重降福").cn);
  const enLines = paragraphsOf(section("隆重降福").en);
  const blessingIndex = cnLines.findIndex((line) => /求你降福|渴求的恩惠/.test(line));
  const finalBlessingIndex = cnLines.findIndex((line) => /愿全能天主/.test(line));

  const introLines = cnLines.slice(0, blessingIndex >= 0 ? blessingIndex : 0);
  const blessingCn = blessingIndex >= 0 ? cnLines[blessingIndex] : "";
  const closingLines = cnLines.slice(finalBlessingIndex > blessingIndex ? blessingIndex + 1 : cnLines.length);
  const blessingEn = enLines.find((line) => /^Bless,\s*O Lord\b/i.test(line)) || enLines[0] || "";

  const specs = [];
  if (introLines.length) {
    specs.push(
      ...twoColSpecs("隆重降福", [246], "Solemn Blessing", introLines.join("\n"), "", "dialogueCn", "twoEn", "global")
    );
  }
  if (blessingCn || blessingEn) {
    specs.push(
      ...twoColSpecs("隆重降福", [247, 248], "Solemn Blessing", blessingCn, blessingEn, "twoCn", "twoEn", "global")
    );
  }
  if (closingLines.length) {
    specs.push(
      ...twoColSpecs("隆重降福", [250], "Solemn Blessing", closingLines.join("\n"), "", "dialogueCn", "twoEn", "global")
    );
  }
  return specs;
}

function boldReadingIntro(line) {
  const normalized = stripReadingCitation(line);
  return { text: normalized, bold: true };
}

function boldEnglishIntro(line) {
  return { text: String(line || "").trim(), bold: true };
}

function introBlock(items) {
  const cleaned = items.filter(Boolean);
  return cleaned.length ? [...cleaned, ""] : [];
}

function normalizeReading(title) {
  const sec = section(title);
  let cnLines = paragraphsOf(sec.cn);
  let enLines = paragraphsOf(sec.en);
  if (title === "读经二") {
    const cnAcclamation = cnLines.findIndex((line) => /^(阿肋路亚|答[:：]\s*阿肋路亚|福音前欢呼)/.test(line));
    if (cnAcclamation >= 0) cnLines = cnLines.slice(0, cnAcclamation);
    const enAcclamation = enLines.findIndex((line) => /^Alleluia\b/i.test(line));
    if (enAcclamation >= 0) enLines = enLines.slice(0, enAcclamation);
  }
  let theme = "";
  let intro = "";
  const cnBody = [];
  for (const line of cnLines) {
    const splitCombined = line.match(/^(（[^）]+）|\([^)]*\))\s*(恭读.+)$/);
    if (splitCombined) {
      if (!theme) theme = splitCombined[1];
      if (!intro) intro = splitCombined[2];
      continue;
    }
    if (!theme && /^[（(].+[）)]$/.test(line)) {
      theme = line;
      continue;
    }
    if (!intro && /^恭读/.test(line)) {
      intro = line;
      continue;
    }
    cnBody.push(line);
  }

  const enIntro = [];
  let enBodyStart = 0;
  if (/^A reading from\b/i.test(enLines[0] || "")) {
    enIntro.push(enLines[0]);
    if (enLines[1] && /^([1-3]\s*)?[A-Z][A-Za-z]+\s+\d/.test(enLines[1])) {
      enBodyStart = 2;
    } else {
      enBodyStart = 1;
    }
  }
  const enBody = enLines.slice(enBodyStart);
  let cnBodyText = cnBody.join("\n");
  let enBodyText = enBody.join("\n");
  if (/上主的圣言/.test(cnBodyText) && !/答[:：]\s*感谢天主/.test(cnBodyText)) cnBodyText = `${cnBodyText}\n答：感谢天主。`;
  if (/The Word of the Lord/i.test(enBodyText) && !/Thanks be to God/i.test(enBodyText)) enBodyText = `${enBodyText}\nR. Thanks be to God.`;

  return {
    cnIntro: [theme, intro ? boldReadingIntro(intro) : ""].filter(Boolean),
    enIntro: enIntro.map(boldEnglishIntro),
    cnBody: splitInlineResponses(cnBodyText),
    enBody: splitInlineResponses(enBodyText),
  };
}

function readingSpecs(title, pattern, titleEn) {
  const reading = normalizeReading(title);
  const cnTail = splitLiturgicalTail(prepareReadingAlignment(reading.cnBody, "cn"), "cn", "reading");
  const enTail = splitLiturgicalTail(prepareReadingAlignment(reading.enBody, "en"), "en", "reading");
  const pairs = appendTailPair(chunkBilingualProgress(cnTail.body, enTail.body, "twoCn", "twoEn"), cnTail.tail, enTail.tail, "twoCn", "twoEn");
  const firstPair = pairs.shift() || { cn: "", en: "" };
  const specs = [
    {
      kind: "two-col",
      source: pattern[0],
      titleCn: title,
      titleEn,
      cnRich: [...introBlock(reading.cnIntro.slice(-1)), ...richifyResponses(firstPair.cn, "cn")],
      enRich: [...introBlock(reading.enIntro), ...richifyResponses(firstPair.en, "en")],
    },
  ];
  specs.push(
    ...pairs.map((pair, i) => ({
      kind: "two-col",
      source: usePattern(pattern.slice(1).length ? pattern.slice(1) : pattern, i),
      titleCn: title,
      titleEn,
      cn: pair.cn,
      en: pair.en,
    }))
  );
  return specs;
}

function gospelSpecs() {
  const sec = section("福音");
  const cnLines = paragraphsOf(sec.cn);
  const enLines = paragraphsOf(sec.en);
  const theme = /^[（(]/.test(cnLines[0] || "") ? cnLines[0] : "";
  const introIndex = cnLines.findIndex((line) => /^恭读/.test(line));
  const intro = introIndex >= 0 ? cnLines[introIndex] : "";
  const cnBody = cnLines.slice(introIndex >= 0 ? introIndex + 1 : theme ? 1 : 0).join("\n");
  const enIntro = [];
  let enBodyStart = 0;
  if (/^A reading from\b/i.test(enLines[0] || "")) {
    enIntro.push(boldEnglishIntro(enLines[0]));
    if (enLines[1] && /^([1-3]\s*)?[A-Z][A-Za-z]+\s+\d/.test(enLines[1])) {
      enBodyStart = 2;
    } else {
      enBodyStart = 1;
    }
  }
  let cnBodyText = cnBody;
  let enBody = enLines.slice(enBodyStart).join("\n");
  if (/基督的福音/.test(cnBodyText) && !/答[:：]\s*基督/.test(cnBodyText)) cnBodyText = `${cnBodyText}\n答：基督，我们赞美你。`;
  if (/The Gospel of the Lord/i.test(enBody) && !/Praise to you/i.test(enBody)) enBody = `${enBody}\nR. Praise to you Lord, Jesus Christ.`;
  const cnTail = splitLiturgicalTail(cnBodyText, "cn", "gospel");
  const enTail = splitLiturgicalTail(enBody, "en", "gospel");
  const pairs = appendTailPair(
    growPairsTo(chunkBilingualProportional(cnTail.body, enTail.body, "gospelCn", "gospelEn"), 7),
    cnTail.tail,
    enTail.tail,
    "gospelCn",
    "gospelEn"
  );
  const firstPair = pairs.shift() || { cn: "", en: "" };
  const introSpec = {
    kind: "two-col",
    source: 50,
    titleCn: "福音",
    titleEn: "Gospel",
    cnRich: [...introBlock([intro ? boldReadingIntro(intro) : ""]), ...richifyResponses(firstPair.cn, "cn")],
    enRich: [...introBlock(enIntro), ...richifyResponses(firstPair.en, "en")],
  };
  return [
    { kind: "preserve", source: 49 },
    introSpec,
    ...pairs.map((pair, i) => ({
      kind: "two-col",
      source: usePattern([51, 52, 53, 54, 55, 56, 57], i),
      titleCn: "福音",
      titleEn: "Gospel",
      cn: pair.cn,
      en: pair.en,
    })),
  ];
}

function faithfulSpecs() {
  const chunks = paragraphsOf(section("信友祷文").cn)
    .map((paragraph) => {
      let text = paragraph.replace(/^主礼[:：]/, "主祭：");
      if (/^领[:：]/.test(text) && !/答[:：]/.test(text)) text = `${text}\n答：求主俯听我们。`;
      if (/^主祭[:：]/.test(text) && /以上所求/.test(text) && !/答[:：]/.test(text)) text = `${text}\n答：阿们。`;
      return splitInlineResponses(text);
    })
    .flatMap((paragraph) => chunkText(paragraph, "cn", "faithfulCn"));
  const pattern = [67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78];
  return chunks.map((cn, i) => ({ kind: "faithful", source: usePattern(pattern, i), cnRich: richifyFaithful(cn) }));
}

function psalmSpecs() {
  const cnLines = paragraphsOf(dropLines(section("答唱咏").cn, [/^咏/]));
  const enLines = paragraphsOf(dropLines(section("答唱咏").en, [/^Responsorial Psalm/]));
  const count = Math.max(cnLines.length, enLines.length, 1);
  return Array.from({ length: count }, (_v, i) => {
    const cn = splitInlineResponses(cnLines[i] || "");
    const en = splitInlineResponses(enLines[i] || "");
    const isResponse = isCnResponse(cn);
    return {
      kind: "two-col",
      source: usePattern([33, 34, 35, 36, 37, 38, 39, 40, 41, 42], i),
      titleCn: "答唱咏",
      titleEn: "Responsorial Psalm",
      cnRich: richifyResponses(cn, "cn"),
      enRich: richifyResponses(en, "en", isResponse),
    };
  });
}

function acclamationSpecs() {
  const cnLines = paragraphsOf(section("福音前欢呼词").cn).filter(
    (line) => !/^(阿肋路亚\s*)?(若|路|玛|谷|Jn|Lk|Mt|Mk)\s*\d/i.test(line)
  );
  const enLines = paragraphsOf(section("福音前欢呼词").en).filter((line) => !/^(Jn|Lk|Mt|Mk|John|Luke|Matthew|Mark)\s*\d/i.test(line));
  const cnRich = [];
  const enRich = [];
  const count = Math.max(cnLines.length, enLines.length, 1);
  for (let i = 0; i < count; i += 1) {
    const cn = splitInlineResponses(cnLines[i] || "");
    const en = splitInlineResponses(enLines[i] || "");
    const isResponse = isCnResponse(cn);
    cnRich.push(...richifyResponses(cn, "cn"));
    enRich.push(...richifyResponses(en, "en", isResponse));
  }
  return [
    {
      kind: "two-col",
      source: 48,
      titleCn: "福音前欢呼",
      titleEn: "Gospel Acclamation",
      cnRich,
      enRich,
    },
  ];
}

function creedSpecs() {
  const chunks = chunkText(section("信经").cn, "cn", "twoCn");
  const pattern = [57, 58, 59, 60, 61, 62, 63, 64, 65, 66];
  const count = Math.max(chunks.length, pattern.length);
  return Array.from({ length: count }, (_v, i) => ({
    kind: "cn-keep-en",
    source: usePattern(pattern, i),
    titleCn: "信经",
    titleEn: "Creed",
    cn: chunks[i] || "",
  }));
}

const replacements = new Map();
const skipRanges = [];

function replaceRange(start, end, specs) {
  replacements.set(start, specs);
  skipRanges.push([start, end]);
}

function inSkipRange(slide) {
  return skipRanges.some(([start, end]) => slide >= start && slide <= end);
}

// ── 绿色「常年期」模板（264 页）的可变段 → 源页码映射 ──
// 固定段（忏悔礼/信经/颂谢词/感恩经一~四/领圣体礼/堂区报告/降福…）一律保留原页。
replaceRange(1, 1, [{ kind: "cover", source: 1 }]);

// 每周圣歌（进堂咏/奉献咏/领主咏/礼成咏）：歌词每周不同、且不在礼仪文档里 →
// 收成一页红色「请手动填入」提示，避免显示上周的旧歌词。标题框（shape10）和站立提示保留。
replaceRange(2, 7, [{ kind: "blank-notice", source: 2, noticeShape: "8", notice: "进堂咏 · 请手动填入本周歌词" }]);
replaceRange(81, 83, [{ kind: "blank-notice", source: 81, noticeShape: "8", notice: "奉献咏 · 请手动填入本周歌词" }]);
replaceRange(240, 244, [{ kind: "blank-notice", source: 240, noticeShape: "3", notice: "领主咏 / 领圣体歌 · 请手动填入" }]);
replaceRange(260, 264, [{ kind: "blank-notice", source: 260, noticeShape: "8", notice: "礼成咏 · 请手动填入本周歌词" }]);

if (!meta.includeGloria) replaceRange(16, 20, []);

const collect = normalizePrayer("集祷经", { letUsPray: true });
replaceRange(21, 22, prayerSpecs("集祷经", [21, 22], "Collect", collect, 2));
replaceRange(25, 32, readingSpecs("读经一", [25, 26, 27, 28, 29, 30, 31, 32], "First Reading"));

replaceRange(33, 42, psalmSpecs());

replaceRange(43, 47, readingSpecs("读经二", [43, 44, 45, 46, 47], "Second Reading"));

replaceRange(48, 48, acclamationSpecs());

replaceRange(49, 57, gospelSpecs());

replaceRange(67, 78, faithfulSpecs());
replaceRange(84, 86, offeringSpecs());

const prayerAfterCommunion = normalizePrayer("领圣体后经", { letUsPray: true });
replaceRange(245, 246, prayerSpecs("领圣体后经", [245, 246], "Prayer after Communion", prayerAfterCommunion, 2));

// 堂区报告（247 分隔页保留；248-252 是上周的具体报告，文档里没有）→ 收成一页红字「请手动填入」。
replaceRange(248, 252, [
  { kind: "blank-notice", source: 248, noticeShape: "4", notice: "本周堂区报告 · 请在此手动填入" },
]);

const specs = [];
for (let i = 1; i <= sourceSlideCount; i += 1) {
  if (replacements.has(i)) specs.push(...replacements.get(i));
  if (inSkipRange(i)) continue;
  specs.push({ kind: "preserve", source: i });
}

const sourceToOutputSlide = new Map();
for (let i = 0; i < specs.length; i += 1) {
  if (specs[i].source && !sourceToOutputSlide.has(specs[i].source)) sourceToOutputSlide.set(specs[i].source, i + 1);
}

function remapSlideTargets(rels) {
  return String(rels || "").replace(/Target="slide(\d+)\.xml"/g, (match, sourceNo) => {
    const outputNo = sourceToOutputSlide.get(Number(sourceNo));
    return outputNo ? `Target="slide${outputNo}.xml"` : match;
  });
}

writeFileSync(
  join(workspace, "slide-map.json"),
  JSON.stringify(
    specs.map((spec, i) => ({
      slide: i + 1,
      source: spec.source,
      kind: spec.kind,
      titleCn: spec.titleCn || "",
      titleEn: spec.titleEn || "",
    })),
    null,
    2
  )
);

for (let i = 0; i < specs.length; i += 1) {
  const slide = renderSpec(specs[i]);
  writeFileSync(join(packageDir, `ppt/slides/slide${i + 1}.xml`), slide.xml);
  writeFileSync(join(packageDir, `ppt/slides/_rels/slide${i + 1}.xml.rels`), remapSlideTargets(slide.rels));
}

for (const file of readdirSync(join(packageDir, "ppt/slides"))) {
  const m = file.match(/^slide(\d+)\.xml$/);
  if (m && Number(m[1]) > specs.length) unlinkSync(join(packageDir, "ppt/slides", file));
}
for (const file of readdirSync(join(packageDir, "ppt/slides/_rels"))) {
  const m = file.match(/^slide(\d+)\.xml\.rels$/);
  if (m && Number(m[1]) > specs.length) unlinkSync(join(packageDir, "ppt/slides/_rels", file));
}

let presXml = readFileSync(join(packageDir, "ppt/presentation.xml"), "utf8");
const sldIds = specs.map((_slide, i) => `<p:sldId id="${2200 + i}" r:id="rIdSlide${i + 1}"/>`).join("");
presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIds}</p:sldIdLst>`);
writeFileSync(join(packageDir, "ppt/presentation.xml"), presXml);

let relsXml = readFileSync(join(packageDir, "ppt/_rels/presentation.xml.rels"), "utf8");
const nonSlideRels = [...relsXml.matchAll(/<Relationship\b[^>]*\/>/g)]
  .map((m) => m[0])
  .filter((r) => !r.includes('/relationships/slide"'));
const slideRels = specs.map(
  (_slide, i) =>
    `<Relationship Id="rIdSlide${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
);
relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nonSlideRels.join("")}${slideRels.join("")}</Relationships>`;
writeFileSync(join(packageDir, "ppt/_rels/presentation.xml.rels"), relsXml);

let contentTypesXml = readFileSync(join(packageDir, "[Content_Types].xml"), "utf8");
contentTypesXml = contentTypesXml.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "");
const slideOverrides = specs
  .map(
    (_slide, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  )
  .join("");
contentTypesXml = contentTypesXml.replace("</Types>", `${slideOverrides}</Types>`);
writeFileSync(join(packageDir, "[Content_Types].xml"), contentTypesXml);

let appXml = readFileSync(join(packageDir, "docProps/app.xml"), "utf8");
appXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${specs.length}</Slides>`);
writeFileSync(join(packageDir, "docProps/app.xml"), appXml);

// 后处理：① 把保留页里残留的 DFKai-SB 字体名统一换成系统楷体（修复缺字乱码）；
// ② 若用户选了主题色，统一替换棕金/浅米填充。两者覆盖所有幻灯片、版式、母版。
for (const dir of ["ppt/slides", "ppt/slideLayouts", "ppt/slideMasters"]) {
  const dirPath = join(packageDir, dir);
  if (!existsSync(dirPath)) continue;
  for (const file of readdirSync(dirPath)) {
    if (!file.endsWith(".xml")) continue;
    const filePath = join(dirPath, file);
    const original = readFileSync(filePath, "utf8");
    let xml = original.replaceAll(LEGACY_CHINESE_FONT, CHINESE_FONT);
    xml = applyThemeRecolor(xml);
    if (xml !== original) writeFileSync(filePath, xml);
  }
}

// 移除已失效的 DFKai-SB 子集嵌入（不再被任何文字引用）。
removeEmbeddedFont(LEGACY_CHINESE_FONT);

mkdirSync(dirname(finalPptx), { recursive: true });
if (existsSync(finalPptx)) renameSync(finalPptx, finalPptx.replace(/\.pptx$/i, `.previous-${Date.now()}.pptx`));
execFileSync("zip", ["-qr", finalPptx, "."], { cwd: packageDir, stdio: "inherit" });

writeFileSync(
  join(workspace, "full-template-build-audit.txt"),
  [
    "Build mode: preserve full source deck structure and replace only mapped variable content ranges.",
    `Source slides: ${sourceSlideCount}. Output slides: ${specs.length}.`,
    "Preserved fixed pages include Introductory Rite, Gloria when enabled, Liturgy of the Word divider, Homily, Creed, Liturgy of the Eucharist, offertory/communion/recessional hymns, Eucharistic Prayer, Communion Rite, parish announcements, blessing, and after-Mass prayers.",
    "Variable pages are the cover title, Entrance Hymn, Collect, readings, psalm, Gospel Acclamation, Gospel, Prayers of the Faithful, Prayer over the Offerings, Communion Antiphon, and Prayer after Communion.",
    "Bilingual sections are paginated with balanced Chinese/English chunks so both columns stay populated and aligned where possible.",
    "First and second readings use a compact intro slide with the reading theme and bold Chinese proclamation line, then continue with the scripture text.",
    `Gloria pages are ${meta.includeGloria ? "preserved from the fixed template" : "omitted for this build"}.`,
  ].join("\n")
);

console.log(JSON.stringify({ finalPptx, slides: specs.length }, null, 2));
