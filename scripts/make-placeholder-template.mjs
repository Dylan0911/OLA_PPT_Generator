// 把一份「绿色常年期成品 deck」就地改造成【占位符模板】，风格对齐用户给的 template.pptx：
//   · 通用占位符 {{CN}} / {{EN}} / {{TITLE_CN}} / {{TITLE_EN}}（不是 {读经一_CN} 这种分段名）
//   · 每段的【所有】幻灯片都保留，只把可变文字框换成占位符（不折叠成一张原型页、不重新编号）
//   · 固定经文（致候礼/忏悔礼/光荣颂/信经/颂谢词/感恩经/天主经/羔羊颂…）原文不动
//   · word 文档里没有的内容（进堂咏/奉献咏/领主咏/礼成咏/堂区报告）→ 红字「请手动填入」
// 形状约定（绿色 OLA deck）：id2=标题、id3=中文、id4=英文；封面标题在 id4。
//
// 用法: node make-placeholder-template.mjs <source.pptx> <out.pptx>
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

const sourcePptx = process.argv[2];
const outPptx = process.argv[3];
if (!sourcePptx || !outPptx) {
  console.error("用法: node make-placeholder-template.mjs <source.pptx> <out.pptx>");
  process.exit(1);
}

// 来自 word 文档、要换成 {{CN}}/{{EN}} 占位符的可变段（按标题文字识别）。
const VAR_SECTIONS = ["集祷经", "读经一", "读经二", "答唱咏", "福音前欢呼", "福音", "信友祷文", "献礼经", "领圣体后经"];
// word 文档里没有、要红字手动填的段。
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

// 取某 id 形状的完整 <p:sp> 块。
function shapeBlock(xml, id) {
  return xml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${id}"(?:(?!</p:sp>)[\\s\\S])*?</p:sp>`))?.[0] || null;
}
function shapeText(sp) {
  return [...(sp || "").matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join("").trim();
}
function deleteShape(xml, id) {
  const sp = shapeBlock(xml, id);
  return sp ? xml.replace(sp, "") : xml;
}
function deleteAllPics(xml) {
  return xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, "");
}

// 正文占位符的【统一】格式（不沿用各框第一段的杂乱格式 —— 生成 deck 里中文框第一段
// 常是加粗的读经出处、字号 5200/5400/5600 不一、还烤进了各页不同的 fontScale）。
// 全部 {{CN}} 用同一种：中文楷体 5400 常规、左对齐(just)、顶端、normAutofit；{{EN}} 用 3200 左对齐。
const CANON_RPR = {
  cn: '<a:rPr lang="zh-CN" altLang="en-US" sz="5400" dirty="0"/>',
  en: '<a:rPr lang="en-HK" altLang="zh-CN" sz="3200" dirty="0"/>',
};
const CANON_ALGN = { cn: "just", en: "l" };
function putToken(xml, shapeId, token, lang) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp || !/<p:txBody>/.test(sp)) return xml;
  const newTx = `<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="${CANON_ALGN[lang]}"><a:buNone/></a:pPr><a:r>${CANON_RPR[lang]}<a:t>${escapeXml(token)}</a:t></a:r></a:p></p:txBody>`;
  return xml.replace(sp, sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx));
}

// 红字「请手动填入」：保留 bodyPr，文字居中红色。
function putRedNotice(xml, shapeId, label) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp) return xml;
  const bodyPr = sp.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/)?.[0] || '<a:bodyPr rtlCol="0" anchor="ctr"/>';
  const newTx = `<p:txBody>${bodyPr}<a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="3600" b="1" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="${CHINESE_FONT}"/><a:ea typeface="${CHINESE_FONT}"/></a:rPr><a:t>${escapeXml(label)}</a:t></a:r></a:p></p:txBody>`;
  return xml.replace(sp, sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx));
}

// 清空某形状的文字（保留形状本身）。
function clearShape(xml, shapeId) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp || !/<p:txBody>/.test(sp)) return xml;
  const newTx = "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr dirty=\"0\"/></a:p></p:txBody>";
  return xml.replace(sp, sp.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTx));
}

// 列出所有承载文字的形状 id（按出现顺序）。
function textShapeIds(xml) {
  const ids = [];
  for (const sp of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const id = sp[0].match(/<p:cNvPr id="(\d+)"/)?.[1];
    if (id && /<a:t>/.test(sp[0])) ids.push(id);
  }
  return ids;
}

// 封面 id4：把中文标题 run → {{TITLE_CN}}、英文标题 run → {{TITLE_EN}}，其余格式不动。
function coverTitle(xml) {
  const sp = shapeBlock(xml, "4");
  if (!sp) return xml;
  let done = { cn: false, en: false };
  const newSp = sp.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (m, inner) => {
    if (!done.cn && hasChinese(inner)) {
      done.cn = true;
      return "<a:t>{{TITLE_CN}}</a:t>";
    }
    if (!done.en && hasLatin(inner) && !hasChinese(inner)) {
      done.en = true;
      return "<a:t>{{TITLE_EN}}</a:t>";
    }
    return m;
  });
  return xml.replace(sp, newSp);
}

// ── 固定经文页排版修复：英文两端对齐→左对齐；文字溢出→烤入 fontScale 缩放 ──
// （Keynote / 预览不会自动应用 normAutofit，必须把缩放比例烤进去，文字才不会超出页面。）
const BOX_FALLBACK = {
  "3": { cx: 4863060, cy: 5254561, sz: 4800, lang: "cn" }, // 左栏中文
  "4": { cx: 3138026, cy: 5254560, sz: 2800, lang: "en" }, // 右栏英文
};
// 视觉宽度单位：中日韩字 = 1 em，其它（英文/数字/标点）≈ 0.5 em。
function emUnits(text) {
  let u = 0;
  for (const ch of String(text)) u += /[　-鿿＀-￯]/.test(ch) ? 1 : 0.5;
  return u;
}
function rowCount(text, capacityUnits) {
  let rows = 0;
  for (const line of String(text).split("\n")) rows += Math.max(1, Math.ceil(emUnits(line) / capacityUnits));
  return rows;
}
// 解析某内容框的文本框尺寸 + 字号（优先框自身，其次版式占位符，再否则兜底）。
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
          if (ext) {
            cx = cx || +ext[1];
            cy = cy || +ext[2];
          }
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
// 修某固定内容框（id3/id4）：英文左对齐 + 溢出时烤入缩放。占位符/红字框跳过。
function fixFixedBox(xml, rels, shapeId) {
  const sp = shapeBlock(xml, shapeId);
  if (!sp || !/<p:txBody>/.test(sp)) return xml;
  const t = shapeText(sp);
  if (!t || t.includes("{{")) return xml;
  const { cx, cy, sz, lang } = resolveBoxSz(sp, rels, shapeId);
  let nsp = sp;
  if (lang === "en") nsp = nsp.replace(/algn="just"/g, 'algn="l"');
  const em = (sz / 100) * 12700;
  const capacity = Math.max(1, (cx * 0.95) / em);
  const maxLines = Math.max(1, Math.floor((cy * 0.96) / (em * 1.22)));
  const rows = rowCount(t, capacity);
  if (rows > maxLines) {
    const needed = Math.max(45, Math.round(Math.sqrt(maxLines / rows) * 100));
    const existing = Number(nsp.match(/normAutofit fontScale="(\d+)"/)?.[1] || 0) / 1000 || 100;
    const finalScale = Math.min(existing, needed);
    if (finalScale < 100) {
      const bodyPr = nsp.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/)?.[0] || "<a:bodyPr/>";
      nsp = nsp.replace(bodyPr, setAutofit(bodyPr, finalScale));
    }
  }
  return xml.replace(sp, nsp);
}

// ───────────────────────── 解包 ─────────────────────────
const work = "/tmp/make-placeholder-template-work";
rmSync(work, { recursive: true, force: true });
mkdirSync(work, { recursive: true });
execFileSync("unzip", ["-q", sourcePptx, "-d", work]);

const slideDir = join(work, "ppt/slides");
const slideFiles = readdirSync(slideDir)
  .filter((f) => /^slide\d+\.xml$/.test(f))
  .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));

const relsDir = join(slideDir, "_rels");
const report = { tokenized: [], red: [], fixed: 0, reflowed: 0 };

for (const file of slideFiles) {
  const n = Number(file.match(/\d+/)[0]);
  const path = join(slideDir, file);
  let xml = readFileSync(path, "utf8");
  const relsPath = join(relsDir, `${file}.rels`);
  const rels = existsSync(relsPath) ? readFileSync(relsPath, "utf8") : "";

  // 封面（第 1 页）：标题 → {{TITLE_CN}}/{{TITLE_EN}}。
  if (n === 1) {
    xml = coverTitle(xml);
    writeFileSync(path, xml);
    report.tokenized.push("封面");
    continue;
  }

  const title = shapeText(shapeBlock(xml, "2"));

  // word 没有的（歌咏 / 堂区报告）→ 红字手动填。
  const redKey = RED_SECTIONS.find((k) => title.includes(k));
  if (redKey) {
    const contentIds = textShapeIds(xml).filter((id) => {
      if (id === "2") return false; // 标题保留
      const t = shapeText(shapeBlock(xml, id));
      return t && !POSTURE.test(t); // 站立/坐下等姿势提示保留
    });
    // 选定红字提示框：已含「请手动填入」的优先，否则取文字最长的那个。
    let noticeId = contentIds.find((id) => shapeText(shapeBlock(xml, id)).includes("请手动填入"));
    if (!noticeId && contentIds.length) {
      noticeId = contentIds.slice().sort((a, b) => shapeText(shapeBlock(xml, b)).length - shapeText(shapeBlock(xml, a)).length)[0];
    }
    if (!noticeId) {
      report.fixed += 1; // 分隔页等没有可填框 → 原样保留
      continue;
    }
    xml = putRedNotice(xml, noticeId, RED_LABEL[redKey]);
    if (redKey === "堂区报告") {
      // 堂区报告：只留标题 + 一条红字提示；删掉上周残留的日期/时间/地点条目框和图标，页面干净。
      for (const id of contentIds) if (id !== noticeId) xml = deleteShape(xml, id);
      xml = deleteAllPics(xml);
    } else {
      // 歌咏：红字提示 + 清空其它残留内容框（保留图标/装饰）。
      for (const id of contentIds) if (id !== noticeId) xml = clearShape(xml, id);
    }
    writeFileSync(path, xml);
    report.red.push(`${redKey}(slide${n})`);
    continue;
  }

  // 来自 word 的可变段 → id3={{CN}}、id4={{EN}}（仅当该框当前有文字）。
  const varKey = VAR_SECTIONS.find((k) => title.includes(k));
  if (varKey) {
    let touched = false;
    if (shapeText(shapeBlock(xml, "3"))) {
      xml = putToken(xml, "3", "{{CN}}", "cn");
      touched = true;
    }
    if (shapeText(shapeBlock(xml, "4"))) {
      xml = putToken(xml, "4", "{{EN}}", "en");
      touched = true;
    }
    if (touched) {
      writeFileSync(path, xml);
      report.tokenized.push(`${varKey}(slide${n})`);
      continue;
    }
  }

  // 固定经文页：文字内容不动，只修排版（英文左对齐 + 溢出缩放）。
  const fixedXml = fixFixedBox(fixFixedBox(xml, rels, "3"), rels, "4");
  if (fixedXml !== xml) {
    writeFileSync(path, fixedXml);
    report.reflowed += 1;
  }
  report.fixed += 1;
}

// ───────────────────────── 重新打包（结构完全不变，只改了 slide xml） ─────────────────────────
mkdirSync(dirname(outPptx), { recursive: true });
if (existsSync(outPptx)) rmSync(outPptx);
execFileSync("zip", ["-qr", outPptx, "."], { cwd: work });

console.log(
  JSON.stringify(
    {
      out: outPptx,
      slides: slideFiles.length,
      tokenized: report.tokenized.length,
      red: report.red,
      fixedKept: report.fixed,
      reflowed: report.reflowed,
    },
    null,
    2
  )
);
