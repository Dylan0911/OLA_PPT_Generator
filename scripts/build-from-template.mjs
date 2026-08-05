// 占位符模板 → 成品 PPTX。
// 加载「占位符版」模板（每个可变段一张原型页，里面是 {段落_CN}/{段落_EN} 占位符），
// 把占位符替换成文档内容：中英按比例配对、内容长时克隆原型页分页、形状的字体/字号/颜色原样保留、
// 溢出时烤入 fontScale 缩放；文档里没有的内容（歌词/堂区报告）显示红字「请手动填入」。
//
// 用法: node build-from-template.mjs <template.pptx> <workDir> <out.pptx> <sections.json>
// sections.json 形如 server.mjs 产出：{ meta:{titleCn,titleEn,...}, sections:[{titleCn,titleEn,cn,en}], missing:[] }
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

const templatePptx = process.argv[2];
const workDir = process.argv[3];
const outPptx = process.argv[4];
const sectionsPath = process.argv[5] || "";
if (!templatePptx || !workDir || !outPptx) {
  console.error("用法: node build-from-template.mjs <template.pptx> <workDir> <out.pptx> <sections.json>");
  process.exit(1);
}

const payload = sectionsPath ? JSON.parse(readFileSync(sectionsPath, "utf8")) : { meta: {}, sections: [] };
const meta = {
  titleCn: payload?.meta?.titleCn || "常年期主日",
  titleEn: payload?.meta?.titleEn || "",
};
const sectionByTitle = new Map((payload.sections || []).map((s) => [s.titleCn, s]));

const CHINESE_FONT = "楷体";

// 占位符 base → 文档段落标题（server.mjs 的 OUTPUT_SECTIONS 用词）。
const TOKEN_TO_SECTION = {
  集祷经: "集祷经",
  读经一: "读经一",
  答唱咏: "答唱咏",
  读经二: "读经二",
  福音前欢呼: "福音前欢呼词",
  福音: "福音",
  献礼经: "献礼经",
  领圣体后经: "领圣体后经",
  信友祷文: "信友祷文",
};
const TWO_COL = new Set(["集祷经", "读经一", "答唱咏", "读经二", "福音前欢呼", "福音", "献礼经", "领圣体后经"]);
const SINGLE_FILL = new Set(["信友祷文"]);
// 有「恭读…/A reading from…」出处行的段落（需要剥出出处单独成页再配对正文）。
const READINGS = new Set(["读经一", "读经二", "福音"]);
// 经文祷词：人工版在「以上所求…/Through…」结礼处分两页。
const PRAYERS = new Set(["集祷经", "献礼经", "领圣体后经"]);

// 祷词分页（照人工版切法）：正文一页 + 「以上所求…/Through … Amen.」结礼一页；
// 集祷经正文前加主礼邀请「请大家祈祷：」。过长时 alignPair 再按 prayer 配置细分。
function prayerPairs(base, cn, en) {
  const cnFlat = String(cn || "").replace(/\n+/g, "").trim();
  const enFlat = String(en || "").replace(/\n+/g, " ").trim();
  // 中文结礼：优先「以上所求」；有的祷文以「因/靠我们的主（耶稣基督）…」「因你圣子…」收尾
  // —— 只认后半段的最后一处，避免正文里的同款措辞误切。
  let cnIdx = cnFlat.indexOf("以上所求");
  if (cnIdx < 0) {
    const m = [...cnFlat.matchAll(/[因靠](我们的主|你圣子|祢圣子)/g)].pop();
    if (m && m.index >= cnFlat.length * 0.55) cnIdx = m.index;
  }
  // 英文结礼：最后一个句号后的 "Through …"（正文中的小写 through 不算）；
  // 也有祷文用 "Who lives and reigns …" 收尾。
  let enIdx = enFlat.lastIndexOf("Through ");
  if (enIdx < 0) {
    const m = [...enFlat.matchAll(/\bWho lives and reigns\b/g)].pop();
    if (m) enIdx = m.index;
  }
  const cnParts = cnIdx > 8 ? [cnFlat.slice(0, cnIdx).trim(), cnFlat.slice(cnIdx).trim()] : [cnFlat];
  const enParts = enIdx > 8 && /[.!?]\s*$/.test(enFlat.slice(0, enIdx).trim()) ? [enFlat.slice(0, enIdx).trim(), enFlat.slice(enIdx).trim()] : [enFlat];
  if (base === "集祷经" && cnParts[0] && !/^请大家祈祷/.test(cnParts[0])) cnParts[0] = `请大家祈祷：${cnParts[0]}`;
  const pairs = [...alignPair(cnParts[0] || "", enParts[0] || "", "prayerCn", "prayerEn")];
  // 领圣体后经首页：「请大家祈祷：/Let us pray.」独立首行（人工版 s245 格式）。
  if (base === "领圣体后经" && pairs.length && !pairs[0].cnLines) {
    const first = pairs[0];
    pairs[0] = {
      cnLines: [{ text: "请大家祈祷：", bold: false }, ...String(first.cn || "").split("\n").filter(Boolean).map((t) => ({ text: t, bold: false }))],
      enLines: [{ text: "Let us pray.", bold: false }, ...String(first.en || "").split("\n").filter(Boolean).map((t) => ({ text: t, bold: false }))],
    };
  }
  // 献礼经的结礼「以上所求…/Through…」是模板固定页（用户 2026-07-09 放入模板），
  // 文档里的这一句只从正文剥掉、不进 PPT，避免和固定页重复。
  if (base === "献礼经") return pairs;
  // 结礼页（人工版标准）：公式（去掉文档尾部的「阿们。/Amen.」）+ 空行 + 加粗会众答句。
  if (cnParts[1] || enParts[1]) {
    const cnConc = (cnParts[1] || "").replace(/阿们[。.！!]?\s*$/, "").trim();
    const enConc = (enParts[1] || "").replace(/\bR?\.?\s*Amen[.!]?\s*$/i, "").trim();
    pairs.push({
      cnLines: [{ text: cnConc, bold: false }, { text: "", bold: false }, { text: "答：阿们。", bold: true }],
      enLines: [{ text: enConc, bold: false }, { text: "", bold: false }, { text: "R. Amen.", bold: true }],
    });
  }
  return pairs;
}

// 答唱咏结构化配对：文档两侧都是「答句/领句」交替（英文答句 = 重复出现的行）。
// 丢掉开头的「咏89:2-3…」出处行；答句页 ↔ 英文叠句页（补 R. 前缀）；领句组按 verse 配置对齐分页。
// 两侧结构对不上时返回 null（回退比例配对）。
function psalmPairs(cn, en, addR = true) {
  const cnLines = paragraphsOf(cn).filter((l) => !/^[咏詠]\s*[\d：:，,.\-–\s]/.test(l));
  const enLines = paragraphsOf(en);
  if (!cnLines.length || !enLines.length) return null;
  const counts = new Map();
  for (const l of enLines) counts.set(l, (counts.get(l) || 0) + 1);
  const isCnRef = (l) => /^答[:：]/.test(l);
  // 英文答句 = 重复出现的叠句；文档自带「R./R:」前缀时（不重复也）认作答句。
  const isEnRef = (l) => (counts.get(l) || 0) >= 2 || /^R[.．:：]\s*/.test(l);
  const grouped = (lines, isRef) => {
    const out = [];
    for (const l of lines) {
      if (isRef(l)) out.push({ ref: true, text: l });
      else if (out.length && !out[out.length - 1].ref) out[out.length - 1].text += `\n${l}`;
      else out.push({ ref: false, text: l });
    }
    return out;
  };
  const cg = grouped(cnLines, isCnRef);
  const eg = grouped(enLines, isEnRef);
  if (!cg.length || cg.length !== eg.length) return null;
  const pairs = [];
  for (let i = 0; i < cg.length; i += 1) {
    if (cg[i].ref !== eg[i].ref) return null;
    if (cg[i].ref) {
      // 答句页：中英都加粗（人工版格式）。
      const enR = addR && !/^R[.．]\s*/.test(eg[i].text) ? `R. ${eg[i].text}` : eg[i].text;
      pairs.push({
        cn: cg[i].text,
        en: enR,
        cnLines: [{ text: cg[i].text, bold: true }],
        enLines: [{ text: enR, bold: true }],
      });
    } else {
      pairs.push(...mergeSmallPairs(alignPair(cg[i].text.replace(/\n+/g, ""), eg[i].text.replace(/\n+/g, " "), "verseCn", "verseEn")));
    }
  }
  return pairs;
}

// 福音前欢呼：丢掉开头的「阿肋路亚 若14:23」出处行；整段（答/领/答）合成【一页】，
// 「答：阿肋路亚」与英文 Alleluia 叠句加粗（人工版格式）；偏长由缩放兜底。
function acclaimPairs(cn, en) {
  const cnLs = paragraphsOf(cn).filter((l, i) => !(i === 0 && /^阿肋路亚/.test(l) && !/^[答领]/.test(l)));
  const enLs = paragraphsOf(en);
  const counts = new Map();
  for (const l of enLs) counts.set(l, (counts.get(l) || 0) + 1);
  return [
    {
      cnLines: cnLs.map((l) => ({ text: l, bold: /^答[:：]/.test(l) })),
      enLines: enLs.map((l) => ({ text: l, bold: (counts.get(l) || 0) >= 2 || /^Alleluia/i.test(l) })),
    },
  ];
}

// 贪心打包：相邻两页合并后中英都仍在每页容量内就并成一页（页面尽量装满——用户要求）。
// 「——上主的圣言/The word of the Lord」结尾页保持独立（人工版单独一页）。
function mergeSmallPairs(pairs) {
  const cnCap = chunkConfig.twoCn.max;
  const enCap = chunkConfig.twoEn.max;
  const isEnding = (p) => /^——/.test((p.cn || "").trim()) || /^The (word|Gospel) of the Lord/i.test((p.en || "").trim());
  const joinTx = (a, b) => (a && b ? `${a}\n${b}` : a || b || "");
  const out = [];
  for (const p of pairs) {
    const prev = out[out.length - 1];
    if (prev && !isEnding(prev) && !isEnding(p)) {
      const cnMerged = joinTx(prev.cn, p.cn);
      const enMerged = joinTx(prev.en, p.en);
      if (units(cnMerged, "cn") <= cnCap && units(enMerged, "en") <= enCap) {
        prev.cn = cnMerged;
        prev.en = enMerged;
        continue;
      }
    }
    out.push({ ...p });
  }
  return out;
}

// 剥离读经开头：中文引题「（…）」行丢弃（人工版不显示）；「恭读…」行 =中文出处；
// 英文「A reading from…」行(+紧随的裸章节行如 "Jeremiah 20:10-13") =英文出处。
function peelReadingIntro(cn, en) {
  const cnLines = String(cn || "").split("\n");
  const enLines = String(en || "").split("\n");
  const cnCit = [];
  const enCit = [];
  // 出处行去掉章节数字：「恭读列王纪下 4:8-11,14-16a」→「恭读列王纪下」（人工版就不显示数字）。
  const stripRef = (l) => l.replace(/[\s　]*[0-9０-９][0-9０-９:：,，;；.．\-–—ab\s]*$/u, "").trim();
  while (cnLines.length) {
    const l = cnLines[0].trim();
    if (!l) { cnLines.shift(); continue; }
    if (/^[（(][^）)]*[)）]$/.test(l)) { cnLines.shift(); continue; } // 引题：丢弃
    if (/^[✠✝†＋+]?\s*恭[读讀]/.test(l)) { cnCit.push(stripRef(l.replace(/^[✠✝†＋+]\s*/, ""))); cnLines.shift(); continue; }
    break;
  }
  while (enLines.length) {
    const l = enLines[0].trim().replace(/^[✠✝†＋+]\s*/, "");
    if (!l) { enLines.shift(); continue; }
    if (/^A reading from/i.test(l)) { enCit.push(stripRef(l)); enLines.shift(); continue; }
    if (enCit.length && l.length < 45 && /^[1-3]?\s?[A-Z][A-Za-z ]+\d+[:.]\d/.test(l)) { enLines.shift(); continue; } // 裸章节行：丢弃
    break;
  }
  return {
    cnCit: cnCit.join("\n"),
    enCit: enCit.join("\n"),
    cnBody: cnLines.join("\n").trim(),
    enBody: enLines.join("\n").trim(),
  };
}
// 文档里没有的内容 → 红字提示（沿用旧生成器的措辞）。
const RED_LABEL = {
  进堂咏: "进堂咏 · 请手动填入本周歌词",
  奉献咏: "奉献咏 · 请手动填入本周歌词",
  领主咏: "领主咏 / 领圣体歌 · 请手动填入",
  礼成咏: "礼成咏 · 请手动填入本周歌词",
  堂区报告: "本周堂区报告 · 请在此手动填入",
};

// ───────────────────────── 文本配对引擎（复制自 build-liturgy-ppt.mjs 的纯函数） ─────────────────────────
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
    else if (lang === "cn") total += /[ -]/.test(ch) ? 0.48 : 1;
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

// 每页文字量上限 = 统一字号下一页的【自然容量】（不靠缩字）：
//   中文 5000 两栏框（383pt 宽 × 6 行 × 每行 ~7.2 字）≈ 43 → 取 42；
//   英文 3000 右栏框（247pt 宽 × 11 行 × 每行 ~14 字）≈ 154 → 取 145；
//   信友祷文全宽框（666pt 宽 × 6 行 × 每行 ~12.5 字）≈ 75 → 取 72。
// 之前允许超容量再按页缩字（79%-100%），导致抽查字号 25/28/30、44/48/54 忽大忽小；
// 现在页容量守住字号：所有程序填充页恒定中文 50pt / 英文 30pt，宁可多一两页。
// 全卷统一字号：中文 4800 / 英文 2800（= 忏悔礼对话页家族的现有规格，固定页拆页最少）。
// 页容量 = 该字号下的自然容量：中文两栏 ~52 → 取 50；英文右栏 ~181 → 取 170；
// 信友祷文全宽 ~91 → 取 86。超容量一律拆页，绝不缩字。
// maxLines 放大到不起作用（历史遗留的「每页 N 句」限制会把页切稀——3 个短句 30 字就翻页）；
// 分页只由 max（字数容量）决定，页面才装得满。
const chunkConfig = {
  singleCn: { max: 96, sentence: 28, maxLines: 5 },
  faithfulCn: { max: 86, sentence: 78, maxLines: 8 },
  twoCn: { max: 50, sentence: 45, maxLines: 8 },
  twoEn: { max: 170, sentence: 155, maxLines: 12 },
  prayerCn: { max: 50, sentence: 45, maxLines: 8 },
  prayerEn: { max: 170, sentence: 158, maxLines: 12 },
  verseCn: { max: 50, sentence: 45, maxLines: 8 },
  verseEn: { max: 170, sentence: 158, maxLines: 12 },
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
  // 严格守住每页上限（上限=统一字号下的自然容量；超一点也会触发缩字、破坏字号统一）。
  return out.flatMap((chunk) => (units(chunk, lang) > config.max ? splitLong(chunk, config.max, lang) : [chunk]));
}

function paragraphsOf(text) {
  return String(text || "")
    .replace(/\r/g, "")
    .split(/\n+/)
    .map((p) => p.trim())
    .filter(Boolean);
}

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
  // 句末引号/括号归回上一句，避免出现「’ The Gospel of the Lord.」这种引号孤行开头。
  for (let i = 1; i < out.length; i += 1) {
    const m = out[i].match(/^[”’」』）)\]]+\s*/);
    if (m) {
      out[i - 1] += m[0].trim();
      out[i] = out[i].slice(m[0].length);
      if (!out[i]) {
        out.splice(i, 1);
        i -= 1;
      }
    }
  }
  return out;
}

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

// 把过长的句子按次级标点从中间劈开（用于比例切分前的整形）；英文没逗号时退而按空格劈。
function splitPieceBySecondary(piece, lang) {
  const secondary = lang === "cn" ? /[，、；：]/g : /[,;:]/g;
  let marks = [...piece.matchAll(secondary)];
  if (!marks.length && lang === "en") marks = [...piece.matchAll(/\s/g)];
  if (!marks.length) return null;
  const cut = marks[Math.floor((marks.length - 1) / 2)].index + 1;
  const left = piece.slice(0, cut).trim();
  const right = piece.slice(cut).trim();
  return left && right ? [left, right] : null;
}

function splitByProportion(text, lang, parts, cap = Infinity) {
  const t = String(text || "").trim();
  if (!t) return [];
  if (parts <= 1) return [t];
  const sents = sentencePieces(t, lang, parts);
  if (sents.length <= 1) return [t];
  // 整形：把长句按逗号劈到「容量−目标」的粒度以内 —— 组才能贴着目标堆满、又不超容量。
  // （英文句子常 80-90 单位，两句 170+ 超容 → 每组只装得下一句、页面稀。）
  const preTotal = sents.reduce((sum, s) => sum + units(s, lang), 0) || 1;
  const preTarget = preTotal / parts;
  const pieceLimit = Math.max(Math.min(preTarget * 1.15, 2 * (cap - preTarget)), 16);
  const unsplittable = new Set();
  for (let guard = 0; guard < 120; guard += 1) {
    let idx = -1;
    let largest = 0;
    for (let i = 0; i < sents.length; i += 1) {
      const u = units(sents[i], lang);
      if (u > pieceLimit && u > largest && !unsplittable.has(sents[i])) {
        largest = u;
        idx = i;
      }
    }
    if (idx < 0) break;
    const split = splitPieceBySecondary(sents[idx], lang);
    if (!split) {
      unsplittable.add(sents[idx]); // 这句劈不开 → 跳过它继续劈别的（不能整个循环放弃）
      continue;
    }
    sents.splice(idx, 1, ...split);
  }
  // 均衡分组：每个边界【全局】取累计量最接近 k×target 的句间断点（单调递增），
  // 组大小 ≈ target ± 半句 —— 之前的顺序贪心会切出 16 字孤句组、页面忽满忽稀。
  const cum = [];
  let acc = 0;
  for (const s of sents) {
    acc += units(s, lang);
    cum.push(acc);
  }
  const total = acc || 1;
  const target = total / parts;
  // 动态规划最优均分：最小化 Σ(组大小−目标)²，且每组 ≤ cap。
  // （此前的贪心切点在句子粒度上会切出 16 字的孤句稀页——组忽满忽稀。）
  const sizes = sents.map((s) => units(s, lang));
  const n = sizes.length;
  const pre = [0];
  for (const u of sizes) pre.push(pre[pre.length - 1] + u);
  const INF = Infinity;
  const dp = Array.from({ length: parts + 1 }, () => new Array(n + 1).fill(INF));
  const choice = Array.from({ length: parts + 1 }, () => new Array(n + 1).fill(-1));
  dp[0][0] = 0;
  for (let k = 1; k <= parts; k += 1) {
    for (let i = k; i <= n - (parts - k); i += 1) {
      for (let j = k - 1; j < i; j += 1) {
        if (dp[k - 1][j] === INF) continue;
        const g = pre[i] - pre[j];
        if (g > cap) continue;
        const cost = dp[k - 1][j] + (g - target) * (g - target);
        if (cost < dp[k][i]) {
          dp[k][i] = cost;
          choice[k][i] = j;
        }
      }
    }
  }
  const joiner = lang === "cn" ? "" : " ";
  if (dp[parts][n] === INF) {
    // 不可行（存在劈不开的超容整句等）→ 首次适应贪心打包（组数可能多于 parts，由调用方重试）。
    const groups = [];
    let cur = [];
    let acc = 0;
    for (let i = 0; i < n; i += 1) {
      if (cur.length && acc + sizes[i] > cap) {
        groups.push(cur);
        cur = [];
        acc = 0;
      }
      cur.push(sents[i]);
      acc += sizes[i];
    }
    if (cur.length) groups.push(cur);
    return groups.map((g) => g.join(joiner).trim());
  }
  const cuts = [];
  let pos = n;
  for (let k = parts; k >= 1; k -= 1) {
    cuts.push(choice[k][pos]);
    pos = choice[k][pos];
  }
  cuts.reverse(); // parts 个起点（首个为 0）
  const groups = [];
  for (let k = 0; k < parts; k += 1) {
    const from = cuts[k];
    const to = k + 1 < parts ? cuts[k + 1] : n;
    groups.push(sents.slice(from, to));
  }
  return groups.map((g) => g.join(joiner).trim());
}

function alignPair(cnText, enText, cnMode, enMode) {
  const cnT = String(cnText || "").trim();
  const enT = String(enText || "").trim();
  const cnCap = chunkConfig[cnMode].max;
  const enCap = chunkConfig[enMode].max;
  // 页数 = ⌈总量/容量⌉ 的最小可行值（贪心分块数会虚高、页面装不满）；
  // 个别组超容量（缩字禁区）就加一页重切，直到全部在容量内。
  let count = Math.max(cnT ? Math.ceil(units(cnT, "cn") / cnCap) : 0, enT ? Math.ceil(units(enT, "en") / enCap) : 0, 1);
  for (let tries = 0; tries < 12; tries += 1) {
    const cn = splitByProportion(cnT, "cn", count, cnCap);
    const en = splitByProportion(enT, "en", count, enCap);
    if ((cn.length > count || en.length > count) && tries < 11) {
      count = Math.max(cn.length, en.length); // DP 不可行时的贪心回退可能多出组 → 双侧按新组数重齐
      continue;
    }
    const over = cn.some((c) => units(c, "cn") > cnCap) || en.some((e) => units(e, "en") > enCap);
    if (!over || tries === 11) {
      return Array.from({ length: Math.max(cn.length, en.length, 1) }, (_v, i) => ({ cn: cn[i] || "", en: en[i] || "" }));
    }
    count += 1;
  }
  return [{ cn: cnT, en: enT }];
}

// 中英按段落数 + 累计长度比例配对（旧生成器里实际应用的对齐方式）。
function chunkBilingualProportional(cn, en, cnMode = "twoCn", enMode = "twoEn") {
  const cnParas = paragraphsOf(cn);
  const enParas = paragraphsOf(en);
  if (cnParas.length <= 1 || enParas.length <= 1) return alignPair(cn, en, cnMode, enMode);

  const cnTotal = cnParas.reduce((sum, p) => sum + units(p, "cn"), 0) || 1;
  const enTotal = enParas.reduce((sum, p) => sum + units(p, "en"), 0) || 1;
  let enSeen = 0;
  let enIndex = 0;
  let cnSeen = 0;
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

// ───────────────────────── 形状填充（保留模板格式 + 烤入 fontScale 防溢出） ─────────────────────────
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
    bold: /\bb="1"/.test(rpr),
    fill: rpr.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0] || "",
  };
}

function buildRunRpr(ref, lang, part) {
  const langAttr = lang === "cn" ? 'lang="zh-CN" altLang="en-US"' : 'lang="en-HK" altLang="zh-CN"';
  const sz = ref.sz ? ` sz="${ref.sz}"` : "";
  const isBold = part.bold ?? ref.bold;
  const bold = isBold ? ' b="1"' : "";
  const fill = part.color ? `<a:solidFill><a:srgbClr val="${part.color}"/></a:solidFill>` : ref.fill;
  // 模板该形状显式指定了字体就沿用；否则不写字体 → 继承占位符/主题的东亚字体（楷体）。
  const tf = ref.typeface ? `<a:latin typeface="${ref.typeface}"/><a:ea typeface="${ref.typeface}"/>` : "";
  return `<a:rPr ${langAttr}${sz}${bold} dirty="0">${fill}${tf}</a:rPr>`;
}

// Keynote 不会自动应用 normAutofit，所以把缩放比例预先烤进 bodyPr。
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

// 估行数：统一用 em 为单位。中日韩字宽 1em、其它（英文/数字/空格）0.55em
// —— 按人工版 deck 实测校准（166 字符英文在 2800 字号 11 行内恰好放下 ≈ 0.54em/字）。
// 之前的公式字符宽按 0.55em 折过一次、行容量又按每字 0.5 再折一次，容量高估约一倍，
// 导致英文页判定「放得下」而不烤缩放 → 直接溢出。
function estimateRows(text, capacityEm) {
  const lines = String(text || "").split("\n");
  let rows = 0;
  for (const line of lines) {
    let em = 0;
    for (const ch of String(line)) em += emOf(ch);
    rows += Math.max(1, Math.ceil(em / capacityEm));
  }
  return rows;
}

// 解析形状文本框尺寸（EMU）：优先形状自身 xfrm，否则查版式占位符，再否则默认。
function resolveShapeBox(slideXml, slideRels, shapeId) {
  const sp = slideXml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?id="${shapeId}"[\\s\\S]*?</p:sp>`)) || [];
  const own = (sp[0] || "").match(/<a:ext cx="(\d+)" cy="(\d+)"/);
  if (own) return { cx: +own[1], cy: +own[2] };
  const phIdx = (sp[0] || "").match(/<p:ph[^>]*idx="(\d+)"/)?.[1];
  const phType = (sp[0] || "").match(/<p:ph[^>]*type="(\w+)"/)?.[1];
  const layoutFile = (slideRels || "").match(/slideLayout\d+\.xml/)?.[0];
  if (layoutFile) {
    try {
      const lx = readFileSync(join(workDir, "package-template/ppt/slideLayouts", layoutFile), "utf8");
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
  return { cx: 8000000, cy: 5300000 };
}

// 用原型形状的参考格式渲染文字，返回 <p:txBody>。
function fillTxBody(slideXml, slideRels, shapeId, text, lang) {
  const tpl = extractTemplateStyle(slideXml, shapeId);
  const ref = refRprParts(tpl.rpr || "");
  const algn = (tpl.pPr || "").match(/algn="(\w+)"/)?.[1];

  const sz = Number(ref.sz) || (lang === "cn" ? 4800 : 2800);
  const box = resolveShapeBox(slideXml, slideRels, shapeId);
  const fontEmu = (sz / 100) * 12700;
  // 行高：英文段落显式 lnSpc 100% ≈ 1.15em；中文楷体默认 ≈ 1.18em。
  const lineHFactor = lang === "en" ? 1.15 : 1.18;
  // 迭代求缩放：从 100% 逐步降，直到按当前缩放估算能放下为止。
  // （sqrt 一步到位在取整边界会差 1 行；迭代自检从构造上保证不溢出。）
  let scalePct = 100;
  while (scalePct > 45) {
    const f = fontEmu * (scalePct / 100);
    const cap = Math.max(1, (box.cx * 0.94) / f);
    const maxl = Math.max(1, Math.floor((box.cy * 0.96) / (f * lineHFactor)));
    if (estimateRows(text, cap) <= maxl) break;
    scalePct -= 3;
  }
  const bodyPr = setAutofit(tpl.bodyPr || "<a:bodyPr/>", scalePct);

  const lines = String(text || "").split("\n");
  if (!lines.length) lines.push("");
  const paras = lines
    .map((line) => {
      const run = `<a:r>${buildRunRpr(ref, lang, {})}<a:t>${escapeXml(line)}</a:t></a:r>`;
      const a = algn ? ` algn="${algn}"` : "";
      const lineSpacing = lang === "en" ? '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>' : "";
      return `<a:p><a:pPr marL="0" indent="0"${a}>${lineSpacing}<a:buNone/></a:pPr>${run}<a:endParaRPr dirty="0"/></a:p>`;
    })
    .join("");
  return `<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody>`;
}

// 红字「请手动填入」（文档里没有的内容）。沿用旧生成器：C00000、居中、3600。
function redNoticeTxBody(text) {
  return `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="3600" b="1" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="${CHINESE_FONT}"/><a:ea typeface="${CHINESE_FONT}"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody>`;
}

function emptyTxBody() {
  return '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr dirty="0"/></a:p></p:txBody>';
}

// 把某形状的 <p:txBody> 整体替换。
function replaceShapeTxBody(slideXml, shapeId, replacement) {
  const re = new RegExp(`(<p:sp>[\\s\\S]*?<p:cNvPr id="${shapeId}"[\\s\\S]*?<p:txBody>)[\\s\\S]*?(<\\/p:txBody>[\\s\\S]*?<\\/p:sp>)`);
  if (!re.test(slideXml)) return slideXml;
  return slideXml.replace(re, (_m, start, end) => {
    const open = start.replace(/<p:txBody>$/, "");
    return `${open}${replacement}${end.replace(/^<\/p:txBody>/, "")}`;
  });
}

// ───────────────────────── 固定页归一化 ─────────────────────────
// 固定经文页（信经/光荣颂/颂谢词/感恩经/天主经…）在源 deck 里格式极乱：中文 2800-5600、
// 英文 2800-3200、粗/常混排、字距 -300/-500、行距被压、缩放 79%-100% 十几档。
// 这里把所有 id3/id4 内容框统一重写为全卷标准（中 4800 就绪对齐 / 英 2800 左对齐、常规、
// 无字距、英文行距 100%），超出自然容量的页用配对引擎拆成多页 —— 不缩字。
function shapeBlockOf(xml, id) {
  return xml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${id}"(?:(?!</p:sp>)[\\s\\S])*?</p:sp>`))?.[0] || null;
}
function deleteShapeOf(xml, id) {
  const sp = shapeBlockOf(xml, id);
  return sp ? xml.replace(sp, "") : xml;
}
function decodeXmlText(s) {
  return String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}
// 按段落取形状文字（<a:p> 之间、<a:br/> 处 = 换行），并解码 XML 实体。
function paraTextOf(sp) {
  if (!sp) return "";
  const raw = [...sp.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
    .map((p) => [...p[0].replace(/<a:br\b[^>]*\/>/g, "<a:t>\n</a:t>").matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join(""))
    .join("\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
  return decodeXmlText(raw);
}
const CANON_FIXED = { cn: { sz: 4800, algn: "just" }, en: { sz: 2800, algn: "l" } };
// 按「行结构」渲染：每行 {text, bold}，空行 = 空段落（保留人工版的空行间隔与粗体答句）。
// opts: { sz 覆盖字号, anchor 垂直对齐, scale 烤入缩放 }
function canonicalTxBodyRuns(lines, lang, opts = {}) {
  const c = CANON_FIXED[lang];
  const sz = opts.sz || c.sz;
  const list = lines && lines.length ? lines : [{ text: "", bold: false }];
  const langAttr = lang === "cn" ? 'lang="zh-CN" altLang="en-US"' : 'lang="en-HK" altLang="zh-CN"';
  // opts.fill：沿用源形状的文字颜色（对话页 id8/9 不继承占位符，丢了颜色会白字白底）。
  const fill = opts.fill || "";
  const paras = list
    .map((l) => {
      const ln = lang === "en" ? '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>' : "";
      // 行 = {text, bold} 或 {runs:[{text,bold}…]}（段内混排：如「领：」加粗 + 正文常规）
      const runsArr = l.runs || [{ text: l.text, bold: l.bold }];
      const runsXml = runsArr
        .filter((r) => r.text)
        .map((r) => `<a:r><a:rPr ${langAttr} sz="${r.sz || sz}"${r.bold ? ' b="1"' : ""} dirty="0"${fill ? `>${fill}</a:rPr>` : "/>"}<a:t>${escapeXml(r.text)}</a:t></a:r>`)
        .join("");
      return `<a:p><a:pPr marL="0" indent="0" algn="${c.algn}">${ln}<a:buNone/></a:pPr>${runsXml}<a:endParaRPr sz="${sz}" dirty="0"/></a:p>`;
    })
    .join("");
  let bodyPr = `<a:bodyPr${opts.anchor ? ` rtlCol="0" anchor="${opts.anchor}"` : ""}><a:normAutofit/></a:bodyPr>`;
  if (opts.scale && opts.scale < 100) bodyPr = setAutofit(bodyPr, opts.scale);
  return `<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody>`;
}
function canonicalTxBody(text, lang, opts = {}) {
  return canonicalTxBodyRuns(String(text || "").split("\n").map((t) => ({ text: t, bold: false })), lang, opts);
}
// 形状文字的颜色（第一个带 solidFill 的 run）：重写 txBody 时原样带回，防止继承色不对。
function firstRunFill(sp) {
  if (!sp) return "";
  for (const r of sp.matchAll(/<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/g)) {
    const f = r[1].match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0];
    if (f) return f;
  }
  return "";
}
// 舞台指示（鞠躬/(Bow)…）：这些 run 本身在源里常是粗体小字，不能让它把整段判成粗体。
const STAGE_RUN_RE = /^\s*[（(]?\s*(鞠躬|Bow)\s*[）)]?\s*$/i;
// 按段落取 {text, bold} 行结构（保留空行）；粗体 = 段内存在【非舞台指示】的粗 run
// （否则信经「…成为人。（鞠躬）祂因…」会因（鞠躬）的 b=1 把正文也误标粗）。
function paraRunsOf(sp) {
  if (!sp) return [];
  const out = [...sp.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((p) => {
    const text = decodeXmlText([...p[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join("")).trim();
    const boldNonStage = [...p[0].matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].some((r) => {
      const t = decodeXmlText([...r[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join(""));
      return t.trim() && /b="1"/.test(r[0]) && !STAGE_RUN_RE.test(t);
    });
    return { text, bold: Boolean(text) && boldNonStage };
  });
  while (out.length && !out[out.length - 1].text) out.pop(); // 去掉尾部空行
  return out;
}
function pageUnitsOfLines(lines, lang) {
  let u = 0;
  for (const l of lines) {
    const t = l.runs ? l.runs.map((r) => r.text).join("") : l.text;
    u += t ? units(t, lang) : lang === "cn" ? 8 : 16;
  }
  return u;
}
// 超容的结构页（含粗体/诗体/空行结构，不能拆页）→ 迭代求缩放兜底。
function fitScaleForLines(lines, lang, slideXml, rels, shapeId) {
  const text = lines.map((l) => (l.runs ? l.runs.map((r) => r.text).join("") : l.text)).join("\n");
  if (!text.trim()) return 100;
  const box = resolveShapeBox(slideXml, rels, shapeId);
  const fontEmu = (CANON_FIXED[lang].sz / 100) * 12700;
  const lineHFactor = lang === "en" ? 1.15 : 1.18;
  let scale = 100;
  while (scale > 45) {
    const f = fontEmu * (scale / 100);
    const cap = Math.max(1, (box.cx * 0.94) / f);
    const maxl = Math.max(1, Math.floor((box.cy * 0.96) / (f * lineHFactor)));
    if (estimateRows(text, cap) <= maxl) break;
    scale -= 3;
  }
  return scale;
}
const POSTURE_FIXED = /^\s*(请?(站立|坐下|跪下|起立)|全体起立|Stand|Sit|Kneel|Please stand|Please sit|Please kneel)\s*$/;
// 把固定页解析成待排条目。两类：
//   struct = 结构保真页（带粗体/空行/额外形状/诗体，或本来就放得下）：保留每行粗体与空行，
//            只统一字号/对齐/字距；超容时烤缩放兜底（不能拆的页）。
//   flow   = 纯常规散文页：进入 flushFixedBuf 的整段连排重排（页面装满）。
function bigContentBox(sp) {
  // 内容框（≥2.5in 高）；小的是「站立/坐下」标签、跳转按钮等，不许动。
  const ext = sp?.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
  return Boolean(ext && +ext[2] >= 2286000);
}
function normalizeFixedEntries(original) {
  let xml = original.xml;
  if (/手动填入/.test(xml)) return { passthrough: original };
  let cnId = "3";
  let enId = "4";
  let sp3 = shapeBlockOf(xml, cnId);
  let sp4 = shapeBlockOf(xml, enId);
  let dialogue = false;
  if (!paraRunsOf(sp3).some((l) => l.text) && !paraRunsOf(sp4).some((l) => l.text)) {
    // 无 id3/id4 内容 → 试对话页家族（id8=中文左栏 / id9=英文右栏，如致候礼/忏悔礼/圣圣圣）。
    // 只认「大内容框」，避免误伤挂在 id8 的姿势标签/按钮（如圣三颂页的「站立」）。
    const sp8 = shapeBlockOf(xml, "8");
    const sp9 = shapeBlockOf(xml, "9");
    if (bigContentBox(sp8) && bigContentBox(sp9) && paraTextOf(sp8) && paraTextOf(sp9)) {
      cnId = "8";
      enId = "9";
      sp3 = sp8;
      sp4 = sp9;
      dialogue = true;
      // 个别页（圣圣圣）有一个和 id8 同文同位的大 id2 幽灵重复形状：重写 id8 后两者字号
      // 会分叉叠影，删掉重复的那个。
      const sp2 = shapeBlockOf(xml, "2");
      if (bigContentBox(sp2) && paraTextOf(sp2).replace(/\s/g, "") === paraTextOf(sp8).replace(/\s/g, "")) {
        xml = deleteShapeOf(xml, "2");
      }
    }
  }
  let lines3 = paraRunsOf(sp3);
  let lines4 = paraRunsOf(sp4);
  const has3 = lines3.some((l) => l.text);
  const has4 = lines4.some((l) => l.text);
  if (!has3 && !has4) return { passthrough: original }; // 分隔页/纯标题页
  const title = dialogue ? paraTextOf(shapeBlockOf(xml, "10")) : paraTextOf(shapeBlockOf(xml, "2"));
  if (!title) return { passthrough: original }; // 无标题 = 分隔大标题页，不动
  // 忏悔礼「上主/基督，求你垂怜」三连句：行间补空行（人工版排版；若源已有空行则保持）。
  const kyrie = (ls, re) => {
    const t = ls.filter((l) => l.text);
    return t.length >= 2 && t.every((l) => re.test(l.text)) && !ls.some((l) => !l.text);
  };
  const interleave = (ls) => ls.filter((l) => l.text).flatMap((l, i, a) => (i < a.length - 1 ? [l, { text: "", bold: false }] : [l]));
  if (kyrie(lines3, /求你垂怜/)) {
    lines3 = interleave(lines3);
    if (kyrie(lines4, /have mercy/i)) lines4 = interleave(lines4);
  }
  // 「（鞠躬）/(Bow)」等动作提示：小号、不加粗（run 级混排；用户指定，人工版原为 28pt 粗体）。
  const bowify = (ls, re2, small) =>
    ls.map((l) => {
      if (!l.text || !re2.test(l.text)) return l;
      const runs = [];
      let rest = l.text;
      let m2;
      while ((m2 = rest.match(re2))) {
        if (m2.index > 0) runs.push({ text: rest.slice(0, m2.index), bold: l.bold });
        runs.push({ text: m2[0], bold: false, sz: small });
        rest = rest.slice(m2.index + m2[0].length);
      }
      if (rest) runs.push({ text: rest, bold: l.bold });
      return { runs };
    });
  const hasBow = lines3.some((l) => /（鞠躬）|\(鞠躬\)/.test(l.text || "")) || lines4.some((l) => /\(bow\)/i.test(l.text || ""));
  const hasBold = lines3.some((l) => l.bold) || lines4.some((l) => l.bold);
  const hasBlank = lines3.some((l) => !l.text && !l.runs) || lines4.some((l) => !l.text && !l.runs);
  lines3 = bowify(lines3, /（鞠躬）|\(鞠躬\)/, 2800);
  lines4 = bowify(lines4, /\(bow\)/i, 2000);
  const hasExtras = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].some((m) => {
    const id = m[0].match(/<p:cNvPr id="(\d+)"/)?.[1];
    return id && !["2", "3", "4", "10"].includes(id) && Boolean(paraTextOf(m[0]));
  });
  const fits = pageUnitsOfLines(lines3, "cn") <= chunkConfig.twoCn.max && pageUnitsOfLines(lines4, "en") <= chunkConfig.twoEn.max;
  const poetic = /光荣颂|羔羊颂/.test(title);
  // 只有带粗体/空行/额外形状/诗体的页才结构保真；普通散文页（无论是否放得下）一律进
  // 整段重排 —— 否则源模板的稀分页会原样保留（一两句占一页）。
  // 对话页（应答结构，一问一答不可跨页重排）一律结构保真。
  // 只有对话页需要带回文字颜色（id8/9 不继承占位符，丢色 → 白字白底）；
  // id3/4 占位符形状继承色本来就对，带回源色反而可能把（鞠躬）等杂色 run 扩散到整页。
  const fill3 = dialogue ? firstRunFill(sp3) : "";
  const fill4 = dialogue ? firstRunFill(sp4) : "";
  if (dialogue || hasBold || hasBlank || hasExtras || poetic || hasBow) {
    return { entries: [{ kind: "struct", lines3, lines4, cnId, enId, fill3, fill4, baseXml: xml, rels: original.rels, title, over: !fits, originFile: original.file }] };
  }
  const cn = lines3.filter((l) => l.text).map((l) => l.text).join("\n");
  const en = lines4.filter((l) => l.text).map((l) => l.text).join("\n");
  return { entries: [{ kind: "flow", cn, en, cnId, enId, fill3, fill4, baseXml: xml, rels: original.rels, title, originFile: original.file }] };
}
// 结构页分页几何：100% 字号下每行的 em 容量 + 每页最大行数。
function structPageGeom(lang, slideXml, rels, shapeId) {
  const box = resolveShapeBox(slideXml, rels, shapeId);
  const f = (CANON_FIXED[lang].sz / 100) * 12700;
  return {
    capEm: Math.max(1, (box.cx * 0.94) / f),
    maxl: Math.max(1, Math.floor((box.cy * 0.96) / (f * (lang === "en" ? 1.15 : 1.18)))),
  };
}
function lineText(l) {
  return l.runs ? l.runs.map((r) => r.text).join("") : l.text || "";
}
function rowsOfLine(l, capEm) {
  const t = lineText(l);
  if (!t) return 1; // 空行占一行
  let em = 0;
  for (const ch of t) em += emOf(ch);
  return Math.max(1, Math.ceil(em / capEm));
}
// CJK 标点在 PowerPoint 里默认参与标点压缩（约半宽），按 0.6em 估——按整字宽（1em）估会把
// 「上主，求你垂怜。」这类短句多算成两行，把明明放得下的页错判成放不下而分页。
const CJK_PUNCT_RE = /[，。、；：？！…—·〈〉《》「」『』【】〔〕（）]/;
const emOf = (ch) => (CJK_PUNCT_RE.test(ch) ? 0.6 : /[　-〿㐀-鿿＀-￯]/.test(ch) ? 1 : 0.55);
// 把一行在给定 em 位置组切成多段（切点吸附句读，保留 runs 粗体等格式）。
// 切点优先级：句末标点 > 逗顿/分号 > 空格 > 任意字符。
function splitLineAtEms(l, capEm, cutEms) {
  const chars = [];
  if (l.runs) {
    for (const r of l.runs) for (const ch of String(r.text || "")) chars.push({ ch, r });
  } else {
    for (const ch of String(l.text || "")) chars.push({ ch, r: null });
  }
  if (chars.length < 4 || !cutEms.length) return [l];
  const grade = (i) => {
    const ch = chars[i]?.ch || "";
    if ("。！？；.!?;".includes(ch)) return 3;
    if ("，、,:：".includes(ch)) return 2;
    if (ch === " ") return 1;
    return 0;
  };
  // 在目标点前后 ~20% 页宽窗口内找最好的标点切点（切在标点之后）。
  const win = Math.max(4, Math.round((capEm * 0.2) / 0.55));
  const cuts = [0];
  let acc = 0;
  let ti = 0;
  for (let i = 0; i < chars.length && ti < cutEms.length; i += 1) {
    acc += emOf(chars[i].ch);
    if (acc < cutEms[ti]) continue;
    let best = i + 1;
    let bestScore = -Infinity;
    for (let j = Math.max(cuts[cuts.length - 1] + 1, i - win); j <= Math.min(chars.length - 2, i + win); j += 1) {
      const score = grade(j) * 1000 - Math.abs(j - i);
      if (score > bestScore) {
        bestScore = score;
        best = j + 1;
      }
    }
    cuts.push(best);
    ti += 1;
  }
  cuts.push(chars.length);
  const pieces = [];
  for (let k = 0; k + 1 < cuts.length; k += 1) {
    const seg = chars.slice(cuts[k], cuts[k + 1]);
    while (seg.length && seg[0].ch === " ") seg.shift();
    if (!seg.length) continue;
    if (!l.runs) {
      pieces.push({ ...l, text: seg.map((c) => c.ch).join("") });
    } else {
      const runs = [];
      for (const c of seg) {
        const last = runs[runs.length - 1];
        if (last && last._r === c.r) last.text += c.ch;
        else runs.push({ ...c.r, text: c.ch, _r: c.r });
      }
      pieces.push({ ...l, runs: runs.map(({ _r, ...rest }) => rest) });
    }
  }
  return pieces.length ? pieces : [l];
}
// K 页【按内容比例同步】切分：第 j 刀落在总行量的 j/K 处，容差内吸附行边界（空行加成），
// 吸附不到就在跨越目标的行内按句读硬切 —— 中英两侧各自按比例下刀，才能页页内容对应
// （独立行平衡会让有现成行边界的一侧懒切：如英文「整段祷文+Amen」切成 [全文][Amen]，
//   中文却切成两半 → 第 1 页中文一半配英文全文、第 2 页另一半配 Amen，完全错位）。
function splitLinesProportional(lines, capEm, K) {
  if (K <= 1 || !lines.length) return [lines.slice()];
  const rws = lines.map((l) => rowsOfLine(l, capEm));
  const total = rws.reduce((a, b) => a + b, 0);
  const cum = [0];
  for (const r of rws) cum.push(cum[cum.length - 1] + r);
  const tol = Math.max(1.5, (total / K) * 0.3);
  const isBlankIdx = (i) => i >= 0 && i < lines.length && !lineText(lines[i]);
  const cutPlan = []; // {line, frac}: frac=0 → 该行前整切；否则在该行内 frac 处切
  let prevPos = 0;
  for (let j = 1; j < K; j += 1) {
    const target = (total * j) / K;
    let bi = -1;
    let bs = Infinity;
    for (let i = 1; i < lines.length; i += 1) {
      if (cum[i] <= prevPos + 0.01) continue; // 刀必须前进
      const d = Math.abs(cum[i] - target) - (isBlankIdx(i - 1) || isBlankIdx(i) ? 0.75 : 0);
      if (d < bs) {
        bs = d;
        bi = i;
      }
    }
    if (bi > 0 && bs <= tol) {
      cutPlan.push({ line: bi, frac: 0 });
      prevPos = cum[bi];
      continue;
    }
    let li = 0;
    while (li + 1 < cum.length - 1 && cum[li + 1] < target) li += 1;
    li = Math.min(li, lines.length - 1);
    const frac = rws[li] > 0 ? Math.min(0.98, Math.max(0.02, (target - cum[li]) / rws[li])) : 0;
    cutPlan.push({ line: li, frac });
    prevPos = cum[li] + frac * rws[li];
  }
  // 物化切分计划：行内刀用 splitLineAtEms。
  const parts = [];
  let cur = [];
  for (let i = 0; i < lines.length; i += 1) {
    const here = cutPlan.filter((c) => c.line === i);
    for (let p = here.filter((c) => c.frac === 0).length; p > 0; p -= 1) {
      parts.push(cur);
      cur = [];
    }
    const fracs = here.filter((c) => c.frac > 0).map((c) => c.frac).sort((a, b) => a - b);
    if (!fracs.length) {
      cur.push(lines[i]);
      continue;
    }
    let lineEm = 0;
    for (const ch of lineText(lines[i])) lineEm += emOf(ch);
    const pieces = splitLineAtEms(lines[i], capEm, fracs.map((f) => f * lineEm));
    for (let p = 0; p < pieces.length; p += 1) {
      cur.push(pieces[p]);
      if (p < pieces.length - 1) {
        parts.push(cur);
        cur = [];
      }
    }
  }
  parts.push(cur);
  const out = parts.map((g) => {
    const h = g.slice();
    while (h.length && !lineText(h[0])) h.shift();
    while (h.length && !lineText(h[h.length - 1])) h.pop();
    return h.length ? h : [{ text: "", bold: false }];
  });
  // 物化组数偏差兜底（行太短切不动等罕见情形）：并尾/补空到恰好 K 组。
  while (out.length > K) {
    const t = out.pop();
    out[out.length - 1].push(...t);
  }
  while (out.length < K) out.push([{ text: "", bold: false }]);
  return out;
}
// 渲染固定页 → 一页或多页。结构页超出自然容量【分页】而不缩字（用户要求，2026-07-09）：
// 按行结构切成 K 份连续克隆页；单页仍自检缩放兜底（正常应为 100%，只有单行超长才触发）。
function renderFixedEntry(e) {
  const cnId = e.cnId || "3";
  const enId = e.enId || "4";
  if (e.kind !== "struct") {
    let x = e.baseXml;
    if (shapeBlockOf(x, cnId)) x = replaceShapeTxBody(x, cnId, canonicalTxBody(e.cn, "cn", { fill: e.fill3 }));
    if (shapeBlockOf(x, enId)) x = replaceShapeTxBody(x, enId, canonicalTxBody(e.en, "en", { fill: e.fill4 }));
    return [{ xml: x, rels: e.rels }];
  }
  const g3 = structPageGeom("cn", e.baseXml, e.rels, cnId);
  const g4 = structPageGeom("en", e.baseXml, e.rels, enId);
  const rows3 = e.lines3.reduce((s, l) => s + rowsOfLine(l, g3.capEm), 0);
  const rows4 = e.lines4.reduce((s, l) => s + rowsOfLine(l, g4.capEm), 0);
  let K = Math.max(1, Math.ceil(rows3 / g3.maxl), Math.ceil(rows4 / g4.maxl));
  let parts3 = splitLinesProportional(e.lines3, g3.capEm, K);
  let parts4 = splitLinesProportional(e.lines4, g4.capEm, K);
  // 切分后仍有组超容（吸附边界偏离比例太多挤了长行）→ 加一页重切，最多试 3 次。
  for (let tries = 0; tries < 3; tries += 1) {
    const over =
      parts3.some((g) => g.reduce((s, l) => s + rowsOfLine(l, g3.capEm), 0) > g3.maxl) ||
      parts4.some((g) => g.reduce((s, l) => s + rowsOfLine(l, g4.capEm), 0) > g4.maxl);
    if (!over || K >= Math.max(e.lines3.length, e.lines4.length) + 4) break;
    K += 1;
    parts3 = splitLinesProportional(e.lines3, g3.capEm, K);
    parts4 = splitLinesProportional(e.lines4, g4.capEm, K);
  }
  const pages = [];
  for (let k = 0; k < K; k += 1) {
    let x = e.baseXml;
    const l3 = parts3[k] || [{ text: "", bold: false }];
    const l4 = parts4[k] || [{ text: "", bold: false }];
    const s3 = fitScaleForLines(l3, "cn", x, e.rels, cnId);
    const s4 = fitScaleForLines(l4, "en", x, e.rels, enId);
    if (shapeBlockOf(x, cnId)) x = replaceShapeTxBody(x, cnId, canonicalTxBodyRuns(l3, "cn", { scale: s3, fill: e.fill3 }));
    if (shapeBlockOf(x, enId)) x = replaceShapeTxBody(x, enId, canonicalTxBodyRuns(l4, "en", { scale: s4, fill: e.fill4 }));
    pages.push({ xml: x, rels: e.rels });
  }
  return pages;
}
// 固定页输出：struct 页原位输出；连续同标题的 flow 页整段连排重排、页面装满。
// 同时登记 firstOut[模板页号]=输出首页索引（跳转按钮重映射用；整段重排的多页映射到组首页）。
function flushFixedBuf(fixedBuf, kept) {
  const out = [];
  let i = 0;
  while (i < fixedBuf.length) {
    const e = fixedBuf[i];
    if (e.kind === "struct") {
      out.push({ ...e, origs: [e.origIdx] });
      i += 1;
      continue;
    }
    let j = i;
    while (j < fixedBuf.length && fixedBuf[j].kind === "flow" && fixedBuf[j].title === e.title && fixedBuf[j].rels === e.rels) j += 1;
    const seg = fixedBuf.slice(i, j);
    const cnAll = seg.map((s) => s.cn).filter(Boolean).join("\n").replace(/\n+/g, "");
    const enAll = seg.map((s) => s.en).filter(Boolean).join("\n").replace(/\n+/g, " ");
    const packed = mergeSmallPairs(alignPair(cnAll, enAll, "twoCn", "twoEn"));
    packed.forEach((p, pi) =>
      out.push({ kind: "flow", cn: p.cn, en: p.en, fill3: e.fill3, fill4: e.fill4, baseXml: e.baseXml, rels: e.rels, title: e.title, origs: pi === 0 ? seg.map((s) => s.origIdx) : [] })
    );
    i = j;
  }
  for (const e of out) {
    const at = kept.length;
    for (const oi of e.origs || []) if (oi !== undefined && firstOut[oi] === undefined) firstOut[oi] = at;
    for (const page of renderFixedEntry(e)) kept.push(page);
  }
  fixedBuf.length = 0;
}

// 找出某页里的占位符 token 及其所在形状。
function findTokens(slideXml) {
  const tokens = [];
  for (const sp of slideXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const shapeId = sp[0].match(/<p:cNvPr id="(\d+)"/)?.[1];
    if (!shapeId) continue;
    for (const m of sp[0].matchAll(/\{\{([^}]+)\}\}/g)) {
      const inner = m[1];
      const suffix = /_(CN|EN)$/.exec(inner)?.[1] || "";
      const base = inner.replace(/_(CN|EN)$/, "");
      tokens.push({ shapeId, token: m[0], base, suffix });
    }
  }
  return tokens;
}

// ───────────────────────── 解包 ─────────────────────────
mkdirSync(workDir, { recursive: true });
const work = join(workDir, "package-template");
rmSync(work, { recursive: true, force: true });
mkdirSync(work, { recursive: true });
execFileSync("unzip", ["-q", templatePptx, "-d", work]);

const slideDir = join(work, "ppt/slides");
const relsDir = join(slideDir, "_rels");

// 页序以 presentation.xml 的 sldIdLst 为准（文件编号 ≠ 真实顺序：模板手术插入的页编号最大但位置在中间）。
const presXml = readFileSync(join(work, "ppt/presentation.xml"), "utf8");
const presRels = readFileSync(join(work, "ppt/_rels/presentation.xml.rels"), "utf8");
const ridToFile = {};
for (const m of presRels.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="slides\/(slide\d+\.xml)"/g)) ridToFile[m[1]] = m[2];
const orderedFiles = [...presXml.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)].map((m) => ridToFile[m[1]]).filter(Boolean);

const originals = [];
for (const file of orderedFiles) {
  const xml = readFileSync(join(slideDir, file), "utf8");
  const relsPath = join(relsDir, `${file}.rels`);
  const rels = existsSync(relsPath) ? readFileSync(relsPath, "utf8") : "";
  originals.push({ xml, rels, file });
}
const slideCount = originals.length;
const fileToPos = Object.fromEntries(orderedFiles.map((f, i) => [f, i]));
const firstOut = {}; // 模板页序号 → 输出首页索引（跳转按钮按内容重映射用）

// 「坐下」姿势标签捐赠体：药丸样式用正牌内容页的绿色「站立」药丸（55725D 白字 24pt，
// 章节卡上的白底坐下药丸不是内容页样式），文字换成「坐下」；图标用「坐下」页的坐姿图标，
// 加 lum 白化滤镜（正牌药丸图标的做法）、去边框、按正牌偏移放进药丸。
function findSitDonor() {
  let pillSp = null;
  for (const o of originals) {
    for (const m of o.xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
      if (paraTextOf(m[0]) === "站立" && m[0].includes('val="55725D"')) {
        pillSp = m[0];
        break;
      }
    }
    if (pillSp) break;
  }
  for (const o of originals) {
    for (const m of o.xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
      if (paraTextOf(m[0]) !== "坐下") continue;
      const off = m[0].match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
      if (!off) continue;
      const px = +off[1];
      const py = +off[2];
      let pic = "";
      for (const pm of o.xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)) {
        const po = pm[0].match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
        if (po && Math.abs(+po[1] - px) < 1500000 && Math.abs(+po[2] - py) < 800000) {
          pic = pm[0];
          break;
        }
      }
      let sp = m[0];
      if (pillSp) {
        sp = pillSp.replace(/<a:t>站立<\/a:t>/, "<a:t>坐下</a:t>");
        if (pic) {
          const pOff = pillSp.match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
          const pExt = pillSp.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
          const iExt = pic.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
          if (pOff && pExt && iExt) {
            const ix = +pOff[1] + 46830; // 正牌药丸图标的左内边距
            const iy = +pOff[2] + Math.round((+pExt[2] - +iExt[2]) / 2);
            pic = pic.replace(/<a:off x="-?\d+" y="-?\d+"\/>/, `<a:off x="${ix}" y="${iy}"/>`);
          }
          pic = pic
            .replace(/<a:biLevel[^>]*\/>/, '<a:lum bright="70000" contrast="-70000"/>')
            .replace(/<a:grpFill\/>/g, "<a:noFill/>")
            .replace(/<a:ln><a:solidFill><a:srgbClr val="[0-9A-F]+"\/><\/a:solidFill><\/a:ln>/g, "<a:ln><a:noFill/></a:ln>");
        }
      }
      const relIds = pic ? [...new Set([...pic.matchAll(/r:(?:embed|link)="([^"]+)"/g)].map((x) => x[1]))] : [];
      const relEntries = relIds
        .map((rid) => o.rels.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*/>`))?.[0])
        .filter(Boolean);
      return { sp, pic, relIds, relEntries };
    }
  }
  return null;
}
const sitDonor = findSitDonor();
// 往页面右下角注入「坐下」标签（pill + 图标），图标 rels 换新 id 并入该页 rels。
function injectSit(xml, rels) {
  if (!sitDonor) return { xml, rels };
  let sp = sitDonor.sp.replace(/<p:cNvPr id="\d+"/, '<p:cNvPr id="9500"');
  let pic = sitDonor.pic ? sitDonor.pic.replace(/<p:cNvPr id="\d+"/, '<p:cNvPr id="9501"') : "";
  let newRels = rels || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  sitDonor.relIds.forEach((rid, i) => {
    const nid = `rIdSit${i + 1}`;
    pic = pic.split(`"${rid}"`).join(`"${nid}"`);
    const entry = sitDonor.relEntries[i]?.replace(`Id="${rid}"`, `Id="${nid}"`);
    if (entry && !newRels.includes(`Id="${nid}"`)) newRels = newRels.replace("</Relationships>", `${entry}</Relationships>`);
  });
  return { xml: xml.replace("</p:spTree>", `${sp}${pic}</p:spTree>`), rels: newRels };
}

// ───────────────────────── 逐页生成 ─────────────────────────
const kept = []; // {xml, rels}
const filled = [];
const stillMissing = [];
const fillList = []; // 需手动填入的页：{page, label}（页码为最终 deck 页码，用于网页「待填清单」）

const fixedBuf = [];
for (let origIdx = 0; origIdx < originals.length; origIdx += 1) {
  const original = originals[origIdx];
  const tokens = findTokens(original.xml);
  if (!tokens.length) {
    const norm = normalizeFixedEntries(original);
    if (norm.passthrough) {
      flushFixedBuf(fixedBuf, kept);
      firstOut[origIdx] = kept.length;
      // 领主咏歌词红字页前，插一页文档里的「领圣体咏(领主咏 antiphon)」中英内容（每周不同）。
      const ptTitle = paraTextOf(shapeBlockOf(original.xml, "2")) || paraTextOf(shapeBlockOf(original.xml, "10"));
      if (/领主咏/.test(ptTitle) && /手动填入/.test(original.xml)) {
        const anti = sectionByTitle.get("领主咏");
        // 去掉开头的经卷出处（各周格式不同：「咏34:9」「若6:56」独立行或内联；英文「Cf. Ps 34:9」）。
        const dropAntiphonRef = (t) => {
          const lines = paragraphsOf(t);
          if (!lines.length) return "";
          const CN_REF = /^[（(]?\s*[一-鿿]{1,4}\s*\d+\s*[:：]\s*\d+[a-z]?(?:[-–,、\s]+\d+[a-z]?)*[）)]?\s*/;
          const EN_REF = /^[（(]?\s*(?:Cf\.?\s*)?(?:[1-3]\s*)?[A-Z][a-z]{1,4}\.?\s+\d+\s*[:：]?\s*\d*[a-z]?(?:[-–,\s]+\d+)*[）)]?\s*/;
          const m = lines[0].match(CN_REF) || lines[0].match(EN_REF);
          if (m) {
            const rest = lines[0].slice(m[0].length).trim();
            if (rest) lines[0] = rest;
            else lines.shift();
          }
          return lines.join("\n");
        };
        const acn = dropAntiphonRef(anti?.cn?.trim() || "");
        const aen = dropAntiphonRef(anti?.en?.trim() || "");
        if (acn || aen) {
          let ax = original.xml;
          if (shapeBlockOf(ax, "3")) ax = replaceShapeTxBody(ax, "3", canonicalTxBody(acn, "cn"));
          if (shapeBlockOf(ax, "4")) ax = replaceShapeTxBody(ax, "4", canonicalTxBody(aen, "en"));
          kept.push({ xml: ax, rels: original.rels, originFile: original.file });
        }
      }
      kept.push(norm.passthrough);
    } else {
      for (const e of norm.entries) e.origIdx = origIdx;
      fixedBuf.push(...norm.entries);
    }
    continue;
  }
  flushFixedBuf(fixedBuf, kept);
  firstOut[origIdx] = kept.length;

  // 封面标题：直接替换两个 token（保留封面原本的居中/字号/白色格式）。
  if (tokens.some((t) => t.base === "主日标题")) {
    let xml = original.xml
      .replace(/\{\{主日标题_CN\}\}/g, escapeXml(meta.titleCn))
      .replace(/\{\{主日标题_EN\}\}/g, escapeXml(meta.titleEn));
    // 防溢出：英文标题按 4400 估行数（框宽约 685pt）；超过 2 行则整个标题降一号（只动标题框）。
    const enLines = Math.max(1, Math.ceil((String(meta.titleEn).length * 0.55 * 44) / 620));
    if (enLines > 2) {
      const spRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<p:cNvPr id="4"(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/;
      xml = xml.replace(spRe, (sp) =>
        sp
          .replace(/sz="4400"/g, 'sz="3600"')
          .replace(/sz="4800"/g, 'sz="4400"')
          .replace(/<a:spcPts val="5200"\/>/, '<a:spcPts val="4400"/>')
      );
    }
    kept.push({ xml, rels: original.rels, originFile: original.file });
    filled.push("主日标题");
    continue;
  }

  const base = tokens[0].base;

  // 文档里没有内容的段落 → 红字提示。
  if (RED_LABEL[base]) {
    let xml = original.xml;
    tokens.forEach((t, i) => {
      xml = replaceShapeTxBody(xml, t.shapeId, i === 0 ? redNoticeTxBody(RED_LABEL[base]) : emptyTxBody());
    });
    kept.push({ xml, rels: original.rels, originFile: original.file });
    continue;
  }

  const section = sectionByTitle.get(TOKEN_TO_SECTION[base]);
  const cn = section?.cn?.trim() || "";
  const en = section?.en?.trim() || "";

  // 文档缺这一段 → 红字提示，保留原型一页。
  if (!cn && !en) {
    let xml = original.xml;
    tokens.forEach((t, i) => {
      xml = replaceShapeTxBody(xml, t.shapeId, i === 0 ? redNoticeTxBody(`${base} · 请手动填入`) : emptyTxBody());
    });
    kept.push({ xml, rels: original.rels, originFile: original.file });
    stillMissing.push(base);
    continue;
  }

  if (SINGLE_FILL.has(base)) {
    // 信友祷文：一段祷词一页（主祭/领），不从中间截断（长了由 fontScale 缩）；
    // 每段「领：…」后插一页「答：求主俯听我们。」（人工版逐段插答句；文档自带答句时不再补）。
    const cnShape = tokens.find((t) => t.suffix !== "EN")?.shapeId || tokens[0].shapeId;
    const paras = paragraphsOf(cn);
    // 会众答句优先从文档探测（各堂区/主日措辞不同：求主俯听我们/求主垂怜…），探测不到才用默认。
    // 文档只写了一次答句时，也统一补到每段「领：」之后（人工版逐段插答句）。
    const docResp = paras.find((l) => /^答[:：]/.test(l));
    const resp = docResp || "答：求主俯听我们。";
    const pages = [];
    for (let i = 0; i < paras.length; i += 1) {
      const p = paras[i];
      // 超过全宽框自然容量的长祷词分两页（保持统一字号，不缩字）。
      const pieces = units(p, "cn") > chunkConfig.faithfulCn.max ? chunkText(p, "cn", "faithfulCn") : [p];
      pages.push(...pieces);
      if (/^领[:：]/.test(p) && !/^答[:：]/.test(paras[i + 1] || "")) pages.push(resp);
    }
    const list = pages.length ? pages : [cn];
    const PREFIX_RE = /^(主祭|主礼|领)([:：])/;
    for (const piece of list) {
      // 答句页：楷体 6000、上下居中、加粗；祷词页只加粗「主祭：/主礼：/领：」前缀，正文常规。
      let body;
      if (/^答[:：]/.test(piece) && units(piece, "cn") < 24) {
        body = canonicalTxBodyRuns([{ text: piece, bold: true }], "cn", { sz: 6000, anchor: "ctr" });
      } else {
        const m = piece.match(PREFIX_RE);
        const lineObj = m
          ? { runs: [{ text: m[0], bold: true }, { text: piece.slice(m[0].length), bold: false }] }
          : { text: piece, bold: false };
        body = canonicalTxBodyRuns([lineObj], "cn", {
          scale: fitScaleForLines([lineObj], "cn", original.xml, original.rels, cnShape),
        });
      }
      const xml = replaceShapeTxBody(original.xml, cnShape, body);
      kept.push({ xml, rels: original.rels, originFile: original.file });
    }
    filled.push(base);
    continue;
  }

  if (TWO_COL.has(base)) {
    const cnShape = tokens.find((t) => t.suffix === "CN")?.shapeId;
    const enShape = tokens.find((t) => t.suffix === "EN")?.shapeId;
    let pairs;
    if (READINGS.has(base)) {
      // 读经/福音：先剥出「恭读…↔A reading from…」出处行单独成页（引题括号行不显示，
      // 与人工版一致），正文再按比例配对 —— 否则这些短行会把比例配对带偏、中英错页。
      const peeled = peelReadingIntro(cn, en);
      pairs = [];
      const isGospel = base === "福音";
      if (isGospel) {
        // 福音第一页 = 人工版固定对话排版，只有出处行来自 word；答句加粗。
        pairs.push({
          cnLines: [
            { text: "愿主与你们同在。", bold: false },
            { text: "答：也与你的心灵同在。", bold: true },
            { text: "", bold: false },
            { text: peeled.cnCit || "恭读福音", bold: false },
            { text: "答：主，愿光荣归于你。", bold: true },
          ],
          enLines: [
            { text: "The Lord be with you.", bold: false },
            { text: "", bold: false },
            { text: "R. And with your spirit.", bold: true },
            { text: "", bold: false },
            { text: peeled.enCit || "A reading from the holy Gospel", bold: false },
            { text: "", bold: false },
            { text: "R. Glory to you, O Lord.", bold: true },
          ],
        });
      } else if (peeled.cnCit || peeled.enCit) {
        // 读经出处页：中文加粗（人工版格式；英文常规）。
        pairs.push({
          cnLines: [{ text: "", bold: false }, { text: peeled.cnCit, bold: true }],
          enLines: [{ text: "", bold: false }, { text: peeled.enCit, bold: false }],
        });
      }
      // 结尾「——上主的圣言/The word of the Lord」从正文剥出（中文常拼在末句尾巴上，按短语剥）。
      const cnLines2 = paragraphsOf(peeled.cnBody);
      const enLines2 = paragraphsOf(peeled.enBody);
      const CN_END = /^(——|[—―─]|答[:：]\s*感谢天主|基督，我们赞美你)/;
      const EN_END = /^(The (word|Gospel) of the Lord|Thanks be to God|Praise to you)/i;
      while (cnLines2.length && CN_END.test(cnLines2[cnLines2.length - 1])) cnLines2.pop();
      while (enLines2.length && EN_END.test(enLines2[enLines2.length - 1])) enLines2.pop();
      // 正文连排成流式文本再分页（人工版每页就是一个连排段落）；保留 \n 会让每段自成一页 → 稀页。
      let cnFlow = cnLines2.join("");
      const enFlow = enLines2.join(" ");
      const endM = cnFlow.match(/[—―─]{1,2}\s*(上主的圣言|基督的福音)[\s\S]*$/);
      if (endM && endM.index > 0) cnFlow = cnFlow.slice(0, endM.index).trim();
      pairs.push(...mergeSmallPairs(alignPair(cnFlow, enFlow, "twoCn", "twoEn")));
      // 结尾页（人工版标准）：圣言/福音 + 空行 + 会众答句（加粗），中英同页。
      pairs.push({
        cnLines: [
          { text: isGospel ? "——基督的福音。" : "——上主的圣言。", bold: false },
          { text: "", bold: false },
          { text: isGospel ? "基督，我们赞美你！" : "答：感谢天主。", bold: true },
        ],
        enLines: [
          { text: isGospel ? "The Gospel of the Lord." : "The word of the Lord.", bold: false },
          { text: "", bold: false },
          { text: isGospel ? "Praise to you Lord, Jesus Christ!" : "Thanks be to God.", bold: true },
        ],
      });
    } else if (base === "答唱咏") {
      pairs = psalmPairs(cn, en) || mergeSmallPairs(chunkBilingualProportional(cn, en));
    } else if (base === "福音前欢呼") {
      pairs = acclaimPairs(cn, en);
    } else if (PRAYERS.has(base)) {
      pairs = prayerPairs(base, cn, en);
    } else {
      pairs = mergeSmallPairs(chunkBilingualProportional(cn, en));
    }
    const list = pairs.length ? pairs : [{ cn, en }];
    for (const pair of list) {
      // 结构化行页（无特殊排版 opts）：走 struct 分页逻辑 —— 放不下按行/句读分页而不缩字。
      let pagesOut;
      if (pair.cnLines && pair.enLines && !pair.cnOpts && !pair.enOpts && cnShape && enShape) {
        pagesOut = renderFixedEntry({
          kind: "struct",
          lines3: pair.cnLines,
          lines4: pair.enLines,
          cnId: cnShape,
          enId: enShape,
          fill3: "",
          fill4: "",
          baseXml: original.xml,
          rels: original.rels,
        });
      } else {
        let xml = original.xml;
        if (cnShape) {
          const body = pair.cnLines
            ? canonicalTxBodyRuns(pair.cnLines, "cn", { scale: fitScaleForLines(pair.cnLines, "cn", original.xml, original.rels, cnShape), ...(pair.cnOpts || {}) })
            : fillTxBody(original.xml, original.rels, cnShape, pair.cn, "cn");
          xml = replaceShapeTxBody(xml, cnShape, body);
        }
        if (enShape) {
          const body = pair.enLines
            ? canonicalTxBodyRuns(pair.enLines, "en", { scale: fitScaleForLines(pair.enLines, "en", original.xml, original.rels, enShape), ...(pair.enOpts || {}) })
            : fillTxBody(original.xml, original.rels, enShape, pair.en, "en");
          xml = replaceShapeTxBody(xml, enShape, body);
        }
        pagesOut = [{ xml, rels: original.rels }];
      }
      for (const pg of pagesOut) {
        // 集祷经最后一页右下角补「坐下」标签（人工版格式：集祷经完坐下听读经）。
        if (base === "集祷经" && pair === list[list.length - 1] && pg === pagesOut[pagesOut.length - 1]) {
          const inj = injectSit(pg.xml, original.rels);
          kept.push({ xml: inj.xml, rels: inj.rels });
        } else {
          kept.push({ xml: pg.xml, rels: pg.rels, originFile: original.file });
        }
      }
    }
    filled.push(base);
    continue;
  }

  // 未知占位符：原样保留（不应发生）。
  kept.push(original);
}
flushFixedBuf(fixedBuf, kept);

// ── 统一所有段落标题字号 ──
// 标题栏(id2；歌咏页在 id10)在模板里字号继承 + 各页 fontScale 不一（感恩经三溢出、颂谢词四~七偏小）。
// 一律写死 3200、去掉 fontScale/autofit 缩放 → 全卷标题一样大、最长标题也放得下（框宽 685pt）。
const TITLE_SZ = 3200;
function setTitleRunSz(sp, sz) {
  return sp.replace(/<a:rPr\b[^>]*?\/?>/g, (m) => {
    let t = m.replace(/\s+sz="\d+"/, "");
    return t.replace(/^<a:rPr\b/, `<a:rPr sz="${sz}"`);
  });
}
function unifyTitle(xml) {
  for (const id of ["2", "10"]) {
    const sp = shapeBlockOf(xml, id);
    if (!sp || !paraTextOf(sp)) continue;
    if (bigContentBox(sp)) continue; // 挂着标题 id 的大内容框（如圣圣圣页的幽灵重复形状）不是标题
    if (POSTURE_FIXED.test(paraTextOf(sp))) continue; // 「站立/坐下」姿势药丸有时占用 id10（如信经页），必须保持 24pt

    let nsp = setTitleRunSz(sp, TITLE_SZ); // 所有标题（含堂区报告）统一 32pt；60pt 会让「Parish Announcements」溢出绿条
    nsp = nsp.replace(/<a:normAutofit\b[^>]*\/>/g, "<a:normAutofit/>"); // 去掉烤入的 fontScale
    xml = xml.replace(sp, nsp);
  }
  return xml;
}
for (const k of kept) k.xml = unifyTitle(k.xml);

// ── 跳转按钮重映射：按【按钮自身文字】把目标重指到匹配内容的输出页 ──
// （模板从折叠 deck 生成时跳转 target 已错位，不能信；按钮文字如「圣、圣、圣」「圣三颂」
//   「常年期主日颂谢词…」本身就说明要去哪，直接在输出里找内容匹配的页最可靠。）
const pageSig = kept.map((k) => {
  const t = (id) => paraTextOf(shapeBlockOf(k.xml, id));
  return { title: (t("2") || t("10")).replace(/\s/g, ""), body: (t("8") || t("3") || t("4") || "").replace(/\s/g, "") };
});
function findJumpTarget(btnText) {
  const b = btnText.replace(/[\s　]/g, "");
  if (b.length < 3) return undefined; // 「继续」等过泛按钮不重指（易误配），保留原样
  const key = b.slice(0, Math.min(6, b.length));
  // ① 正文首句精确匹配（圣、圣、圣 等独特正文页）；② 标题/正文包含匹配（Doxology　圣三颂、颂谢词、悼念亡者…）。
  for (let i = 0; i < pageSig.length; i += 1) if (pageSig[i].body && pageSig[i].body.startsWith(key)) return i + 1;
  for (let i = 0; i < pageSig.length; i += 1) if ((pageSig[i].title && pageSig[i].title.includes(key)) || pageSig[i].body.includes(key)) return i + 1;
  return undefined;
}
// 命名靶点：全部按标题/正文首句在【输出 deck】里解析 —— 页码每次生成都不同，绝不能按页码跳。
const CN_NUMS = ["一", "二", "三", "四", "五", "六", "七", "八"];
const findFirst = (pred) => {
  const i = pageSig.findIndex(pred);
  return i >= 0 ? i + 1 : undefined;
};
const prefaceFirst = CN_NUMS.map((n) => findFirst((p) => p.title.includes(`常年期主日颂谢词（${n}）`)));
const epFirst = CN_NUMS.slice(0, 4).map((n) => findFirst((p) => p.title.includes(`感恩经（${n}）`)));
const sanctusPage = findFirst((p) => p.body.startsWith("圣、圣、圣"));
const doxologyPage = findFirst((p) => p.title.includes("圣三颂") || p.title.includes("Doxology"));
// 感恩经（三）的两条支线：悼念亡者（殡葬弥撒插段）与正常路径（跳过插段）。
const memorialPage = findFirst((p) => p.title.includes("感恩经（三）") && p.body.startsWith("求祢垂念祢的仆人"));
const memorialSkipPage = findFirst((p) => p.title.includes("感恩经（三）") && p.body.startsWith("求祢垂怜我们的祖先"));

// 文字按钮 → 靶点。explicitOnly=true 时不走泛匹配兜底（用于激活模板里 noaction 的断链按钮，避免误激活）。
const textTarget = (btn, pageTitle, explicitOnly) => {
  const b = btn.replace(/[\s　]/g, "");
  if (!b) return undefined;
  if (b.includes("圣三颂")) return doxologyPage;
  if (b.replace(/、/g, "").startsWith("圣圣圣")) return sanctusPage;
  if (b.includes("悼念亡者")) return memorialPage;
  if (b === "继续" && pageTitle.includes("感恩经（三）")) return memorialSkipPage;
  if (b.startsWith("常年期主日颂谢词")) return prefaceFirst[0];
  return explicitOnly ? undefined : findJumpTarget(btn);
};

let jmpSeq = 0;
const setRelTarget = (k, rid, pos) => {
  k.rels = k.rels.replace(new RegExp(`(<Relationship\\b[^>]*Id="${rid}"[^>]*Target=")[^"]*(")`), `$1slide${pos}.xml$2`);
};
const addSlideRel = (k, rid, pos) => {
  k.rels = k.rels.replace(
    "</Relationships>",
    `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${pos}.xml"/></Relationships>`
  );
};

// ── 圣圣圣页补齐 4 个感恩经跳转按钮 ──
// 模板圣圣圣页只有 2 个按钮（→ 感恩经一、三）。这里克隆「请大家感谢主」页的 4 个数字按钮
// （视觉从左到右 = 一二三四），替换圣圣圣页原有 2 个按钮，随后由下面的重映射循环把它们指到
// 感恩经（一~四）首页。用户要求：圣圣圣页导航要和「请大家感谢主」页一样有 1/2/3/4。
(() => {
  const fullText = (xml) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join("").replace(/\s/g, "");
  const jumpPicsOf = (xml) => [...xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)].map((m) => m[0]).filter((p) => p.includes("hlinksldjump"));
  const relTargetOf = (rels, rid) => (new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(rels) || [])[1];
  const donor = kept.find((k) => /请大家感谢主/.test(fullText(k.xml)) && jumpPicsOf(k.xml).length >= 4);
  if (!donor) return; // 没有 4 按钮供体（异常文档）→ 保持原状，圣圣圣页仍走 2 按钮兜底
  const specs = jumpPicsOf(donor.xml)
    .map((p) => {
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"/.exec(p);
      return {
        block: p,
        png: relTargetOf(donor.rels, /<a:blip r:embed="([^"]+)"/.exec(p)?.[1]),
        svg: relTargetOf(donor.rels, /<asvg:svgBlip[^>]*r:embed="([^"]+)"/.exec(p)?.[1]),
        x: +(off?.[1] || 0),
      };
    })
    .filter((s) => s.png)
    .sort((a, b) => a.x - b.x) // 左→右 = 感恩经一二三四
    .slice(0, 4);
  if (specs.length < 4) return;
  // 右下角紧凑布局（slide 9144000×6858000 EMU），按用户 mockup：
  //   底行：[感恩经标签] [1][2][3][4] 右对齐、贴底；「下跪」按钮放在数字行上方。
  //   三者都缩小并加半透明填充,让后面的英文透出来、互不重叠、且不超出页面。
  // 所有控件都收进绿色英文框内（English box 5605670..8923284 × 1109893..6634129,留内边距）。
  const BOX_R = 8863284; // 框右 8923284 - 60000
  const BOX_B = 6584129; // 框底 6634129 - 50000
  const NB = 340000; // 数字按钮边长（再缩小以塞进框内）
  const NP = 372000; // 数字按钮间距
  const numW = 3 * NP + NB; // 4 个数字总宽
  const numX0 = BOX_R - numW; // 第一个数字 x（底行右对齐到框内右边）
  const numY = BOX_B - NB; // 底行贴框内底
  const CAP_W = 920000; // 加宽,保证「感恩经」三字一行不换行
  const CAP_H = 340000;
  const capX = numX0 - 60000 - CAP_W; // 感恩经标签在数字左侧
  const capY = numY; // 与数字同高
  const KN_W = CAP_W; // 下跪药丸与感恩经同尺寸（视觉一致）
  const KN_H = CAP_H;
  const knX = BOX_R - KN_W; // 数字行上方、右对齐（框内）
  const knY = numY - 40000 - KN_H;
  const ICON_W = 360000; // 下跪图标
  const ICON_H = 330000;
  const iconX = 7433919 + 90000; // 图标在下跪组左侧（子坐标系）
  const iconY = 5004791 + Math.round((599878 - ICON_H) / 2); // 子坐标系内垂直居中
  const addAlpha = (s) => s.replace(/(<a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}")\/>(<\/a:solidFill>)/g, '$1><a:alpha val="52000"/></a:srgbClr>$2');
  for (let ki = 0; ki < kept.length; ki += 1) {
    const k = kept[ki];
    if (!k.rels || !pageSig[ki].body.startsWith("圣、圣、圣")) continue;
    const old = jumpPicsOf(k.xml);
    if (old.length !== 2) continue; // 只补带 2 按钮的那张圣圣圣页；无按钮的圣圣圣页保持原样
    for (const p of old) k.xml = k.xml.replace(p, ""); // 删除原有 2 个按钮
    let cid = 9100;
    let seq = 0;
    const newPics = specs
      .map((s, i) => {
        const pngRid = `rIdEP${(seq += 1)}`;
        const svgRid = `rIdEP${(seq += 1)}`;
        const hRid = `rIdEP${(seq += 1)}`;
        k.rels = k.rels.replace(
          "</Relationships>",
          `<Relationship Id="${pngRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${s.png}"/>` +
            (s.svg ? `<Relationship Id="${svgRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${s.svg}"/>` : "") +
            `<Relationship Id="${hRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide1.xml"/></Relationships>`
        );
        return s.block
          .replace(/(<p:cNvPr id=")\d+("\s+name=")[^"]*(")/, `$1${(cid += 1)}$2感恩经按钮$3`)
          .replace(/(<a:hlinkClick r:id=")[^"]*(")/, `$1${hRid}$2`)
          .replace(/(<a:blip r:embed=")[^"]*(")/, `$1${pngRid}$2`)
          .replace(/(<a:blip r:embed="[^"]*">)/, '$1<a:alphaModFix amt="78000"/>') // 半透明,透出英文
          .replace(/(<asvg:svgBlip[^>]*r:embed=")[^"]*(")/, `$1${svgRid}$2`)
          .replace(/<a:off x="-?\d+" y="-?\d+"\/>/, `<a:off x="${numX0 + i * NP}" y="${numY}"/>`) // 底行右对齐
          .replace(/<a:ext cx="\d+" cy="\d+"\/>/, `<a:ext cx="${NB}" cy="${NB}"/>`); // 缩小
      })
      .join("");
    k.xml = k.xml.replace("</p:spTree>", `${newPics}</p:spTree>`);
    // 「感恩经」标签：移到数字左侧、缩小、缩字、半透明。
    const capSp = [...k.xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)]
      .map((m) => m[0])
      .find((sp) => [...sp.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join("").replace(/\s/g, "") === "感恩经");
    if (capSp) {
      let n = capSp
        .replace(/<a:off x="-?\d+" y="-?\d+"\/>/, `<a:off x="${capX}" y="${capY}"/>`)
        .replace(/<a:ext cx="\d+" cy="\d+"\/>/, `<a:ext cx="${CAP_W}" cy="${CAP_H}"/>`)
        .replace(/(<a:rPr lang="zh-TW" altLang="en-US" )(b="1")/, '$1sz="1500" $2'); // 缩字
      n = addAlpha(n);
      k.xml = k.xml.replace(capSp, n);
    }
    // 「下跪」按钮（grpSp）：缩小并收进绿框、移到数字行上方；绿底半透明 + 白字 + 字号都与
    // 「感恩经」一致（用户要求）；只去掉图标那圈绿描边,保留提亮(浅色图标配白字)。
    const knGrp = [...k.xml.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)]
      .map((m) => m[0])
      .find((g) => /下跪/.test(g));
    if (knGrp) {
      const n = knGrp
        .replace(/(<p:grpSpPr><a:xfrm[^>]*>)<a:off x="-?\d+" y="-?\d+"\/><a:ext cx="\d+" cy="\d+"\/>/, `$1<a:off x="${knX}" y="${knY}"/><a:ext cx="${KN_W}" cy="${KN_H}"/>`) // 缩小 + 移进框内
        .replace(/(<a:chExt cx="\d+" cy="\d+"\/><\/a:xfrm>)<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"\/><\/a:solidFill>/, '$1<a:solidFill><a:srgbClr val="$2"><a:alpha val="52000"/></a:srgbClr></a:solidFill>') // 绿底半透明,与感恩经一致
        .replace(/<a:ln><a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"\/><\/a:solidFill><\/a:ln>/g, "<a:ln><a:noFill/></a:ln>") // 去掉图标那圈绿描边
        .replace(/(<a:rPr lang="zh-TW" altLang="en-US" )sz="2400"( b="1")/, '$1sz="1500"$2') // 字号与感恩经一致(15pt)
        .replace(/(<p:pic>[\s\S]*?<a:off )x="-?\d+" y="-?\d+"(\/><a:ext )cx="\d+" cy="\d+"/, `$1x="${iconX}" y="${iconY}"$2cx="${ICON_W}" cy="${ICON_H}"`); // 图标尺寸
      k.xml = k.xml.replace(knGrp, n);
    }
  }
})();

for (let ki = 0; ki < kept.length; ki += 1) {
  const k = kept[ki];
  if (!k.rels) continue;
  const sig = pageSig[ki];
  const allText = [...k.xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join("").replace(/\s/g, "");

  // ── 图片按钮（数字图标，无文字，文字匹配无从下手）：按页面身份 + 视觉从左到右顺序解析。
  const pics = [...k.xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)]
    .filter((m) => m[0].includes("hlinksldjump"))
    .map((m) => {
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"/.exec(m[0]);
      return { block: m[0], rid: /<a:hlinkClick[^>]*r:id="([^"]+)"/.exec(m[0])?.[1], x: +(off?.[1] || 0), y: +(off?.[2] || 0) };
    })
    .sort((a, b) => (Math.abs(a.y - b.y) > 457200 ? a.y - b.y : a.x - b.x));
  let picTargets;
  if (pics.length >= 8 && allText.includes("常年期主日颂谢词")) picTargets = prefaceFirst; // 颂谢词目录：1~8 → 颂谢词（一~八）首页
  else if (pics.length === 4 && sig.body.startsWith("圣、圣、圣")) picTargets = epFirst; // 圣圣圣页补齐后：4 个 → 感恩经（一~四）首页
  else if (pics.length === 2 && sig.body.startsWith("圣、圣、圣")) picTargets = [epFirst[0], epFirst[2]]; // 兜底：未补齐时仍映感恩经（一）/（三）
  else if (pics.length === 4 && sig.title.startsWith("感恩经")) picTargets = epFirst; // 感恩经目录：1~4 → 感恩经（一~四）首页
  if (picTargets) {
    const assigned = new Map(); // rid → pos（模板有两个按钮共用一个 rId 的笔误，需拆开补新关系）
    for (let i = 0; i < pics.length; i += 1) {
      const p = pics[i];
      const pos = picTargets[i];
      if (!pos || !p.rid) continue;
      if (!assigned.has(p.rid)) {
        setRelTarget(k, p.rid, pos);
        assigned.set(p.rid, pos);
      } else if (assigned.get(p.rid) !== pos) {
        jmpSeq += 1;
        const nid = `rIdJmp${jmpSeq}`;
        k.xml = k.xml.replace(p.block, p.block.replace(/(<a:hlinkClick[^>]*r:id=")[^"]*(")/, `$1${nid}$2`));
        addSlideRel(k, nid, pos);
      }
    }
  }

  // ── 文字按钮：有 rId 的改写关系；模板断链的（noaction / 空 rId，如圣三颂箭头）显式命中才激活。
  for (const sp of k.xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const hl = /<a:hlinkClick[^>]*\/>/.exec(sp[0])?.[0];
    if (!hl) continue;
    const isJump = hl.includes("hlinksldjump");
    const isNoact = hl.includes("ppaction://noaction");
    if (!isJump && !isNoact) continue;
    const rid = /r:id="([^"]*)"/.exec(hl)?.[1];
    const btn = paraTextOf(sp[0]);
    const pos = textTarget(btn, sig.title, !(isJump && rid));
    if (!pos) continue;
    if (isJump && rid) {
      setRelTarget(k, rid, pos);
    } else {
      jmpSeq += 1;
      const nid = `rIdJmp${jmpSeq}`;
      k.xml = k.xml.replace(sp[0], sp[0].replace(hl, `<a:hlinkClick r:id="${nid}" action="ppaction://hlinksldjump"/>`));
      addSlideRel(k, nid, pos);
    }
  }
}

// ── 待填清单：扫描最终 deck，凡带红字「请手动填入」提示的页 → {最终页码, 段落名} ──
// （这些红字页多为模板静态页 passthrough，不走 token 分支，故在成品定稿后统一扫描最可靠。）
const FILL_KEYS = ["进堂咏", "奉献咏", "领主咏", "礼成咏", "堂区报告"];
kept.forEach((s, idx) => {
  const t = [...s.xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join("");
  if (!/请手动填入|请在此手动填入|手动填入/.test(t)) return;
  const label = FILL_KEYS.find((k) => t.includes(k)) || "请手动填入";
  fillList.push({ page: idx + 1, label });
});

// ───────────────────────── 重新打包（重新编号 + 重建 rels/presentation/content-types/app） ─────────────────────────
for (const f of readdirSync(slideDir)) if (/^slide\d+\.xml$/.test(f)) rmSync(join(slideDir, f));
for (const f of readdirSync(relsDir)) if (/^slide\d+\.xml\.rels$/.test(f)) rmSync(join(relsDir, f));
kept.forEach((s, idx) => {
  const n = idx + 1;
  writeFileSync(join(slideDir, `slide${n}.xml`), s.xml);
  writeFileSync(
    join(relsDir, `slide${n}.xml.rels`),
    s.rels || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
  );
});

let pres = readFileSync(join(work, "ppt/presentation.xml"), "utf8");
const sldIds = kept.map((_s, i) => `<p:sldId id="${2147483648 - kept.length + i}" r:id="rIdSlide${i + 1}"/>`).join("");
pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIds}</p:sldIdLst>`);
writeFileSync(join(work, "ppt/presentation.xml"), pres);

let prels = readFileSync(join(work, "ppt/_rels/presentation.xml.rels"), "utf8");
const nonSlide = [...prels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).filter((r) => !r.includes('/relationships/slide"'));
const slideRelList = kept.map((_s, i) => `<Relationship Id="rIdSlide${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`);
prels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nonSlide.join("")}${slideRelList.join("")}</Relationships>`;
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

console.log(JSON.stringify({ out: outPptx, slides: kept.length, from: slideCount, filled, stillMissing, fillList }, null, 2));
