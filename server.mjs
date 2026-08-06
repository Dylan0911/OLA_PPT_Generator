import { createServer } from "node:http";
import { execFileSync } from "node:child_process";
import { createReadStream, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { basename, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { liturgyForDate } from "./liturgy.mjs";

const ROOT = resolve(fileURLToPath(new URL(".", import.meta.url)));
const PUBLIC = join(ROOT, "public");
const ASSETS = join(ROOT, "assets");
const RUNS = join(ROOT, "runs");
const PORT = Number(process.env.PORT || 4177);
const MAX_UPLOAD = 25 * 1024 * 1024;
const FINAL_TEMPLATE_NAME = "OLA_Liturgy_Final_Template.pptx";
// 原型占位符模板：每段一张原型页 + 命名占位符 {{读经一_CN}}…；程序自动克隆分页填充。
// 生成（上传文档→成品）和前端「下载模板」都用它。
const PROTOTYPE_TEMPLATE_NAME = "常年期原型模板_命名占位符.pptx";

mkdirSync(RUNS, { recursive: true });

const SECTION_TITLES = [
  "导言",
  "进堂咏唱",
  "进堂咏",
  "光荣颂",
  "集祷经",
  "读经一",
  "答唱咏",
  "读经二",
  "福音前欢呼词",
  "福音前欢呼",
  "福音",
  "信经",
  "信友祷文",
  "奉献咏",
  "献礼经",
  "领主咏",
  "领圣体后经",
  "隆重降福",
];

const OUTPUT_SECTIONS = [
  ["导言", "Introduction"],
  ["集祷经", "Collect"],
  ["读经一", "First Reading"],
  ["答唱咏", "Responsorial Psalm"],
  ["读经二", "Second Reading"],
  ["福音前欢呼词", "Gospel Acclamation"],
  ["福音", "Gospel"],
  ["信经", "Creed"],
  ["信友祷文", "Prayers of the Faithful"],
  ["献礼经", "Prayer over the Offerings"],
  ["领主咏", "Communion Antiphon"],
  ["领圣体后经", "Prayer after Communion"],
  ["隆重降福", "Solemn Blessing"],
];

function send(res, code, body, type = "text/plain; charset=utf-8") {
  res.writeHead(code, { "content-type": type });
  res.end(body);
}

function json(res, code, body) {
  send(res, code, JSON.stringify(body), "application/json; charset=utf-8");
}

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function readBody(req) {
  return new Promise((resolveBody, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_UPLOAD) {
        reject(new Error("文件太大，请控制在 25MB 内。"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolveBody(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function parseMultipart(buffer, contentType) {
  const boundary = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[1] || contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[2];
  if (!boundary) throw new Error("上传格式不正确。");
  const raw = buffer.toString("latin1");
  const parts = raw.split(`--${boundary}`);
  const fields = {};
  const files = {};

  for (const part of parts) {
    const splitAt = part.indexOf("\r\n\r\n");
    if (splitAt < 0) continue;
    const header = part.slice(0, splitAt);
    let body = part.slice(splitAt + 4);
    if (body.endsWith("\r\n")) body = body.slice(0, -2);
    const disposition = header.match(/content-disposition:[^\n]*/i)?.[0] || "";
    const name = disposition.match(/name="([^"]+)"/)?.[1];
    const filename = disposition.match(/filename="([^"]*)"/)?.[1];
    if (!name) continue;
    if (filename) {
      const decodedFilename = Buffer.from(filename, "latin1").toString("utf8");
      files[name] = {
        filename: basename(decodedFilename),
        data: Buffer.from(body, "latin1"),
      };
    } else {
      fields[name] = Buffer.from(body, "latin1").toString("utf8").trim();
    }
  }
  return { fields, files };
}

function extractDocxText(filePath) {
  const xml = execFileSync("unzip", ["-p", filePath, "word/document.xml"], {
    encoding: "utf8",
    maxBuffer: 80 * 1024 * 1024,
  });
  const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
    .map((p) =>
      [...p[0].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:br\b[^>]*\/>|<w:tab\b[^>]*\/>/g)]
        .map((m) => {
          if (m[0].startsWith("<w:br")) return "\n";
          if (m[0].startsWith("<w:tab")) return " ";
          return decodeXml(m[1]);
        })
        .join("")
    )
    .flatMap((p) => p.split("\n"))
    .map((p) => p.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  return paragraphs.join("\n");
}

function extractText(filePath, originalName) {
  const ext = extname(originalName).toLowerCase();
  if (ext === ".docx") return extractDocxText(filePath);
  return readFileSync(filePath, "utf8");
}

function hasChinese(text) {
  return /[\u3400-\u9fff]/.test(text);
}

const ORDINAL_WORDS = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth", "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth", "Twentieth", "Twenty-first", "Twenty-second", "Twenty-third", "Twenty-fourth", "Twenty-fifth", "Twenty-sixth", "Twenty-seventh", "Twenty-eighth", "Twenty-ninth", "Thirtieth", "Thirty-first", "Thirty-second", "Thirty-third", "Thirty-fourth"];

function cnNumToInt(cn) {
  const map = { 零: 0, 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
  if (!cn) return null;
  if (cn === "十") return 10;
  if (cn.length === 1) return map[cn] ?? null;
  if (cn[0] === "十") return 10 + (map[cn[1]] ?? 0);
  if (cn.includes("十")) {
    const [a, b] = cn.split("十");
    return (map[a] ?? 1) * 10 + (b ? map[b] ?? 0 : 0);
  }
  return map[cn] ?? null;
}

// 从中文主日标题推导英文标题：季节 + 第N主日 → "Nth Sunday in/of <Season>"。
function titleEnglish(titleCn) {
  const t = String(titleCn || "");
  const seasons = [
    [/常年期/, "in Ordinary Time"],
    [/四旬期/, "of Lent"],
    [/复活期/, "of Easter"],
    [/将临期/, "of Advent"],
    [/圣诞期/, "of Christmas"],
  ];
  const season = seasons.find(([re]) => re.test(t));
  const m = t.match(/第([一二三四五六七八九十]+)主日/);
  if (season && m) {
    const n = cnNumToInt(m[1]);
    if (n && ORDINAL_WORDS[n]) return `${ORDINAL_WORDS[n]} Sunday ${season[1]}`;
  }
  return "";
}

function normalizeHeading(line) {
  const clean = line.replace(/\u00a0/g, " ").trim();
  for (const title of SECTION_TITLES) {
    const canonical = title === "福音前欢呼" ? "福音前欢呼词" : title;
    if (clean === title) return { title: canonical, rest: "" };
    if (clean.startsWith(title)) {
      const next = clean.slice(title.length);
      if (!next || /^[\s　（(]/.test(next) || hasChinese(next)) {
        return {
          title: canonical,
          rest: next.trim(),
        };
      }
    }
  }
  return null;
}

function parseDocument(text) {
  const lines = String(text)
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const buckets = new Map();
  let current = null;
  let titleCn = "";
  let titleEn = "";

  for (const line of lines) {
    const hit = normalizeHeading(line);
    if (hit) {
      current = hit.title;
      if (!buckets.has(current)) buckets.set(current, []);
      if (hit.rest) buckets.get(current).push(hit.rest);
      continue;
    }
    if (!current && !titleCn && hasChinese(line)) {
      titleCn = line
        .replace(/[（(].*?[）)]/g, "") // 去（甲年）等括注
        .replace(/[—-]{1,2}.+$/u, "")
        .replace(/[\s ]*[甲乙丙丁]\s*年\s*$/u, "") // 去结尾无括号的「甲年/乙年/丙年」
        .trim();
      titleEn = titleEnglish(line);
      continue;
    }
    if (!current && !titleEn && /[A-Za-z]/.test(line)) {
      titleEn = line;
      continue;
    }
    if (current) buckets.get(current).push(line);
  }

  splitAcclamationFromSecondReading(buckets);

  const sections = OUTPUT_SECTIONS.map(([titleCnSection, titleEnSection]) => {
    const rawLines = buckets.get(titleCnSection) || [];
    const cn = [];
    const en = [];
    for (const line of rawLines) {
      if (!hasChinese(line) && /[A-Za-z]/.test(line)) en.push(line);
      else cn.push(line);
    }
    return {
      titleCn: titleCnSection,
      titleEn: titleEnSection,
      cn: cn.join("\n"),
      en: en.join("\n"),
    };
  });

  const missing = OUTPUT_SECTIONS.map(([title]) => title).filter((title) => !buckets.has(title));
  const noGloria = /不念[“"']?光荣颂|不唱[“"']?光荣颂/.test(text);
  return {
    meta: {
      titleCn: titleCn || "常年期主日",
      titleEn: titleEn || titleEnglish(titleCn) || "",
      includeGloria: buckets.has("光荣颂") && !noGloria,
    },
    sections,
    missing,
  };
}

function splitAcclamationFromSecondReading(buckets) {
  const lines = buckets.get("读经二");
  if (!lines?.length) return;
  if (buckets.get("福音前欢呼词")?.length) return;

  const firstAcclamation = lines.findIndex((line) => /^(阿肋路亚|答[:：]\s*阿肋路亚|福音前欢呼|Alleluia\b)/i.test(line.trim()));
  if (firstAcclamation < 0) return;

  const readingLines = lines.slice(0, firstAcclamation);
  const acclamationLines = lines.slice(firstAcclamation);
  if (!readingLines.length || !acclamationLines.length) return;

  buckets.set("读经二", readingLines);
  buckets.set("福音前欢呼词", acclamationLines);
}

function safeRunName(filename) {
  const stamp = new Date().toISOString().replace(/[-:T.Z]/g, "").slice(0, 14);
  const stem = filename.replace(/\.[^.]+$/, "").replace(/[^\p{Letter}\p{Number}]+/gu, "-").replace(/^-|-$/g, "") || "liturgy";
  return `${stamp}-${stem}`;
}

function normalizeThemeColor(value) {
  if (!value) return "";
  let hex = String(value).trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{3}$/.test(hex)) hex = hex.split("").map((c) => c + c).join("");
  return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : "";
}

async function handleGenerate(req, res) {
  const body = await readBody(req);
  const { fields, files } = parseMultipart(body, req.headers["content-type"] || "");
  const uploaded = files.document;
  if (!uploaded?.data?.length) throw new Error("请选择一个文档。");

  const runName = safeRunName(uploaded.filename);
  const runDir = join(RUNS, runName);
  mkdirSync(runDir, { recursive: true });

  const inputPath = join(runDir, uploaded.filename || "input.txt");
  writeFileSync(inputPath, uploaded.data);

  const text = extractText(inputPath, uploaded.filename || "");
  const payload = parseDocument(text);
  if (Object.hasOwn(fields, "includeGloria")) {
    payload.meta.includeGloria = fields.includeGloria === "true" || fields.includeGloria === "on";
  }
  const themeColor = normalizeThemeColor(fields.themeColor);
  if (themeColor) payload.meta.themeColor = themeColor;
  const payloadPath = join(runDir, "sections.json");
  writeFileSync(payloadPath, JSON.stringify(payload, null, 2));

  // 成品命名 = 文档标题 + 「_生成版」（节期/主日与上传文档一致），如「常年期第十二主日_生成版.pptx」。
  const titleForName =
    String(payload.meta.titleCn || "")
      .replace(/[（(][^）)]*[）)]/g, "") // 去（甲年）等括注
      .replace(/[\\/:*?"<>|]/g, "") // 去文件系统非法字符
      .trim() || "弥撒";
  const outputName = `${titleForName}_生成版.pptx`;
  const outputPath = join(runDir, outputName);
  const stdout = execFileSync(
    process.execPath,
    [join(ROOT, "scripts/build-from-template.mjs"), join(ASSETS, PROTOTYPE_TEMPLATE_NAME), runDir, outputPath, payloadPath],
    { encoding: "utf8", maxBuffer: 20 * 1024 * 1024 }
  );
  const generated = JSON.parse(stdout.match(/\{[\s\S]*\}/)?.[0] || "{}");

  json(res, 200, {
    ok: true,
    filename: outputName,
    downloadUrl: `/download/${encodeURIComponent(runName)}/${encodeURIComponent(outputName)}`,
    slides: generated.slides,
    missing: payload.missing,
    fillList: generated.fillList || [], // 待填清单：[{page, label}]，前端提示「第 X 页需手动填入」
    title: `${payload.meta.titleCn} / ${payload.meta.titleEn}`,
    includeGloria: payload.meta.includeGloria,
  });
}

function download(req, res) {
  const [, , run = "", file = ""] = req.url.split("/");
  const runName = decodeURIComponent(run);
  const fileName = decodeURIComponent(file.split("?")[0]);
  const filePath = join(RUNS, basename(runName), basename(fileName));
  if (!existsSync(filePath)) {
    send(res, 404, "File not found");
    return;
  }
  const filename = basename(filePath);
  const asciiFallback = filename.replace(/[^\x20-\x7e]/g, "_");
  const stat = statSync(filePath);
  res.writeHead(200, {
    "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "content-length": stat.size,
    "cache-control": "no-store",
    "content-disposition": `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encodeURIComponent(filename)}`,
  });
  createReadStream(filePath).pipe(res);
}

function downloadTemplate(_req, res) {
  const filePath = join(ASSETS, PROTOTYPE_TEMPLATE_NAME);
  if (!existsSync(filePath)) {
    send(res, 404, "Template not found");
    return;
  }
  const stat = statSync(filePath);
  const asciiFallback = PROTOTYPE_TEMPLATE_NAME.replace(/[^\x20-\x7e]/g, "_");
  res.writeHead(200, {
    "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "content-length": stat.size,
    "cache-control": "no-store",
    "content-disposition": `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encodeURIComponent(PROTOTYPE_TEMPLATE_NAME)}`,
  });
  createReadStream(filePath).pipe(res);
}

function serveStatic(req, res) {
  const pathname = req.url.split("?")[0];
  const path = pathname === "/" ? "/index.html" : pathname;
  const filePath = join(PUBLIC, basename(path));
  const ext = extname(filePath);
  const types = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
  };
  if (!existsSync(filePath)) {
    send(res, 404, "Not found");
    return;
  }
  res.setHeader("Cache-Control", "no-cache, must-revalidate");
  send(res, 200, readFileSync(filePath), types[ext] || "application/octet-stream");
}

async function handleLiturgy(req, res) {
  const date = new URL(req.url, "http://x").searchParams.get("date") || "";
  const data = await liturgyForDate(date);
  res.setHeader("Cache-Control", "no-cache");
  json(res, 200, data);
}

const server = createServer(async (req, res) => {
  try {
    if (req.method === "POST" && req.url === "/generate") return await handleGenerate(req, res);
    if (req.method === "GET" && req.url.startsWith("/api/liturgy")) return await handleLiturgy(req, res);
    if (req.method === "GET" && req.url.startsWith("/download/")) return download(req, res);
    if (req.method === "GET" && req.url === "/template") return downloadTemplate(req, res);
    if (req.method === "GET") return serveStatic(req, res);
    send(res, 405, "Method not allowed");
  } catch (error) {
    json(res, 500, { ok: false, error: error.message });
  }
});

const HOST = process.env.HOST || "127.0.0.1"; // 部署时由主机设为 0.0.0.0 对外；本地默认仅本机
server.listen(PORT, HOST, () => {
  console.log(`Liturgy PPT tool running at http://${HOST}:${PORT}`);
});
