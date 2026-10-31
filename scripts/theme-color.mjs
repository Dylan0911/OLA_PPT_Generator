import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { basename, dirname, extname, join, resolve } from "node:path";
import { PNG } from "pngjs";

function normalizeColor(value) {
  let hex = String(value || "").trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) hex = hex.split("").map((c) => c + c).join("");
  return /^[0-9a-f]{6}$/i.test(hex) ? hex.toUpperCase() : "";
}

function mixColor(hex, target = 255, amount = 0.8) {
  return [0, 2, 4].map((offset) => {
    const channel = parseInt(hex.slice(offset, offset + 2), 16);
    return Math.round(channel + (target - channel) * amount).toString(16).padStart(2, "0");
  }).join("").toUpperCase();
}

function pngFallbacks(partPath, xml, themedSvgs) {
  const relsPath = join(dirname(partPath), "_rels", `${basename(partPath)}.rels`);
  if (!existsSync(relsPath)) return [];
  const targets = new Map();
  for (const match of readFileSync(relsPath, "utf8").matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = match[0].match(/\bId="([^"]+)"/)?.[1];
    const target = match[0].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target && !/TargetMode="External"/.test(match[0])) targets.set(id, resolve(dirname(partPath), target));
  }
  const fallbackPaths = [];
  for (const picture of xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)) {
    const assets = [...picture[0].matchAll(/\br:embed="([^"]+)"/g)].map((match) => targets.get(match[1])).filter(Boolean);
    if (assets.some((asset) => themedSvgs.has(asset))) fallbackPaths.push(...assets.filter((asset) => extname(asset).toLowerCase() === ".png"));
  }
  return fallbackPaths;
}

function recolorPng(filePath, palette) {
  const image = PNG.sync.read(readFileSync(filePath));
  const channels = (hex) => [0, 2, 4].map((offset) => parseInt(hex.slice(offset, offset + 2), 16));
  const colors = [...palette].map(([from, to]) => [channels(from), channels(to)]);
  let changed = false;
  for (let i = 0; i < image.data.length; i += 4) {
    const alpha = image.data[i + 3];
    if (alpha < 16) continue;
    // Transparent antialiased edges can have small RGB rounding differences.
    const tolerance = Math.max(2, Math.ceil(255 / alpha));
    for (const [from, to] of colors) {
      if (!from.every((channel, index) => Math.abs(image.data[i + index] - channel) <= tolerance)) continue;
      image.data.set(to, i);
      changed = true;
      break;
    }
  }
  if (changed) writeFileSync(filePath, PNG.sync.write(image));
}

export function applyThemeToPackage(packageDir, requestedColor) {
  const primary = normalizeColor(requestedColor);
  if (!primary || primary === "55725D") return;
  const palette = new Map([
    ["55725D", primary],
    ["BCCCBE", mixColor(primary)],
    ["767F77", mixColor(primary, 128)],
    ["BF913B", primary],
    ["F2E1AC", mixColor(primary)],
  ]);

  const themedSvgs = new Set();
  const mediaDir = join(packageDir, "ppt/media");
  if (existsSync(mediaDir)) {
    for (const file of readdirSync(mediaDir)) {
      if (!file.endsWith(".svg")) continue;
      const filePath = resolve(mediaDir, file);
      const original = readFileSync(filePath, "utf8");
      const recolored = original.replace(/#([0-9a-f]{6})\b/gi, (match, hex) => {
        const replacement = palette.get(hex.toUpperCase());
        return replacement ? `#${replacement}` : match;
      });
      if (recolored === original) continue;
      writeFileSync(filePath, recolored);
      themedSvgs.add(filePath);
    }
  }

  const fallbacks = new Set();
  for (const dir of ["ppt/slides", "ppt/slideLayouts", "ppt/slideMasters", "ppt/theme"]) {
    const dirPath = join(packageDir, dir);
    if (!existsSync(dirPath)) continue;
    for (const file of readdirSync(dirPath)) {
      if (!file.endsWith(".xml")) continue;
      const filePath = join(dirPath, file);
      const original = readFileSync(filePath, "utf8");
      for (const fallback of pngFallbacks(filePath, original, themedSvgs)) fallbacks.add(fallback);
      let recolored = original.replace(/(<a:srgbClr\b[^>]*\bval=")([0-9a-f]{6})(")/gi, (match, before, hex, after) => {
        const replacement = palette.get(hex.toUpperCase());
        return replacement ? `${before}${replacement}${after}` : match;
      });
      // Accent 6 supplies the pale English panels; keep their luminance transforms.
      recolored = recolored.replace(/(<a:accent[16]>\s*<a:srgbClr\b[^>]*\bval=")[0-9a-f]{6}(")/gi,
        (_match, before, after) => `${before}${primary}${after}`);
      if (recolored !== original) writeFileSync(filePath, recolored);
    }
  }
  for (const filePath of fallbacks) recolorPng(filePath, palette);
}
