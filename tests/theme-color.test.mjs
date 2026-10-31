import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";
import { PNG } from "pngjs";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const template = join(root, "assets/常年期原型模板_命名占位符.pptx");

function themeColors(pptx) {
  const xml = execFileSync("unzip", ["-p", pptx, "ppt/slides/*.xml", "ppt/slideLayouts/*.xml", "ppt/slideMasters/*.xml", "ppt/theme/*.xml"], {
    encoding: "utf8",
    maxBuffer: 30 * 1024 * 1024,
  });
  return [...xml.matchAll(/<a:srgbClr\b[^>]*\bval="([0-9a-f]{6})"/gi)].map((match) => match[1].toUpperCase());
}

function generateColors(themeColor, inspect = themeColors) {
  const work = mkdtempSync(join(tmpdir(), "ola-theme-color-"));
  try {
    const payload = join(work, "sections.json");
    const output = join(work, "blue.pptx");
    writeFileSync(payload, JSON.stringify({
      meta: { titleCn: "Color regression", titleEn: "Color regression", themeColor },
      sections: [],
    }));
    execFileSync(process.execPath, [join(root, "scripts/build-from-template.mjs"), template, work, output, payload], {
      encoding: "utf8",
      maxBuffer: 10 * 1024 * 1024,
    });
    return inspect(output);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

for (const [input, primary, light] of [
  ["#1F4E79", "1F4E79", "D2DCE4"],
  ["B3262A", "B3262A", "F0D4D4"],
  ["#a1b2c3", "A1B2C3", "ECF0F3"],
  ["#abc", "AABBCC", "EEF1F5"],
]) {
  test(`the active template generator applies ${input} and its light tint`, () => {
    const original = themeColors(template);
    assert.ok(original.includes("55725D"), "fixture must contain the original green theme");
    assert.ok(original.includes("BCCCBE"), "fixture must contain the original light green");
    const generated = generateColors(input);
    assert.ok(generated.includes(primary), "selected color must be present in the generated PPTX");
    assert.ok(generated.includes(light), "light accents must follow the selected color");
    assert.ok(!generated.includes("55725D"), "original green must be replaced");
    assert.ok(!generated.includes("BCCCBE"), "original light green must be replaced");
    assert.ok(generated.includes("282120"), "body text color must be preserved");
    assert.ok(generated.includes("C00000"), "manual-entry warning color must be preserved");
  });
}

for (const color of [undefined, "#55725D", "not-a-color"]) {
  test(`default/invalid color ${color} preserves the original palette`, () => {
    const generated = generateColors(color);
    assert.ok(generated.includes("55725D"));
    assert.ok(generated.includes("BCCCBE"));
  });
}

test("the inherited side-panel color follows the selected gold theme", () => {
  generateColors("#C9A227", (pptx) => {
    const themes = execFileSync("unzip", ["-p", pptx, "ppt/theme/*.xml"], { encoding: "utf8" });
    const panels = execFileSync("unzip", ["-p", pptx, "ppt/slides/*.xml", "ppt/slideLayouts/*.xml"], { encoding: "utf8", maxBuffer: 30 * 1024 * 1024 });
    assert.match(panels, /<a:schemeClr val="accent6"><a:lumMod val="20000"\/><a:lumOff val="80000"\/><\/a:schemeClr>/);
    const accents = [...themes.matchAll(/<a:accent6><a:srgbClr val="([0-9A-F]{6})"\/><\/a:accent6>/g)];
    assert.ok(accents.length > 0, "fixture must contain the theme color used by the side panels");
    assert.ok(accents.every((match) => match[1] === "C9A227"), "the side panels must inherit gold, not the original green");
  });
});

test("muted decoration and embedded vector icons follow the selected blue", () => {
  generateColors("#1F4E79", (pptx) => {
    assert.ok(!themeColors(pptx).includes("767F77"), "muted green decoration must also follow the palette");
    const svg = execFileSync("unzip", ["-p", pptx, "ppt/media/*.svg"], { encoding: "utf8" });
    assert.doesNotMatch(svg, /#(?:55725D|BCCCBE|BF913B)\b/i, "themed vector icons must not keep green or old gold");
  });
});

test("PNG icon fallbacks are recolored without changing transparency or background artwork", () => {
  const readPart = (pptx, name) => execFileSync("unzip", ["-p", pptx, `ppt/media/${name}`], { maxBuffer: 10 * 1024 * 1024 });
  generateColors("#1F4E79", (pptx) => {
    for (const [name, from, to] of [
      ["image17.png", "BCCCBE", "D2DCE4"],
      ["image40.png", "BF913B", "1F4E79"],
    ]) {
      const source = PNG.sync.read(readPart(template, name));
      const output = PNG.sync.read(readPart(pptx, name));
      assert.equal(output.width, source.width);
      assert.equal(output.height, source.height);
      let matched = 0;
      for (let i = 0; i < source.data.length; i += 4) {
        assert.equal(output.data[i + 3], source.data[i + 3], `${name}: alpha must be unchanged`);
        if (source.data[i + 3] !== 255 || source.data.subarray(i, i + 3).toString("hex").toUpperCase() !== from) continue;
        assert.equal(output.data.subarray(i, i + 3).toString("hex").toUpperCase(), to, `${name}: icon retains the old palette`);
        matched += 1;
      }
      assert.ok(matched > 1000, `${name}: fixture must contain a substantial themed area`);
    }
    assert.deepEqual(readPart(pptx, "image1.png"), readPart(template, "image1.png"), "background artwork must remain byte-identical");
  });
});

test("recoloring leaves all text, fonts, shapes, links and other XML unchanged", () => {
  const readXml = (pptx) => execFileSync("unzip", ["-p", pptx, "*.xml"], {
    encoding: "utf8",
    maxBuffer: 30 * 1024 * 1024,
  });
  const original = generateColors(undefined, readXml);
  const recolored = generateColors("#1F4E79", readXml);
  // Recent-color metadata outside the slide hierarchy is deliberately unchanged.
  const normalizePalette = (xml) => xml
    .replace(/(<a:accent[16]>\s*<a:srgbClr\b[^>]*\bval=")[0-9A-F]{6}(")/g, "$1ACCENT$2")
    .replace(/(<a:srgbClr\b[^>]*\bval=")(55725D|1F4E79)(")/g, "$1PRIMARY$3")
    .replace(/(<a:srgbClr\b[^>]*\bval=")(BCCCBE|D2DCE4)(")/g, "$1LIGHT$3")
    .replace(/(<a:srgbClr\b[^>]*\bval=")(767F77|6D767F)(")/g, "$1MUTED$3");
  const expected = normalizePalette(original);
  const actual = normalizePalette(recolored);
  let firstDifference = 0;
  while (firstDifference < expected.length && actual[firstDifference] === expected[firstDifference]) firstDifference += 1;
  assert.ok(actual === expected, `unexpected XML change at ${firstDifference}: ${JSON.stringify(actual.slice(firstDifference - 80, firstDifference + 120))} vs ${JSON.stringify(expected.slice(firstDifference - 80, firstDifference + 120))}`);
});
