import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

test("each Kyrie response stays on a single line in the English column", () => {
  const work = mkdtempSync(join(tmpdir(), "ola-kyrie-layout-"));
  try {
    const payload = join(work, "sections.json");
    const output = join(work, "kyrie.pptx");
    writeFileSync(payload, JSON.stringify({ meta: {}, sections: [] }));
    execFileSync(process.execPath, [join(root, "scripts/build-from-template.mjs"), join(root, "assets/常年期原型模板_命名占位符.pptx"), work, output, payload]);
    const xml = execFileSync("unzip", ["-p", output, "ppt/slides/*.xml"], { encoding: "utf8", maxBuffer: 30 * 1024 * 1024 });
    const shapes = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((match) => match[0]);
    const english = shapes.filter((shape) => shape.includes("Lord, have mercy."));
    assert.equal(english.length, 1);
    const paragraphs = [...english[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((paragraph) => [...paragraph[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1]).join("")).filter(Boolean);
    assert.deepEqual(paragraphs, ["Lord, have mercy.", "Christ, have mercy.", "Lord, have mercy."]);
    assert.match(english[0], /<a:bodyPr\b[^>]*\bwrap="none"/, "the three short responses must not wrap inside the narrow English column");
    assert.doesNotMatch(english[0], /<a:br\b/);
    const otherEnglish = shapes.find((shape) => shape.includes("In the name of the Father"));
    assert.ok(otherEnglish, "fixture must include a normal English paragraph");
    assert.doesNotMatch(otherEnglish, /wrap="none"/, "other paragraphs must retain their existing wrapping behavior");
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
});
