import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

test("page scrim keeps the church background faint behind the tool", async () => {
  const css = await readFile(new URL("../public/styles.css", import.meta.url), "utf8");
  const overlay = css.match(/body::after\s*\{([\s\S]*?)\}/)?.[1] ?? "";
  const alphaValues = [...overlay.matchAll(/rgba\(4,\s*12,\s*24,\s*(0?\.\d+)\)/g)]
    .map((match) => Number(match[1]));

  assert.equal(alphaValues.length, 2, "the overlay should use a two-stop navy gradient");
  assert.ok(Math.min(...alphaValues) >= 0.84, "the background should be visible only faintly");
  assert.ok(Math.max(...alphaValues) <= 0.94, "the background should not disappear completely");
  assert.ok(alphaValues[1] > alphaValues[0], "the overlay should deepen toward the bottom");
});
