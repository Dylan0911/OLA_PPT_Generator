import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

test("page scrim keeps the church background faint behind the tool", async () => {
  const css = await readFile(new URL("../public/styles.css", import.meta.url), "utf8");
  const overlay = css.match(/body::after\s*\{([\s\S]*?)\}/)?.[1] ?? "";
  const stops = [...overlay.matchAll(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*(0?\.\d+)\)/g)]
    .map((match) => match.slice(1).map(Number));
  const alphaValues = stops.map(([, , , alpha]) => alpha);

  assert.equal(stops.length, 2, "the overlay should use a two-stop black gradient");
  assert.ok(stops.every(([r, g, b]) => r === 0 && g === 0 && b === 0), "the overlay should be black");
  assert.ok(Math.min(...alphaValues) >= 0.58, "the background should stay visually secondary");
  assert.ok(Math.max(...alphaValues) <= 0.74, "the church should remain faintly visible");
  assert.ok(alphaValues[1] > alphaValues[0], "the overlay should deepen toward the bottom");
});
