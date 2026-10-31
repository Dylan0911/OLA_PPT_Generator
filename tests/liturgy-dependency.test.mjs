import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { liturgyForDate } from "../liturgy.mjs";

const require = createRequire(import.meta.url);

test("the calendar uses a Moment version patched for GHSA-4p3w-j4w9-5jqw", () => {
  const [major, minor] = require("moment/package.json").version.split(".").map(Number);
  assert.ok(major > 2 || (major === 2 && minor >= 31), "Moment must be at least 2.31.0");
});

test("calendar dependency updates preserve holiday names, seasons and colors", async () => {
  for (const [date, nameCn, seasonCn, colorHex] of [
    ["2026-10-31", "常年期平日", "常年期", "#4F7E63"],
    ["2026-12-25", "圣诞节", "圣诞期", "#C9A227"],
    ["2026-04-05", "复活节主日", "复活期", "#C9A227"],
  ]) {
    const result = await liturgyForDate(date);
    assert.deepEqual([result.date, result.nameCn, result.seasonCn, result.colorHex], [date, nameCn, seasonCn, colorHex]);
  }
});
