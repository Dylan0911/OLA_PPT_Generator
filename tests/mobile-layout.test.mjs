import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

test("mobile layout keeps the primary controls horizontal and compact", async () => {
  const css = await readFile(new URL("../public/styles.css", import.meta.url), "utf8");
  const mobile = css.slice(css.indexOf("@media (max-width: 760px)"));

  assert.match(mobile, /\.brandLockup\s*\{[\s\S]*?flex-direction:\s*row;/);
  assert.match(mobile, /\.upload\s*\{[\s\S]*?grid-template-columns:\s*minmax\(0,\s*1fr\)\s+108px;[\s\S]*?align-items:\s*center;/);
  assert.match(mobile, /\.dropzone\s*\{[\s\S]*?min-height:\s*0;[\s\S]*?height:\s*112px;[\s\S]*?box-sizing:\s*border-box;/);
  assert.match(mobile, /\.actions\s*\{[\s\S]*?height:\s*112px;/);
  assert.match(mobile, /\.seasonRow\s*\{[\s\S]*?align-items:\s*center;[\s\S]*?flex-wrap:\s*nowrap;[\s\S]*?overflow-x:\s*auto;/);
  assert.match(mobile, /\.swatches\s*\{[\s\S]*?align-items:\s*center;[\s\S]*?flex-wrap:\s*nowrap;[\s\S]*?overflow-x:\s*auto;/);
});
