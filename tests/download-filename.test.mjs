import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { once } from "node:events";
import { rmSync } from "node:fs";
import { createServer } from "node:net";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

test("generated downloads use the document title without a version suffix", { timeout: 20000 }, async (t) => {
  const reservation = createServer();
  await new Promise((resolve) => reservation.listen(0, "127.0.0.1", resolve));
  const port = reservation.address().port;
  await new Promise((resolve) => reservation.close(resolve));
  const child = spawn(process.execPath, [join(root, "server.mjs")], {
    cwd: root,
    env: { ...process.env, HOST: "127.0.0.1", PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
  });
  t.after(async () => {
    if (child.exitCode === null) {
      child.kill();
      await once(child, "exit");
    }
  });
  await Promise.race([
    once(child.stdout, "data"),
    once(child, "exit").then(([code]) => { throw new Error(`Server exited with code ${code}`); }),
  ]);

  const form = new FormData();
  form.set("document", new Blob(["常年期第十五主日（甲年）\n集祷经\n测试祷文。"]), `filename-test-${randomUUID()}.txt`);
  const response = await fetch(`http://127.0.0.1:${port}/generate`, { method: "POST", body: form });
  assert.equal(response.status, 200);
  const result = await response.json();
  assert.equal(result.ok, true);
  const runName = basename(decodeURIComponent(result.downloadUrl.split("/")[2]));
  t.after(() => rmSync(join(root, "runs", runName), { recursive: true, force: true }));

  assert.equal(result.filename, "常年期第十五主日.pptx");
  assert.equal(decodeURIComponent(result.downloadUrl.split("/").at(-1)), result.filename);
  const download = await fetch(`http://127.0.0.1:${port}${result.downloadUrl}`);
  assert.equal(download.status, 200);
  const disposition = download.headers.get("content-disposition");
  assert.equal(decodeURIComponent(disposition.split("filename*=UTF-8''")[1]), result.filename);
  assert.ok((await download.arrayBuffer()).byteLength > 0);
});
