const form = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#document");
const fileName = document.querySelector("#fileName");
const button = document.querySelector("#generateButton");
const statusText = document.querySelector("#statusText");
const stateDot = document.querySelector("#stateDot");
const downloadLink = document.querySelector("#downloadLink");
const details = document.querySelector("#details");
const titleValue = document.querySelector("#titleValue");
const slidesValue = document.querySelector("#slidesValue");
const missingValue = document.querySelector("#missingValue");
const fillListWrap = document.querySelector("#fillListWrap");
const fillListValue = document.querySelector("#fillListValue");
const swatches = document.querySelector("#swatches");
const seasons = document.querySelector("#seasons");
const colorInput = document.querySelector("#themeColor");
const colorWell = document.querySelector("#colorWell");
const hexInput = document.querySelector("#hexInput");
let currentDownloadUrl = "";

function setState(kind, text) {
  stateDot.className = `dot ${kind}`;
  statusText.textContent = text;
}

function normalizeHexInput(raw) {
  let hex = String(raw || "").trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{3}$/.test(hex)) hex = hex.split("").map((c) => c + c).join("");
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex.toUpperCase()}` : "";
}

function syncActiveSwatch(hex) {
  const target = hex.toUpperCase();
  swatches.querySelectorAll(".swatch").forEach((btn) => {
    btn.classList.toggle("is-active", (btn.dataset.color || "").toUpperCase() === target);
  });
}

function syncActiveSeason(hex) {
  const target = hex.toUpperCase();
  seasons?.querySelectorAll(".season").forEach((btn) => {
    btn.classList.toggle("is-active", (btn.dataset.color || "").toUpperCase() === target);
  });
}

function applyColor(hex, { updateHexField = true } = {}) {
  colorInput.value = hex.toLowerCase();
  colorWell.style.setProperty("--c", hex);
  if (updateHexField) hexInput.value = hex.toUpperCase();
  hexInput.classList.remove("invalid");
  syncActiveSwatch(hex);
  syncActiveSeason(hex);
}

swatches.addEventListener("click", (event) => {
  const btn = event.target.closest(".swatch");
  if (btn) applyColor(btn.dataset.color);
});

seasons?.addEventListener("click", (event) => {
  const btn = event.target.closest(".season");
  if (btn) applyColor(btn.dataset.color);
});

colorInput.addEventListener("input", () => applyColor(colorInput.value));

hexInput.addEventListener("input", () => {
  const hex = normalizeHexInput(hexInput.value);
  if (hex) applyColor(hex, { updateHexField: false });
  else hexInput.classList.add("invalid");
});

hexInput.addEventListener("blur", () => {
  const hex = normalizeHexInput(hexInput.value);
  applyColor(hex || colorInput.value);
});

// 从礼仪日历跳转过来时带的 ?color=#xxxxxx，自动套用为主题色。
const urlColor = normalizeHexInput(new URLSearchParams(location.search).get("color") || "");
if (urlColor) applyColor(urlColor);

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  fileName.textContent = file ? file.name : "选择中英文文档";
  currentDownloadUrl = "";
  downloadLink.classList.add("hidden");
  details.classList.add("hidden");
  setState(file ? "ready" : "idle", file ? "文档已选择" : "等待文档");
});

downloadLink.addEventListener("click", (event) => {
  if (!currentDownloadUrl) return;
  event.preventDefault();
  window.location.assign(currentDownloadUrl);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = fileInput.files?.[0];
  if (!file) {
    setState("error", "请选择文档");
    return;
  }

  button.disabled = true;
  downloadLink.classList.add("hidden");
  details.classList.add("hidden");
  setState("busy", "正在生成");

  try {
    const data = new FormData(form);
    const response = await fetch("/generate", { method: "POST", body: data });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.error || "生成失败");

    currentDownloadUrl = new URL(result.downloadUrl, window.location.href).href;
    downloadLink.href = currentDownloadUrl;
    downloadLink.removeAttribute("download");
    downloadLink.classList.remove("hidden");
    titleValue.textContent = result.title || "-";
    slidesValue.textContent = result.slides ? `${result.slides}` : "-";
    missingValue.textContent = result.missing?.length ? result.missing.join("、") : "无";
    // 待填清单：直接列出「第 X 页：段落」，不必逐页翻找。
    fillListValue.innerHTML = "";
    if (result.fillList?.length) {
      for (const item of result.fillList) {
        const li = document.createElement("li");
        const num = document.createElement("span");
        num.className = "pillNum";
        num.textContent = item.page;
        li.appendChild(num);
        li.appendChild(document.createTextNode(item.label));
        fillListValue.appendChild(li);
      }
      fillListWrap.classList.remove("hidden");
    } else {
      fillListWrap.classList.add("hidden");
    }
    details.classList.remove("hidden");
    setState("ok", "生成完成");
  } catch (error) {
    setState("error", error.message);
  } finally {
    button.disabled = false;
  }
});
