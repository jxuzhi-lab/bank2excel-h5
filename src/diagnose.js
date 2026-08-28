// diagnose.js —— 未识别格式的诊断包导出引导
import { downloadText } from "./download.js";

export function showDiagnoseOffer(container, { pdfBytes, filename, rawError, engine, contactHTML }) {
  const box = document.createElement("div");
  box.className = "diag-offer";
  box.innerHTML = `
    <p><b>未能识别该对账单格式。</b></p>
    <p class="muted">原因: ${escapeHtml(rawError || "表头/记录识别失败")}</p>
    <p>你可以导出一份<b>脱敏诊断包</b>(几 KB JSON, 金额与账号已替换为占位符, 不含真实数据),
       发给维护者适配后, 下个版本即可支持该格式。</p>
    <div class="row">
      <button class="btn diag-btn">导出脱敏诊断包</button>
      <span class="muted contact">${contactHTML || ""}</span>
    </div>
    <p class="diag-done muted" style="display:none">✅ 已导出: <span class="diag-name"></span></p>
  `;
  container.appendChild(box);
  box.querySelector(".diag-btn").addEventListener("click", async (ev) => {
    ev.target.disabled = true;
    ev.target.textContent = "生成中...";
    try {
      engine.diagnose(pdfBytes, filename, {
        onDone: (json) => {
          const name = `bank2excel-diagnosis-${Date.now()}.json`;
          downloadText(json, name);
          const done = box.querySelector(".diag-done");
          done.style.display = "block";
          box.querySelector(".diag-name").textContent = name;
          ev.target.textContent = "已导出 ✓";
        },
        onError: (e) => {
          ev.target.disabled = false;
          ev.target.textContent = "导出失败, 点击重试";
          alert("诊断包生成失败: " + e.message);
        },
      });
    } catch (e) {
      ev.target.disabled = false;
      ev.target.textContent = "导出失败, 点击重试";
    }
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}
