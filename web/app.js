/**
 * app.js
 * Main application script for FST 16 Drivetrain Tracker Web Portal.
 * 24/7 web application on GitHub Pages with two-tier access, multi-image gallery,
 * SheetJS Excel export, and complete feature parity.
 */
(function () {
  "use strict";

  // State
  let partsData = [];
  let assembliesData = [];
  let requirementsData = [];
  let trashData = [];

  let currentView = "assembly"; // 'assembly', 'requirements', 'trash', or part name (e.g. 'UP_01')
  let currentPartName = null;
  let selectedEvalVersion = null;
  let selectedRefVersion = null;
  let selectedEvalAssVersion = null;
  let selectedRefAssVersion = null;
  let activeGalleryIndex = 0;
  let reqActiveTab = "list"; // 'list' or 'diagrams'
  let activeFeaLc = "LC1";
  let activePartTab = "tabOverview";
  let newIterLcImages = {};
  let editIterLcImages = {};
  let newIterSchematic = null;
  let editIterSchematic = null;

  // DOM Elements
  const elSidebarNav = document.getElementById("sidebarNav");
  const elContentArea = document.getElementById("contentArea");
  const elUserRoleBar = document.getElementById("userRoleBar");
  const elBtnOpenSync = document.getElementById("btnOpenSync");
  const elModalContainer = document.getElementById("modalContainer");

  // Helper: Two-tier permissions
  function isEditor() {
    return window.DTCloud && window.DTCloud.isConfigured();
  }

  // Multi-image schema helpers
  function parseImagePaths(val) {
    if (!val) return [];
    if (Array.isArray(val)) {
      return val.filter(Boolean).map(s => String(s).trim().replace(/\\/g, "/"));
    }
    const s = String(val).trim();
    if (!s || s === "null" || s === "nan" || s === "[]") return [];
    if (s.startsWith("[") && s.endsWith("]")) {
      try {
        const parsed = JSON.parse(s);
        if (Array.isArray(parsed)) {
          return parsed.filter(Boolean).map(x => String(x).trim().replace(/\\/g, "/"));
        }
      } catch (e) {}
    }
    return [s.replace(/\\/g, "/")];
  }

  function formatImagePaths(arr) {
    if (!arr || !arr.length) return null;
    const clean = arr.map(x => String(x).trim().replace(/\\/g, "/")).filter(Boolean);
    if (!clean.length) return null;
    if (clean.length === 1) return clean[0];
    return JSON.stringify(clean);
  }

  // Version formatting helper
  function formatVersion(v) {
    if (!v) return "";
    const str = String(v).trim();
    if (str.startsWith("v") || str.startsWith("V") || str.startsWith("FST") || str.startsWith("OEM")) {
      return str;
    }
    return `v${str}`;
  }

  // Client-side downscaling for image uploads (max 1920x1080)
  function downscaleImage(fileOrDataUrl, maxW = 1920, maxH = 1080) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        let w = img.width;
        let h = img.height;
        if (w > maxW || h > maxH) {
          const ratio = Math.min(maxW / w, maxH / h);
          w = Math.round(w * ratio);
          h = Math.round(h * ratio);
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/png"));
      };
      img.onerror = reject;

      if (typeof fileOrDataUrl === "string") {
        img.src = fileOrDataUrl;
      } else {
        const reader = new FileReader();
        reader.onload = (e) => (img.src = e.target.result);
        reader.onerror = reject;
        reader.readAsDataURL(fileOrDataUrl);
      }
    });
  }

  // Toast notification helper
  function showToast(message, duration = 3500) {
    const existing = document.getElementById("ocrToastNotification");
    if (existing) existing.remove();

    const toast = document.createElement("div");
    toast.id = "ocrToastNotification";
    toast.className = "ocr-toast";
    toast.innerHTML = `<span>${message}</span>`;
    document.body.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(20px)";
      setTimeout(() => toast.remove(), 300);
    }, duration);
  }

  // Preprocess image on canvas and run OCR via Tesseract.js
  async function extractNtopFeaFromImage(fileOrDataUrl, targetPart) {
    if (typeof Tesseract === "undefined") {
      showToast("OCR Error: Tesseract.js library not loaded.");
      return { detectedLc: null, detectedDisp: null, fullImage: null };
    }

    showToast("Analyzing FEA screenshot with OCR...", 6000);

    return new Promise((resolve) => {
      const img = new Image();
      img.onload = async () => {
        try {
          const w = img.width;
          const h = img.height;

          // Crop top-right legend (w * 0.60 to w, and 0.12 to 0.75 * h)
          const cropX = Math.floor(w * 0.60);
          const cropY = Math.floor(h * 0.12);
          const cropW = w - cropX;
          const cropH = Math.floor(h * 0.65);

          const canvas = document.createElement("canvas");
          const scale = 2.5;
          canvas.width = Math.round(cropW * scale);
          canvas.height = Math.round(cropH * scale);
          const ctx = canvas.getContext("2d");

          // Draw cropped & scaled region
          ctx.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, canvas.width, canvas.height);

          // Invert & enhance contrast to optimize for OCR (matches Python ImageOps.invert)
          const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
          const d = imgData.data;
          for (let i = 0; i < d.length; i += 4) {
            const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
            // Invert: dark background with bright text -> light background with dark text
            const inv = 255 - gray;
            const contrast = (inv - 128) * 1.6 + 128;
            const finalVal = Math.min(255, Math.max(0, contrast));
            d[i] = finalVal;
            d[i + 1] = finalVal;
            d[i + 2] = finalVal;
          }
          ctx.putImageData(imgData, 0, 0);

          // Run OCR recognition
          const res = await Tesseract.recognize(canvas, "eng");
          const text = (res && res.data && res.data.text) ? res.data.text : "";
          console.log("OCR Recognized Text:\n", text);

          // Match Load Case keywords
          const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
          const firstLines = lines.slice(0, 5).join(" ").toLowerCase();
          const nospace = firstLines.replace(/[\s\-_]/g, "");

          let detectedLc = null;
          if (firstLines.includes("pure") && firstLines.includes("brak")) {
            detectedLc = "LC1";
          } else if (firstLines.includes("pure") && firstLines.includes("acc")) {
            detectedLc = "LC2";
          } else if (firstLines.includes("pure") && (firstLines.includes("corn") || firstLines.includes("com"))) {
            detectedLc = "LC3";
          } else if (nospace.includes("b+c2") || nospace.includes("bc2") || (firstLines.includes("brak") && (nospace.includes("cornering2") || nospace.includes("comering2")))) {
            detectedLc = "LC5";
          } else if (nospace.includes("b+c1") || nospace.includes("bc1") || nospace.includes("b+c") || (firstLines.includes("brak") && (first_lines.includes("corn") || first_lines.includes("com")))) {
            detectedLc = "LC4";
          } else if (nospace.includes("a+c2") || nospace.includes("ac2") || (firstLines.includes("acc") && (nospace.includes("cornering2") || nospace.includes("comering2")))) {
            detectedLc = "LC7";
          } else if (nospace.includes("a+c1") || nospace.includes("ac1") || nospace.includes("a+c") || (firstLines.includes("acc") && (first_lines.includes("corn") || first_lines.includes("com")))) {
            detectedLc = "LC6";
          } else if (targetPart !== "HB_01") {
            if (nospace.includes("-1kn") || nospace.includes("-ikn") || nospace.includes("result2") || nospace.includes("res2")) {
              detectedLc = "LC9";
            } else if (nospace.includes("1kn") || nospace.includes("ikn") || nospace.includes("result1")) {
              detectedLc = "LC8";
            }
          }

          // Extract scientific notation or decimal displacement numbers
          const nums = [];
          const sciRegex = /([1-9])[.\-]?(\d{3,5})\s*(?:e\s*[\-+]|[\+&]|\s*-\s*)0([0-4])/gi;
          let match;
          while ((match = sciRegex.exec(text)) !== null) {
            const mant = parseFloat(match[1] + "." + match[2]);
            const exp = parseInt(match[3], 10);
            nums.push(mant * Math.pow(10, -exp));
          }

          if (!nums.length) {
            const decRegex = /\b0\.(\d{3,5})\b/g;
            while ((match = decRegex.exec(text)) !== null) {
              nums.push(parseFloat(match[0]));
            }
          }

          let detectedDisp = null;
          const isHB = targetPart === "HB_01";
          const partDecimals = isHB ? 5 : 4;

          if (nums.length) {
            const rawMax = nums[0]; // Top value of legend scale is peak displacement
            const factor = Math.pow(10, partDecimals);
            detectedDisp = Math.floor(rawMax * factor + 1e-9) / factor;
          }

          // Full image as compressed data URL for saving as LC contour
          const fullCanvas = document.createElement("canvas");
          const maxDim = 1200;
          let outW = w;
          let outH = h;
          if (w > maxDim || h > maxDim) {
            const r = Math.min(maxDim / w, maxDim / h);
            outW = Math.round(w * r);
            outH = Math.round(h * r);
          }
          fullCanvas.width = outW;
          fullCanvas.height = outH;
          const fullCtx = fullCanvas.getContext("2d");
          fullCtx.drawImage(img, 0, 0, outW, outH);
          const fullDataUrl = fullCanvas.toDataURL("image/png");

          if (detectedLc && detectedDisp !== null) {
            showToast(`Detected ${detectedLc}: ${detectedDisp.toFixed(partDecimals)} mm`);
          } else if (detectedDisp !== null) {
            showToast(`Detected displacement: ${detectedDisp.toFixed(partDecimals)} mm`);
          } else if (detectedLc) {
            showToast(`Detected ${detectedLc}, but could not read displacement scale.`);
          }

          resolve({ detectedLc, detectedDisp, fullImage: fullDataUrl });
        } catch (err) {
          console.error("OCR Exception:", err);
          showToast("OCR processing failed: " + err.message);
          resolve({ detectedLc: null, detectedDisp: null, fullImage: null });
        }
      };
      img.onerror = () => {
        showToast("Error loading image for OCR.");
        resolve({ detectedLc: null, detectedDisp: null, fullImage: null });
      };

      if (typeof fileOrDataUrl === "string") {
        img.src = fileOrDataUrl;
      } else {
        const reader = new FileReader();
        reader.onload = (e) => (img.src = e.target.result);
        reader.onerror = () => resolve({ detectedLc: null, detectedDisp: null, fullImage: null });
        reader.readAsDataURL(fileOrDataUrl);
      }
    });
  }

  // Handle OCR from pasted or dropped image
  async function handleFeaScreenshotPasted(fileOrDataUrl, targetPart, mode = "new") {
    const res = await extractNtopFeaFromImage(fileOrDataUrl, targetPart);
    const isHB = targetPart === "HB_01";
    const partDecimals = isHB ? 5 : 4;

    if (res.detectedLc && res.detectedDisp !== null) {
      const inpId = mode === "new" ? `newLC_${res.detectedLc}` : `editLC_${res.detectedLc}`;
      const inp = document.getElementById(inpId);
      if (inp) {
        inp.value = res.detectedDisp.toFixed(partDecimals);
        inp.style.borderColor = "var(--delta-green)";
        inp.style.backgroundColor = "rgba(34, 197, 94, 0.1)";
        setTimeout(() => {
          inp.style.borderColor = "";
          inp.style.backgroundColor = "";
        }, 1800);
      }

      if (mode === "new") {
        newIterLcImages[res.detectedLc] = res.fullImage;
      } else {
        editIterLcImages[res.detectedLc] = res.fullImage;
      }

      const badgeId = `${mode}_badge_${res.detectedLc}`;
      const badge = document.getElementById(badgeId);
      if (badge) {
        badge.style.display = "inline-block";
        badge.textContent = "Attached";
      }
    } else if (res.detectedDisp !== null) {
      showToast(`Detected ${res.detectedDisp.toFixed(partDecimals)} mm. Enter into target LC.`);
    } else if (!res.detectedLc && res.fullImage) {
      // Non-FEA image (e.g. CAD drawing or technical screenshot) -> Attach as CAD schematic!
      showToast("Attaching image as CAD schematic...");
      if (mode === "edit") {
        const evalVer = selectedEvalVersion;
        const row = partsData.find((p) => p["Part Name"] === targetPart && String(p.Version) === String(evalVer));
        if (row) {
          await uploadImagesToIteration([fileOrDataUrl], row, targetPart);
          showToast("CAD schematic attached!");
        } else {
          showToast("Iteration record not found.");
        }
      } else {
        newIterSchematic = res.fullImage;
        const previewCont = document.getElementById("newSchematicPreview");
        if (previewCont) {
          previewCont.innerHTML = `
            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-sm); margin-top: 0.5rem;">
              <img src="${res.fullImage}" style="max-height: 60px; border-radius: 4px;">
              <span class="tag" style="color: var(--delta-green); border-color: var(--delta-green);">Staged</span>
              <button type="button" class="btn-action danger" onclick="window.clearNewSchematic()">Remove</button>
            </div>
          `;
        }
        showToast("CAD schematic staged for new iteration!");
      }
    }
  }

  // Client-side Markdown and KaTeX math renderer
  function renderMarkdown(md) {
    if (!md) return "";
    let str = String(md);

    // 1. Extract and protect Display Math: $$ ... $$
    const mathBlocks = [];
    str = str.replace(/\$\$([\s\S]*?)\$\$/g, (match, expr) => {
      let rendered = "";
      if (typeof katex !== "undefined" && katex.renderToString) {
        try {
          rendered = katex.renderToString(expr.trim(), { displayMode: true, throwOnError: false });
        } catch (e) {
          rendered = '<div class="katex-display-error">$$' + expr + '$$</div>';
        }
      } else {
        rendered = '<div class="katex-display-raw">$$' + expr + '$$</div>';
      }
      const token = '@@MATHBLOCK_' + mathBlocks.length + '@@';
      mathBlocks.push(rendered);
      return '\n\n' + token + '\n\n';
    });

    // 2. Extract and protect Inline Math: $ ... $ (excluding escaped \$)
    const mathInlines = [];
    str = str.replace(/(^|[^\\])\$([^\$\n]+?)\$/g, (match, prefix, expr) => {
      let rendered = "";
      if (typeof katex !== "undefined" && katex.renderToString) {
        try {
          rendered = katex.renderToString(expr.trim(), { displayMode: false, throwOnError: false });
        } catch (e) {
          rendered = '<span class="katex-inline-error">$' + expr + '$</span>';
        }
      } else {
        rendered = '<span class="katex-inline-raw">$' + expr + '$</span>';
      }
      const token = '@@MATHINLINE_' + mathInlines.length + '@@';
      mathInlines.push(rendered);
      return prefix + token;
    });

    // 3. Parse Markdown with Marked.js (GFM tables, headers, lists)
    let html = "";
    const parseFn = typeof marked === "function" ? marked : (typeof marked !== "undefined" && typeof marked.parse === "function" ? marked.parse.bind(marked) : null);
    if (parseFn) {
      try {
        if (typeof marked.setOptions === "function") {
          marked.setOptions({ gfm: true, breaks: true });
        }
        html = parseFn(str);
      } catch (err) {
        console.warn("Marked parse error, using fallback:", err);
        html = fallbackMarkdown(str);
      }
    } else {
      html = fallbackMarkdown(str);
    }

    // 4. Restore math blocks and inlines
    mathBlocks.forEach((blockHtml, i) => {
      const pRegex = new RegExp('<p>\\s*@@MATHBLOCK_' + i + '@@\\s*<\\/p>', 'g');
      if (pRegex.test(html)) {
        html = html.replace(pRegex, blockHtml);
      } else {
        html = html.replace(new RegExp('@@MATHBLOCK_' + i + '@@', 'g'), blockHtml);
      }
    });
    mathInlines.forEach((inlineHtml, i) => {
      html = html.replace(new RegExp('@@MATHINLINE_' + i + '@@', 'g'), inlineHtml);
    });

    return html;
  }

  function fallbackMarkdown(raw) {
    let out = String(raw)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    out = out.replace(/^###### (.*$)/gim, '<h6 style="margin: 0.5rem 0 0.2rem; font-size: 0.85rem; font-weight: 700; color: var(--text-muted);">$1</h6>');
    out = out.replace(/^##### (.*$)/gim, '<h5 style="margin: 0.6rem 0 0.3rem; font-size: 0.92rem; font-weight: 700; color: var(--text-muted);">$1</h5>');
    out = out.replace(/^#### (.*$)/gim, '<h4 style="margin: 0.8rem 0 0.4rem; font-size: 1.0rem; font-weight: 700; color: var(--text-main);">$1</h4>');
    out = out.replace(/^### (.*$)/gim, '<h3 style="margin: 1rem 0 0.5rem; font-size: 1.1rem; font-weight: 700; color: var(--text-main);">$1</h3>');
    out = out.replace(/^## (.*$)/gim, '<h2 style="margin: 1.2rem 0 0.6rem; font-size: 1.25rem; font-weight: 700; color: var(--text-main);">$1</h2>');
    out = out.replace(/^# (.*$)/gim, '<h1 style="margin: 1.4rem 0 0.7rem; font-size: 1.4rem; font-weight: 700; color: var(--text-main);">$1</h1>');
    out = out.replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>');
    out = out.replace(/\*(.*?)\*/gim, '<em>$1</em>');
    out = out.replace(/^[-*]{3,}$/gim, '<hr>');
    out = out.replace(/^(\d+)\.\s+(.*$)/gim, '<div style="margin-left: 1.25rem; margin-bottom: 0.25rem;"><strong>$1.</strong> $2</div>');
    out = out.replace(/^[-*]\s+(.*$)/gim, '<div style="margin-left: 1.25rem; margin-bottom: 0.25rem;">&bull; $1</div>');
    out = out.replace(/\n\n+/g, '<div style="height: 0.5rem;"></div>');
    out = out.replace(/\n/g, '<br>');
    return out;
  }

  // Direct paste from system clipboard via Clipboard API
  window.pasteFeaFromClipboard = async function (targetPart, mode = "new") {
    if (!navigator.clipboard || !navigator.clipboard.read) {
      showToast("Clipboard API unavailable. Please press Ctrl+V directly.");
      return;
    }
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        for (const type of item.types) {
          if (type.startsWith("image/")) {
            const blob = await item.getType(type);
            await handleFeaScreenshotPasted(blob, targetPart, mode);
            return;
          }
        }
      }
      showToast("No screenshot on clipboard. Take a capture (Win+Shift+S) first.");
    } catch (err) {
      console.warn("Clipboard read error:", err);
      showToast("Clipboard access denied. Use Ctrl+V or browse file.");
    }
  };

  // Direct paste schematic from system clipboard via Clipboard API
  window.pasteSchematicFromClipboard = async function (targetPart, mode = "edit") {
    if (!navigator.clipboard || !navigator.clipboard.read) {
      showToast("Clipboard API unavailable. Please use Browse File.");
      return;
    }
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        for (const type of item.types) {
          if (type.startsWith("image/")) {
            const blob = await item.getType(type);
            if (mode === "edit") {
              const evalVer = selectedEvalVersion;
              const row = partsData.find((p) => p["Part Name"] === targetPart && String(p.Version) === String(evalVer));
              if (row) {
                await uploadImagesToIteration([blob], row, targetPart);
                showToast("Schematic attached successfully!");
              } else {
                showToast("Iteration not found.");
              }
            } else {
              const dataUrl = await downscaleImage(blob, 1920, 1080);
              newIterSchematic = dataUrl;
              const previewCont = document.getElementById("newSchematicPreview");
              if (previewCont) {
                previewCont.innerHTML = `
                  <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-sm); margin-top: 0.5rem;">
                    <img src="${dataUrl}" style="max-height: 60px; border-radius: 4px;">
                    <span class="tag" style="color: var(--delta-green); border-color: var(--delta-green);">Staged</span>
                    <button type="button" class="btn-action danger" onclick="window.clearNewSchematic()">Remove</button>
                  </div>
                `;
              }
              showToast("Schematic staged for new iteration!");
            }
            return;
          }
        }
      }
      showToast("No image found on clipboard. Capture (Win+Shift+S) first.");
    } catch (err) {
      console.warn("Clipboard read error:", err);
      showToast("Clipboard access denied. Please use Browse File.");
    }
  };

  window.clearNewSchematic = function () {
    newIterSchematic = null;
    const previewCont = document.getElementById("newSchematicPreview");
    if (previewCont) previewCont.innerHTML = "";
  };

  function renderSchematicThumbnails(evalRow, partName) {
    const images = parseImagePaths(evalRow["Image Path"]);
    if (!images.length) {
      return '<p style="font-size: 0.8rem; color: var(--text-dim); margin-top: 0.5rem;">No CAD schematics currently attached.</p>';
    }
    return `
      <div style="display: flex; gap: 0.75rem; flex-wrap: wrap; align-items: center; margin-top: 0.5rem;">
        ${images.map((imgP, i) => {
          const cleanThumb = `../${imgP.replace(/\\/g, "/")}`;
          return `
            <div style="position: relative; display: inline-block;">
              <img src="${cleanThumb}" style="max-height: 75px; max-width: 130px; object-fit: cover; border-radius: var(--radius-sm); border: 1px solid var(--border-color); cursor: pointer;" onclick="window.openLightbox('${cleanThumb}')">
              <button type="button" style="position: absolute; top: -6px; right: -6px; background: var(--delta-red); color: #fff; border: 1px solid #fff; border-radius: 50%; width: 20px; height: 20px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; justify-content: center;" onclick="window.deleteGalleryImage('${partName}', '${evalRow.Version}', ${i})" title="Remove schematic">x</button>
            </div>
          `;
        }).join("")}
      </div>
    `;
  }

  window.switchToTab = function (tabId) {
    const btn = document.querySelector(`.main-tabs-header .tab-btn[data-tab="${tabId}"]`);
    if (btn) btn.click();
  };

  // URL Hash token auto-configuration
  function processUrlHashToken() {
    const hash = window.location.hash;
    if (hash && hash.includes("setup_token=")) {
      const m = hash.match(/setup_token=([^&]+)/);
      if (m && m[1]) {
        const token = decodeURIComponent(m[1]).trim();
        if (token) {
          window.DTCloud.saveConfig(token);
          // Clean hash from address bar without reloading
          window.history.replaceState(null, "", window.location.pathname + window.location.search);
        }
      }
    }
  }

  // Initialize data loading
  async function init() {
    processUrlHashToken();
    updateRoleIndicator();

    try {
      const [resParts, resAss, resReq, resTrash] = await Promise.all([
        window.DTCloud.fetchFile("data/parts.json"),
        window.DTCloud.fetchFile("data/assemblies.json").catch(() => ({ data: [] })),
        window.DTCloud.fetchFile("data/requirements.json").catch(() => ({ data: [] })),
        window.DTCloud.fetchFile("data/trash.json").catch(() => ({ data: [] }))
      ]);

      partsData = resParts.data || [];
      assembliesData = resAss.data || [];
      requirementsData = resReq.data || [];
      trashData = resTrash.data || [];

      buildSidebar();
      renderCurrentView();
    } catch (err) {
      console.error("Data load error:", err);
      elContentArea.innerHTML = `
        <div style="padding: 2rem; background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-color);">
          <h2 style="color: var(--accent-red); margin-bottom: 0.5rem;">Load Error</h2>
          <p style="color: var(--text-muted);">${err.message}</p>
          <button onclick="location.reload()" class="btn-primary" style="margin-top: 1rem;">Retry</button>
        </div>
      `;
    }
  }

  // Update role indicator in footer
  function updateRoleIndicator() {
    if (!elUserRoleBar) return;
    if (isEditor()) {
      elUserRoleBar.innerHTML = `
        <span class="role-pill editor">Editor</span>
        <button id="btnLockEditor" class="role-action-btn">Lock Editor</button>
      `;
      const btnLock = document.getElementById("btnLockEditor");
      if (btnLock) {
        btnLock.onclick = () => {
          window.DTCloud.clearToken();
          updateRoleIndicator();
          buildSidebar();
          renderCurrentView();
        };
      }
    } else {
      elUserRoleBar.innerHTML = `
        <span class="role-pill viewer">Viewer</span>
        <button id="btnLoginEditor" class="role-action-btn">Editor Login</button>
      `;
      const btnLogin = document.getElementById("btnLoginEditor");
      if (btnLogin) {
        btnLogin.onclick = openCloudSettingsModal;
      }
    }
  }

  // Sidebar Builder
  function buildSidebar() {
    const partsBySub = {};
    partsData.forEach((p) => {
      const sub = p.Subassembly || "Other";
      if (!partsBySub[sub]) partsBySub[sub] = new Set();
      partsBySub[sub].add(p["Part Name"]);
    });

    let html = `
      <div class="sidebar-section">
        <div class="section-label">General</div>
        <div class="nav-item ${currentView === "assembly" ? "active" : ""}" data-view="assembly">
          <span>Wheel Assembly</span>
          <span class="tag">Overview</span>
        </div>
        <div class="nav-item ${currentView === "requirements" ? "active" : ""}" data-view="requirements">
          <span>Requirements</span>
          <span class="tag">${requirementsData.length}</span>
        </div>
        <div class="nav-item" id="btnExportExcel">
          <span>Export Excel</span>
          <span class="tag">.xlsx</span>
        </div>
        ${isEditor() ? `
        <div class="nav-item ${currentView === "trash" ? "active" : ""}" data-view="trash">
          <span>Recycle Bin</span>
          <span class="tag">${trashData.length}</span>
        </div>
        ` : ""}
      </div>
    `;

    const subOrder = ["UP", "HB", "GB", "Other"];
    const subDisplayNames = {
      "UP": "UP",
      "HB": "HB",
      "GB": "GB",
      "Other": "OTHER"
    };
    const allSubs = Array.from(new Set([...subOrder, ...Object.keys(partsBySub)]));

    allSubs.forEach((sub) => {
      const parts = Array.from(partsBySub[sub] || []).sort();
      if (!parts.length) return;

      const displayName = subDisplayNames[sub] || (sub.toUpperCase ? sub.toUpperCase() : sub);

      html += `
        <div class="sidebar-section">
          <div class="section-label">
            <span>${displayName}</span>
            ${isEditor() ? `<button class="role-action-btn" onclick="window.openNewPartModal('${sub}')">+ Part</button>` : ""}
          </div>
      `;

      parts.forEach((partName) => {
        const isActive = currentView === partName ? "active" : "";
        const iters = partsData.filter((p) => p["Part Name"] === partName);
        const latest = iters[iters.length - 1];
        const vLabel = latest ? formatVersion(latest.Version) : "";

        html += `
          <div class="nav-item ${isActive}" data-view="${partName}">
            <span>${partName}</span>
            <span class="tag">${vLabel}</span>
          </div>
        `;
      });
      html += `</div>`;
    });

    elSidebarNav.innerHTML = html;

    // Attach click events
    elSidebarNav.querySelectorAll(".nav-item[data-view]").forEach((item) => {
      item.addEventListener("click", () => {
        const nextView = item.getAttribute("data-view");
        if (nextView !== currentView) {
          currentPartName = null;
          selectedEvalVersion = null;
          selectedRefVersion = null;
        }
        currentView = nextView;
        activeGalleryIndex = 0;
        buildSidebar();
        renderCurrentView();
      });
    });

    const elBtnExport = document.getElementById("btnExportExcel");
    if (elBtnExport) {
      elBtnExport.addEventListener("click", exportExcelWorkbook);
    }
  }

  // Client-side Excel export using SheetJS (0 seconds)
  function exportExcelWorkbook() {
    if (typeof XLSX === "undefined") {
      alert("SheetJS library not loaded.");
      return;
    }

    const wb = XLSX.utils.book_new();

    // 1. Parts Sheet
    const cleanParts = partsData.map((p) => ({
      "Subassembly": p.Subassembly || "",
      "Part Name": p["Part Name"] || "",
      "Version": p.Version || "",
      "Date": p.Date || "",
      "Weight (g)": parseFloat(p["Weight (g)"] || 0),
      "Description": p.Description || "",
      "Specifications": p.Specifications || "",
      "Image Path": typeof p["Image Path"] === "object" ? JSON.stringify(p["Image Path"]) : (p["Image Path"] || ""),
      "Stiffness Data": typeof p["Stiffness Data"] === "object" ? JSON.stringify(p["Stiffness Data"]) : (p["Stiffness Data"] || ""),
      "LC Images": typeof p["LC Images"] === "object" ? JSON.stringify(p["LC Images"]) : (p["LC Images"] || ""),
      "Tolerancing Data": typeof p["Tolerancing Data"] === "object" ? JSON.stringify(p["Tolerancing Data"]) : (p["Tolerancing Data"] || ""),
      "Justification": p.Justification || ""
    }));
    const wsParts = XLSX.utils.json_to_sheet(cleanParts);
    XLSX.utils.book_append_sheet(wb, wsParts, "Parts");

    // 2. Assemblies Sheet
    if (assembliesData && assembliesData.length) {
      const wsAss = XLSX.utils.json_to_sheet(assembliesData);
      XLSX.utils.book_append_sheet(wb, wsAss, "Assemblies");
    }

    // 3. Requirements Sheet
    if (requirementsData && requirementsData.length) {
      const wsReq = XLSX.utils.json_to_sheet(requirementsData);
      XLSX.utils.book_append_sheet(wb, wsReq, "Requirements");
    }

    XLSX.writeFile(wb, "FST16_DT_Parts_Export.xlsx");
  }

  // Render View Router
  function renderCurrentView() {
    if (currentView === "assembly") {
      currentPartName = null;
      renderAssemblyView();
    } else if (currentView === "requirements") {
      currentPartName = null;
      renderRequirementsView();
    } else if (currentView === "trash") {
      currentPartName = null;
      renderTrashView();
    } else {
      renderPartView(currentView);
    }
  }

  function resolvePartInfo(partKey, verStr) {
    if (!verStr) return { version: "", weight: 0, partName: partKey, subassembly: "-" };
    const vClean = normVer(verStr).toLowerCase();

    // 1. Try custom part matching Part Name === partKey
    const customIters = partsData.filter((p) => p["Part Name"] === partKey);
    if (customIters.length) {
      const matched = customIters.find((p) => normVer(p.Version).toLowerCase() === vClean);
      if (matched) {
        return {
          version: matched.Version,
          weight: parseFloat(matched["Weight (g)"]) || 0,
          partName: partKey,
          subassembly: matched.Subassembly || "Drivetrain",
          description: matched.Description || ""
        };
      }
      return {
        version: verStr,
        weight: 0,
        partName: partKey,
        subassembly: customIters[0].Subassembly || "-"
      };
    }

    // 2. Try OEM part (Bearing / Seal / Other) matching Part Name === verStr
    const oemParts = partsData.filter(
      (p) => (p.Subassembly || "") === "Other" || partKey.includes("Bearing") || partKey.includes("Seal")
    );
    const matchedOem = oemParts.find((p) => normVer(p["Part Name"]).toLowerCase() === vClean);
    if (matchedOem) {
      return {
        version: matchedOem["Part Name"],
        weight: parseFloat(matchedOem["Weight (g)"]) || 0,
        partName: partKey,
        subassembly: "Other",
        description: matchedOem.Description || ""
      };
    }

    return { version: verStr, weight: 0, partName: partKey, subassembly: "-" };
  }

  function renderAssemblyBreakdownRows(config) {
    if (!config || !Object.keys(config).length) {
      return `<tr><td colspan="3" style="color: var(--text-dim); text-align: center;">No parts configured</td></tr>`;
    }
    return Object.keys(config)
      .map((k) => {
        const v = config[k];
        const info = resolvePartInfo(k, v);
        return `
          <tr>
            <td><strong style="color: #ffffff;">${k}</strong></td>
            <td>${formatVersion(v)}</td>
            <td class="mono">${info.weight > 0 ? info.weight.toFixed(1) + " g" : "-"}</td>
          </tr>
        `;
      })
      .join("");
  }

  window.changeAssemblyCompare = function () {
    const elEval = document.getElementById("selectEvalAss");
    const elRef = document.getElementById("selectRefAss");
    if (elEval) selectedEvalAssVersion = elEval.value;
    if (elRef) selectedRefAssVersion = elRef.value;
    renderAssemblyView();
  };

  // 1. Wheel Assembly View
  function renderAssemblyView() {
    if (!assembliesData || !assembliesData.length) {
      let emptyHtml = `
        <div class="page-header">
          <div>
            <h1 class="page-title">Wheel Assembly</h1>
            <p class="page-subtitle">Assembly comparison and subsystem breakdown</p>
          </div>
          <div class="header-actions">
            <button class="btn-action" onclick="window.exportExcelWorkbook()">Export Excel</button>
            ${isEditor() ? `
            <button class="btn-action primary" onclick="window.openManageAssembliesModal()">Manage Assemblies</button>
            <button class="btn-action" onclick="window.openNewPartModal('')">New Part</button>
            ` : ""}
          </div>
        </div>
        <div class="inline-tab-card" style="padding: 2rem; text-align: center;">
          <p style="color: var(--text-muted); margin-bottom: 1rem;">No assemblies configured in tracker.</p>
          ${isEditor() ? `<button class="btn-action primary" onclick="window.createNewAssembly()">Create Assembly</button>` : ""}
        </div>
      `;
      elContentArea.innerHTML = emptyHtml;
      return;
    }

    // Ensure selected versions exist
    if (!selectedEvalAssVersion || !assembliesData.some((a) => a["Assembly Version"] === selectedEvalAssVersion)) {
      selectedEvalAssVersion = assembliesData[0]["Assembly Version"];
    }
    if (!selectedRefAssVersion || !assembliesData.some((a) => a["Assembly Version"] === selectedRefAssVersion)) {
      const fst15Ass = assembliesData.find((a) => (a["Assembly Version"] || "").toUpperCase().includes("FST15"));
      selectedRefAssVersion = fst15Ass ? fst15Ass["Assembly Version"] : assembliesData[assembliesData.length - 1]["Assembly Version"];
    }

    const evalAss = assembliesData.find((a) => a["Assembly Version"] === selectedEvalAssVersion) || assembliesData[0];
    const refAss = assembliesData.find((a) => a["Assembly Version"] === selectedRefAssVersion) || assembliesData[assembliesData.length - 1];

    const evalConfig = parseAssemblyConfig(evalAss["Part Configuration"]);
    const refConfig = parseAssemblyConfig(refAss["Part Configuration"]);

    const standardKeys = [
      { key: "UP_01", label: "Upright Body", sub: "UP" },
      { key: "UP_02", label: "Upright Lid", sub: "UP" },
      { key: "UP_03", label: "Front Plate", sub: "UP" },
      { key: "UP_04", label: "Backplate", sub: "UP" },
      { key: "HB_01", label: "Wheel Hub", sub: "HB" },
      { key: "HB_02", label: "Hub Lid", sub: "HB" },
      { key: "HB_07", label: "Ring Spacer", sub: "HB" },
      { key: "TD", label: "Transmission", sub: "TD" },
      { key: "Inner Bearing", label: "Inner Bearing", sub: "Other" },
      { key: "Outer Bearing", label: "Outer Bearing", sub: "Other" },
      { key: "Seal", label: "Shaft Seal", sub: "Other" }
    ];

    const configuredKeys = new Set([...Object.keys(evalConfig), ...Object.keys(refConfig)]);
    const allRows = [];
    standardKeys.forEach((item) => {
      if (configuredKeys.has(item.key)) {
        allRows.push(item);
        configuredKeys.delete(item.key);
      }
    });
    configuredKeys.forEach((k) => {
      allRows.push({ key: k, label: k, sub: "Other" });
    });

    let evalTotalWeight = 0;
    let refTotalWeight = 0;
    const rowComparisons = allRows.map((r) => {
      const ep = resolvePartInfo(r.key, evalConfig[r.key]);
      const rp = resolvePartInfo(r.key, refConfig[r.key]);
      evalTotalWeight += ep.weight;
      refTotalWeight += rp.weight;
      const diff = ep.weight - rp.weight;
      const pct = rp.weight > 0 ? (diff / rp.weight) * 100 : 0;
      return {
        ...r,
        evalPart: ep,
        refPart: rp,
        diff,
        pct
      };
    });

    if (evalTotalWeight === 0 && evalAss["Total Weight (g)"]) {
      evalTotalWeight = parseFloat(evalAss["Total Weight (g)"]) || 0;
    }
    if (refTotalWeight === 0 && refAss["Total Weight (g)"]) {
      refTotalWeight = parseFloat(refAss["Total Weight (g)"]) || 0;
    }

    const totalDiff = evalTotalWeight - refTotalWeight;
    const isTotalZero = Math.abs(totalDiff) < 1e-7;
    const totalDeltaClass = isTotalZero ? "delta-neutral" : (totalDiff <= 0 ? "delta-good" : "delta-bad");

    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">Wheel Assembly</h1>
          <p class="page-subtitle">Assembly comparison and subsystem breakdown</p>
        </div>
        <div class="header-actions">
          <button class="btn-action" onclick="window.exportExcelWorkbook()">Export Excel</button>
          ${isEditor() ? `
          <button class="btn-action primary" onclick="window.openManageAssembliesModal()">Manage Assemblies</button>
          <button class="btn-action" onclick="window.openNewPartModal('')">New Part</button>
          ` : ""}
        </div>
      </div>

      <div class="metrics-grid">
        <div class="metric-card">
          <div class="metric-label">Evaluated Assembly</div>
          <div class="metric-value">${evalTotalWeight.toFixed(1)} g</div>
          <div class="metric-sub delta-neutral">${(evalTotalWeight / 1000).toFixed(3)} kg &bull; ${evalAss["Assembly Version"]}</div>
        </div>
        <div class="metric-card">
          <div class="metric-label">Baseline Assembly</div>
          <div class="metric-value">${refTotalWeight.toFixed(1)} g</div>
          <div class="metric-sub delta-neutral">${(refTotalWeight / 1000).toFixed(3)} kg &bull; ${refAss["Assembly Version"]}</div>
        </div>
        <div class="metric-card">
          <div class="metric-label">Assembly Delta</div>
          <div class="metric-value ${totalDeltaClass}">
            ${totalDiff <= 0 ? "" : "+"}${totalDiff.toFixed(1)} g
          </div>
          <div class="metric-sub ${totalDeltaClass}">
            ${refTotalWeight > 0 ? ((totalDiff / refTotalWeight) * 100).toFixed(1) + "%" : "0.0%"} vs Baseline
          </div>
        </div>
      </div>

      <div style="display: flex; gap: 1.5rem; align-items: center; margin-bottom: 1.5rem; flex-wrap: wrap;">
        <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
          Evaluate Assembly:
          <select id="selectEvalAss" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;" onchange="window.changeAssemblyCompare()">
            ${assembliesData.map((a) => `<option value="${a["Assembly Version"]}" ${a["Assembly Version"] === evalAss["Assembly Version"] ? "selected" : ""}>${a["Assembly Version"]} (${parseFloat(a["Total Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
          </select>
        </label>
        <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
          Baseline Assembly:
          <select id="selectRefAss" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;" onchange="window.changeAssemblyCompare()">
            ${assembliesData.map((a) => `<option value="${a["Assembly Version"]}" ${a["Assembly Version"] === refAss["Assembly Version"] ? "selected" : ""}>${a["Assembly Version"]} (${parseFloat(a["Total Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
          </select>
        </label>
      </div>

      <h2 style="font-size: 1.1rem; margin-bottom: 0.75rem;">Assembly Comparison</h2>
      <div class="card-table-wrapper">
        <table class="styled-table">
          <thead>
            <tr>
              <th>Subsystem / Part</th>
              <th>${evalAss["Assembly Version"]}</th>
              <th>${refAss["Assembly Version"]}</th>
              <th>Diff (g)</th>
              <th>%</th>
            </tr>
          </thead>
          <tbody>
    `;

    rowComparisons.forEach((r) => {
      const isCustom = partsData.some((p) => p["Part Name"] === r.key);
      const isDiffZero = Math.abs(r.diff) < 1e-7;
      const diffClass = isDiffZero ? "delta-neutral" : (r.diff <= 0 ? "delta-good" : "delta-bad");
      const navClick = isCustom ? `style="cursor: pointer;" onclick="window.navigateToPart('${r.key}')"` : "";

      let pctLabel = "--";
      if (r.refPart.weight > 0) {
        pctLabel = `${r.diff <= 0 ? "" : "+"}${r.pct.toFixed(1)}%`;
      } else if (r.evalPart.weight > 0) {
        pctLabel = "New";
      }

      html += `
        <tr ${navClick}>
          <td>
            <strong>${r.sub}</strong> &bull; <span style="color: #ffffff; font-weight: 600;">${r.key}</span>
            <span style="font-size: 0.78rem; color: var(--text-muted); margin-left: 0.35rem;">(${r.label})</span>
          </td>
          <td class="mono">
            ${r.evalPart.weight > 0 ? `${formatVersion(r.evalPart.version)} (${r.evalPart.weight.toFixed(1)} g)` : `<span style="color: var(--text-dim);">--</span>`}
          </td>
          <td class="mono">
            ${r.refPart.weight > 0 ? `${formatVersion(r.refPart.version)} (${r.refPart.weight.toFixed(1)} g)` : `<span style="color: var(--text-dim);">--</span>`}
          </td>
          <td class="mono ${diffClass}">
            <strong>${r.diff <= 0 ? "" : "+"}${r.diff.toFixed(1)} g</strong>
          </td>
          <td class="mono ${diffClass}">
            <strong>${pctLabel}</strong>
          </td>
        </tr>
      `;
    });

    html += `
          </tbody>
          <tfoot>
            <tr style="background: rgba(255, 255, 255, 0.04); font-weight: 700; border-top: 2px solid var(--border-color);">
              <td><strong>Total Assembly Weight</strong></td>
              <td class="mono"><strong>${evalTotalWeight.toFixed(1)} g</strong></td>
              <td class="mono"><strong>${refTotalWeight.toFixed(1)} g</strong></td>
              <td class="mono ${totalDeltaClass}"><strong>${totalDiff <= 0 ? "" : "+"}${totalDiff.toFixed(1)} g</strong></td>
              <td class="mono ${totalDeltaClass}"><strong>${refTotalWeight > 0 ? ((totalDiff / refTotalWeight) * 100).toFixed(1) + "%" : "0.0%"}</strong></td>
            </tr>
          </tfoot>
        </table>
      </div>

      <h3 style="font-size: 1rem; margin: 1.5rem 0 0.75rem;">Side Comparison</h3>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 1.25rem;">
        <div class="inline-tab-card">
          <h4 style="margin-bottom: 0.35rem; color: var(--text-muted);">Baseline: ${refAss["Assembly Version"]}</h4>
          <div style="font-size: 0.8rem; color: var(--text-dim); margin-bottom: 0.75rem;">Total: ${refTotalWeight.toFixed(1)} g | Date: ${refAss.Date || "-"}</div>
          <p style="font-size: 0.88rem; margin-bottom: 0.75rem;">${refAss.Description || "No description recorded."}</p>
          <h5 style="font-size: 0.85rem; margin-bottom: 0.35rem; color: var(--text-muted);">Part Breakdown</h5>
          <div class="card-table-wrapper">
            <table class="styled-table">
              <thead>
                <tr><th>Component</th><th>Version</th><th>Weight</th></tr>
              </thead>
              <tbody>
                ${renderAssemblyBreakdownRows(refConfig)}
              </tbody>
            </table>
          </div>
        </div>
        <div class="inline-tab-card">
          <h4 style="margin-bottom: 0.35rem; color: var(--accent-red);">Evaluated: ${evalAss["Assembly Version"]}</h4>
          <div style="font-size: 0.8rem; color: var(--text-dim); margin-bottom: 0.75rem;">Total: ${evalTotalWeight.toFixed(1)} g | Date: ${evalAss.Date || "-"}</div>
          <p style="font-size: 0.88rem; margin-bottom: 0.75rem;">${evalAss.Description || "No description recorded."}</p>
          <h5 style="font-size: 0.85rem; margin-bottom: 0.35rem; color: var(--text-muted);">Part Breakdown</h5>
          <div class="card-table-wrapper">
            <table class="styled-table">
              <thead>
                <tr><th>Component</th><th>Version</th><th>Weight</th></tr>
              </thead>
              <tbody>
                ${renderAssemblyBreakdownRows(evalConfig)}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;

    elContentArea.innerHTML = html;
  }

  // 2. Requirements View
  function renderRequirementsView() {
    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">Requirements</h1>
          <p class="page-subtitle">Formula Student regulations and internal team hardpoints</p>
        </div>
        <div class="header-actions">
          <button class="btn-action" onclick="window.exportExcelWorkbook()">Export Excel</button>
        </div>
      </div>

      <div class="tabs-header">
        <button class="tab-btn ${reqActiveTab === "list" ? "active" : ""}" id="tabBtnReqList">Requirements List</button>
        <button class="tab-btn ${reqActiveTab === "diagrams" ? "active" : ""}" id="tabBtnReqDiag">Source Diagrams</button>
      </div>

      <div class="tab-pane ${reqActiveTab === "list" ? "active" : ""}" id="paneReqList">
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>ID</th>
                <th>Category</th>
                <th>Group</th>
                <th>Requirement</th>
                <th>Explanation</th>
                <th>Verification</th>
                <th>Validation</th>
                <th>Related Parts</th>
              </tr>
            </thead>
            <tbody>
    `;

    requirementsData.forEach((r) => {
      const reqId = r.ID || r["Requirement ID"] || "REQ";
      const cat = r.Category || r.Subassembly || "—";
      const grp = r.Group || "—";
      const reqText = r.Requirement || "—";
      const exp = r.Explanation || "—";
      const ver = r.Verification || "—";
      const val = r.Validation || "—";
      const parts = r["Related Parts"] || "—";

      html += `
        <tr>
          <td><strong style="color: var(--accent-red);">${reqId}</strong></td>
          <td>${cat}</td>
          <td><span class="tag">${grp}</span></td>
          <td><strong>${reqText}</strong></td>
          <td style="font-size: 0.82rem; color: var(--text-muted);">${exp}</td>
          <td style="font-size: 0.82rem; color: var(--text-muted);">${ver}</td>
          <td style="font-size: 0.82rem;">${val}</td>
          <td style="font-size: 0.82rem; color: var(--text-muted);">${parts}</td>
        </tr>
      `;
    });

    html += `
            </tbody>
          </table>
        </div>
      </div>

      <div class="tab-pane ${reqActiveTab === "diagrams" ? "active" : ""}" id="paneReqDiag">
        <h3 style="font-size: 1rem; margin-bottom: 0.5rem;">Source Diagrams</h3>
        <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1rem;">
          Formula Student rulebook excerpts and technical specification sheets. Click any diagram to zoom.
        </p>
        <div class="diagrams-grid">
          ${renderDiagramCard("images/requirements/fsg_requirements_p1.png", "FSG Rules P1", "Formula Student rules excerpt part 1")}
          ${renderDiagramCard("images/requirements/fsg_requirements_p2.png", "FSG Rules P2", "Formula Student rules excerpt part 2")}
          ${renderDiagramCard("images/requirements/team_requirements_p1.png", "Team Targets P1", "System weight targets and motor sizing")}
          ${renderDiagramCard("images/requirements/team_requirements_p2.png", "Team Targets P2", "Thermal cooling and bearing limits")}
          ${renderDiagramCard("images/requirements/team_requirements_p3.png", "Team Targets P3", "Structural stiffness requirements")}
        </div>
      </div>
    `;

    elContentArea.innerHTML = html;

    // Attach Requirements tabs
    const btnList = document.getElementById("tabBtnReqList");
    const btnDiag = document.getElementById("tabBtnReqDiag");
    if (btnList && btnDiag) {
      btnList.onclick = () => {
        reqActiveTab = "list";
        renderRequirementsView();
      };
      btnDiag.onclick = () => {
        reqActiveTab = "diagrams";
        renderRequirementsView();
      };
    }
  }

  function renderDiagramCard(path, title, desc) {
    const cleanPath = `../${path.replace(/\\/g, "/")}`;
    return `
      <div class="diagram-card" onclick="window.openLightbox('${cleanPath}')">
        <img src="${cleanPath}" alt="${title}" class="diagram-preview" onerror="this.src='';">
        <div class="diagram-body">
          <div class="diagram-title">${title}</div>
          <div class="diagram-desc">${desc}</div>
        </div>
      </div>
    `;
  }

  // 3. Recycle Bin View
  function renderTrashView() {
    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">Recycle Bin</h1>
          <p class="page-subtitle">Deleted component iterations stored in data/trash.json</p>
        </div>
        <div class="header-actions">
          ${trashData.length && isEditor() ? `
          <button class="btn-action danger" onclick="window.emptyTrash()">Empty Bin</button>
          ` : ""}
        </div>
      </div>
    `;

    if (!trashData.length) {
      html += `
        <div style="padding: 2.5rem; text-align: center; background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-color); color: var(--text-muted);">
          Recycle Bin is empty.
        </div>
      `;
    } else {
      html += `
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>Subassembly</th>
                <th>Part Name</th>
                <th>Version</th>
                <th>Weight (g)</th>
                <th>Description</th>
                ${isEditor() ? `<th>Actions</th>` : ""}
              </tr>
            </thead>
            <tbody>
      `;

      trashData.forEach((item, idx) => {
        const w = parseFloat(item["Weight (g)"]) || 0;
        html += `
          <tr>
            <td><strong>${item.Subassembly || "—"}</strong></td>
            <td><strong style="color: var(--accent-red);">${item["Part Name"]}</strong></td>
            <td>${formatVersion(item.Version)}</td>
            <td class="mono">${w.toFixed(1)} g</td>
            <td style="color: var(--text-muted); font-size: 0.82rem;">${(item.Description || "—").slice(0, 70)}</td>
            ${isEditor() ? `
            <td>
              <button class="btn-action primary" onclick="window.restoreTrashItem(${idx})">Restore Part</button>
            </td>
            ` : ""}
          </tr>
        `;
      });

      html += `
            </tbody>
          </table>
        </div>
      `;
    }

    elContentArea.innerHTML = html;
  }

  // Restore trash item
  window.restoreTrashItem = async function (idx) {
    if (!isEditor()) return;
    const item = trashData[idx];
    if (!item) return;

    trashData.splice(idx, 1);
    partsData.push(item);

    try {
      await Promise.all([
        window.DTCloud.commitFile("data/trash.json", trashData, `Restore ${item["Part Name"]} from trash`),
        window.DTCloud.commitFile("data/parts.json", partsData, `Restore ${item["Part Name"]} to parts`)
      ]);
      alert("Part restored successfully.");
      buildSidebar();
      renderCurrentView();
    } catch (e) {
      alert("Error restoring part: " + e.message);
    }
  };

  // Empty trash
  window.emptyTrash = async function () {
    if (!isEditor()) return;
    if (!confirm("Permanently empty all items in the Recycle Bin?")) return;

    trashData = [];
    try {
      await window.DTCloud.commitFile("data/trash.json", trashData, "Empty Recycle Bin");
      alert("Recycle Bin emptied.");
      buildSidebar();
      renderCurrentView();
    } catch (e) {
      alert("Error emptying trash: " + e.message);
    }
  };

  function getDefaultBaseline(partName, iters) {
    if (partName === "UP_01") {
      const up01Body = iters.find((p) => String(p.Version).toLowerCase().includes("1 body"));
      if (up01Body) return up01Body.Version;
    }
    const fst15Exact = iters.find((p) => String(p.Version) === "FST15");
    if (fst15Exact) return fst15Exact.Version;
    const fst15Any = iters.find((p) => String(p.Version).toUpperCase().startsWith("FST15"));
    if (fst15Any) return fst15Any.Version;
    return iters[0] ? iters[0].Version : "FST15";
  }

  // 4. Component / Part Detail View
  function renderPartView(partName) {
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    if (!iters.length) {
      elContentArea.innerHTML = `<p style="padding: 2rem;">No data found for ${partName}</p>`;
      return;
    }

    if (currentPartName !== partName) {
      currentPartName = partName;
      selectedEvalVersion = iters[iters.length - 1].Version;
      selectedRefVersion = getDefaultBaseline(partName, iters);
    } else {
      if (!selectedEvalVersion || !iters.some((p) => String(p.Version) === String(selectedEvalVersion))) {
        selectedEvalVersion = iters[iters.length - 1].Version;
      }
      if (!selectedRefVersion || !iters.some((p) => String(p.Version) === String(selectedRefVersion))) {
        selectedRefVersion = getDefaultBaseline(partName, iters);
      }
    }

    const evalRow = iters.find((p) => String(p.Version) === String(selectedEvalVersion)) || iters[iters.length - 1];
    const refRow = iters.find((p) => String(p.Version) === String(selectedRefVersion)) || iters[0];

    const evalWeight = parseFloat(evalRow["Weight (g)"]) || 0;
    const refWeight = parseFloat(refRow["Weight (g)"]) || 0;
    const weightDiff = evalWeight - refWeight;

    const isHB = partName === "HB_01";
    const isUP = partName === "UP_01";
    const partLcs = isUP
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]
      : isHB
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"]
      : [];

    const descMap = {
      LC1: "Pure Braking",
      LC2: "Pure Acceleration",
      LC3: "Pure Cornering",
      LC4: "Braking + Cornering 1",
      LC5: "Braking + Cornering 2",
      LC6: "Acceleration + Cornering 1",
      LC7: "Acceleration + Cornering 2",
      LC8: "1kN Fy",
      LC9: "-1kN Fy"
    };

    const latest = iters[iters.length - 1];
    const nextVer = incrementVersion(latest.Version);

    // Parse stiffness data for new/edit
    let evalStiff = {};
    try {
      if (evalRow["Stiffness Data"]) evalStiff = JSON.parse(evalRow["Stiffness Data"]);
    } catch (e) {}

    // Check active status of Justification and Tolerances
    const currJust = String(evalRow.Justification || "").trim();
    const hasJust = Boolean(currJust && currJust !== "nan" && currJust !== "null");

    let currTols = [];
    try {
      if (typeof evalRow["Tolerancing Data"] === "string") {
        currTols = JSON.parse(evalRow["Tolerancing Data"]) || [];
      } else if (Array.isArray(evalRow["Tolerancing Data"])) {
        currTols = evalRow["Tolerancing Data"];
      }
    } catch (e) {}
    const hasTol = Boolean(currTols && currTols.length > 0);

    // Related requirements
    const relatedReqs = requirementsData.filter((r) => {
      const rel = String(r["Related Parts"] || "");
      return rel.includes(partName);
    });

    const isWeightZero = Math.abs(weightDiff) < 1e-7;
    const weightDeltaClass = isWeightZero ? "delta-neutral" : (weightDiff < 0 ? "delta-good" : "delta-bad");

    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">${partName}</h1>
          <p class="page-subtitle">${evalRow.Subassembly || "Drivetrain"} &bull; Latest: ${formatVersion(iters[iters.length - 1].Version)}</p>
        </div>
        <div class="header-actions">
          <button class="btn-action" onclick="window.exportExcelWorkbook()">Export Excel</button>
        </div>
      </div>

      <div class="metrics-grid">
        <div class="metric-card">
          <div class="metric-label">Current Weight (${formatVersion(evalRow.Version)})</div>
          <div class="metric-value">${evalWeight.toFixed(1)} g</div>
          <div class="metric-sub delta-neutral">${(evalWeight / 1000).toFixed(3)} kg</div>
        </div>
        <div class="metric-card">
          <div class="metric-label">Reference Weight (${formatVersion(refRow.Version)})</div>
          <div class="metric-value">${refWeight.toFixed(1)} g</div>
          <div class="metric-sub delta-neutral">Baseline</div>
        </div>
        <div class="metric-card">
          <div class="metric-label">Weight Delta</div>
          <div class="metric-value ${weightDeltaClass}">
            ${weightDiff <= 0 ? "" : "+"}${weightDiff.toFixed(1)} g
          </div>
          <div class="metric-sub ${weightDeltaClass}">
            ${refWeight > 0 ? ((weightDiff / refWeight) * 100).toFixed(1) + "%" : "0.0%"}
          </div>
        </div>
      </div>

      <div class="main-tabs-header tabs-header">
        <button class="tab-btn ${activePartTab === "tabOverview" ? "active" : ""}" data-tab="tabOverview">Overview</button>
        <button class="tab-btn ${activePartTab === "tabJustification" ? "active" : ""}" data-tab="tabJustification">Justification</button>
        <button class="tab-btn ${activePartTab === "tabTolerances" ? "active" : ""}" data-tab="tabTolerances">Tolerances</button>
        <button class="tab-btn ${activePartTab === "tabNewIter" ? "active" : ""}" data-tab="tabNewIter">New Iteration</button>
        <button class="tab-btn ${activePartTab === "tabEdit" ? "active" : ""}" data-tab="tabEdit">Edit</button>
        <button class="tab-btn ${activePartTab === "tabSchematic" ? "active" : ""}" data-tab="tabSchematic">Schematic</button>
      </div>

      <!-- Tab: Overview -->
      <div class="tab-pane ${activePartTab === "tabOverview" ? "active" : ""}" id="tabOverview">
        ${hasJust ? `
        <div class="inline-tab-card" style="margin-bottom: 1.5rem; border-left: 3px solid var(--accent-red); background: rgba(225, 29, 72, 0.03);">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.6rem;">
            <span style="font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: var(--accent-red);">
              Component Justification
            </span>
            ${isEditor() ? `<button class="role-action-btn" onclick="window.switchToTab('tabJustification')">Edit</button>` : ""}
          </div>
          <div class="markdown-body" style="font-size: 0.9rem; line-height: 1.6; color: var(--text-main);">
            ${renderMarkdown(currJust)}
          </div>
        </div>
        ` : ""}

        ${iters.length > 1 ? `
        <div style="display: flex; gap: 1rem; align-items: center; margin-bottom: 1.5rem; flex-wrap: wrap;">
          <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
            Evaluate Iteration:
            <select id="selectEvalVer" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
              ${iters.map((it) => `<option value="${it.Version}" ${String(it.Version) === String(evalRow.Version) ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
            </select>
          </label>
          <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
            Baseline Iteration:
            <select id="selectRefVer" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
              ${iters.map((it) => `<option value="${it.Version}" ${String(it.Version) === String(refRow.Version) ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
            </select>
          </label>
        </div>
        ` : `
        <div style="margin-bottom: 1.5rem; padding: 0.6rem 0.9rem; background: rgba(225, 29, 72, 0.08); border-left: 3px solid var(--accent-red); border-radius: 4px; font-size: 0.85rem; color: var(--text-muted);">
          Single iteration available (${formatVersion(evalRow.Version)}).
        </div>
        `}

        <h3 style="font-size: 1rem; margin-bottom: 0.5rem;">Mass Comparison</h3>
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>Metric</th>
                <th>${formatVersion(evalRow.Version)}</th>
                <th>${formatVersion(refRow.Version)}</th>
                <th>Diff (g)</th>
                <th>%</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><strong>Total Weight</strong></td>
                <td class="mono">${evalWeight.toFixed(1)} g</td>
                <td class="mono">${refWeight.toFixed(1)} g</td>
                <td class="mono ${weightDeltaClass}">
                  <strong>${weightDiff <= 0 ? "" : "+"}${weightDiff.toFixed(1)} g</strong>
                </td>
                <td class="mono ${weightDeltaClass}">
                  <strong>${refWeight > 0 ? ((weightDiff / refWeight) * 100).toFixed(1) + "%" : "0.0%"}</strong>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        ${renderFEATable(evalRow, refRow, true)}

        ${String(evalRow.Version) !== String(refRow.Version) ? `
        <h3 style="font-size: 1rem; margin: 1.5rem 0 0.75rem;">Side Comparison</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 1.25rem;">
          <div class="inline-tab-card">
            <h4 style="margin-bottom: 0.35rem; color: var(--text-muted);">Baseline: ${formatVersion(refRow.Version)}</h4>
            <div style="font-size: 0.8rem; color: var(--text-dim); margin-bottom: 0.75rem;">Weight: ${refWeight.toFixed(1)} g | Date: ${refRow.Date || "—"}</div>
            <p style="font-size: 0.88rem; margin-bottom: 0.75rem;">${refRow.Description || "No description recorded."}</p>
            <h5 style="font-size: 0.85rem; margin-bottom: 0.35rem;">Specifications</h5>
            <div class="card-table-wrapper" style="margin-bottom: 0.75rem;">
              <table class="styled-table"><tbody>${renderSpecsRows(refRow.Specifications)}</tbody></table>
            </div>
            ${renderSideSchematic(refRow)}
          </div>
          <div class="inline-tab-card">
            <h4 style="margin-bottom: 0.35rem; color: var(--accent-red);">Evaluated: ${formatVersion(evalRow.Version)}</h4>
            <div style="font-size: 0.8rem; color: var(--text-dim); margin-bottom: 0.75rem;">Weight: ${evalWeight.toFixed(1)} g | Date: ${evalRow.Date || "—"}</div>
            <p style="font-size: 0.88rem; margin-bottom: 0.75rem;">${evalRow.Description || "No description recorded."}</p>
            <h5 style="font-size: 0.85rem; margin-bottom: 0.35rem;">Specifications</h5>
            <div class="card-table-wrapper" style="margin-bottom: 0.75rem;">
              <table class="styled-table"><tbody>${renderSpecsRows(evalRow.Specifications)}</tbody></table>
            </div>
            ${renderSideSchematic(evalRow)}
          </div>
        </div>
        ` : `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1.25rem; margin-top: 1.5rem;">
          <div class="inline-tab-card">
            <h4 style="margin-bottom: 0.5rem;">Specifications</h4>
            <p style="font-size: 0.88rem; margin-bottom: 0.75rem;">${evalRow.Description || "—"}</p>
            <div class="card-table-wrapper">
              <table class="styled-table"><tbody>${renderSpecsRows(evalRow.Specifications)}</tbody></table>
            </div>
          </div>
          <div class="inline-tab-card">
            <h4 style="margin-bottom: 0.5rem;">Schematic</h4>
            ${renderSideSchematic(evalRow)}
          </div>
        </div>
        `}

        ${hasTol ? `
        <h3 style="font-size: 1rem; margin: 1.5rem 0 0.5rem;">Manufacturing GD&T</h3>
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>Interface</th>
                <th>Nominal (mm)</th>
                <th>Fit / Tolerance</th>
                <th>Roughness / Notes</th>
              </tr>
            </thead>
            <tbody>
              ${renderTolRows(evalRow["Tolerancing Data"])}
            </tbody>
          </table>
        </div>
        ` : ""}

        ${relatedReqs.length ? `
        <h3 style="font-size: 1rem; margin: 1.5rem 0 0.5rem;">Related Requirements</h3>
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>ID</th>
                <th>Category</th>
                <th>Requirement</th>
                <th>Verification</th>
                <th>Validation</th>
              </tr>
            </thead>
            <tbody>
              ${relatedReqs.map(r => `
                <tr>
                  <td><strong style="color: var(--accent-red);">${r.ID || "—"}</strong></td>
                  <td>${r.Category || "—"}</td>
                  <td>${r.Requirement || "—"}</td>
                  <td style="color: var(--text-muted); font-size: 0.82rem;">${r.Verification || "—"}</td>
                  <td style="color: var(--text-muted); font-size: 0.82rem;">${r.Validation || "—"}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        ` : ""}
      </div>

      <!-- Tab: Justification -->
      <div class="tab-pane ${activePartTab === "tabJustification" ? "active" : ""}" id="tabJustification">
        <div class="inline-tab-card">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: 0.75rem;">
            <label class="checkbox-label">
              <input type="checkbox" id="cbIncludeJust" ${hasJust ? "checked" : ""}>
              Include Justification
            </label>
            ${isEditor() ? `<button class="btn-primary" id="btnSaveJustInline">Save Justification</button>` : ""}
          </div>
          <div id="justEditorContainer" style="display: ${hasJust ? "block" : "none"};">
            ${isEditor() ? `
            <div class="form-group">
              <label for="textareaJustInline">Engineering Justification</label>
              <textarea id="textareaJustInline" class="form-textarea" style="min-height: 220px;" placeholder="Document mechanical justification in Markdown format...">${currJust}</textarea>
            </div>
            ` : `
            <div class="markdown-body" style="font-size: 0.92rem; line-height: 1.6; color: var(--text-main);">
              ${renderMarkdown(currJust) || "<span style='color: var(--text-dim);'>No justification recorded.</span>"}
            </div>
            `}
          </div>
        </div>
      </div>

      <!-- Tab: Tolerances -->
      <div class="tab-pane ${activePartTab === "tabTolerances" ? "active" : ""}" id="tabTolerances">
        <div class="inline-tab-card">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: 0.75rem;">
            <label class="checkbox-label">
              <input type="checkbox" id="cbIncludeTol" ${hasTol ? "checked" : ""}>
              Include Tolerances
            </label>
            <div style="display: flex; gap: 0.5rem;">
              ${isEditor() ? `
              <button class="btn-action" id="btnEditTol" style="display: ${hasTol ? "inline-block" : "none"};" onclick="window.openTolerancesModal('${partName}')">Edit Tolerances</button>
              <button class="btn-primary" id="btnSaveTolInline">Save Tolerances</button>
              ` : ""}
            </div>
          </div>
          <div id="tolTableContainer" style="display: ${hasTol ? "block" : "none"};">
            <div class="card-table-wrapper">
              <table class="styled-table">
                <thead>
                  <tr>
                    <th>Interface</th>
                    <th>Nominal (mm)</th>
                    <th>Fit / Tolerance</th>
                    <th>Roughness / Notes</th>
                  </tr>
                </thead>
                <tbody>
                  ${renderTolRows(evalRow["Tolerancing Data"])}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      <!-- Tab: New Iteration -->
      <div class="tab-pane ${activePartTab === "tabNewIter" ? "active" : ""}" id="tabNewIter">
        <div class="inline-tab-card">
          <div style="margin-bottom: 1.25rem;">
            <h3 style="font-size: 1.1rem; margin-bottom: 0.25rem;">New Iteration</h3>
            <p style="font-size: 0.82rem; color: var(--text-muted);">Create new design iteration for ${partName}</p>
          </div>

          ${!isEditor() ? `
          <div style="padding: 2rem; text-align: center; color: var(--text-muted);">
            Editor access required to create iterations. Enter fine-grained PAT in Cloud Settings.
          </div>
          ` : `
          <div style="display: flex; gap: 1rem; align-items: center; margin-bottom: 1rem;">
            <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
              Copy From:
              <select id="selectNewIterSource" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
                ${iters.map((it, idx) => `<option value="${it.Version}" ${idx === iters.length - 1 ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
              </select>
            </label>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label for="newIterVer">Version ID</label>
              <input type="text" id="newIterVer" class="form-input" value="${nextVer}">
            </div>
            <div class="form-group">
              <label for="newIterWeight">Weight (g)</label>
              <input type="number" step="0.1" id="newIterWeight" class="form-input" value="${parseFloat(latest["Weight (g)"] || 0)}">
            </div>
          </div>

          <div class="form-group">
            <label for="newIterDesc">Description</label>
            <input type="text" id="newIterDesc" class="form-input" placeholder="Summary of changes" value="${latest.Description || ""}">
          </div>

          <div class="form-group">
            <label for="newIterSpecs">Specifications</label>
            <textarea id="newIterSpecs" class="form-textarea" placeholder="Feature: Value">${latest.Specifications || ""}</textarea>
          </div>

          <!-- CAD Schematics Section -->
          <div style="margin: 1.5rem 0 0.75rem;">
            <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">CAD Schematics</h4>
            <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.75rem;">
              Capture in CAD (Win+Shift+S) and click Paste Schematic, or browse files.
            </p>
          </div>

          <div class="paste-action-bar">
            <button type="button" class="btn-paste-clipboard" onclick="window.pasteSchematicFromClipboard('${partName}', 'new')">Paste Schematic</button>
            <button type="button" class="btn-browse-file" onclick="document.getElementById('fileInpNewIterSchematic').click()">Browse File</button>
            <span class="paste-hint-text">Or drag and drop files below</span>
          </div>
          <input type="file" id="fileInpNewIterSchematic" accept="image/*" multiple style="display: none;">

          <div id="newSchematicPreview" style="margin-bottom: 1.25rem;">
            ${newIterSchematic ? `
              <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-sm); margin-top: 0.5rem;">
                <img src="${newIterSchematic}" style="max-height: 60px; border-radius: 4px;">
                <span class="tag" style="color: var(--delta-green); border-color: var(--delta-green);">Staged</span>
                <button type="button" class="btn-action danger" onclick="window.clearNewSchematic()">Remove</button>
              </div>
            ` : ""}
          </div>

          <!-- FEA Displacements Section -->
          ${partLcs.length ? `
          <div style="margin: 1.5rem 0 0.75rem;">
            <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">FEA Displacements</h4>
            <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.75rem;">
              Capture in nTop (Win+Shift+S) and click Paste from Clipboard, or press Ctrl+V.
            </p>
          </div>

          <div class="paste-action-bar">
            <button type="button" class="btn-paste-clipboard" onclick="window.pasteFeaFromClipboard('${partName}', 'new')">Paste from Clipboard</button>
            <button type="button" class="btn-browse-file" onclick="document.getElementById('fileInpNewIterFea').click()">Browse File</button>
            <span class="paste-hint-text">Or press Ctrl+V anywhere on this tab</span>
          </div>
          <input type="file" id="fileInpNewIterFea" accept="image/*" style="display: none;">

          <div class="card-table-wrapper" style="margin-top: 0.75rem; margin-bottom: 1.5rem; max-width: 650px;">
            <table class="styled-table">
              <thead>
                <tr>
                  <th style="width: 60%;">Load Case</th>
                  <th style="width: 40%;">Displacement (mm)</th>
                </tr>
              </thead>
              <tbody>
                ${partLcs.map((lc) => {
                  const curVal = typeof evalStiff[lc] === "object" ? evalStiff[lc].disp : (evalStiff[lc] || 0.0);
                  const desc = descMap[lc] || "";
                  const hasImg = newIterLcImages[lc];
                  return `
                    <tr>
                      <td>
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                          <div>
                            <strong style="color: #ffffff;">${lc}</strong>
                            <span style="color: var(--text-muted); font-size: 0.82rem; margin-left: 0.5rem;">${desc}</span>
                          </div>
                          <span id="new_badge_${lc}" class="tag" style="display: ${hasImg ? "inline-block" : "none"}; font-size: 0.65rem; color: var(--delta-green); border-color: var(--delta-green);">Attached</span>
                        </div>
                      </td>
                      <td>
                        <input type="number" step="${isHB ? "0.00001" : "0.0001"}" id="newLC_${lc}" class="form-input" style="padding: 0.35rem 0.6rem; font-family: monospace;" value="${curVal}">
                      </td>
                    </tr>
                  `;
                }).join("")}
              </tbody>
            </table>
          </div>
          ` : ""}

          <div style="display: flex; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-primary" id="btnSaveNewIterInline">Save Iteration</button>
          </div>
          `}
        </div>
      </div>

      <!-- Tab: Edit -->
      <div class="tab-pane ${activePartTab === "tabEdit" ? "active" : ""}" id="tabEdit">
        <div class="inline-tab-card">
          <div style="margin-bottom: 1.25rem;">
            <h3 style="font-size: 1.1rem; margin-bottom: 0.25rem;">Edit Iteration</h3>
            <p style="font-size: 0.82rem; color: var(--text-muted);">Modify recorded parameters for ${partName}</p>
          </div>

          ${!isEditor() ? `
          <div style="padding: 2rem; text-align: center; color: var(--text-muted);">
            Editor access required to edit iterations. Enter fine-grained PAT in Cloud Settings.
          </div>
          ` : `
          <div style="display: flex; gap: 1rem; align-items: center; margin-bottom: 1rem;">
            <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
              Select Iteration:
              <select id="editSelectVer" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
                ${iters.map((it) => `<option value="${it.Version}" ${String(it.Version) === String(evalRow.Version) ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
              </select>
            </label>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label>Version</label>
              <input type="text" class="form-input" value="${formatVersion(evalRow.Version)}" disabled>
            </div>
            <div class="form-group">
              <label for="editIterWeight">Weight (g)</label>
              <input type="number" step="0.1" id="editIterWeight" class="form-input" value="${parseFloat(evalRow["Weight (g)"] || 0)}">
            </div>
          </div>

          <div class="form-group">
            <label for="editIterDesc">Description</label>
            <input type="text" id="editIterDesc" class="form-input" value="${evalRow.Description || ""}">
          </div>

          <div class="form-group">
            <label for="editIterSpecs">Specifications</label>
            <textarea id="editIterSpecs" class="form-textarea">${evalRow.Specifications || ""}</textarea>
          </div>

          <!-- CAD Schematics Section -->
          <div style="margin: 1.5rem 0 0.75rem;">
            <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">CAD Schematics</h4>
            <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.75rem;">
              Capture in CAD (Win+Shift+S) and click Paste Schematic, or browse files.
            </p>
          </div>

          <div class="paste-action-bar">
            <button type="button" class="btn-paste-clipboard" onclick="window.pasteSchematicFromClipboard('${partName}', 'edit')">Paste Schematic</button>
            <button type="button" class="btn-browse-file" onclick="document.getElementById('fileInpEditIterSchematic').click()">Browse File</button>
            <span class="paste-hint-text">Or drag and drop files below</span>
          </div>
          <input type="file" id="fileInpEditIterSchematic" accept="image/*" multiple style="display: none;">

          <div id="editSchematicPreview" style="margin-bottom: 1.25rem;">
            ${renderSchematicThumbnails(evalRow, partName)}
          </div>

          <!-- FEA Displacements Section -->
          ${partLcs.length ? `
          <div style="margin: 1.5rem 0 0.75rem;">
            <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">FEA Displacements</h4>
            <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.75rem;">
              Capture in nTop and click Paste from Clipboard, or press Ctrl+V.
            </p>
          </div>

          <div class="paste-action-bar">
            <button type="button" class="btn-paste-clipboard" onclick="window.pasteFeaFromClipboard('${partName}', 'edit')">Paste from Clipboard</button>
            <button type="button" class="btn-browse-file" onclick="document.getElementById('fileInpEditIterFea').click()">Browse File</button>
            <span class="paste-hint-text">Or press Ctrl+V anywhere on this tab</span>
          </div>
          <input type="file" id="fileInpEditIterFea" accept="image/*" style="display: none;">

          <div class="card-table-wrapper" style="margin-top: 0.75rem; margin-bottom: 1.5rem; max-width: 650px;">
            <table class="styled-table">
              <thead>
                <tr>
                  <th style="width: 60%;">Load Case</th>
                  <th style="width: 40%;">Displacement (mm)</th>
                </tr>
              </thead>
              <tbody>
                ${partLcs.map((lc) => {
                  const curVal = typeof evalStiff[lc] === "object" ? evalStiff[lc].disp : (evalStiff[lc] || 0.0);
                  const desc = descMap[lc] || "";
                  const hasImg = editIterLcImages[lc] || (evalRow["LC Images"] && parseLcImages(evalRow["LC Images"])[lc]);
                  return `
                    <tr>
                      <td>
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                          <div>
                            <strong style="color: #ffffff;">${lc}</strong>
                            <span style="color: var(--text-muted); font-size: 0.82rem; margin-left: 0.5rem;">${desc}</span>
                          </div>
                          <span id="edit_badge_${lc}" class="tag" style="display: ${hasImg ? "inline-block" : "none"}; font-size: 0.65rem; color: var(--delta-green); border-color: var(--delta-green);">Attached</span>
                        </div>
                      </td>
                      <td>
                        <input type="number" step="${isHB ? "0.00001" : "0.0001"}" id="editLC_${lc}" class="form-input" style="padding: 0.35rem 0.6rem; font-family: monospace;" value="${curVal}">
                      </td>
                    </tr>
                  `;
                }).join("")}
              </tbody>
            </table>
          </div>
          ` : ""}

          <div style="display: flex; gap: 0.5rem; justify-content: space-between; align-items: center; margin-top: 1.5rem; flex-wrap: wrap;">
            <div style="display: flex; gap: 0.5rem;">
              <button class="btn-action danger" onclick="window.deleteIteration('${partName}', '${evalRow.Version}')">Delete Iteration</button>
              <button class="btn-action danger" onclick="window.deletePart('${partName}')">Delete Part</button>
            </div>
            <button class="btn-primary" id="btnSaveEditIterInline">Save Iteration</button>
          </div>
          `}
        </div>
      </div>

      <!-- Tab: Schematic Gallery -->
      <div class="tab-pane ${activePartTab === "tabSchematic" ? "active" : ""}" id="tabSchematic">
        ${renderGallery(evalRow, partName)}
      </div>
    `;

    elContentArea.innerHTML = html;

    // Attach Version Selectors
    const selEval = document.getElementById("selectEvalVer");
    const selRef = document.getElementById("selectRefVer");
    if (selEval) {
      selEval.addEventListener("change", (e) => {
        selectedEvalVersion = e.target.value;
        activeGalleryIndex = 0;
        renderPartView(partName);
      });
    }
    if (selRef) {
      selRef.addEventListener("change", (e) => {
        selectedRefVersion = e.target.value;
        renderPartView(partName);
      });
    }

    // Attach Tab Navigation strictly to main-tabs-header
    elContentArea.querySelectorAll(".main-tabs-header .tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        elContentArea.querySelectorAll(".main-tabs-header .tab-btn").forEach((b) => b.classList.remove("active"));
        elContentArea.querySelectorAll(".tab-pane").forEach((p) => p.classList.remove("active"));
        btn.classList.add("active");
        const targetId = btn.getAttribute("data-tab");
        activePartTab = targetId;
        const targetPane = document.getElementById(targetId);
        if (targetPane) targetPane.classList.add("active");
      });
    });

    // Attach listeners for inline actions
    attachInlineFormListeners(partName, evalRow, iters);
    attachGalleryEventListeners(evalRow, partName);
    attachFeaListeners(evalRow, partName);
  }

  function renderSideSchematic(row) {
    const imgs = parseImagePaths(row["Image Path"]);
    if (!imgs.length) {
      return `<div style="color: var(--text-dim); font-size: 0.82rem; padding: 1rem; text-align: center;">No schematic image.</div>`;
    }
    const clean = imgs[0].startsWith("data:") ? imgs[0] : `../${imgs[0]}`;
    return `
      <div style="cursor: pointer; text-align: center; margin-top: 0.5rem;" onclick="window.openLightbox('${clean}')">
        <img src="${clean}" alt="Schematic" style="max-width: 100%; max-height: 240px; object-fit: contain; border-radius: var(--radius-sm); border: 1px solid var(--border-color);" onerror="this.src='';">
      </div>
    `;
  }

  function attachInlineFormListeners(partName, evalRow, iters) {
    // 1. Justification Toggle & Save
    const cbJust = document.getElementById("cbIncludeJust");
    const justCont = document.getElementById("justEditorContainer");
    if (cbJust && justCont) {
      cbJust.onchange = () => {
        justCont.style.display = cbJust.checked ? "block" : "none";
      };
    }

    const btnSaveJust = document.getElementById("btnSaveJustInline");
    if (btnSaveJust) {
      btnSaveJust.onclick = async () => {
        const isEnabled = cbJust ? cbJust.checked : false;
        const txt = isEnabled && document.getElementById("textareaJustInline")
          ? document.getElementById("textareaJustInline").value.trim()
          : null;

        partsData.forEach((p) => {
          if (p["Part Name"] === partName) {
            p.Justification = txt || null;
          }
        });
        try {
          await window.DTCloud.commitFile("data/parts.json", partsData, `Update justification for ${partName}`);
          showToast(isEnabled && txt ? "Justification saved!" : "Justification disabled!");
          renderPartView(partName);
        } catch (e) {
          alert("Error saving justification: " + e.message);
        }
      };
    }

    // 2. Tolerances Toggle & Save
    const cbTol = document.getElementById("cbIncludeTol");
    const tolCont = document.getElementById("tolTableContainer");
    const btnEditTol = document.getElementById("btnEditTol");
    if (cbTol) {
      cbTol.onchange = () => {
        if (tolCont) tolCont.style.display = cbTol.checked ? "block" : "none";
        if (btnEditTol) btnEditTol.style.display = cbTol.checked ? "inline-block" : "none";
      };
    }

    const btnSaveTol = document.getElementById("btnSaveTolInline");
    if (btnSaveTol) {
      btnSaveTol.onclick = async () => {
        const isEnabled = cbTol ? cbTol.checked : false;
        if (!isEnabled) {
          partsData.forEach((p) => {
            if (p["Part Name"] === partName) {
              p["Tolerancing Data"] = null;
            }
          });
          try {
            await window.DTCloud.commitFile("data/parts.json", partsData, `Remove tolerances for ${partName}`);
            showToast("Tolerances disabled!");
            renderPartView(partName);
          } catch (e) {
            alert("Error removing tolerances: " + e.message);
          }
        } else {
          window.openTolerancesModal(partName);
        }
      };
    }

    // 3. New Iteration Copy From Selector
    const selCopy = document.getElementById("selectNewIterSource");
    if (selCopy) {
      selCopy.onchange = (e) => {
        const srcVer = e.target.value;
        const srcRow = iters.find((p) => String(p.Version) === String(srcVer));
        if (srcRow) {
          const wInp = document.getElementById("newIterWeight");
          const dInp = document.getElementById("newIterDesc");
          const sInp = document.getElementById("newIterSpecs");
          if (wInp) wInp.value = parseFloat(srcRow["Weight (g)"] || 0);
          if (dInp) dInp.value = srcRow.Description || "";
          if (sInp) sInp.value = srcRow.Specifications || "";
        }
      };
    }

    // 4. New Iteration File Browse Input
    const fileNewFea = document.getElementById("fileInpNewIterFea");
    if (fileNewFea) {
      fileNewFea.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
          handleFeaScreenshotPasted(e.target.files[0], partName, "new");
        }
      };
    }

    // 5. Edit Iteration File Browse Input (FEA & Schematic)
    const fileEditFea = document.getElementById("fileInpEditIterFea");
    if (fileEditFea) {
      fileEditFea.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
          handleFeaScreenshotPasted(e.target.files[0], partName, "edit");
        }
      };
    }

    const fileEditSch = document.getElementById("fileInpEditIterSchematic");
    if (fileEditSch) {
      fileEditSch.onchange = async (e) => {
        if (e.target.files && e.target.files.length) {
          await uploadImagesToIteration(Array.from(e.target.files), evalRow, partName);
          showToast("Schematic attached successfully!");
        }
      };
    }

    // 6. New Iteration Schematic Upload
    const dropNewSch = document.getElementById("dropzoneNewIterSchematic");
    const fileNewSch = document.getElementById("fileInpNewIterSchematic");
    const labelNewSch = document.getElementById("labelNewIterSchematic");
    if (dropNewSch && fileNewSch) {
      dropNewSch.onclick = () => fileNewSch.click();
      fileNewSch.onchange = async (e) => {
        if (e.target.files && e.target.files[0]) {
          const downscaled = await downscaleImage(e.target.files[0]);
          newIterSchematic = downscaled;
          if (labelNewSch) labelNewSch.textContent = `Schematic loaded (${e.target.files[0].name})`;
          dropNewSch.style.borderColor = "var(--delta-green)";
        }
      };
    }

    // 7. Save New Iteration Button
    const btnSaveNew = document.getElementById("btnSaveNewIterInline");
    if (btnSaveNew) {
      btnSaveNew.onclick = async () => {
        const v = document.getElementById("newIterVer").value.trim();
        const w = parseFloat(document.getElementById("newIterWeight").value) || 0;
        const desc = document.getElementById("newIterDesc").value.trim();
        const specs = document.getElementById("newIterSpecs").value.trim();

        if (!v) {
          alert("Version string is required.");
          return;
        }

        if (iters.some((it) => String(it.Version) === v)) {
          alert("Version " + v + " already exists for this part.");
          return;
        }

        const isHB = partName === "HB_01";
        const isUP = partName === "UP_01";
        const partLcs = isUP
          ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]
          : isHB
          ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"]
          : [];

        const stiffObj = {};
        if (partLcs.length) {
          partLcs.forEach((lc) => {
            const inp = document.getElementById("newLC_" + lc);
            stiffObj[lc] = inp ? parseFloat(inp.value) || 0 : 0;
          });
        }

        const latest = iters[iters.length - 1];
        const now = new Date().toISOString().replace("T", " ").slice(0, 19);

        let finalLcImages = {};
        try {
          if (latest["LC Images"]) finalLcImages = JSON.parse(latest["LC Images"]);
        } catch (e) {}
        Object.assign(finalLcImages, newIterLcImages);

        const newRow = {
          Subassembly: latest.Subassembly || "Drivetrain",
          "Part Name": partName,
          Version: v,
          Date: now,
          "Weight (g)": w,
          Description: desc,
          Specifications: specs,
          "Image Path": newIterSchematic || latest["Image Path"] || null,
          "Stiffness Data": partLcs.length ? JSON.stringify(stiffObj) : latest["Stiffness Data"],
          "LC Images": Object.keys(finalLcImages).length ? JSON.stringify(finalLcImages) : null,
          "Tolerancing Data": latest["Tolerancing Data"] || null,
          Justification: latest["Justification"] || null
        };

        partsData.push(newRow);

        try {
          await window.DTCloud.commitFile("data/parts.json", partsData, `Add ${partName} v${v}`);
          selectedEvalVersion = v;
          newIterLcImages = {};
          newIterSchematic = null;
          activePartTab = "tabOverview";
          buildSidebar();
          renderPartView(partName);
          showToast(`Iteration v${v} saved!`);
        } catch (e) {
          alert("Error saving iteration: " + e.message);
        }
      };
    }

    // 8. Edit Iteration Selector & Save Button
    const editSelVer = document.getElementById("editSelectVer");
    if (editSelVer) {
      editSelVer.onchange = (e) => {
        selectedEvalVersion = e.target.value;
        renderPartView(partName);
      };
    }

    const btnSaveEdit = document.getElementById("btnSaveEditIterInline");
    if (btnSaveEdit) {
      btnSaveEdit.onclick = async () => {
        const targetVer = editSelVer ? editSelVer.value : evalRow.Version;
        const row = partsData.find((p) => p["Part Name"] === partName && String(p.Version) === String(targetVer));
        if (!row) return;

        const w = parseFloat(document.getElementById("editIterWeight").value) || 0;
        const desc = document.getElementById("editIterDesc").value.trim();
        const specs = document.getElementById("editIterSpecs").value.trim();

        row["Weight (g)"] = w;
        row.Description = desc;
        row.Specifications = specs;

        const isHB = partName === "HB_01";
        const isUP = partName === "UP_01";
        const partLcs = isUP
          ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]
          : isHB
          ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"]
          : [];

        if (partLcs.length) {
          const stiffObj = {};
          partLcs.forEach((lc) => {
            const inp = document.getElementById("editLC_" + lc);
            stiffObj[lc] = inp ? parseFloat(inp.value) || 0 : 0;
          });
          row["Stiffness Data"] = JSON.stringify(stiffObj);
        }

        if (Object.keys(editIterLcImages).length) {
          let curLc = {};
          try {
            if (row["LC Images"]) curLc = JSON.parse(row["LC Images"]);
          } catch (e) {}
          Object.assign(curLc, editIterLcImages);
          row["LC Images"] = JSON.stringify(curLc);
        }

        try {
          await window.DTCloud.commitFile("data/parts.json", partsData, `Update ${partName} v${targetVer}`);
          editIterLcImages = {};
          activePartTab = "tabOverview";
          renderPartView(partName);
          showToast(`Iteration ${formatVersion(targetVer)} updated!`);
        } catch (e) {
          alert("Error saving iteration: " + e.message);
        }
      };
    }
  }

  // Global paste listener for FEA Screenshot OCR
  window.addEventListener("paste", async (e) => {
    if (!currentView || currentView === "assembly" || currentView === "requirements" || currentView === "trash") {
      return;
    }
    if (!isEditor()) return;
    if (activePartTab !== "tabNewIter" && activePartTab !== "tabEdit") {
      return;
    }

    if (e.clipboardData && e.clipboardData.items) {
      for (const item of e.clipboardData.items) {
        if (item.type.indexOf("image") !== -1) {
          e.preventDefault();
          const file = item.getAsFile();
          if (file) {
            handleFeaScreenshotPasted(file, currentView, activePartTab === "tabNewIter" ? "new" : "edit");
          }
          break;
        }
      }
    }
  });

  // Specifications parser
  function renderSpecsRows(specStr) {
    if (!specStr || !String(specStr).trim()) {
      return `<tr><td colspan="2" style="color: var(--text-dim);">No specifications defined.</td></tr>`;
    }
    const lines = String(specStr).split("\n");
    return lines
      .map((l) => {
        if (l.includes(":")) {
          const parts = l.split(":");
          return `<tr><td><strong>${parts[0].trim()}</strong></td><td>${parts.slice(1).join(":").trim()}</td></tr>`;
        }
        return `<tr><td colspan="2">${l.trim()}</td></tr>`;
      })
      .join("");
  }

  // Tolerancing parser
  function renderTolRows(tolData) {
    if (!tolData) return `<tr><td colspan="4" style="color: var(--text-dim);">No GD&T tolerances specified.</td></tr>`;
    let records = [];
    try {
      if (typeof tolData === "string") {
        records = JSON.parse(tolData);
      } else if (Array.isArray(tolData)) {
        records = tolData;
      }
    } catch (e) {
      return `<tr><td colspan="4" style="color: var(--text-dim);">Invalid GD&T format.</td></tr>`;
    }

    if (!records.length) {
      return `<tr><td colspan="4" style="color: var(--text-dim);">No GD&T tolerances specified.</td></tr>`;
    }

    return records
      .map(
        (r) => `
        <tr>
          <td><strong>${r["Interface / Feature"] || r["Interface / Elemento"] || "—"}</strong></td>
          <td class="mono">${r["Nominal (mm)"] || "—"}</td>
          <td class="mono">${r["Fit / Tolerance"] || r["Ajuste / Tolerância"] || "—"}</td>
          <td style="color: var(--text-muted);">${r["Roughness / Notes"] || r["Rugosidade / Notas"] || "—"}</td>
        </tr>
      `
      )
      .join("");
  }

  // Parse LC Images dictionary
  function parseLcImages(raw) {
    if (!raw) return {};
    if (typeof raw === "object") return raw;
    try {
      return JSON.parse(raw) || {};
    } catch (e) {
      return {};
    }
  }

  // FEA Displacements Table Renderer
  function renderFEATable(evalRow, refRow, showFull = false) {
    let evalStiff = {};
    let refStiff = {};
    try {
      if (evalRow["Stiffness Data"]) evalStiff = JSON.parse(evalRow["Stiffness Data"]);
      if (refRow["Stiffness Data"]) refStiff = JSON.parse(refRow["Stiffness Data"]);
    } catch (e) {}

    let lcKeys = Object.keys(evalStiff).length ? Object.keys(evalStiff) : Object.keys(refStiff);
    if (!lcKeys.length) {
      if (evalRow["Part Name"] === "HB_01") {
        lcKeys = ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"];
      } else if (evalRow["Part Name"] === "UP_01") {
        lcKeys = ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"];
      } else {
        const fallbackMsg = `
          <h3 style="font-size: 1rem; margin: 1.5rem 0 0.5rem;">FEA Displacements</h3>
          <p style="color: var(--text-dim); font-size: 0.88rem;">No FEA load-case displacement data recorded for this iteration.</p>
        `;
        return showFull ? fallbackMsg + renderFeaScreenshots(evalRow, refRow) : fallbackMsg;
      }
    }

    const descMap = {
      LC1: "Pure Braking",
      LC2: "Pure Acceleration",
      LC3: "Pure Cornering",
      LC4: "Braking + Cornering 1",
      LC5: "Braking + Cornering 2",
      LC6: "Acceleration + Cornering 1",
      LC7: "Acceleration + Cornering 2",
      LC8: "1kN Fy",
      LC9: "-1kN Fy"
    };

    const isHB = evalRow["Part Name"] === "HB_01";
    const decimals = isHB ? 5 : 4;

    let rowsHtml = "";
    lcKeys.forEach((lc) => {
      const eVal = typeof evalStiff[lc] === "object" ? evalStiff[lc].disp : parseFloat(evalStiff[lc] || 0);
      const rVal = typeof refStiff[lc] === "object" ? refStiff[lc].disp : parseFloat(refStiff[lc] || 0);
      const diff = eVal - rVal;
      const pct = rVal > 0 ? (diff / rVal) * 100 : 0;
      const isZero = Math.abs(diff) < 1e-7;
      const deltaClass = isZero ? "delta-neutral" : (diff < 0 ? "delta-good" : "delta-bad");

      rowsHtml += `
        <tr>
          <td><strong>${lc}</strong></td>
          <td style="color: var(--text-muted); font-size: 0.82rem;">${descMap[lc] || lc}</td>
          <td class="mono">${eVal.toFixed(decimals)} mm</td>
          <td class="mono">${rVal.toFixed(decimals)} mm</td>
          <td class="mono ${deltaClass}">
            <strong>${isZero ? "0.0" : (diff <= 0 ? "" : "+")}${diff.toFixed(decimals)} mm</strong>
          </td>
          <td class="mono ${deltaClass}">
            <strong>${isZero ? "0.0%" : (diff <= 0 ? "" : "+") + pct.toFixed(1) + "%"}</strong>
          </td>
        </tr>
      `;
    });

    let html = `
      <h3 style="font-size: 1rem; margin: 1.5rem 0 0.5rem;">FEA Displacements</h3>
      <div class="card-table-wrapper">
        <table class="styled-table">
          <thead>
            <tr>
              <th>Load Case</th>
              <th>Description</th>
              <th>${formatVersion(evalRow.Version)}</th>
              <th>${formatVersion(refRow.Version)}</th>
              <th>Diff (mm)</th>
              <th>%</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>
    `;

    if (showFull) {
      html += renderFeaScreenshots(evalRow, refRow);
    }

    return html;
  }

  function renderFeaScreenshots(evalRow, refRow) {
    const evalLcMap = parseLcImages(evalRow["LC Images"]);
    const refLcMap = parseLcImages(refRow["LC Images"]);
    const partName = evalRow["Part Name"];

    let allLcs = [];
    if (partName === "UP_01") {
      allLcs = ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"];
    } else if (partName === "HB_01") {
      allLcs = ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"];
    } else {
      allLcs = Array.from(new Set([...Object.keys(evalLcMap), ...Object.keys(refLcMap)]));
    }

    const availLcs = allLcs.filter(lc => evalLcMap[lc] || refLcMap[lc]);
    if (!availLcs.length) {
      return `
        <div style="margin-top: 1.5rem; padding: 2rem; text-align: center; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-md); color: var(--text-dim);">
          No FEA screenshots attached.
        </div>
      `;
    }

    if (!availLcs.includes(activeFeaLc)) {
      activeFeaLc = availLcs[0];
    }

    const evalImg = evalLcMap[activeFeaLc];
    const refImg = refLcMap[activeFeaLc];

    const cleanEval = evalImg ? `../${evalImg.replace(/\\/g, "/")}` : null;
    const cleanRef = refImg ? `../${refImg.replace(/\\/g, "/")}` : null;
    const isDistinct = String(evalRow.Version) !== String(refRow.Version);

    let html = `
      <div style="margin-top: 1.5rem;" id="feaScreenshotsContainer">
        <h3 style="font-size: 1rem; margin-bottom: 0.75rem;">FEA Screenshots</h3>
        <div class="fea-lc-header">
          ${availLcs.map(lc => `
            <button class="fea-lc-btn ${lc === activeFeaLc ? "active" : ""}" onclick="window.selectFeaLc('${lc}')">${lc}</button>
          `).join("")}
        </div>
    `;

    if (isDistinct && cleanRef && cleanEval) {
      html += `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 1rem;">
          <div class="card-table-wrapper" style="padding: 0.75rem; text-align: center;">
            <div style="font-size: 0.82rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.5rem;">
              Baseline: ${formatVersion(refRow.Version)} (${activeFeaLc})
            </div>
            <div class="gallery-viewport" style="min-height: 280px; max-height: 380px;" onclick="window.openLightbox('${cleanRef}')">
              <img src="${cleanRef}" alt="${activeFeaLc} Baseline" class="gallery-main-img" onerror="this.src='';">
            </div>
          </div>
          <div class="card-table-wrapper" style="padding: 0.75rem; text-align: center;">
            <div style="font-size: 0.82rem; font-weight: 600; color: var(--accent-red); margin-bottom: 0.5rem;">
              Evaluated: ${formatVersion(evalRow.Version)} (${activeFeaLc})
            </div>
            <div class="gallery-viewport" style="min-height: 280px; max-height: 380px;" onclick="window.openLightbox('${cleanEval}')">
              <img src="${cleanEval}" alt="${activeFeaLc} Evaluated" class="gallery-main-img" onerror="this.src='';">
            </div>
          </div>
        </div>
      `;
    } else if (cleanEval) {
      html += `
        <div class="card-table-wrapper" style="padding: 0.75rem; text-align: center;">
          <div style="font-size: 0.82rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.5rem;">
            ${formatVersion(evalRow.Version)} (${activeFeaLc})
          </div>
          <div class="gallery-viewport" style="min-height: 320px; max-height: 480px;" onclick="window.openLightbox('${cleanEval}')">
            <img src="${cleanEval}" alt="${activeFeaLc}" class="gallery-main-img" onerror="this.src='';">
          </div>
        </div>
      `;
    } else if (cleanRef) {
      html += `
        <div class="card-table-wrapper" style="padding: 0.75rem; text-align: center;">
          <div style="font-size: 0.82rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.5rem;">
            ${formatVersion(refRow.Version)} (${activeFeaLc})
          </div>
          <div class="gallery-viewport" style="min-height: 320px; max-height: 480px;" onclick="window.openLightbox('${cleanRef}')">
            <img src="${cleanRef}" alt="${activeFeaLc}" class="gallery-main-img" onerror="this.src='';">
          </div>
        </div>
      `;
    }

    if (isEditor()) {
      html += `
        <div style="margin-top: 1rem; display: flex; justify-content: flex-end;">
          <input type="file" id="inputUploadFeaLc" accept="image/*" style="display: none;">
          <button class="btn-action" onclick="document.getElementById('inputUploadFeaLc').click()">Upload ${activeFeaLc}</button>
        </div>
      `;
    }

    html += `</div>`;
    return html;
  }

  function attachFeaListeners(evalRow, partName) {
    const fileInp = document.getElementById("inputUploadFeaLc");
    if (!fileInp) return;

    fileInp.onchange = async () => {
      const file = fileInp.files[0];
      if (!file) return;

      try {
        const dataUrl = await downscaleImage(file, 1920, 1080);
        const timestamp = new Date().toISOString().replace(/[-:T.]/g, "").slice(0, 14);
        const vClean = formatVersion(evalRow.Version).replace(/[^a-zA-Z0-9_-]/g, "_");
        const filename = `images/${partName}_${vClean}_${activeFeaLc}_${timestamp}.png`;

        if (window.DTCloud.isConfigured()) {
          await window.DTCloud.commitBinaryFile(filename, dataUrl, `Upload ${filename}`);
        }

        const lcMap = parseLcImages(evalRow["LC Images"]);
        lcMap[activeFeaLc] = filename;
        evalRow["LC Images"] = JSON.stringify(lcMap);

        await window.DTCloud.commitFile("data/parts.json", partsData, `Update ${activeFeaLc} for ${partName} ${evalRow.Version}`);
        renderPartView(partName);
      } catch (err) {
        alert("Upload error: " + err.message);
      }
    };
  }

  window.selectFeaLc = function (lc) {
    activeFeaLc = lc;
    const iters = partsData.filter((p) => p["Part Name"] === currentView);
    if (!iters.length) return;
    const evalRow = iters.find((p) => String(p.Version) === String(selectedEvalVersion)) || iters[iters.length - 1];
    const refRow = iters.find((p) => String(p.Version) === String(selectedRefVersion)) || iters[0];
    const container = document.getElementById("feaScreenshotsContainer");
    if (container) {
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = renderFeaScreenshots(evalRow, refRow);
      const inner = tempDiv.querySelector("#feaScreenshotsContainer");
      if (inner) {
        container.innerHTML = inner.innerHTML;
        attachFeaListeners(evalRow, currentView);
        return;
      }
    }
    renderCurrentView();
  };

  // Multi-image Gallery Renderer
  function renderGallery(evalRow, partName) {
    const images = parseImagePaths(evalRow["Image Path"]);
    if (activeGalleryIndex >= images.length) {
      activeGalleryIndex = Math.max(0, images.length - 1);
    }

    const currentImg = images[activeGalleryIndex];
    const currentClean = currentImg ? `../${currentImg.replace(/\\/g, "/")}` : "";

    let html = `
      <div class="gallery-container">
    `;

    if (images.length > 0) {
      html += `
        <div class="gallery-viewport" onclick="window.openLightbox('${currentClean}')">
          <img src="${currentClean}" alt="CAD Schematic" class="gallery-main-img" onerror="this.src='';">
          <span class="gallery-count-badge">${activeGalleryIndex + 1} / ${images.length}</span>
        </div>
      `;

      if (images.length > 1 || isEditor()) {
        html += `
          <div class="gallery-filmstrip">
            ${images.map((imgP, i) => {
              const cleanThumb = `../${imgP.replace(/\\/g, "/")}`;
              const isActive = i === activeGalleryIndex ? "active" : "";
              return `
                <div class="gallery-thumb-item">
                  <img src="${cleanThumb}" alt="Thumbnail ${i + 1}" class="gallery-thumb ${isActive}" onclick="window.selectGalleryImage(${i})">
                  ${isEditor() ? `<button class="gallery-thumb-delete" onclick="window.deleteGalleryImage('${partName}', '${evalRow.Version}', ${i})">X</button>` : ""}
                </div>
              `;
            }).join("")}
          </div>
        `;
      }
    } else {
      html += `
        <div style="padding: 3rem; text-align: center; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-md); color: var(--text-dim);">
          No CAD schematic or render attached for this iteration.
        </div>
      `;
    }

    if (isEditor()) {
      html += `
        <div class="gallery-dropzone" id="galleryDropzone">
          <input type="file" id="fileGalleryUpload" accept="image/*" multiple style="display: none;">
          <div>Drag and drop images, click to upload, or paste (Ctrl+V)</div>
          <div class="form-hint">Images auto-scaled to max 1920x1080 to prevent repository bloat</div>
        </div>
      `;
    }

    html += `</div>`;
    return html;
  }

  // Select gallery index
  window.selectGalleryImage = function (index) {
    activeGalleryIndex = index;
    renderCurrentView();
  };

  // Delete gallery image
  window.deleteGalleryImage = async function (partName, version, imgIndex) {
    if (!isEditor()) return;
    if (!confirm("Remove this image from iteration?")) return;

    const row = partsData.find((p) => p["Part Name"] === partName && String(p.Version) === String(version));
    if (!row) return;

    const images = parseImagePaths(row["Image Path"]);
    images.splice(imgIndex, 1);
    row["Image Path"] = formatImagePaths(images);

    try {
      await window.DTCloud.commitFile("data/parts.json", partsData, `Remove image from ${partName} ${version}`);
      if (activeGalleryIndex >= images.length) {
        activeGalleryIndex = Math.max(0, images.length - 1);
      }
      renderPartView(partName);
    } catch (e) {
      alert("Error removing image: " + e.message);
    }
  };

  // Attach gallery listeners for upload & clipboard paste
  function attachGalleryEventListeners(evalRow, partName) {
    const dropzone = document.getElementById("galleryDropzone");
    const fileInput = document.getElementById("fileGalleryUpload");
    if (!dropzone || !fileInput) return;

    dropzone.onclick = () => fileInput.click();

    fileInput.onchange = async () => {
      const files = Array.from(fileInput.files);
      if (!files.length) return;
      await uploadImagesToIteration(files, evalRow, partName);
    };

    dropzone.ondragover = (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    };

    dropzone.ondragleave = () => {
      dropzone.classList.remove("dragover");
    };

    dropzone.ondrop = async (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
      const files = Array.from(e.dataTransfer.files).filter((f) => f.type.startsWith("image/"));
      if (files.length) {
        await uploadImagesToIteration(files, evalRow, partName);
      }
    };

    // Global paste listener when viewing this part
    window.onpaste = async (e) => {
      if (currentView !== partName) return;
      const items = (e.clipboardData || e.originalEvent.clipboardData).items;
      const imageFiles = [];
      for (const item of items) {
        if (item.kind === "file" && item.type.startsWith("image/")) {
          imageFiles.push(item.getAsFile());
        }
      }
      if (imageFiles.length) {
        await uploadImagesToIteration(imageFiles, evalRow, partName);
      }
    };
  }

  // Upload multiple images with downscaling
  async function uploadImagesToIteration(files, evalRow, partName) {
    const existingImages = parseImagePaths(evalRow["Image Path"]);
    const vStr = formatVersion(evalRow.Version);
    const versionClean = vStr.replace(/[^a-zA-Z0-9_-]/g, "_");

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const dataUrl = await downscaleImage(file, 1920, 1080);
        const timestamp = new Date().toISOString().replace(/[-:T.]/g, "").slice(0, 14);
        const filename = `images/${partName}_${versionClean}_${timestamp}_${i + 1}.png`;

        if (window.DTCloud.isConfigured()) {
          await window.DTCloud.commitBinaryFile(filename, dataUrl, `Upload ${filename}`);
        }
        existingImages.push(filename);
      } catch (err) {
        console.error("Image upload failed:", err);
        alert("Image upload error: " + err.message);
      }
    }

    evalRow["Image Path"] = formatImagePaths(existingImages);
    try {
      await window.DTCloud.commitFile("data/parts.json", partsData, `Attach images to ${partName} ${evalRow.Version}`);
      activeGalleryIndex = existingImages.length - 1;
      renderPartView(partName);
    } catch (err) {
      alert("Error saving iteration images: " + err.message);
    }
  }

  // Lightbox Modal (Click-to-zoom)
  window.openLightbox = function (imgSrc) {
    if (!imgSrc) return;
    elModalContainer.innerHTML = `
      <div class="lightbox-modal" onclick="window.closeLightbox()">
        <button class="lightbox-close" onclick="window.closeLightbox()">Close</button>
        <img src="${imgSrc}" class="lightbox-img" onclick="event.stopPropagation();">
      </div>
    `;

    window.onkeydown = (e) => {
      if (e.key === "Escape") window.closeLightbox();
    };
  };

  window.closeLightbox = function () {
    elModalContainer.innerHTML = "";
    window.onkeydown = null;
  };

  // Cloud Settings Modal
  function openCloudSettingsModal() {
    const cfg = window.DTCloud.getConfig();
    elModalContainer.innerHTML = `
      <div class="modal-overlay open" id="modalCloud">
        <div class="modal-box">
          <div class="modal-header">
            <h2 class="modal-title">Cloud Settings</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1rem;">
            Configure your GitHub Personal Access Token to unlock editor permissions and 24/7 synchronization.
          </p>
          <div class="form-group">
            <label for="inputToken">Access Token</label>
            <input type="password" id="inputToken" class="form-input" placeholder="ghp_xxxxxxxxxxxxxxxxxxxx" value="${cfg.token || ""}">
          </div>
          <div class="form-group">
            <label for="inputRepo">Repository</label>
            <input type="text" id="inputRepo" class="form-input" value="${cfg.repo || "diogoribeir07/fst16-drivetrain"}">
          </div>
          <div class="form-group">
            <label for="inputBranch">Branch</label>
            <input type="text" id="inputBranch" class="form-input" value="${cfg.branch || "main"}">
          </div>
          <div style="display: flex; gap: 0.5rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
            <button class="btn-primary" id="btnSaveCloudCfg">Save Settings</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveCloudCfg").onclick = () => {
      const tok = document.getElementById("inputToken").value.trim();
      const rep = document.getElementById("inputRepo").value.trim();
      const br = document.getElementById("inputBranch").value.trim();
      window.DTCloud.saveConfig(tok, rep, br);
      window.closeModal();
      alert("Settings saved successfully.");
      init();
    };
  }

  // Modal Closer
  window.closeModal = function () {
    elModalContainer.innerHTML = "";
  };

  // Modal: + New Iteration (Clones previous iteration)
  window.openNewIterationModal = function (partName) {
    if (!isEditor()) return;
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    const latest = iters[iters.length - 1] || {};

    const nextVer = incrementVersion(latest.Version || "1.0");
    const isHB = partName === "HB_01";
    const isUP = partName === "UP_01";
    const partLcs = isUP
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]
      : isHB
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"]
      : [];

    let latestStiff = {};
    try {
      if (latest["Stiffness Data"]) latestStiff = JSON.parse(latest["Stiffness Data"]);
    } catch (e) {}

    let lcInputsHtml = "";
    if (partLcs.length) {
      lcInputsHtml = `
        <h4 style="font-size: 0.9rem; margin: 1rem 0 0.5rem;">FEA Displacements</h4>
        <div class="form-row">
          ${partLcs.map((lc) => {
            const curVal = typeof latestStiff[lc] === "object" ? latestStiff[lc].disp : (latestStiff[lc] || 0.0);
            return `
              <div class="form-group">
                <label for="inputLC_${lc}">${lc} (mm)</label>
                <input type="number" step="${isHB ? "0.00001" : "0.0001"}" id="inputLC_${lc}" class="form-input" value="${curVal}">
              </div>
            `;
          }).join("")}
        </div>
      `;
    }

    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-lg">
          <div class="modal-header">
            <h2 class="modal-title">New Iteration</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Part Name</label>
              <input type="text" class="form-input" value="${partName}" disabled>
            </div>
            <div class="form-group">
              <label for="inputNewVer">Version</label>
              <input type="text" id="inputNewVer" class="form-input" value="${nextVer}">
            </div>
            <div class="form-group">
              <label for="inputNewWeight">Weight (g)</label>
              <input type="number" step="0.1" id="inputNewWeight" class="form-input" value="${parseFloat(latest["Weight (g)"] || 0)}">
            </div>
          </div>
          <div class="form-group">
            <label for="inputNewDesc">Description</label>
            <input type="text" id="inputNewDesc" class="form-input" placeholder="Iteration description" value="${latest.Description || ""}">
          </div>
          <div class="form-group">
            <label for="inputNewSpecs">Specifications</label>
            <textarea id="inputNewSpecs" class="form-textarea" placeholder="Feature: Value">${latest.Specifications || ""}</textarea>
          </div>
          ${lcInputsHtml}
          <div style="display: flex; gap: 0.5rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
            <button class="btn-primary" id="btnSaveNewIter">Save Iteration</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveNewIter").onclick = async () => {
      const v = document.getElementById("inputNewVer").value.trim();
      const w = parseFloat(document.getElementById("inputNewWeight").value) || 0;
      const desc = document.getElementById("inputNewDesc").value.trim();
      const specs = document.getElementById("inputNewSpecs").value.trim();

      if (!v) {
        alert("Version string is required.");
        return;
      }

      // Check if version already exists
      if (iters.some((it) => String(it.Version) === v)) {
        alert("Version " + v + " already exists for this part.");
        return;
      }

      const stiffObj = {};
      if (partLcs.length) {
        partLcs.forEach((lc) => {
          const inp = document.getElementById(`inputLC_${lc}`);
          stiffObj[lc] = inp ? parseFloat(inp.value) || 0 : 0;
        });
      }

      const now = new Date().toISOString().replace("T", " ").slice(0, 19);
      const newRow = {
        "Subassembly": latest.Subassembly || "Drivetrain",
        "Part Name": partName,
        "Version": v,
        "Date": now,
        "Weight (g)": w,
        "Description": desc,
        "Specifications": specs,
        "Image Path": latest["Image Path"] || null,
        "Stiffness Data": partLcs.length ? JSON.stringify(stiffObj) : latest["Stiffness Data"],
        "LC Images": latest["LC Images"] || null,
        "Tolerancing Data": latest["Tolerancing Data"] || null,
        "Justification": latest["Justification"] || null
      };

      partsData.push(newRow);
      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Add ${partName} v${v}`);
        selectedEvalVersion = v;
        window.closeModal();
        buildSidebar();
        renderPartView(partName);
      } catch (e) {
        alert("Error saving new iteration: " + e.message);
      }
    };
  };

  // Helper to increment version string (e.g. 1.0 -> 1.1)
  function incrementVersion(verStr) {
    const clean = String(verStr).replace(/^v/i, "").trim();
    const parts = clean.split(".");
    if (parts.length >= 2) {
      const last = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(last)) {
        parts[parts.length - 1] = String(last + 1);
        return parts.join(".");
      }
    }
    const num = parseFloat(clean);
    if (!isNaN(num)) {
      return (num + 0.1).toFixed(1);
    }
    return clean + "_new";
  }

  // Modal: Edit Iteration
  window.openEditIterationModal = function (partName, version) {
    if (!isEditor()) return;
    const row = partsData.find((p) => p["Part Name"] === partName && String(p.Version) === String(version));
    if (!row) return;

    const isHB = partName === "HB_01";
    const isUP = partName === "UP_01";
    const partLcs = isUP
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]
      : isHB
      ? ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7"]
      : [];

    let currentStiff = {};
    try {
      if (row["Stiffness Data"]) currentStiff = JSON.parse(row["Stiffness Data"]);
    } catch (e) {}

    let lcInputsHtml = "";
    if (partLcs.length) {
      lcInputsHtml = `
        <h4 style="font-size: 0.9rem; margin: 1rem 0 0.5rem;">FEA Displacements</h4>
        <div class="form-row">
          ${partLcs.map((lc) => {
            const curVal = typeof currentStiff[lc] === "object" ? currentStiff[lc].disp : (currentStiff[lc] || 0.0);
            return `
              <div class="form-group">
                <label for="editLC_${lc}">${lc} (mm)</label>
                <input type="number" step="${isHB ? "0.00001" : "0.0001"}" id="editLC_${lc}" class="form-input" value="${curVal}">
              </div>
            `;
          }).join("")}
        </div>
      `;
    }

    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-lg">
          <div class="modal-header">
            <h2 class="modal-title">Edit Iteration</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Part Name</label>
              <input type="text" class="form-input" value="${partName}" disabled>
            </div>
            <div class="form-group">
              <label>Version</label>
              <input type="text" class="form-input" value="${formatVersion(version)}" disabled>
            </div>
            <div class="form-group">
              <label for="editWeight">Weight (g)</label>
              <input type="number" step="0.1" id="editWeight" class="form-input" value="${parseFloat(row["Weight (g)"] || 0)}">
            </div>
          </div>
          <div class="form-group">
            <label for="editDesc">Description</label>
            <input type="text" id="editDesc" class="form-input" value="${row.Description || ""}">
          </div>
          <div class="form-group">
            <label for="editSpecs">Specifications</label>
            <textarea id="editSpecs" class="form-textarea">${row.Specifications || ""}</textarea>
          </div>
          ${lcInputsHtml}
          <div style="display: flex; gap: 0.5rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
            <button class="btn-primary" id="btnSaveEditIter">Save Iteration</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveEditIter").onclick = async () => {
      row["Weight (g)"] = parseFloat(document.getElementById("editWeight").value) || 0;
      row.Description = document.getElementById("editDesc").value.trim();
      row.Specifications = document.getElementById("editSpecs").value.trim();

      if (partLcs.length) {
        const stiffObj = {};
        partLcs.forEach((lc) => {
          const inp = document.getElementById(`editLC_${lc}`);
          stiffObj[lc] = inp ? parseFloat(inp.value) || 0 : 0;
        });
        row["Stiffness Data"] = JSON.stringify(stiffObj);
      }

      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Update ${partName} v${version}`);
        window.closeModal();
        renderPartView(partName);
      } catch (e) {
        alert("Error saving iteration: " + e.message);
      }
    };
  };

  // Delete Iteration
  window.deleteIteration = async function (partName, version) {
    if (!isEditor()) return;
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    if (iters.length <= 1) {
      alert("Cannot delete sole iteration. Delete part instead.");
      return;
    }

    if (!confirm(`Delete iteration ${formatVersion(version)} for ${partName}?`)) return;

    const delIdx = partsData.findIndex((p) => p["Part Name"] === partName && String(p.Version) === String(version));
    if (delIdx !== -1) {
      partsData.splice(delIdx, 1);
      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Delete ${partName} v${version}`);
        const remaining = partsData.filter((p) => p["Part Name"] === partName);
        selectedEvalVersion = remaining[remaining.length - 1].Version;
        buildSidebar();
        renderPartView(partName);
      } catch (e) {
        alert("Error deleting iteration: " + e.message);
      }
    }
  };

  // Delete Part (Moves to Recycle Bin)
  window.deletePart = async function (partName) {
    if (!isEditor()) return;
    if (!confirm(`Delete part ${partName} and move all iterations to Recycle Bin?`)) return;

    const toDelete = partsData.filter((p) => p["Part Name"] === partName);
    trashData.push(...toDelete);
    partsData = partsData.filter((p) => p["Part Name"] !== partName);

    try {
      await Promise.all([
        window.DTCloud.commitFile("data/trash.json", trashData, `Move ${partName} to trash`),
        window.DTCloud.commitFile("data/parts.json", partsData, `Delete ${partName}`)
      ]);
      alert("Part moved to Recycle Bin.");
      currentView = "assembly";
      buildSidebar();
      renderCurrentView();
    } catch (e) {
      alert("Error deleting part: " + e.message);
    }
  };

  // Modal: + New Part
  window.openNewPartModal = function (prefilledSub) {
    if (!isEditor()) return;
    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-lg">
          <div class="modal-header">
            <h2 class="modal-title">New Part</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label for="newPartSub">Subassembly</label>
              <input type="text" id="newPartSub" class="form-input" placeholder="Upright / Hub / Gearbox" value="${prefilledSub || "Upright"}">
            </div>
            <div class="form-group">
              <label for="newPartName">Part Name</label>
              <input type="text" id="newPartName" class="form-input" placeholder="e.g. UP_05">
            </div>
            <div class="form-group">
              <label for="newPartVer">Version</label>
              <input type="text" id="newPartVer" class="form-input" value="1.0">
            </div>
          </div>
          <div class="form-group">
            <label for="newPartWeight">Weight (g)</label>
            <input type="number" step="0.1" id="newPartWeight" class="form-input" value="0.0">
          </div>
          <div class="form-group">
            <label for="newPartDesc">Description</label>
            <input type="text" id="newPartDesc" class="form-input" placeholder="Part description">
          </div>
          <div class="form-group">
            <label for="newPartSpecs">Specifications</label>
            <textarea id="newPartSpecs" class="form-textarea" placeholder="Feature: Value"></textarea>
          </div>
          <div style="display: flex; gap: 0.5rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
            <button class="btn-primary" id="btnSaveNewPart">Save Part</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveNewPart").onclick = async () => {
      const sub = document.getElementById("newPartSub").value.trim();
      const pName = document.getElementById("newPartName").value.trim();
      const ver = document.getElementById("newPartVer").value.trim() || "1.0";
      const w = parseFloat(document.getElementById("newPartWeight").value) || 0;
      const desc = document.getElementById("newPartDesc").value.trim();
      const specs = document.getElementById("newPartSpecs").value.trim();

      if (!pName) {
        alert("Part Name is required.");
        return;
      }

      if (partsData.some((p) => p["Part Name"] === pName)) {
        alert("Part " + pName + " already exists. Open that part to create a new iteration.");
        return;
      }

      const now = new Date().toISOString().replace("T", " ").slice(0, 19);
      const newPartRow = {
        "Subassembly": sub || "Drivetrain",
        "Part Name": pName,
        "Version": ver,
        "Date": now,
        "Weight (g)": w,
        "Description": desc,
        "Specifications": specs,
        "Image Path": null,
        "Stiffness Data": null,
        "LC Images": null,
        "Tolerancing Data": null,
        "Justification": null
      };

      partsData.push(newPartRow);
      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Create part ${pName}`);
        window.closeModal();
        currentView = pName;
        selectedEvalVersion = ver;
        buildSidebar();
        renderPartView(pName);
      } catch (e) {
        alert("Error creating part: " + e.message);
      }
    };
  };

  // Modal: Edit Tolerances (Updates all iterations of part)
  window.openTolerancesModal = function (partName) {
    if (!isEditor()) return;
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    const firstWithTol = iters.find((p) => p["Tolerancing Data"]) || iters[0];

    let tolRecords = [];
    try {
      if (typeof firstWithTol["Tolerancing Data"] === "string") {
        tolRecords = JSON.parse(firstWithTol["Tolerancing Data"]);
      } else if (Array.isArray(firstWithTol["Tolerancing Data"])) {
        tolRecords = firstWithTol["Tolerancing Data"];
      }
    } catch (e) {}

    renderTolModalContent(partName, tolRecords);
  };

  function renderTolModalContent(partName, tolRecords) {
    let rowsHtml = tolRecords.map((r, idx) => `
      <tr>
        <td><input type="text" class="form-input" value="${r["Interface / Feature"] || ""}" id="tolFeat_${idx}"></td>
        <td><input type="text" class="form-input" value="${r["Nominal (mm)"] || ""}" id="tolNom_${idx}"></td>
        <td><input type="text" class="form-input" value="${r["Fit / Tolerance"] || ""}" id="tolFit_${idx}"></td>
        <td><input type="text" class="form-input" value="${r["Roughness / Notes"] || ""}" id="tolNotes_${idx}"></td>
        <td><button class="btn-action danger" onclick="window.removeTolRow('${partName}', ${idx})">Remove</button></td>
      </tr>
    `).join("");

    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-xl">
          <div class="modal-header">
            <h2 class="modal-title">Edit Tolerances</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1rem;">
            Manufacturing GD&T specifications for ${partName}. Changes sync across all iterations.
          </p>
          <div class="card-table-wrapper" style="margin-bottom: 1rem;">
            <table class="styled-table">
              <thead>
                <tr>
                  <th>Interface</th>
                  <th>Nominal (mm)</th>
                  <th>Fit / Tolerance</th>
                  <th>Roughness / Notes</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody id="tolTableBody">
                ${rowsHtml || `<tr><td colspan="5" style="text-align: center; color: var(--text-dim);">No tolerances defined. Click Add Row below.</td></tr>`}
              </tbody>
            </table>
          </div>
          <div style="display: flex; gap: 0.5rem; justify-content: space-between; align-items: center;">
            <button class="btn-action" onclick="window.addTolRow('${partName}')">Add Row</button>
            <div style="display: flex; gap: 0.5rem;">
              <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
              <button class="btn-primary" id="btnSaveTols">Save Tolerances</button>
            </div>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveTols").onclick = async () => {
      const currentModalTols = getCurrentModalTols();
      const updatedTols = currentModalTols.filter(
        (r) =>
          (r["Interface / Feature"] || "").trim() ||
          (r["Nominal (mm)"] || "").trim() ||
          (r["Fit / Tolerance"] || "").trim()
      );

      const tolJson = updatedTols.length ? JSON.stringify(updatedTols) : null;
      partsData.forEach((p) => {
        if (p["Part Name"] === partName) {
          p["Tolerancing Data"] = tolJson;
        }
      });

      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Update GD&T for ${partName}`);
        window.closeModal();
        renderPartView(partName);
      } catch (e) {
        alert("Error saving tolerances: " + e.message);
      }
    };
  }

  window.addTolRow = function (partName) {
    const currentTols = getCurrentModalTols();
    currentTols.push({ "Interface / Feature": "", "Nominal (mm)": "", "Fit / Tolerance": "", "Roughness / Notes": "" });
    renderTolModalContent(partName, currentTols);
  };

  window.removeTolRow = function (partName, idx) {
    const currentTols = getCurrentModalTols();
    currentTols.splice(idx, 1);
    renderTolModalContent(partName, currentTols);
  };

  function getCurrentModalTols() {
    const tbody = document.getElementById("tolTableBody");
    if (!tbody) return [];
    const rows = tbody.querySelectorAll("tr");
    const result = [];
    rows.forEach((r, i) => {
      const f = document.getElementById(`tolFeat_${i}`);
      const n = document.getElementById(`tolNom_${i}`);
      const fit = document.getElementById(`tolFit_${i}`);
      const notes = document.getElementById(`tolNotes_${i}`);
      if (f && n && fit) {
        result.push({
          "Interface / Feature": f.value,
          "Nominal (mm)": n.value,
          "Fit / Tolerance": fit.value,
          "Roughness / Notes": notes ? notes.value : ""
        });
      }
    });
    return result;
  }

  // Modal: Edit Justification (Updates all iterations of part)
  window.openJustificationModal = function (partName) {
    if (!isEditor()) return;
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    const firstWithJust = iters.find((p) => p.Justification) || iters[0];
    const currJust = firstWithJust.Justification || "";

    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-lg">
          <div class="modal-header">
            <h2 class="modal-title">Edit Justification</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1rem;">
            Design rationale and technical justification for ${partName}. Changes sync across all iterations.
          </p>
          <div class="form-group">
            <label for="inputJustText">Justification</label>
            <textarea id="inputJustText" class="form-textarea" style="min-height: 180px;">${currJust}</textarea>
          </div>
          <div style="display: flex; gap: 0.5rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
            <button class="btn-primary" id="btnSaveJust">Save Justification</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveJust").onclick = async () => {
      const newJust = document.getElementById("inputJustText").value.trim() || null;
      partsData.forEach((p) => {
        if (p["Part Name"] === partName) {
          p.Justification = newJust;
        }
      });

      try {
        await window.DTCloud.commitFile("data/parts.json", partsData, `Update justification for ${partName}`);
        window.closeModal();
        renderPartView(partName);
      } catch (e) {
        alert("Error saving justification: " + e.message);
      }
    };
  };

  // -------------------------------------------------------------
  // Visual Master-Detail Assembly Builder Modal
  // -------------------------------------------------------------
  let activeAssemblyIdx = 0;

  function parseAssemblyConfig(configStr) {
    const config = {};
    if (!configStr) return config;
    const lines = String(configStr).split("\n");
    for (const line of lines) {
      const idx = line.indexOf(":");
      if (idx !== -1) {
        const k = line.substring(0, idx).trim();
        let v = line.substring(idx + 1).trim();
        if (v.startsWith("v") || v.startsWith("V")) {
          v = v.substring(1).trim();
        }
        config[k] = v;
      }
    }
    return config;
  }

  function normVer(v) {
    let s = String(v || "").trim();
    if (s.startsWith("v") || s.startsWith("V")) s = s.substring(1).trim();
    return s;
  }

  window.openManageAssembliesModal = function () {
    if (!isEditor()) return;
    activeAssemblyIdx = 0;
    renderManageAssembliesModal();
  };

  window.selectAssemblyIndex = function (idx) {
    activeAssemblyIdx = idx;
    renderManageAssembliesModal();
  };

  window.createNewAssembly = function () {
    const nowStr = new Date().toISOString().replace("T", " ").substring(0, 19);
    assembliesData.push({
      "Assembly Version": `New Assembly ${assembliesData.length + 1}`,
      "Date": nowStr,
      "Total Weight (g)": 0,
      "Description": "New wheel assembly configuration",
      "Part Configuration": ""
    });
    activeAssemblyIdx = assembliesData.length - 1;
    renderManageAssembliesModal();
  };

  window.duplicateAssembly = function () {
    const curr = assembliesData[activeAssemblyIdx];
    if (!curr) return;
    const nowStr = new Date().toISOString().replace("T", " ").substring(0, 19);
    assembliesData.push({
      "Assembly Version": `${curr["Assembly Version"]} (Copy)`,
      "Date": nowStr,
      "Total Weight (g)": parseFloat(curr["Total Weight (g)"] || 0),
      "Description": curr.Description || "",
      "Part Configuration": curr["Part Configuration"] || ""
    });
    activeAssemblyIdx = assembliesData.length - 1;
    renderManageAssembliesModal();
  };

  window.deleteActiveAssembly = async function () {
    const curr = assembliesData[activeAssemblyIdx];
    if (!curr) return;
    if (!confirm(`Delete assembly "${curr["Assembly Version"]}"?`)) return;
    assembliesData.splice(activeAssemblyIdx, 1);
    activeAssemblyIdx = Math.max(0, activeAssemblyIdx - 1);
    try {
      await window.DTCloud.commitFile("data/assemblies.json", assembliesData, `Delete assembly ${curr["Assembly Version"]}`);
      showToast("Assembly deleted!");
      if (assembliesData.length === 0) {
        window.closeModal();
      } else {
        renderManageAssembliesModal();
      }
      renderAssemblyView();
    } catch (e) {
      alert("Error deleting assembly: " + e.message);
    }
  };

  window.updateLiveAssemblyMass = function () {
    const liveWeightEl = document.getElementById("liveAssWeight");
    const liveDeltaEl = document.getElementById("liveAssDelta");
    if (!liveWeightEl) return;

    let sum = 0;
    document.querySelectorAll(".ass-part-select").forEach((sel) => {
      const selectedOpt = sel.options[sel.selectedIndex];
      if (selectedOpt && selectedOpt.dataset.weight) {
        sum += parseFloat(selectedOpt.dataset.weight) || 0;
      }
    });

    liveWeightEl.textContent = `${sum.toFixed(1)} g (${(sum / 1000).toFixed(3)} kg)`;
    if (liveDeltaEl) {
      const fst15Base = 2220.2;
      const diff = sum - fst15Base;
      const pct = (diff / fst15Base) * 100;
      const isGood = diff <= 0;
      liveDeltaEl.className = isGood ? "delta-good" : "delta-bad";
      liveDeltaEl.textContent = `${isGood ? "" : "+"}${diff.toFixed(1)} g (${pct.toFixed(1)}% vs FST15)`;
    }
  };

  window.syncLiveMassToInput = function () {
    let sum = 0;
    document.querySelectorAll(".ass-part-select").forEach((sel) => {
      const selectedOpt = sel.options[sel.selectedIndex];
      if (selectedOpt && selectedOpt.dataset.weight) {
        sum += parseFloat(selectedOpt.dataset.weight) || 0;
      }
    });
    const inp = document.getElementById("inpAssTotalWeight");
    if (inp) {
      inp.value = sum.toFixed(1);
    }
  };

  window.saveActiveAssembly = async function () {
    const curr = assembliesData[activeAssemblyIdx];
    if (!curr) return;

    const verInp = document.getElementById("inpAssVersion");
    const descInp = document.getElementById("inpAssDescription");
    const weightInp = document.getElementById("inpAssTotalWeight");

    const newVer = verInp ? verInp.value.trim() : curr["Assembly Version"];
    const newDesc = descInp ? descInp.value.trim() : (curr.Description || "");
    const newWeight = weightInp ? (parseFloat(weightInp.value) || 0) : (parseFloat(curr["Total Weight (g)"]) || 0);

    const configLines = [];
    document.querySelectorAll(".ass-part-select").forEach((sel) => {
      const partKey = sel.dataset.partKey;
      const val = sel.value;
      if (val && val !== "-- Not Included --") {
        configLines.push(`${partKey}: v${val}`);
      }
    });

    curr["Assembly Version"] = newVer;
    curr["Description"] = newDesc;
    curr["Total Weight (g)"] = newWeight;
    curr["Part Configuration"] = configLines.join("\n");
    if (!curr.Date) {
      curr.Date = new Date().toISOString().replace("T", " ").substring(0, 19);
    }

    try {
      await window.DTCloud.commitFile("data/assemblies.json", assembliesData, `Update assembly ${newVer}`);
      showToast("Assembly saved!");
      renderManageAssembliesModal();
      renderAssemblyView();
    } catch (e) {
      alert("Error saving assembly: " + e.message);
    }
  };

  function renderManageAssembliesModal() {
    if (!assembliesData || !assembliesData.length) {
      window.createNewAssembly();
      return;
    }

    if (activeAssemblyIdx >= assembliesData.length) {
      activeAssemblyIdx = 0;
    }

    const curr = assembliesData[activeAssemblyIdx];
    const currConfig = parseAssemblyConfig(curr["Part Configuration"]);

    const renderCustomPartPicker = (partKey, label) => {
      const partIters = partsData.filter((p) => p["Part Name"] === partKey);
      const chosenVer = currConfig[partKey];
      return `
        <div class="assembly-part-row">
          <div>
            <strong style="font-size: 0.88rem; color: var(--text-main);">${partKey}</strong>
            <span style="font-size: 0.78rem; color: var(--text-muted); margin-left: 0.4rem;">${label}</span>
          </div>
          <select class="form-input ass-part-select" data-part-key="${partKey}" style="width: 250px; font-size: 0.82rem;" onchange="window.updateLiveAssemblyMass()">
            <option value="" data-weight="0">-- Not Included --</option>
            ${partIters.map((it) => {
              const itNorm = normVer(it.Version);
              const isSel = chosenVer && normVer(chosenVer) === itNorm;
              const w = parseFloat(it["Weight (g)"] || 0);
              return `<option value="${it.Version}" data-weight="${w}" ${isSel ? "selected" : ""}>v${it.Version} (${w.toFixed(1)} g)</option>`;
            }).join("")}
          </select>
        </div>
      `;
    };

    const renderOemPicker = (partKey, label, filterFn) => {
      const oemParts = partsData.filter((p) => (p.Subassembly || "") === "Other" && filterFn(p["Part Name"]));
      const chosenVer = currConfig[partKey];
      return `
        <div class="assembly-part-row">
          <div>
            <strong style="font-size: 0.88rem; color: var(--text-main);">${partKey}</strong>
            <span style="font-size: 0.78rem; color: var(--text-muted); margin-left: 0.4rem;">${label}</span>
          </div>
          <select class="form-input ass-part-select" data-part-key="${partKey}" style="width: 250px; font-size: 0.82rem;" onchange="window.updateLiveAssemblyMass()">
            <option value="" data-weight="0">-- Not Included --</option>
            ${oemParts.map((p) => {
              const isSel = chosenVer && normVer(chosenVer) === normVer(p["Part Name"]);
              const w = parseFloat(p["Weight (g)"] || 0);
              return `<option value="${p["Part Name"]}" data-weight="${w}" ${isSel ? "selected" : ""}>${p["Part Name"]} (${w.toFixed(1)} g)</option>`;
            }).join("")}
          </select>
        </div>
      `;
    };

    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-xl" style="max-width: 1040px; height: 86vh; display: flex; flex-direction: column;">
          
          <!-- Modal Header -->
          <div class="modal-header" style="flex-shrink: 0; margin-bottom: 0.75rem;">
            <div>
              <h2 class="modal-title">Manage Assemblies</h2>
              <p style="color: var(--text-muted); font-size: 0.82rem; margin-top: 2px;">
                Select component versions for each wheel assembly.
              </p>
            </div>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>

          <!-- Two-Pane Layout -->
          <div class="assembly-builder-container" style="flex: 1; min-height: 0; overflow: hidden;">
            
            <!-- Left Pane: Assemblies List -->
            <div class="assembly-nav-pane">
              <button class="btn-action primary" style="width: 100%; margin-bottom: 0.75rem;" onclick="window.createNewAssembly()">+ New Assembly</button>
              <div style="flex: 1; overflow-y: auto;">
                ${assembliesData.map((a, i) => `
                  <div class="assembly-nav-card ${i === activeAssemblyIdx ? "active" : ""}" onclick="window.selectAssemblyIndex(${i})">
                    <div style="font-weight: 700; font-size: 0.92rem; color: var(--text-main); margin-bottom: 3px;">
                      ${a["Assembly Version"] || "Untitled"}
                    </div>
                    <div style="font-size: 0.78rem; color: var(--text-muted); display: flex; justify-content: space-between;">
                      <span style="font-weight: 600; color: var(--text-main);">${parseFloat(a["Total Weight (g)"] || 0).toFixed(1)} g</span>
                      <span style="color: var(--text-dim);">${(a.Date || "").substring(0, 10)}</span>
                    </div>
                  </div>
                `).join("")}
              </div>
            </div>

            <!-- Right Pane: Active Assembly Details -->
            <div class="assembly-editor-pane">
              
              <!-- General Info Card -->
              <div class="assembly-section-card">
                <div class="form-row" style="margin-bottom: 0.75rem;">
                  <div class="form-group" style="flex: 2;">
                    <label for="inpAssVersion">Assembly Version ID</label>
                    <input type="text" id="inpAssVersion" class="form-input" style="font-weight: 700;" value="${curr["Assembly Version"] || ""}">
                  </div>
                  <div class="form-group" style="flex: 1;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.25rem;">
                      <label for="inpAssTotalWeight" style="margin: 0;">Total Mass (g)</label>
                      <button type="button" class="btn-action" style="font-size: 0.68rem; padding: 1px 6px;" onclick="window.syncLiveMassToInput()">Auto Sum</button>
                    </div>
                    <input type="number" step="0.1" id="inpAssTotalWeight" class="form-input" value="${parseFloat(curr["Total Weight (g)"] || 0).toFixed(1)}">
                  </div>
                </div>

                <div class="form-group" style="margin-bottom: 0.75rem;">
                  <label for="inpAssDescription">Description</label>
                  <input type="text" id="inpAssDescription" class="form-input" placeholder="Technical purpose and differences of this assembly" value="${curr.Description || ""}">
                </div>

                <!-- Live Mass Status Banner -->
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(0, 0, 0, 0.25); border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 0.5rem 0.85rem;">
                  <div>
                    <span style="font-size: 0.76rem; text-transform: uppercase; color: var(--text-muted); font-weight: 600;">Selected Components Sum:</span>
                    <strong id="liveAssWeight" style="margin-left: 0.5rem; font-size: 0.95rem; color: var(--text-main);">0.0 g</strong>
                  </div>
                  <div>
                    <span id="liveAssDelta" style="font-size: 0.82rem; font-weight: 600;"></span>
                  </div>
                </div>
              </div>

              <!-- Component Pickers: Upright (UP) -->
              <div class="assembly-section-card">
                <div style="font-size: 0.76rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: var(--accent-red); margin-bottom: 0.6rem;">
                  Upright Assembly (UP)
                </div>
                ${renderCustomPartPicker("UP_01", "Main Upright Body")}
                ${renderCustomPartPicker("UP_02", "Upright Lid")}
                ${renderCustomPartPicker("UP_03", "Internal Front Plate")}
                ${renderCustomPartPicker("UP_04", "Backplate")}
              </div>

              <!-- Component Pickers: Hub (HB) -->
              <div class="assembly-section-card">
                <div style="font-size: 0.76rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: var(--accent-red); margin-bottom: 0.6rem;">
                  Hub Assembly (HB)
                </div>
                ${renderCustomPartPicker("HB_01", "Wheel Hub")}
                ${renderCustomPartPicker("HB_02", "Hub Lid")}
                ${renderCustomPartPicker("HB_07", "Outer Ring Spacer")}
              </div>

              <!-- Component Pickers: Gearbox (GB) -->
              <div class="assembly-section-card">
                <div style="font-size: 0.76rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: var(--accent-red); margin-bottom: 0.6rem;">
                  Transmission (GB)
                </div>
                ${renderCustomPartPicker("TD", "Transmission Drivetrain")}
              </div>

              <!-- Component Pickers: OEM Bearings & Seals -->
              <div class="assembly-section-card">
                <div style="font-size: 0.76rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: var(--accent-red); margin-bottom: 0.6rem;">
                  OEM Bearings & Seals
                </div>
                ${renderOemPicker("Inner Bearing", "Motor-side Bearing", (name) => name.includes("618") || name.includes("718"))}
                ${renderOemPicker("Outer Bearing", "Rim-side Bearing", (name) => name.includes("618") || name.includes("718"))}
                ${renderOemPicker("Seal", "Transmission Oil Seal", (name) => name.includes("HMSA") || name.includes("X"))}
              </div>

              <!-- Actions Toolbar -->
              <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 0.5rem; margin-bottom: 1rem; flex-wrap: wrap; gap: 0.5rem;">
                <div style="display: flex; gap: 0.5rem;">
                  <button type="button" class="btn-action" onclick="window.duplicateAssembly()">Duplicate</button>
                  <button type="button" class="btn-action danger" onclick="window.deleteActiveAssembly()">Delete Assembly</button>
                </div>
                <div style="display: flex; gap: 0.5rem;">
                  <button type="button" class="btn-secondary" onclick="window.closeModal()">Close</button>
                  <button type="button" class="btn-primary" onclick="window.saveActiveAssembly()">Save Assembly</button>
                </div>
              </div>

            </div>
          </div>
        </div>
      </div>
    `;

    window.updateLiveAssemblyMass();
  }

  // Global Navigation Shortcut
  window.navigateToPart = function (partName) {
    if (partName !== currentView) {
      currentPartName = null;
      selectedEvalVersion = null;
      selectedRefVersion = null;
    }
    currentView = partName;
    activeGalleryIndex = 0;
    buildSidebar();
    renderCurrentView();
  };

  // Export functions to window
  window.exportExcelWorkbook = exportExcelWorkbook;

  // Cloud settings trigger in footer
  if (elBtnOpenSync) {
    elBtnOpenSync.onclick = openCloudSettingsModal;
  }

  // Launch App
  init();
})();
