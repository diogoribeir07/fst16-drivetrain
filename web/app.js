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
  let selectedEvalVersion = null;
  let selectedRefVersion = "FST15";
  let activeGalleryIndex = 0;
  let reqActiveTab = "list"; // 'list' or 'diagrams'
  let activeFeaLc = "LC1";

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

    const subOrder = ["Upright", "Hub", "Gearbox", "OEM Components"];
    const allSubs = Array.from(new Set([...subOrder, ...Object.keys(partsBySub)]));

    allSubs.forEach((sub) => {
      const parts = Array.from(partsBySub[sub] || []).sort();
      if (!parts.length) return;

      html += `
        <div class="sidebar-section">
          <div class="section-label">
            <span>${sub}</span>
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
        currentView = item.getAttribute("data-view");
        selectedEvalVersion = null;
        selectedRefVersion = "FST15";
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
      renderAssemblyView();
    } else if (currentView === "requirements") {
      renderRequirementsView();
    } else if (currentView === "trash") {
      renderTrashView();
    } else {
      renderPartView(currentView);
    }
  }

  // 1. Wheel Assembly View
  function renderAssemblyView() {
    const partLatestMap = new Map();
    partsData.forEach((p) => {
      partLatestMap.set(p["Part Name"], p);
    });

    let totalWeight = 0;
    partLatestMap.forEach((p) => {
      const w = parseFloat(p["Weight (g)"]) || 0;
      totalWeight += w;
    });

    const targetWeight = 4800; // 4.8 kg target
    const fst15Baseline = 5040; // 5.04 kg baseline
    const diffVsFst15 = totalWeight - fst15Baseline;
    const diffVsTarget = totalWeight - targetWeight;

    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">Wheel Assembly</h1>
          <p class="page-subtitle">Global mass tracking and subsystem breakdown</p>
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
          <div class="metric-label">Unsprung Mass</div>
          <div class="metric-value">${(totalWeight / 1000).toFixed(3)} kg</div>
          <div class="metric-sub delta-neutral">${totalWeight.toFixed(1)} g total</div>
        </div>
        <div class="metric-card">
          <div class="metric-label">FST15 Baseline</div>
          <div class="metric-value ${diffVsFst15 <= 0 ? "delta-good" : "delta-bad"}">
            ${diffVsFst15 <= 0 ? "" : "+"}${(diffVsFst15 / 1000).toFixed(3)} kg
          </div>
          <div class="metric-sub ${diffVsFst15 <= 0 ? "delta-good" : "delta-bad"}">
            ${((diffVsFst15 / fst15Baseline) * 100).toFixed(1)}% vs 5.040 kg
          </div>
        </div>
        <div class="metric-card">
          <div class="metric-label">Target Margin</div>
          <div class="metric-value ${diffVsTarget <= 0 ? "delta-good" : "delta-bad"}">
            ${diffVsTarget <= 0 ? "" : "+"}${(diffVsTarget / 1000).toFixed(3)} kg
          </div>
          <div class="metric-sub ${diffVsTarget <= 0 ? "delta-good" : "delta-bad"}">
            Target: ${(targetWeight / 1000).toFixed(3)} kg
          </div>
        </div>
      </div>

      <h2 style="font-size: 1.1rem; margin-bottom: 0.75rem;">Component Mass</h2>
      <div class="card-table-wrapper">
        <table class="styled-table">
          <thead>
            <tr>
              <th>Subassembly</th>
              <th>Part Name</th>
              <th>Version</th>
              <th>Weight (g)</th>
              <th>Description</th>
            </tr>
          </thead>
          <tbody>
    `;

    Array.from(partLatestMap.values())
      .sort((a, b) => (a.Subassembly || "").localeCompare(b.Subassembly || "") || a["Part Name"].localeCompare(b["Part Name"]))
      .forEach((p) => {
        const w = parseFloat(p["Weight (g)"]) || 0;
        html += `
          <tr style="cursor: pointer;" onclick="window.navigateToPart('${p["Part Name"]}')">
            <td><strong>${p.Subassembly || "—"}</strong></td>
            <td><span style="color: var(--accent-red); font-weight: 600;">${p["Part Name"]}</span></td>
            <td>${formatVersion(p.Version)}</td>
            <td class="mono"><strong>${w.toFixed(1)} g</strong></td>
            <td style="color: var(--text-muted); font-size: 0.82rem;">${(p.Description || "").slice(0, 60)}...</td>
          </tr>
        `;
      });

    html += `
          </tbody>
        </table>
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

  // 4. Component / Part Detail View
  function renderPartView(partName) {
    const iters = partsData.filter((p) => p["Part Name"] === partName);
    if (!iters.length) {
      elContentArea.innerHTML = `<p style="padding: 2rem;">No data found for ${partName}</p>`;
      return;
    }

    if (!selectedEvalVersion) {
      selectedEvalVersion = iters[iters.length - 1].Version;
    }

    const evalRow = iters.find((p) => String(p.Version) === String(selectedEvalVersion)) || iters[iters.length - 1];
    const refRow = iters.find((p) => String(p.Version) === String(selectedRefVersion)) || iters[0];

    const evalWeight = parseFloat(evalRow["Weight (g)"]) || 0;
    const refWeight = parseFloat(refRow["Weight (g)"]) || 0;
    const weightDiff = evalWeight - refWeight;

    let html = `
      <div class="page-header">
        <div>
          <h1 class="page-title">${partName}</h1>
          <p class="page-subtitle">${evalRow.Subassembly || "Drivetrain"} &bull; Latest: ${formatVersion(iters[iters.length - 1].Version)}</p>
        </div>
        <div class="header-actions">
          ${isEditor() ? `
          <button class="btn-action primary" onclick="window.openNewIterationModal('${partName}')">New Iteration</button>
          <button class="btn-action" onclick="window.openEditIterationModal('${partName}', '${evalRow.Version}')">Edit Iteration</button>
          <button class="btn-action danger" onclick="window.deleteIteration('${partName}', '${evalRow.Version}')">Delete Iteration</button>
          <button class="btn-action danger" onclick="window.deletePart('${partName}')">Delete Part</button>
          ` : ""}
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
          <div class="metric-value ${weightDiff <= 0 ? "delta-good" : "delta-bad"}">
            ${weightDiff <= 0 ? "" : "+"}${weightDiff.toFixed(1)} g
          </div>
          <div class="metric-sub ${weightDiff <= 0 ? "delta-good" : "delta-bad"}">
            ${refWeight > 0 ? ((weightDiff / refWeight) * 100).toFixed(1) + "%" : "0.0%"}
          </div>
        </div>
      </div>

      <div class="tabs-header">
        <button class="tab-btn active" data-tab="tabOverview">Overview</button>
        <button class="tab-btn" data-tab="tabSpecs">Specifications</button>
        <button class="tab-btn" data-tab="tabFEA">FEA Stiffness</button>
        <button class="tab-btn" data-tab="tabImage">Schematic</button>
      </div>

      <!-- Tab: Overview -->
      <div class="tab-pane active" id="tabOverview">
        ${iters.length > 1 ? `
        <div style="display: flex; gap: 1rem; align-items: center; margin-bottom: 1.5rem; flex-wrap: wrap;">
          <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
            Compare:
            <select id="selectEvalVer" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
              ${iters.map((it) => `<option value="${it.Version}" ${String(it.Version) === String(evalRow.Version) ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
            </select>
          </label>
          <label style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
            Against Ref:
            <select id="selectRefVer" class="form-input" style="width: auto; display: inline-block; margin-left: 0.5rem;">
              ${iters.map((it) => `<option value="${it.Version}" ${String(it.Version) === String(refRow.Version) ? "selected" : ""}>${formatVersion(it.Version)} (${parseFloat(it["Weight (g)"] || 0).toFixed(1)} g)</option>`).join("")}
            </select>
          </label>
        </div>
        ` : `
        <div style="margin-bottom: 1.5rem; padding: 0.6rem 0.9rem; background: rgba(255, 75, 75, 0.08); border-left: 3px solid var(--accent-red); border-radius: 4px; font-size: 0.85rem; color: var(--text-muted);">
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
                <td class="mono ${weightDiff <= 0 ? "delta-good" : "delta-bad"}">
                  <strong>${weightDiff <= 0 ? "" : "+"}${weightDiff.toFixed(1)} g</strong>
                </td>
                <td class="mono ${weightDiff <= 0 ? "delta-good" : "delta-bad"}">
                  <strong>${refWeight > 0 ? ((weightDiff / refWeight) * 100).toFixed(1) + "%" : "0.0%"}</strong>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        ${renderFEATable(evalRow, refRow)}
      </div>

      <!-- Tab: Specifications -->
      <div class="tab-pane" id="tabSpecs">
        <h3 style="font-size: 1rem; margin-bottom: 0.5rem;">Specifications</h3>
        <div class="card-table-wrapper">
          <table class="styled-table">
            <thead>
              <tr>
                <th>Property</th>
                <th>Value</th>
              </tr>
            </thead>
            <tbody>
              ${renderSpecsRows(evalRow.Specifications)}
            </tbody>
          </table>
        </div>

        <div style="display: flex; align-items: center; justify-content: space-between; margin: 1.5rem 0 0.5rem;">
          <h3 style="font-size: 1rem;">Manufacturing GD&T</h3>
          ${isEditor() ? `<button class="btn-action" onclick="window.openTolerancesModal('${partName}')">Edit Tolerances</button>` : ""}
        </div>
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

        <div style="display: flex; align-items: center; justify-content: space-between; margin: 1.5rem 0 0.5rem;">
          <h3 style="font-size: 1rem;">Justification</h3>
          ${isEditor() ? `<button class="btn-action" onclick="window.openJustificationModal('${partName}')">Edit Justification</button>` : ""}
        </div>
        <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 1.25rem; font-size: 0.9rem; line-height: 1.6; color: var(--text-main);">
          ${evalRow.Justification && String(evalRow.Justification).trim() ? String(evalRow.Justification).replace(/\n/g, "<br>") : "<span style='color: var(--text-dim);'>No justification recorded.</span>"}
        </div>
      </div>

      <!-- Tab: FEA Stiffness -->
      <div class="tab-pane" id="tabFEA">
        ${renderFEATable(evalRow, refRow, true)}
      </div>

      <!-- Tab: Schematic Gallery -->
      <div class="tab-pane" id="tabImage">
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

    // Attach Tab Navigation
    elContentArea.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        elContentArea.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
        elContentArea.querySelectorAll(".tab-pane").forEach((p) => p.classList.remove("active"));
        btn.classList.add("active");
        const targetId = btn.getAttribute("data-tab");
        const targetPane = document.getElementById(targetId);
        if (targetPane) targetPane.classList.add("active");
      });
    });

    // Attach dropzone & paste events for gallery in Editor mode
    attachGalleryEventListeners(evalRow, partName);
    attachFeaListeners(evalRow, partName);
  }

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
      const isGood = diff <= 0;

      rowsHtml += `
        <tr>
          <td><strong>${lc}</strong></td>
          <td style="color: var(--text-muted); font-size: 0.82rem;">${descMap[lc] || lc}</td>
          <td class="mono">${eVal.toFixed(decimals)} mm</td>
          <td class="mono">${rVal.toFixed(decimals)} mm</td>
          <td class="mono ${isGood ? "delta-good" : "delta-bad"}">
            <strong>${diff <= 0 ? "" : "+"}${diff.toFixed(decimals)} mm</strong>
          </td>
          <td class="mono ${isGood ? "delta-good" : "delta-bad"}">
            <strong>${diff <= 0 ? "" : "+"}${pct.toFixed(1)}%</strong>
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
      <div style="margin-top: 1.5rem;">
        <h3 style="font-size: 1rem; margin-bottom: 0.75rem;">FEA Screenshots</h3>
        <div class="tabs-header" style="margin-bottom: 1rem;">
          ${availLcs.map(lc => `
            <button class="tab-btn ${lc === activeFeaLc ? "active" : ""}" onclick="window.selectFeaLc('${lc}')">${lc}</button>
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

  // Modal: Manage Assemblies
  window.openManageAssembliesModal = function () {
    if (!isEditor()) return;
    renderAssembliesModalContent(assembliesData);
  };

  function renderAssembliesModalContent(assembliesList) {
    elModalContainer.innerHTML = `
      <div class="modal-overlay open">
        <div class="modal-box modal-xl">
          <div class="modal-header">
            <h2 class="modal-title">Manage Assemblies</h2>
            <button class="modal-close-btn" onclick="window.closeModal()">x</button>
          </div>
          <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1rem;">
            Subassembly mass baselines, targets, and included parts configurations.
          </p>
          <div class="card-table-wrapper" style="margin-bottom: 1.5rem;">
            <table class="styled-table">
              <thead>
                <tr>
                  <th>Assembly Name</th>
                  <th>Baseline (g)</th>
                  <th>Target (g)</th>
                  <th>Included Parts</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody id="assTableBody">
                ${assembliesList.map((a, i) => `
                  <tr>
                    <td><input type="text" class="form-input" value="${a["Assembly Name"] || a.Assembly || ""}" id="assName_${i}"></td>
                    <td><input type="number" step="0.1" class="form-input" value="${parseFloat(a["FST15 Baseline (g)"] || a.Baseline || 0)}" id="assBase_${i}"></td>
                    <td><input type="number" step="0.1" class="form-input" value="${parseFloat(a["Target Weight (g)"] || a.Target || 0)}" id="assTarget_${i}"></td>
                    <td><input type="text" class="form-input" value="${a["Included Parts"] || a.Parts || ""}" id="assParts_${i}"></td>
                    <td><button class="btn-action danger" onclick="window.removeAssemblyRow(${i})">Remove</button></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
          <div style="display: flex; gap: 0.5rem; justify-content: space-between; align-items: center;">
            <button class="btn-action" onclick="window.addAssemblyRow()">Add Assembly</button>
            <div style="display: flex; gap: 0.5rem;">
              <button class="btn-secondary" onclick="window.closeModal()">Cancel</button>
              <button class="btn-primary" id="btnSaveAss">Save Assemblies</button>
            </div>
          </div>
        </div>
      </div>
    `;

    document.getElementById("btnSaveAss").onclick = async () => {
      assembliesData = getCurrentModalAssemblies();
      try {
        await window.DTCloud.commitFile("data/assemblies.json", assembliesData, "Update assemblies configuration");
        window.closeModal();
        renderAssemblyView();
      } catch (e) {
        alert("Error saving assemblies: " + e.message);
      }
    };
  }

  function getCurrentModalAssemblies() {
    const tbody = document.getElementById("assTableBody");
    if (!tbody) return assembliesData;
    const rows = tbody.querySelectorAll("tr");
    const result = [];
    rows.forEach((r, i) => {
      const name = document.getElementById(`assName_${i}`);
      const base = document.getElementById(`assBase_${i}`);
      const target = document.getElementById(`assTarget_${i}`);
      const parts = document.getElementById(`assParts_${i}`);
      if (name) {
        result.push({
          "Assembly Name": name.value.trim(),
          "FST15 Baseline (g)": parseFloat(base?.value) || 0,
          "Target Weight (g)": parseFloat(target?.value) || 0,
          "Included Parts": parts ? parts.value.trim() : ""
        });
      }
    });
    return result;
  }

  window.addAssemblyRow = function () {
    const list = getCurrentModalAssemblies();
    list.push({ "Assembly Name": "New Assembly", "FST15 Baseline (g)": 0, "Target Weight (g)": 0, "Included Parts": "" });
    renderAssembliesModalContent(list);
  };

  window.removeAssemblyRow = function (idx) {
    const list = getCurrentModalAssemblies();
    list.splice(idx, 1);
    renderAssembliesModalContent(list);
  };

  // Global Navigation Shortcut
  window.navigateToPart = function (partName) {
    currentView = partName;
    selectedEvalVersion = null;
    selectedRefVersion = "FST15";
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
