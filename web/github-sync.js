/**
 * github-sync.js
 * Direct GitHub REST API client for FST 16 Drivetrain Tracker.
 * Allows 24/7 read/write synchronization with the GitHub repository.
 */
(function (root) {
  "use strict";

  const STORAGE_KEY = "fst16_cloud_config";
  const DEFAULT_REPO = "diogoribeir07/fst16-drivetrain";
  const DEFAULT_BRANCH = "main";

  function getConfig() {
    let token = "";
    let repo = DEFAULT_REPO;
    let branch = DEFAULT_BRANCH;
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === "object") {
          token = (parsed.token || "").trim();
          repo = (parsed.repo || DEFAULT_REPO).trim();
          branch = (parsed.branch || DEFAULT_BRANCH).trim();
        }
      }
    } catch (e) {
      console.warn("Could not read fst16_cloud_config", e);
    }
    return {
      token: token,
      repo: repo || DEFAULT_REPO,
      branch: branch || DEFAULT_BRANCH
    };
  }

  function saveConfig(token, repo, branch) {
    const current = getConfig();
    const cfg = {
      token: token !== undefined ? (token || "").trim() : current.token,
      repo: (repo || current.repo || DEFAULT_REPO).trim(),
      branch: (branch || current.branch || DEFAULT_BRANCH).trim()
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    return cfg;
  }

  function clearToken() {
    const cfg = getConfig();
    cfg.token = "";
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    return cfg;
  }

  function isConfigured() {
    const cfg = getConfig();
    return Boolean(cfg.token && cfg.repo);
  }

  function utf8ToBase64(str) {
    const bytes = new TextEncoder().encode(str);
    let binary = "";
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  function base64ToUtf8(b64) {
    const binary = atob(b64.replace(/\s/g, ""));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new TextDecoder().decode(bytes);
  }

  async function fetchFile(path) {
    const cfg = getConfig();
    const cleanPath = path.startsWith("/") ? path.slice(1) : path;

    // If configured with token, query GitHub API to get the absolute latest commit
    if (cfg.token && cfg.repo) {
      try {
        const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}?ref=${encodeURIComponent(cfg.branch)}&_cb=${Date.now()}`;
        const res = await fetch(url, {
          headers: {
            "Accept": "application/vnd.github.v3+json",
            "Authorization": `token ${cfg.token}`
          }
        });
        if (res.ok) {
          const json = await res.json();
          const content = base64ToUtf8(json.content);
          return { data: JSON.parse(content), sha: json.sha };
        }
      } catch (err) {
        console.warn(`GitHub API fetch failed for ${cleanPath}, falling back to static path:`, err);
      }
    }

    // Fallback: static fetch relative to web folder (e.g. ../data/parts.json)
    const fallbackUrl = `../${cleanPath}?v=${Date.now()}`;
    const staticRes = await fetch(fallbackUrl);
    if (!staticRes.ok) {
      throw new Error(`Failed to load ${fallbackUrl} (Status: ${staticRes.status})`);
    }
    const data = await staticRes.json();
    return { data, sha: null };
  }

  async function commitFile(path, dataObject, commitMessage, retryCount = 0) {
    const cfg = getConfig();
    if (!cfg.token || !cfg.repo) {
      throw new Error("GitHub Token not configured. Please enter access token in Cloud Settings.");
    }
    const cleanPath = path.startsWith("/") ? path.slice(1) : path;
    const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}`;

    // 1. Get current SHA if exists
    let sha = null;
    try {
      const getUrl = `${url}?ref=${encodeURIComponent(cfg.branch)}&_cb=${Date.now()}`;
      const headRes = await fetch(getUrl, {
        headers: {
          "Accept": "application/vnd.github.v3+json",
          "Authorization": `token ${cfg.token}`
        }
      });
      if (headRes.ok) {
        const headJson = await headRes.json();
        sha = headJson.sha;
      } else if (headRes.status === 404) {
        sha = null;
      } else if (headRes.status === 401) {
        throw new Error("Invalid or expired GitHub token. Please update your token in Cloud Settings.");
      } else if (headRes.status === 403) {
        const errJson = await headRes.json().catch(() => ({}));
        throw new Error(`GitHub access denied (403): ${errJson.message || "Token missing repository write permissions"}`);
      } else {
        const errJson = await headRes.json().catch(() => ({}));
        throw new Error(`Failed to verify ${cleanPath} status (${headRes.status}): ${errJson.message || headRes.statusText}`);
      }
    } catch (e) {
      if (e.message && (e.message.includes("GitHub") || e.message.includes("token") || e.message.includes("Failed to verify"))) {
        throw e;
      }
      console.warn(`Could not fetch sha for ${cleanPath}:`, e);
    }

    // 2. Put file
    const contentStr = JSON.stringify(dataObject, null, 2);
    const contentB64 = utf8ToBase64(contentStr);

    const bodyPayload = {
      message: commitMessage || `Update ${cleanPath} via web portal`,
      content: contentB64,
      branch: cfg.branch
    };
    if (sha) {
      bodyPayload.sha = sha;
    }

    const putRes = await fetch(url, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/vnd.github.v3+json",
        "Authorization": `token ${cfg.token}`
      },
      body: JSON.stringify(bodyPayload)
    });

    if (putRes.status === 409 && retryCount < 3) {
      console.warn(`409 Conflict encountered on ${cleanPath}. Retrying commit (attempt ${retryCount + 1})...`);
      await new Promise(r => setTimeout(r, 800));
      return commitFile(path, dataObject, commitMessage, retryCount + 1);
    }

    if (!putRes.ok) {
      const errJson = await putRes.json().catch(() => ({}));
      throw new Error(errJson.message || `GitHub commit failed (${putRes.status})`);
    }

    return await putRes.json();
  }

  async function commitBinaryFile(path, base64Content, commitMessage, retryCount = 0) {
    const cfg = getConfig();
    if (!cfg.token || !cfg.repo) {
      throw new Error("GitHub Token not configured. Please enter access token in Cloud Settings.");
    }
    const cleanPath = path.startsWith("/") ? path.slice(1) : path;
    const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}`;

    let sha = null;
    try {
      const getUrl = `${url}?ref=${encodeURIComponent(cfg.branch)}&_cb=${Date.now()}`;
      const headRes = await fetch(getUrl, {
        headers: {
          "Accept": "application/vnd.github.v3+json",
          "Authorization": `token ${cfg.token}`
        }
      });
      if (headRes.ok) {
        const headJson = await headRes.json();
        sha = headJson.sha;
      } else if (headRes.status === 404) {
        sha = null;
      } else if (headRes.status === 401) {
        throw new Error("Invalid or expired GitHub token. Please update your token in Cloud Settings.");
      } else if (headRes.status === 403) {
        const errJson = await headRes.json().catch(() => ({}));
        throw new Error(`GitHub access denied (403): ${errJson.message || "Token missing repository write permissions"}`);
      } else {
        const errJson = await headRes.json().catch(() => ({}));
        throw new Error(`Failed to verify ${cleanPath} status (${headRes.status}): ${errJson.message || headRes.statusText}`);
      }
    } catch (e) {
      if (e.message && (e.message.includes("GitHub") || e.message.includes("token") || e.message.includes("Failed to verify"))) {
        throw e;
      }
      console.warn(`Could not fetch sha for ${cleanPath}:`, e);
    }

    const cleanB64 = base64Content.includes(",") ? base64Content.split(",")[1] : base64Content;

    const bodyPayload = {
      message: commitMessage || `Upload ${cleanPath} via web portal`,
      content: cleanB64.replace(/\s/g, ""),
      branch: cfg.branch
    };
    if (sha) {
      bodyPayload.sha = sha;
    }

    const putRes = await fetch(url, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/vnd.github.v3+json",
        "Authorization": `token ${cfg.token}`
      },
      body: JSON.stringify(bodyPayload)
    });

    if (putRes.status === 409 && retryCount < 3) {
      console.warn(`409 Conflict on binary ${cleanPath}. Retrying (attempt ${retryCount + 1})...`);
      await new Promise(r => setTimeout(r, 800));
      return commitBinaryFile(path, base64Content, commitMessage, retryCount + 1);
    }

    if (!putRes.ok) {
      const errJson = await putRes.json().catch(() => ({}));
      throw new Error(errJson.message || `Upload failed (${putRes.status})`);
    }

    return await putRes.json();
  }

  async function commitBatch(fileEntries, commitMessage, onProgress) {
    const cfg = getConfig();
    if (!cfg.token || !cfg.repo) {
      throw new Error("GitHub Token not configured. Please enter access token in Cloud Settings.");
    }
    if (!fileEntries || !fileEntries.length) return null;

    const headers = {
      "Content-Type": "application/json",
      "Accept": "application/vnd.github.v3+json",
      "Authorization": `token ${cfg.token}`
    };
    const baseUrl = `https://api.github.com/repos/${cfg.repo}`;

    try {
      if (onProgress) onProgress(`Syncing (${fileEntries.length} files)...`);
      const refRes = await fetch(`${baseUrl}/git/ref/heads/${encodeURIComponent(cfg.branch)}?_cb=${Date.now()}`, { headers });
      if (!refRes.ok) throw new Error(`Ref fetch failed (${refRes.status})`);
      const refData = await refRes.json();
      const latestCommitSha = refData.object.sha;

      const commitRes = await fetch(`${baseUrl}/git/commits/${latestCommitSha}`, { headers });
      if (!commitRes.ok) throw new Error(`Commit fetch failed (${commitRes.status})`);
      const commitData = await commitRes.json();
      const baseTreeSha = commitData.tree.sha;

      let completed = 0;
      const treeItems = await Promise.all(fileEntries.map(async (entry) => {
        const cleanPath = entry.path.startsWith("/") ? entry.path.slice(1) : entry.path;
        let contentB64 = "";
        if (entry.json !== undefined) {
          contentB64 = utf8ToBase64(JSON.stringify(entry.json, null, 2));
        } else if (entry.base64) {
          const raw = entry.base64.includes(",") ? entry.base64.split(",")[1] : entry.base64;
          contentB64 = raw.replace(/\s/g, "");
        }
        const blobRes = await fetch(`${baseUrl}/git/blobs`, {
          method: "POST",
          headers,
          body: JSON.stringify({ content: contentB64, encoding: "base64" })
        });
        if (!blobRes.ok) throw new Error(`Blob creation failed for ${cleanPath} (${blobRes.status})`);
        const blobData = await blobRes.json();
        completed++;
        if (onProgress) onProgress(`Uploading (${completed}/${fileEntries.length})...`);
        return {
          path: cleanPath,
          mode: "100644",
          type: "blob",
          sha: blobData.sha
        };
      }));

      if (onProgress) onProgress("Finalizing commit...");
      const treeRes = await fetch(`${baseUrl}/git/trees`, {
        method: "POST",
        headers,
        body: JSON.stringify({ base_tree: baseTreeSha, tree: treeItems })
      });
      if (!treeRes.ok) throw new Error(`Tree creation failed (${treeRes.status})`);
      const treeData = await treeRes.json();

      const newCommitRes = await fetch(`${baseUrl}/git/commits`, {
        method: "POST",
        headers,
        body: JSON.stringify({
          message: commitMessage || "Batch update via web portal",
          tree: treeData.sha,
          parents: [latestCommitSha]
        })
      });
      if (!newCommitRes.ok) throw new Error(`Commit creation failed (${newCommitRes.status})`);
      const newCommitData = await newCommitRes.json();

      const updateRefRes = await fetch(`${baseUrl}/git/refs/heads/${encodeURIComponent(cfg.branch)}`, {
        method: "PATCH",
        headers,
        body: JSON.stringify({ sha: newCommitData.sha, force: false })
      });
      if (!updateRefRes.ok) throw new Error(`Ref update failed (${updateRefRes.status})`);
      return await updateRefRes.json();
    } catch (batchErr) {
      console.warn("Batch commit failed, falling back to sequential commits:", batchErr);
      let idx = 0;
      for (const entry of fileEntries) {
        idx++;
        if (onProgress) onProgress(`Saving (${idx}/${fileEntries.length})...`);
        if (entry.json !== undefined) {
          await commitFile(entry.path, entry.json, commitMessage);
        } else if (entry.base64) {
          await commitBinaryFile(entry.path, entry.base64, `Upload ${entry.path}`);
        }
      }
      return { fallback: true };
    }
  }

  root.DTCloud = {
    getConfig,
    saveConfig,
    clearToken,
    isConfigured,
    fetchFile,
    commitFile,
    commitBinaryFile,
    commitBatch
  };
})(window);
