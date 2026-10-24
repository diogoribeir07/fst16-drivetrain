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
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) {
      console.warn("Could not read fst16_cloud_config", e);
    }
    return {
      token: "",
      repo: DEFAULT_REPO,
      branch: DEFAULT_BRANCH
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
        const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}?ref=${cfg.branch}&t=${Date.now()}`;
        const res = await fetch(url, {
          headers: {
            "Accept": "application/vnd.github.v3+json",
            "Authorization": `token ${cfg.token}`,
            "Cache-Control": "no-cache",
            "Pragma": "no-cache"
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
      throw new Error("GitHub Token not configured. Please enter access token.");
    }
    const cleanPath = path.startsWith("/") ? path.slice(1) : path;
    const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}`;

    // 1. Get current SHA if exists
    let sha = null;
    try {
      const headRes = await fetch(`${url}?ref=${cfg.branch}&t=${Date.now()}`, {
        headers: {
          "Accept": "application/vnd.github.v3+json",
          "Authorization": `token ${cfg.token}`,
          "Cache-Control": "no-cache",
          "Pragma": "no-cache"
        }
      });
      if (headRes.ok) {
        const headJson = await headRes.json();
        sha = headJson.sha;
      }
    } catch (e) {
      // file might not exist yet
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

    if (putRes.status === 409 && retryCount < 2) {
      console.warn(`409 Conflict encountered on ${cleanPath}. Retrying commit (attempt ${retryCount + 1})...`);
      await new Promise(r => setTimeout(r, 600));
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
      throw new Error("GitHub Token not configured. Please enter access token.");
    }
    const cleanPath = path.startsWith("/") ? path.slice(1) : path;
    const url = `https://api.github.com/repos/${cfg.repo}/contents/${cleanPath}`;

    let sha = null;
    try {
      const headRes = await fetch(`${url}?ref=${cfg.branch}&t=${Date.now()}`, {
        headers: {
          "Accept": "application/vnd.github.v3+json",
          "Authorization": `token ${cfg.token}`,
          "Cache-Control": "no-cache",
          "Pragma": "no-cache"
        }
      });
      if (headRes.ok) {
        const headJson = await headRes.json();
        sha = headJson.sha;
      }
    } catch (e) {}

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

    if (putRes.status === 409 && retryCount < 2) {
      console.warn(`409 Conflict on binary ${cleanPath}. Retrying (attempt ${retryCount + 1})...`);
      await new Promise(r => setTimeout(r, 600));
      return commitBinaryFile(path, base64Content, commitMessage, retryCount + 1);
    }

    if (!putRes.ok) {
      const errJson = await putRes.json().catch(() => ({}));
      throw new Error(errJson.message || `Upload failed (${putRes.status})`);
    }

    return await putRes.json();
  }

  root.DTCloud = {
    getConfig,
    saveConfig,
    clearToken,
    isConfigured,
    fetchFile,
    commitFile,
    commitBinaryFile
  };
})(window);
