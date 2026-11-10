const SITE = "https://echodrive.qinyibo-2013-9.workers.dev";
const CALLBACK = SITE + "/auth/callback";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const ROOT = "/me/drive/special/approot";
const SCOPE = "https://graph.microsoft.com/Files.ReadWrite.AppFolder https://graph.microsoft.com/User.Read";
const OAUTH = "__Host-echo-oauth";
const SESSION = "__Host-echo-session";
const COOKIE = "Path=/; Secure; HttpOnly; SameSite=Lax";
const LIMIT = 20 * 1024 * 1024;
const encoder = new TextEncoder();

function base64(bytes) {
  return btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function unbase64(value) {
  return Uint8Array.from(atob(value.replace(/-/g, "+").replace(/_/g, "/")), c => c.charCodeAt(0));
}

function readCookie(request, name) {
  return (request.headers.get("Cookie") || "").split(/;\s*/)
    .find(value => value.startsWith(name + "="))?.slice(name.length + 1) || "";
}

function setCookie(headers, name, value, seconds) {
  headers.append("Set-Cookie", `${name}=${value}; ${COOKIE}; Max-Age=${seconds}`);
}

async function sessionKey(env) {
  const material = encoder.encode("EchoDrive session v1\0" + env.MS_CLIENT_SECRET);
  return crypto.subtle.importKey("raw", await crypto.subtle.digest("SHA-256", material),
    "AES-GCM", false, ["encrypt", "decrypt"]);
}

async function writeSession(headers, env, token, expiresIn) {
  const seconds = Math.min(3600, Math.floor(Number(expiresIn)) - 30);
  if (!(seconds > 0)) throw new Problem("微软授权有效期异常。", 502);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const data = encoder.encode(JSON.stringify({ token, owner: env.MS_OWNER_ID.trim(), expires: Date.now() + seconds * 1000 }));
  const encrypted = new Uint8Array(await crypto.subtle.encrypt(
    { name: "AES-GCM", iv }, await sessionKey(env), data));
  const value = base64(iv) + "." + base64(encrypted);
  if (value.length > 9000) throw new Problem("登录会话太长，无法保存。", 502);
  for (let i = 0; i < 3; i++) {
    const part = value.slice(i * 3000, (i + 1) * 3000);
    setCookie(headers, SESSION + i, part, part ? seconds : 0);
  }
}

async function readSession(request, env) {
  const value = [0, 1, 2].map(i => readCookie(request, SESSION + i)).join("");
  if (!value || value.length > 9000) return null;
  try {
    const [iv, encrypted] = value.split(".").map(unbase64);
    const bytes = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv }, await sessionKey(env), encrypted);
    const session = JSON.parse(new TextDecoder().decode(bytes));
    return session.owner === env.MS_OWNER_ID.trim() && session.expires > Date.now()
      ? session.token : null;
  } catch {
    return null;
  }
}

const escapeHtml = value => String(value).replace(/[&<>"']/g,
  c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

function response(body, status = 200, extra) {
  const headers = new Headers(extra);
  headers.set("Cache-Control", "private, no-store");
  headers.set("Referrer-Policy", "no-referrer");
  headers.set("X-Content-Type-Options", "nosniff");
  headers.set("Content-Security-Policy", "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'");
  return new Response(body, { status, headers });
}

function page(title, content, status = 200, extra) {
  const headers = new Headers(extra);
  headers.set("Content-Type", "text/html; charset=utf-8");
  return response(`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)} · EchoDrive</title><style>
  body{font:16px/1.5 system-ui,sans-serif;background:#f5f7fb;color:#182238;margin:0}
  main{max-width:850px;margin:auto;padding:24px 16px} h1{margin:0 0 16px}
  header,li,.tools{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
  header{justify-content:space-between} li,.panel{background:white;border:1px solid #dbe3ef;border-radius:12px;padding:12px;margin:8px 0}
  ul{padding:0;list-style:none} li strong{flex:1;min-width:120px;overflow-wrap:anywhere}
  a{color:#1752ab} button,input{font:inherit;max-width:100%} button{cursor:pointer}
  label{display:block} form{margin:0} small{color:#41516c} details{max-width:100%}
  </style><main><header><h1>${escapeHtml(title)}</h1></header>${content}</main></html>`, status, headers);
}

function redirect(location, headers = new Headers()) {
  headers.set("Location", location);
  return response(null, 303, headers);
}

class Problem extends Error {
  constructor(message, status) {
    super(message);
    this.status = status;
  }
}

async function graph(token, path, options = {}) {
  const headers = new Headers(options.headers);
  headers.set("Authorization", "Bearer " + token);
  const result = await fetch(GRAPH + path, { ...options, headers, redirect: "manual" });
  if (result.ok || (options.download && result.status === 302)) return result;
  if (result.status === 401) throw new Problem("微软会话已过期，请重新登录。", 401);
  const data = await result.json().catch(() => ({}));
  const detail = typeof data.error?.code === "string" ? ` / ${data.error.code}` : "";
  throw new Problem(`OneDrive 操作失败（HTTP ${result.status}${detail}）。`, result.status < 500 ? result.status : 502);
}

function jsonOptions(method, body, headers = {}) {
  return { method, body: JSON.stringify(body), headers: { ...headers, "Content-Type": "application/json" } };
}

function validName(name) {
  return typeof name === "string" && name.length > 0 && name.length <= 255 &&
    name !== "." && name !== ".." && !/[\u0000-\u001f\u007f\\/:*?"<>|]/.test(name);
}

function checkedName(name) {
  if (!validName(name) || name !== name.trim() || name.endsWith(".")) {
    throw new Problem("文件名或文件夹名无效。", 400);
  }
  return name;
}

function folderPath(url) {
  const raw = url.searchParams.get("folder");
  const parts = raw ? raw.split("/") : [];
  if (parts.length > 50 || parts.some(name => !validName(name))) {
    throw new Problem("文件夹路径无效。", 400);
  }
  return parts;
}

function appPath(parts) {
  return parts.length ? ROOT + ":/" + parts.map(encodeURIComponent).join("/") + ":" : ROOT;
}

function localUrl(route, parts, fields = {}) {
  const query = new URLSearchParams(fields);
  if (parts.length) query.set("folder", parts.join("/"));
  return route + (query.size ? "?" + query : "");
}

async function getItem(token, parts) {
  return (await graph(token, appPath(parts) + "?$select=id,name,folder,file,eTag,remoteItem")).json();
}

function itemPath(item) {
  return "/me/drive/items/" + encodeURIComponent(item.id);
}

function signedUrl(value) {
  let url;
  try { url = new URL(value); } catch {}
  if (url?.protocol !== "https:" || url.username || url.password) {
    throw new Problem("OneDrive 返回的临时地址无效。", 502);
  }
  return url;
}

async function readForm(request) {
  const size = Number(request.headers.get("Content-Length"));
  if (!request.headers.has("Content-Length") || !Number.isSafeInteger(size) || size < 0) {
    throw new Problem("请求缺少有效的长度，请从网页重新提交。", 411);
  }
  if (size > LIMIT + 65536) throw new Problem("单个文件最多 20 MiB。", 413);
  try { return await request.formData(); }
  catch { throw new Problem("无法读取表单。", 400); }
}

async function login(request, url, env) {
  if (url.pathname === "/auth/login") {
    const state = crypto.randomUUID();
    const verifier = crypto.randomUUID() + crypto.randomUUID();
    const challenge = base64(new Uint8Array(await crypto.subtle.digest("SHA-256", encoder.encode(verifier))));
    const query = new URLSearchParams({
      client_id: env.MS_CLIENT_ID, response_type: "code", redirect_uri: CALLBACK,
      scope: SCOPE, state, code_challenge: challenge, code_challenge_method: "S256",
      prompt: "select_account",
    });
    const headers = new Headers();
    setCookie(headers, OAUTH, state + "." + verifier, 600);
    return redirect(AUTH + "/authorize?" + query, headers);
  }

  const headers = new Headers();
  setCookie(headers, OAUTH, "", 0);
  const [state, verifier] = readCookie(request, OAUTH).split(".");
  const code = url.searchParams.get("code");
  if (!state || !verifier || state !== url.searchParams.get("state") || !code) {
    return page("登录未完成", '<p><a href="/auth/login">重新登录</a></p>', 400, headers);
  }
  const result = await fetch(AUTH + "/token", {
    method: "POST", redirect: "error",
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID, client_secret: env.MS_CLIENT_SECRET,
      grant_type: "authorization_code", code, redirect_uri: CALLBACK, code_verifier: verifier,
    }),
  });
  const { access_token: token, expires_in: expiresIn, error } = await result.json();
  if (!result.ok || !token) {
    return page("微软登录失败", `<p>${escapeHtml(error || result.status)}</p><a href="/">回到首页</a>`, 502, headers);
  }
  const profile = await (await graph(token, "/me?$select=id")).json();
  if (profile.id !== env.MS_OWNER_ID.trim()) {
    return page("此账号没有访问权限", '<a href="/auth/login">切换微软账号</a>', 403, headers);
  }
  await writeSession(headers, env, token, expiresIn);
  return redirect("/", headers);
}

async function list(token, url, parts) {
  const query = new URLSearchParams({ $top: "100", $select: "name,folder,file,size,remoteItem,eTag" });
  const cursor = url.searchParams.get("next");
  if (cursor) query.set("$skiptoken", cursor);
  const data = await (await graph(token, appPath(parts) + "/children?" + query)).json();
  const next = data["@odata.nextLink"] ? new URL(data["@odata.nextLink"]).searchParams.get("$skiptoken") : null;
  const href = (route, path, fields) => escapeHtml(localUrl(route, path, fields));
  const rows = data.value.map(item => {
    const name = escapeHtml(item.name);
    const blocked = item.remoteItem || !validName(item.name);
    const primary = blocked ? name : item.folder
      ? `<a href="${href("/", [...parts, item.name])}">${name}</a>`
      : item.file ? `<a href="${href("/download", parts, { name: item.name })}">${name}</a>` : name;
    const etag = escapeHtml(item.eTag || "");
    const inputs = `<input type="hidden" name="name" value="${name}"><input type="hidden" name="etag" value="${etag}">`;
    return `<li><strong>${primary}</strong><small>${item.folder ? "文件夹" : Number(item.size || 0).toLocaleString() + " B"}</small>
      ${blocked ? "" : `<details><summary>管理</summary>
      <form method="post" action="${href("/rename", parts)}">${inputs}
      <label>新名称 <input name="newName" value="${name}" maxlength="255" required></label><button>重命名</button></form>
      <form method="post" action="${href("/delete", parts)}">${inputs}
      <label><input type="checkbox" name="confirm" required> 移到 OneDrive 回收站</label><button>删除</button></form>
      </details>`}</li>`;
  }).join("");
  return page("EchoDrive", `<div class="tools"><a href="/">根目录</a>
    ${parts.length ? `<a href="${href("/", parts.slice(0,-1))}">返回上一级</a>` : ""}
    <form method="post" action="/logout"><button>退出</button></form></div>
    <p>当前目录 · ${escapeHtml(parts.join(" / ") || "根目录")}</p>
    <div class="panel tools">
    <form method="post" action="${href("/mkdir", parts)}"><label>文件夹名称 <input name="name" maxlength="255" required></label><button>新建文件夹</button></form>
    <form method="post" enctype="multipart/form-data" action="${href("/upload", parts)}"><label>文件（最多 20 MiB） <input type="file" name="file" required></label><button>上传</button></form></div>
    <ul>${rows || "<li>这个文件夹是空的。</li>"}</ul>
    ${next ? `<a href="${href("/", parts, { next })}">下一页</a>` : ""}`);
}

async function download(token, url, parts) {
  const name = url.searchParams.get("name");
  if (!validName(name)) throw new Problem("文件名无效。", 400);
  const item = await getItem(token, [...parts, name]);
  if (!item.file || item.remoteItem) throw new Problem("请选择应用文件夹里的普通文件。", 400);
  let file = await graph(token, itemPath(item) + "/content", { download: true });
  if (file.status === 302) file = await fetch(signedUrl(file.headers.get("Location")));
  if (!file.ok) throw new Problem(`下载失败（HTTP ${file.status}）。`, 502);
  const encoded = encodeURIComponent(item.name).replace(/['()*]/g,
    c => "%" + c.charCodeAt(0).toString(16));
  return response(file.body, 200, {
    "Content-Type": "application/octet-stream",
    "Content-Disposition": "attachment; filename*=UTF-8''" + encoded,
  });
}

async function modify(request, token, route, parts) {
  const form = await readForm(request);
  if (route === "/mkdir") {
    await graph(token, appPath(parts) + "/children", jsonOptions("POST", {
      name: checkedName(form.get("name")), folder: {}, "@microsoft.graph.conflictBehavior": "fail",
    }));
  } else if (route === "/upload") {
    const file = form.get("file");
    if (!(file instanceof File) || !file.size || file.size > LIMIT) {
      throw new Problem("请选择一个大于 0 字节且不超过 20 MiB 的文件。", 400);
    }
    const name = checkedName(file.name);
    const path = appPath([...parts, name]) + "/createUploadSession";
    const session = await (await graph(token, path, jsonOptions("POST", {
      item: { "@microsoft.graph.conflictBehavior": "fail", name, fileSize: file.size },
    }))).json();
    const result = await fetch(signedUrl(session.uploadUrl), {
      method: "PUT", redirect: "error", body: file,
      headers: { "Content-Range": `bytes 0-${file.size - 1}/${file.size}` },
    });
    if (![200, 201].includes(result.status)) {
      throw new Problem(`上传未完成（HTTP ${result.status}）。如果同名文件已存在，请改名后重试。`, result.status === 409 ? 409 : 502);
    }
  } else {
    const name = form.get("name");
    if (!validName(name)) throw new Problem("文件名无效。", 400);
    const item = await getItem(token, [...parts, name]);
    if (item.remoteItem || !item.eTag || item.eTag !== form.get("etag")) {
      throw new Problem("项目已变更或无法管理。请刷新列表后再试。", 409);
    }
    const headers = { "If-Match": item.eTag };
    if (route === "/rename") {
      await graph(token, itemPath(item), jsonOptions("PATCH", { name: checkedName(form.get("newName")) }, headers));
    } else {
      if (form.get("confirm") !== "on") throw new Problem("请先确认删除。", 400);
      await graph(token, itemPath(item), { method: "DELETE", headers });
    }
  }
  return redirect(localUrl("/", parts));
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.origin !== SITE) return response("Not found", 404);
    if (!env.MS_CLIENT_ID || !env.MS_CLIENT_SECRET || !env.MS_OWNER_ID?.trim()) {
      return page("缺少 Worker 配置", "<p>检查 MS_CLIENT_ID、MS_CLIENT_SECRET 和 MS_OWNER_ID。</p>", 503);
    }

    try {
      const route = url.pathname;
      const isGet = request.method === "GET";
      const isPost = request.method === "POST";
      if (isGet && ["/auth/login", "/auth/callback"].includes(route)) return await login(request, url, env);

      if (isPost && request.headers.get("Origin") !== SITE) throw new Problem("请求来源无效。", 403);
      if (isPost && route === "/logout") {
        const headers = new Headers();
        for (let i = 0; i < 3; i++) setCookie(headers, SESSION + i, "", 0);
        return redirect("/", headers);
      }
      const token = await readSession(request, env);
      if (!token) {
        return page("EchoDrive", '<p>这是私人网盘。</p><a href="/auth/login">登录微软账号</a>', 401);
      }

      const parts = folderPath(url);
      if (isGet && route === "/") return await list(token, url, parts);
      if (isGet && route === "/download") return await download(token, url, parts);
      if (isPost && ["/mkdir", "/upload", "/rename", "/delete"].includes(route)) {
        return await modify(request, token, route, parts);
      }
      throw new Problem("页面不存在。", 404);
    } catch (error) {
      const known = error instanceof Problem;
      return page("操作失败", `<p>${escapeHtml(known ? error.message : "请求失败，请重试。")}</p>
        <p><a href="/">回到首页</a> · <a href="/auth/login">重新登录</a></p>`, known ? error.status : 502);
    }
  },
};
