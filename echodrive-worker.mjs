import { DurableObject } from "cloudflare:workers";

const SITE = "https://echodrive.qinyibo-2013-9.workers.dev";
const CALLBACK = SITE + "/auth/callback";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const ROOT = "/me/drive/special/approot";
const SCOPE = "https://graph.microsoft.com/Files.ReadWrite.AppFolder https://graph.microsoft.com/User.Read offline_access";
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

const escapeHtml = value => String(value).replace(/[&<>"']/g,
  c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

function response(body, status = 200, extra) {
  const headers = new Headers(extra);
  headers.set("Cache-Control", "private, no-store");
  headers.set("Referrer-Policy", "strict-origin");
  headers.set("X-Content-Type-Options", "nosniff");
  headers.set("Content-Security-Policy", "default-src 'none'; script-src 'self'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'");
  return new Response(body, { status, headers });
}

const ICONS = {
  cloud: '<path d="M6.5 18a4.5 4.5 0 0 1-.4-9 6 6 0 0 1 11.5 1.6A3.8 3.8 0 0 1 18 18z"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
  file: '<path d="M6 3h8l4 4v14H6z M14 3v5h4 M9 12h6m-6 4h6"/>',
  search: '<circle cx="10.5" cy="10.5" r="5.5"/><path d="m15 15 5 5"/>',
  back: '<path d="M20 12H5m6-6-6 6 6 6"/>',
  exit: '<path d="M12 4H5v16h7m-2-8h11m-4-4 4 4-4 4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  upload: '<path d="M12 16V4m-5 5 5-5 5 5M4 15v5h16v-5"/>',
  lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 1 1 8 0v3m-4 5v2"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>',
  dots: '<circle cx="5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/>',
};

function icon(name) {
  return `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">${ICONS[name]}</svg>`;
}

const UI_SCRIPT = `
document.addEventListener("click", function (event) {
  const trigger = event.target.closest("[data-open]");
  if (trigger) {
    const dialog = document.getElementById(trigger.dataset.open);
    if (dialog && !dialog.open) {
      if (typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute("open", "");
      document.querySelectorAll("details[open]").forEach(menu => menu.open = false);
    }
  }
  const close = event.target.closest("[data-close]");
  if (close) {
    const dialog = close.closest("dialog");
    if (typeof dialog.close === "function") dialog.close();
    else dialog.removeAttribute("open");
  }
  if (!event.target.closest("details")) {
    document.querySelectorAll("details[open]").forEach(menu => menu.open = false);
  }
  if (event.target.closest("[data-dismiss]")) event.target.closest(".toast").remove();
});
document.querySelectorAll("time[datetime]").forEach(time => {
  const date = new Date(time.dateTime);
  if (!Number.isNaN(date.getTime())) time.textContent = new Intl.DateTimeFormat("zh-CN", {
    year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit"
  }).format(date);
});
const search = document.querySelector("[data-search]");
const fileList = document.querySelector("[data-file-list]");
if (search && fileList) {
  const rows = Array.from(fileList.querySelectorAll("[data-file-row]"));
  const chips = Array.from(document.querySelectorAll("[data-kind]"));
  const count = document.querySelector("[data-count]");
  const noMatches = document.querySelector("[data-no-matches]");
  let kind = "all";

  function filterFiles() {
    const term = search.value.trim().toLocaleLowerCase();
    let visible = 0;
    for (const row of rows) {
      const matchesKind = kind === "all" || kind === row.dataset.fileRow;
      const matchesName = row.querySelector(".filename").textContent.toLocaleLowerCase().includes(term);
      row.hidden = !(matchesKind && matchesName);
      if (!row.hidden) visible++;
    }
    count.textContent = visible + " 项";
    noMatches.hidden = !rows.length || visible > 0;
  }

  search.addEventListener("input", filterFiles);
  for (const chip of chips) {
    chip.addEventListener("click", () => {
      kind = chip.dataset.kind;
      for (const choice of chips) choice.setAttribute("aria-pressed", String(choice === chip));
      filterFiles();
    });
  }
}
const toast = document.querySelector(".toast");
if (toast) {
  const url = new URL(location.href);
  url.searchParams.delete("done");
  history.replaceState(null, "", url);
  setTimeout(() => { if (!toast.matches(":focus-within")) toast.remove(); }, 6000);
}
`;

function page(title, content, status = 200, extra, layout = "center") {
  const headers = new Headers(extra);
  headers.set("Content-Type", "text/html; charset=utf-8");
  return response(`<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title === "EchoDrive" ? title : escapeHtml(title) + " · EchoDrive"}</title><script src="/ui.js" defer></script><style>
  :root{--bg:#f8f8f8;--paper:#fff;--ink:#242424;--muted:#616161;--line:#e5e5e5;--blue:#0f6cbd;--pale:#f2f2f2;--shadow:0 4px 16px #00000022}
  *,*::before,*::after{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 "Segoe UI",system-ui,-apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
  a{color:var(--ink);text-decoration:none}a:hover{text-decoration:underline}button,input,summary{font:inherit}button,summary{cursor:pointer}form{margin:0}
  :focus-visible{outline:3px solid #0f6cbd;outline-offset:3px}svg{width:20px;height:20px;flex:none;vertical-align:middle}
  .bar{height:64px;border-bottom:1px solid var(--line);background:var(--paper);display:flex;align-items:center;padding:0 max(20px,calc((100vw - 1360px)/2));gap:14px}
  .brand{display:inline-flex;align-items:center;gap:10px;color:var(--ink);font-size:19px;font-weight:750;white-space:nowrap}
  .mark{display:grid;place-items:center;background:var(--pale);color:var(--ink);width:36px;height:36px;border-radius:10px}.mark svg{width:24px;height:24px}
  .brand:hover{text-decoration:none}.shell{min-width:0;padding:26px 28px 32px}.center{display:grid;place-items:center;min-height:calc(100vh - 64px);padding:28px 16px}
  .card{width:min(100%,420px);background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:34px;box-shadow:0 8px 24px #00000012}.card h1{font-size:26px;margin:13px 0 6px}
  .lock{color:var(--ink);background:var(--pale);width:48px;height:48px;border-radius:14px;display:grid;place-items:center}.lock svg{width:24px;height:24px}
  p{margin:0 0 16px}.muted,.help{color:var(--muted)}.help{font-size:13px}.help a{color:var(--blue)}.card form{margin:22px 0 14px}.card form .primary{width:100%;justify-content:center;margin-top:14px}
  .toolbar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px;margin-bottom:18px}.actions,.crumbs,.name,.dialog-actions{display:flex;align-items:center;gap:10px}
  .crumbs{flex-wrap:wrap;min-width:0;font-size:21px;overflow-wrap:anywhere}.crumbs a,.crumbs .sep{color:var(--muted)}.crumbs strong{font-weight:720}.sep{font-size:17px}
  .btn,summary{display:inline-flex;align-items:center;justify-content:center;gap:8px;color:var(--ink);min-height:40px;border:1px solid var(--line);border-radius:10px;padding:7px 13px;background:var(--paper);white-space:nowrap;font-weight:600}
  .btn:hover,summary:hover{background:#f0f0f0;text-decoration:none}.primary{background:var(--blue);border-color:var(--blue);color:white}.primary:hover{background:#115ea3}.quiet{border-color:transparent;background:transparent}.icon-btn{padding:6px;width:38px;min-height:36px}
  details{position:relative}summary{list-style:none}summary::-webkit-details-marker{display:none}.menu{position:absolute;z-index:2;right:0;top:calc(100% + 7px);width:max-content;min-width:170px;max-width:calc(100vw - 32px);padding:6px;background:var(--paper);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow)}
  .menu button{width:100%;justify-content:flex-start;border:0;border-radius:7px}.menu .danger{color:#a62936}
  .files{border:1px solid var(--line);border-radius:14px;background:var(--paper);min-height:380px}.grid{width:100%;border-spacing:0;table-layout:fixed;text-align:left}
  .grid th,.grid td{padding:12px 14px;border-bottom:1px solid var(--line)}.grid th{font-weight:600;color:var(--muted);font-size:13px;background:#fafafa}
  .grid th:first-child{border-top-left-radius:14px}.grid th:last-child{border-top-right-radius:14px}.grid tr:last-child td{border-bottom:0}.grid tbody tr:hover{background:#f6f6f6}
  .grid .date{width:24%}.grid .size{width:15%}.grid .row-actions{width:64px;text-align:right}.name{font-weight:600;overflow-wrap:anywhere}.name a{color:var(--ink)}
  .file-icon{display:grid;place-items:center;width:34px;height:34px;flex:none;border-radius:8px;color:#4a4a4a;background:var(--pale)}.file-icon.folder{color:#975600;background:#fff0cf}.size,.date{color:var(--muted);font-size:13px}
  .empty{text-align:center;min-height:380px;padding:55px 18px;display:grid;place-content:center;justify-items:center}.empty-art{display:grid;place-items:center;width:110px;height:90px;border-radius:24px;background:var(--pale);color:#616161;margin-bottom:20px}.empty-art svg{width:54px;height:54px}
  .empty h2{margin:0 0 7px;font-size:20px}.empty p{margin-bottom:16px}.pager{padding:16px 0;text-align:right}
  dialog{width:min(460px,calc(100% - 32px));padding:24px;border:1px solid var(--line);border-radius:15px;color:var(--ink);box-shadow:var(--shadow)}dialog::backdrop{background:#00000066}
  .dialog-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:16px}.dialog-head h2{margin:0;font-size:20px}.dialog-actions{justify-content:flex-end;margin-top:20px}
  label{display:block;font-weight:600}input:not([type="hidden"]):not([type="checkbox"]){font:inherit;display:block;width:100%;min-width:0;margin-top:6px;padding:10px 11px;border:1px solid #b5b5b5;border-radius:8px;background:white;color:var(--ink)}
  .danger-btn{background:#a62936;border-color:#a62936;color:white}.danger-btn:hover{background:#85212c}.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:12px;width:max-content;max-width:calc(100% - 28px);padding:9px 12px 9px 16px;background:#242424;color:white;border-radius:10px;box-shadow:var(--shadow);z-index:3}
  .toast .icon-btn{background:transparent;border:0;color:white}.toast>svg{color:#a6d871}
  @media(max-width:640px){.bar{height:56px}.center{min-height:calc(100vh - 56px)}.shell{padding:20px 12px}.toolbar{align-items:flex-start}.crumbs{width:100%;font-size:19px}.grid .date{display:none}.grid .size{width:78px}.grid th,.grid td{padding:9px 8px}.grid .row-actions{width:52px}.card{padding:25px}.empty{min-height:320px}.files{min-height:320px}}
  .bar.drive-bar{display:grid;grid-template-columns:minmax(160px,1fr) minmax(200px,480px) minmax(110px,1fr);padding:0 24px;gap:18px}
  .drive-bar .bar-end{justify-self:end}.bar-end .btn{gap:7px}
  .bar-search{display:flex;align-items:center;gap:9px;padding:0 13px;border:1px solid var(--line);background:#f6f6f6;border-radius:9px;color:var(--muted)}
  .bar-search input[type="search"]{padding:8px 0;margin:0;border:0;background:transparent;border-radius:0}
  .bar-search:focus-within{border-color:var(--blue);background:white}.bar-search input:focus-visible{outline:none}
  .drive-layout{display:grid;grid-template-columns:208px minmax(0,1fr);min-height:calc(100vh - 64px)}
  .side{padding:24px 13px 18px 15px;border-right:1px solid var(--line);display:flex;flex-direction:column;gap:10px}
  .side-caption{padding:0 11px;color:var(--muted);font-size:12px;font-weight:650}
  .side-link{display:flex;align-items:center;gap:10px;border-radius:7px;padding:9px 11px;font-weight:650}.side-link.active{background:#ececec;box-shadow:inset 3px 0 var(--blue)}
  .storage-note{margin-top:auto;padding:12px;border-top:1px solid var(--line);color:var(--muted);font-size:13px}.storage-note strong{display:block;color:var(--ink);margin-bottom:3px}
  .filters{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 13px}.filters .btn{min-height:33px;padding:4px 11px;border-radius:18px;font-size:13px;font-weight:550}.filters .btn[aria-pressed="true"]{background:#ececec;border-color:#d5d5d5;font-weight:700}
  .item-count{margin-left:auto;color:var(--muted);font-size:13px}.grid th,.grid td{padding-block:9px}.grid .name{min-width:0}
  .files{min-height:clamp(330px,65vh,740px)}.empty{min-height:inherit}.empty-art{position:relative;width:118px;height:94px;background:transparent}.empty-art>svg:first-child{width:104px;height:104px;color:#c9c9c9;fill:#efefef;stroke-width:1.15}
  .empty-art>svg:last-child{position:absolute;bottom:6px;right:15px;width:45px;height:45px;color:#a56b17;fill:#ffda87;filter:drop-shadow(0 3px 3px #0002)}
  .no-matches{padding:36px 14px;text-align:center;color:var(--muted)}[hidden]{display:none!important}
  @media(max-width:760px){.drive-layout{grid-template-columns:1fr}.side{display:none}.bar.drive-bar{padding:0 12px;grid-template-columns:auto minmax(0,1fr) auto;gap:8px}.bar.drive-bar .brand{font-size:0}.bar-end .btn{font-size:0;min-width:38px;padding:6px}.bar-search{padding:0 8px;gap:5px}.bar-search input[type="search"]{font-size:14px}.shell{padding:18px 12px}.files{min-height:360px}}
  </style></head><body><header class="bar ${layout === "drive" ? "drive-bar" : ""}"><a class="brand" href="/"><span class="mark">${icon("cloud")}</span>EchoDrive</a>
  ${layout === "drive" ? `<label class="bar-search">${icon("search")}<input type="search" data-search placeholder="搜索本页文件" aria-label="搜索本页文件" autocomplete="off"></label>
  <form method="post" action="/logout" class="bar-end"><button class="btn quiet" aria-label="退出网盘">${icon("exit")}退出</button></form>` : ""}</header>
  ${layout === "drive" ? content : `<main class="center"><section class="card"><span class="lock">${icon("lock")}</span><h1>${escapeHtml(title)}</h1>${content}</section></main>`}</body></html>`, status, headers);
}

function dialog(id, title, content) {
  return `<dialog id="${id}" aria-labelledby="${id}-title"><div class="dialog-head"><h2 id="${id}-title">${title}</h2>
    <button class="btn quiet icon-btn" type="button" data-close aria-label="关闭">${icon("close")}</button></div>${content}</dialog>`;
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
  if (result.status === 401) throw new Problem("OneDrive 授权不可用，请返回首页重新绑定。", 401);
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
  const grant = await exchange(env, {
    grant_type: "authorization_code", code, redirect_uri: CALLBACK, code_verifier: verifier,
  });
  if (!grant.access_token || !grant.refresh_token) {
    return page("微软授权未完成", '<p>请重新登录并允许 EchoDrive 保持 OneDrive 授权。</p><a href="/auth/login">重试</a>', 502, headers);
  }
  let profile;
  try {
    profile = await (await graph(grant.access_token, "/me?$select=id")).json();
  } catch (error) {
    if (error instanceof Problem) throw error;
    throw new Problem("读取微软账号资料失败（MS-PROFILE）。", 502);
  }
  if (profile.id !== env.MS_OWNER_ID.trim()) {
    return page("此账号没有访问权限", '<a href="/auth/login">切换微软账号</a>', 403, headers);
  }
  await vault(env, "bind", { refresh: grant.refresh_token });
  setCookie(headers, SESSION, "", 0);
  return redirect("/?bound=1", headers);
}

const SUCCESS = {
  mkdir: "已创建文件夹",
  upload: "文件已上传",
  rename: "已重命名",
  delete: "已移到 OneDrive 回收站",
};

function sizeLabel(size) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 ** 2) return `${(size / 1024).toFixed(1)} KiB`;
  return `${(size / 1024 ** 2).toFixed(1)} MiB`;
}

async function list(token, url, parts) {
  const query = new URLSearchParams({ $top: "100", $select: "name,folder,file,size,remoteItem,eTag,lastModifiedDateTime" });
  const cursor = url.searchParams.get("next");
  if (cursor) query.set("$skiptoken", cursor);
  const data = await (await graph(token, appPath(parts) + "/children?" + query)).json();
  const next = data["@odata.nextLink"] ? new URL(data["@odata.nextLink"]).searchParams.get("$skiptoken") : null;
  const href = (route, path, fields) => escapeHtml(localUrl(route, path, fields));
  const open = (id, label, glyph = "") =>
    `<button type="button" class="btn" data-open="${id}">${glyph ? icon(glyph) : ""}${label}</button>`;
  const submit = (label, danger = false) =>
    `<div class="dialog-actions"><button class="btn" type="button" data-close>取消</button>
    <button class="btn ${danger ? "danger-btn" : "primary"}">${label}</button></div>`;
  const dialogs = [
    dialog("new-folder", "新建文件夹", `<form method="post" action="${href("/mkdir", parts)}">
      <label>名称<input name="name" placeholder="请输入文件夹名称" maxlength="255" required autofocus></label>${submit("创建")}</form>`),
    dialog("upload-file", "上传文件", `<form method="post" enctype="multipart/form-data" action="${href("/upload", parts)}">
      <p class="help">选择一个文件，单个文件最多 20 MiB。</p><label>选择文件<input type="file" name="file" required></label>${submit("上传")}</form>`),
  ];
  const rows = data.value.map((item, index) => {
    const name = escapeHtml(item.name);
    const blocked = item.remoteItem || !validName(item.name);
    const link = blocked ? "" : item.folder ? href("/", [...parts, item.name])
      : item.file ? href("/download", parts, { name: item.name }) : "";
    if (!blocked) {
      const hidden = `<input type="hidden" name="name" value="${name}"><input type="hidden" name="etag" value="${escapeHtml(item.eTag || "")}">`;
      dialogs.push(
        dialog(`rename-${index}`, "重命名", `<form method="post" action="${href("/rename", parts)}">${hidden}
          <label>新名称<input name="newName" value="${name}" maxlength="255" required autofocus></label>${submit("保存")}</form>`),
        dialog(`delete-${index}`, "删除项目", `<form method="post" action="${href("/delete", parts)}">${hidden}
          <p>将 <strong>${name}</strong> 移到 OneDrive 回收站？</p><input type="hidden" name="confirm" value="on">${submit("删除", true)}</form>`),
      );
    }
    const modified = escapeHtml(item.lastModifiedDateTime || "");
    return `<tr data-file-row="${item.folder ? "folder" : "file"}"><td><div class="name"><span class="file-icon ${item.folder ? "folder" : ""}">${icon(item.folder ? "folder" : "file")}</span>
      ${link ? `<a class="filename" href="${link}">${name}</a>` : `<span class="filename">${name}</span>`}</div></td>
      <td class="date">${modified ? `<time datetime="${modified}">${modified.slice(0, 10)}</time>` : "—"}</td>
      <td class="size">${item.folder ? "—" : sizeLabel(Number(item.size || 0))}</td>
      <td class="row-actions">${blocked ? "" : `<details><summary class="icon-btn quiet" aria-label="管理 ${name}">${icon("dots")}</summary>
        <div class="menu">${open(`rename-${index}`, "重命名")}
        <button class="btn danger" type="button" data-open="delete-${index}">删除</button></div></details>`}</td></tr>`;
  }).join("");
  const crumbs = parts.map((part, index) => {
    const name = escapeHtml(part);
    return `<span class="sep" aria-hidden="true">›</span>${index === parts.length - 1
      ? `<strong aria-current="page">${name}</strong>`
      : `<a href="${href("/", parts.slice(0, index + 1))}">${name}</a>`}`;
  }).join("");
  const done = url.searchParams.get("done");
  const toast = Object.hasOwn(SUCCESS, done) ? `<div class="toast" role="status">${icon("check")}${SUCCESS[done]}
    <button class="btn icon-btn" type="button" data-dismiss aria-label="关闭提示">${icon("close")}</button></div>` : "";
  return page("EchoDrive", `<div class="drive-layout"><aside class="side" aria-label="网盘导航">
    <div class="side-caption">文件</div>
    <a class="side-link active" href="/">${icon("folder")}我的文件</a>
    ${parts.length ? `<a class="side-link" href="${href("/", parts.slice(0, -1))}">${icon("back")}上一级文件夹</a>` : ""}
    <div class="storage-note"><strong>OneDrive</strong>EchoDrive 文件存储</div>
    </aside><main class="shell"><div class="toolbar"><nav class="crumbs" aria-label="当前位置">
    ${parts.length ? '<a href="/">我的文件</a>' : '<strong aria-current="page">我的文件</strong>'}${crumbs}</nav>
    <div class="actions"><details><summary class="primary">${icon("plus")}创建或上传</summary><div class="menu">
    ${open("new-folder", "新建文件夹", "folder")}${open("upload-file", "文件上传", "upload")}</div></details></div></div>
    <div class="filters" role="group" aria-label="按类型筛选">
      <button class="btn" type="button" data-kind="all" aria-pressed="true">全部</button>
      <button class="btn" type="button" data-kind="folder" aria-pressed="false">${icon("folder")}文件夹</button>
      <button class="btn" type="button" data-kind="file" aria-pressed="false">${icon("file")}文件</button>
      <span class="item-count" data-count aria-live="polite">${data.value.length} 项</span>
    </div>
    <section class="files" data-file-list aria-label="文件列表">${rows ? `<table class="grid"><thead><tr>
      <th scope="col">名称</th><th scope="col" class="date">修改时间</th><th scope="col" class="size">大小</th><th scope="col" class="row-actions">操作</th>
      </tr></thead><tbody>${rows}</tbody></table><div class="no-matches" data-no-matches hidden>没有符合条件的项目</div>`
      : `<div class="empty"><span class="empty-art">${icon("cloud")}${icon("folder")}</span><h2>此文件夹为空</h2>
        <p class="muted">上传文件或新建文件夹来整理你的资料。</p>${open("upload-file", "上传文件", "upload")}</div><div data-no-matches hidden></div>`}</section>
    ${next ? `<div class="pager"><a class="btn" href="${href("/", parts, { next })}">下一页 →</a></div>` : ""}
    ${dialogs.join("")}${toast}</main></div>`, 200, undefined, "drive");
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
      method: "PUT", redirect: "manual", body: file,
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
  return redirect(localUrl("/", parts, { done: route.slice(1) }));
}

const SESSION_SECONDS = 12 * 60 * 60;
const ATTEMPT_WINDOW = 10 * 60 * 1000;
const MAX_ATTEMPTS = 8;

const randomToken = () => base64(crypto.getRandomValues(new Uint8Array(32)));
const digest = async value => base64(new Uint8Array(await crypto.subtle.digest("SHA-256", encoder.encode(value))));
const validTicket = value => /^[A-Za-z0-9_-]{43}$/.test(value);

async function vaultKey(env) {
  const material = encoder.encode("EchoDrive vault v1\0" + env.MS_OWNER_ID.trim() + "\0" + env.MS_CLIENT_SECRET);
  return crypto.subtle.importKey("raw", await crypto.subtle.digest("SHA-256", material),
    "AES-GCM", false, ["encrypt", "decrypt"]);
}

async function encryptToken(env, value) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const data = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await vaultKey(env), encoder.encode(value));
  return base64(iv) + "." + base64(new Uint8Array(data));
}

async function decryptToken(env, value) {
  const [iv, data] = value.split(".").map(unbase64);
  const plaintext = await crypto.subtle.decrypt({ name: "AES-GCM", iv }, await vaultKey(env), data);
  return new TextDecoder().decode(plaintext);
}

async function passwordVersion(env) {
  const material = encoder.encode("EchoDrive password v1\0" + env.MS_CLIENT_SECRET);
  const key = await crypto.subtle.importKey("raw", material,
    { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(env.ECHO_PASSWORD));
  return base64(new Uint8Array(signature));
}

function matchesHash(actual, expected) {
  let difference = actual.length ^ expected.length;
  for (let i = 0; i < actual.length; i++) difference |= actual[i] ^ expected[i];
  return difference === 0;
}

async function exchange(env, fields) {
  let result;
  try {
    result = await fetch(AUTH + "/token", {
      method: "POST", redirect: "manual", signal: AbortSignal.timeout(10000),
      body: new URLSearchParams({ client_id: env.MS_CLIENT_ID, client_secret: env.MS_CLIENT_SECRET, ...fields }),
    });
  } catch {
    throw new Problem("连接微软授权服务失败（MS-CONNECT）。", 502);
  }
  if (result.headers.get("Content-Type")?.includes("application/json")) {
    return result.json();
  }
  throw new Problem(`微软授权服务返回了意外响应（MS-TOKEN / HTTP ${result.status}）。`, 502);
}

export class EchoDriveVault extends DurableObject {
  constructor(ctx, env) {
    super(ctx, env);
    this.storage = ctx.storage;
    this.queue = Promise.resolve();
    this.access = null;
  }

  fetch(request) {
    const operation = new URL(request.url).pathname;
    const job = this.queue.then(() => this.handle(request)).catch(error => {
      if (error instanceof Problem) {
        return Response.json({ serviceError: error.message }, { status: error.status });
      }
      console.error("EchoDrive vault operation failed", operation, error?.name || "Error");
      return Response.json({ code: "vault_error" }, { status: 502 });
    });
    this.queue = job.catch(() => {});
    return job;
  }

  async handle(request) {
    const operation = new URL(request.url).pathname;
    const data = await request.json();
    const now = Date.now();
    const owner = this.env.MS_OWNER_ID.trim();
    const saved = await this.storage.get("vault");
    const state = saved?.owner === owner ? saved : { owner, sessions: [], failures: [] };
    const save = () => this.storage.put("vault", state);
    const sessions = () => state.sessions.filter(session => session.until > now);
    const result = (body, status = 200) => Response.json(body, { status });

    const version = await passwordVersion(this.env);
    if (state.version !== version) {
      state.version = version;
      state.sessions = [];
      state.failures = [];
      delete state.password;
      delete state.pending;
      if (saved?.owner === owner) await save();
    }

    if (operation === "/bind") {
      state.refresh = await encryptToken(this.env, data.refresh);
      state.sessions = [];
      state.failures = [];
      this.access = null;
      await save();
      return result({});
    }

    if (operation === "/unlock") {
      if (!state.refresh) return result({ code: "relink" }, 409);
      state.failures = state.failures.filter(time => time > now - ATTEMPT_WINDOW);
      if (state.failures.length >= MAX_ATTEMPTS) return result({ code: "limited" }, 429);
      const valid = typeof data.password === "string" && data.password.length <= 128 &&
        matchesHash(unbase64(await digest(data.password)), unbase64(await digest(this.env.ECHO_PASSWORD)));
      if (!valid) {
        state.failures.push(now);
        await save();
        return result({ code: "bad_password" }, 401);
      }
      const session = randomToken();
      state.sessions = [...sessions().slice(-9), { hash: await digest(session), until: now + SESSION_SECONDS * 1000 }];
      state.failures = [];
      await save();
      return result({ session });
    }

    const hash = await digest(data.session);
    const active = sessions();
    if (operation === "/logout") {
      state.sessions = active.filter(session => session.hash !== hash);
      if (state.sessions.length !== active.length) await save();
      return result({});
    }

    if (operation !== "/token" || !active.some(session => session.hash === hash)) {
      return result({ code: "locked" }, 401);
    }
    if (!state.refresh) return result({ code: "relink" }, 409);
    if (this.access?.owner === owner && this.access.until > now) return result({ token: this.access.token });

    let refresh;
    try {
      refresh = await decryptToken(this.env, state.refresh);
    } catch {
      return result({ code: "relink" }, 409);
    }
    const grant = await exchange(this.env, { grant_type: "refresh_token", refresh_token: refresh, scope: SCOPE });
    if (grant.error === "invalid_grant" || grant.error === "interaction_required") {
      delete state.refresh;
      state.sessions = [];
      this.access = null;
      await save();
      return result({ code: "relink" }, 409);
    }
    if (!grant.access_token || !(Number(grant.expires_in) > 30)) {
      return result({ code: "upstream" }, 502);
    }
    if (grant.refresh_token && grant.refresh_token !== refresh) {
      state.refresh = await encryptToken(this.env, grant.refresh_token);
      await save();
    }
    this.access = { owner, token: grant.access_token, until: now + (Number(grant.expires_in) - 30) * 1000 };
    return result({ token: this.access.token });
  }
}

async function vault(env, operation, payload) {
  let result;
  try {
    const stub = env.ECHO_VAULT.getByName("owner");
    const request = new Request("https://vault.internal/" + operation, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload),
    });
    const reply = await stub.fetch(request);
    result = await reply.json();
  } catch (error) {
    console.error("EchoDrive vault connection failed", operation, error?.name || "Error");
    throw new Problem(`Cloudflare 无法连接密码存储（VAULT-CONNECT / ${operation}）。`, 502);
  }
  if (result.serviceError) throw new Problem(result.serviceError, 502);
  if (result.code === "vault_error") {
    throw new Problem(`Cloudflare 密码存储操作失败（VAULT-STORE / ${operation}）。`, 502);
  }
  return result;
}

function passwordPage(message = "", status = 200, headers) {
  return page("EchoDrive", `<p class="muted">${escapeHtml(message || "输入网盘密码，访问你的私人文件。")}</p>
    <form method="post" action="/unlock"><label>网盘密码
    <input type="password" name="password" placeholder="请输入网盘密码" autocomplete="current-password" maxlength="128" required></label>
    <button class="btn primary">${icon("lock")}解锁</button></form>
    <div class="help"><a href="/auth/login">首次绑定或重新授权 OneDrive</a></div>`, status, headers);
}

async function passwordForm(request) {
  const size = Number(request.headers.get("Content-Length"));
  if (!size || size > 4096) throw new Problem("密码表单无效。", 400);
  return readForm(request);
}

async function unlock(env, password) {
  const result = await vault(env, "unlock", { password });
  if (result.code === "limited") return passwordPage("尝试次数过多，请稍后再试（最多等待 10 分钟）。", 429);
  if (result.code === "relink") return passwordPage("请先绑定微软账号，或重新绑定已失效的授权。", 409);
  if (!result.session) return passwordPage("密码错误。", 401);
  const headers = new Headers();
  setCookie(headers, SESSION, result.session, SESSION_SECONDS);
  return redirect("/", headers);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.origin !== SITE) return response("Not found", 404);
    if (request.method === "GET" && url.pathname === "/ui.js") {
      return response(UI_SCRIPT, 200, { "Content-Type": "text/javascript; charset=utf-8" });
    }
    if (!env.MS_CLIENT_ID || !env.MS_CLIENT_SECRET || !env.MS_OWNER_ID?.trim() || !env.ECHO_VAULT) {
      return page("缺少 Worker 配置", "<p>检查 MS_CLIENT_ID、MS_CLIENT_SECRET 、MS_OWNER_ID 和 ECHO_VAULT。</p>", 503);
    }

    if (!env.ECHO_PASSWORD || env.ECHO_PASSWORD.length > 128) {
      return page("缺少有效的网盘密码", "<p>请在 Worker 的运行时密钥中设置 ECHO_PASSWORD（最多 128 个字符）。</p>", 503);
    }

    const route = url.pathname;
    try {
      const isGet = request.method === "GET";
      const isPost = request.method === "POST";
      if (isGet && ["/auth/login", "/auth/callback"].includes(route)) return await login(request, url, env);

      if (isPost && request.headers.get("Origin") !== SITE) throw new Problem("请求来源无效。", 403);
      const session = readCookie(request, SESSION);
      if (isPost && route === "/unlock") {
        const form = await passwordForm(request);
        return await unlock(env, form.get("password"));
      }
      if (isPost && route === "/logout") {
        if (validTicket(session)) await vault(env, "logout", { session });
        const headers = new Headers();
        setCookie(headers, SESSION, "", 0);
        return redirect("/", headers);
      }
      if (!validTicket(session)) {
        const bound = isGet && route === "/" && url.searchParams.get("bound") === "1";
        return passwordPage(bound ? "OneDrive 已绑定。请输入网盘密码。" : "");
      }
      const { token, code } = await vault(env, "token", { session });
      if (code === "relink") return passwordPage("OneDrive 授权需要重新绑定。", 409);
      if (code === "upstream") throw new Problem("微软授权服务暂时不可用。", 502);
      if (!token) return passwordPage();
      const parts = folderPath(url);
      if (isGet && route === "/") return await list(token, url, parts);
      if (isGet && route === "/download") return await download(token, url, parts);
      if (isPost && ["/mkdir", "/upload", "/rename", "/delete"].includes(route)) {
        return await modify(request, token, route, parts);
      }
      throw new Problem("页面不存在。", 404);
    } catch (error) {
      const known = error instanceof Problem;
      if (!known) console.error("EchoDrive request failed", route, error?.name || "Error");
      return page("操作失败", `<p>${escapeHtml(known ? error.message : `请求失败（ECHO-REQUEST / ${route}）。`)}</p>
        <p><a href="/">回到首页</a> · <a href="/auth/login">重新绑定微软账号</a></p>`, known ? error.status : 502);
    }
  },
};
