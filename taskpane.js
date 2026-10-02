/* HAS Filing – Outlook add-in (pilot) */
const CFG = {
  tenant: "e477b781-89bb-444f-951c-b746325df73a",
  clientId: (window.HAS_CONFIG || {}).clientId,
  sitePath: "haslegal.sharepoint.com:/sites/MatterFiling",
  lists: { index: "Folder Index", log: "Filing Log", requests: "Filing Requests" },
  inboxLibrary: "Filing Inbox",
  inlineSkipBytes: 30000,
  timeZone: "Asia/Dubai",
};
const G = "https://graph.microsoft.com/v1.0";
const SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/Files.ReadWrite.All", "https://graph.microsoft.com/User.Read"];
const TAG_RE = /\[(\d{6})([A-Za-z]{0,2})\]/g;

let pca, siteId, listIds = {}, inboxDriveId;
const $ = (id) => document.getElementById(id);
const show = (el, html, cls) => { el.className = "hint " + (cls || ""); el.innerHTML = html; };
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const odataStr = (s) => encodeURIComponent(String(s).replace(/'/g, "''"));

function fail(e) {
  console.error(e);
  const box = $("err"); box.hidden = false;
  box.textContent = (e && (e.message || e.errorMessage)) ? (e.message || e.errorMessage) : String(e);
}

/* ---------- auth ---------- */
const withTimeout = (p, ms, label) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(label + " timed out")), ms))]);
const setLoading = (t) => { const l = $("loading"); l.hidden = false; l.textContent = t; };
let naa = false, cached = null;
async function initAuth() {
  const conf = { auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenant}`,
                         redirectUri: location.origin + location.pathname },
                 cache: { cacheLocation: "localStorage" } };
  naa = !!(Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
  try {
    pca = await withTimeout(msal.createNestablePublicClientApplication(conf), 10000, "Sign-in setup");
  } catch (e) {
    naa = false;
    pca = new msal.PublicClientApplication(conf);
    await pca.initialize();
  }
}
// interactive=true only from a button click (pop-ups need a user click)
async function token(interactive) {
  if (cached && cached.exp > Date.now() + 60000) return cached.t;
  const login = Office.context.mailbox.userProfile.emailAddress;
  const req = { scopes: SCOPES, loginHint: login };
  const keep = (r) => { cached = { t: r.accessToken, exp: r.expiresOn ? new Date(r.expiresOn).getTime() : Date.now() + 30 * 60000 }; return r.accessToken; };
  try {
    const acct = pca.getAllAccounts().find(a => (a.username || "").toLowerCase() === login.toLowerCase());
    return keep(await withTimeout(pca.acquireTokenSilent({ ...req, account: acct }), 12000, "Silent sign-in"));
  } catch (e) {
    if (!interactive) { const n = new Error("SIGNIN_NEEDED: " + (e.errorCode || e.message || e)); n.needSignIn = true; throw n; }
    return keep(await pca.acquireTokenPopup(req));
  }
}
async function gfetch(url, opt = {}) {
  const t = await token(false);
  const headers = { Authorization: "Bearer " + t, ...(opt.headers || {}) };
  const r = await fetch(url.startsWith("http") ? url : G + url, { ...opt, headers });
  if (!r.ok && r.status !== 404) {
    let msg = r.status + " " + r.statusText;
    try { const j = await r.json(); msg += " – " + (j.error && j.error.message || ""); } catch (_) {}
    const err = new Error(msg); err.status = r.status; throw err;
  }
  return r;
}
const gjson = async (url, opt) => { const r = await gfetch(url, opt); return r.status === 404 ? null : r.json(); };

/* ---------- SharePoint lookups ---------- */
async function initSite() {
  siteId = (await gjson(`/sites/${CFG.sitePath}?$select=id`)).id;
  const lists = (await gjson(`/sites/${siteId}/lists?$select=id,displayName&$top=100`)).value;
  for (const [k, name] of Object.entries(CFG.lists)) listIds[k] = (lists.find(l => l.displayName === name) || {}).id;
  const drives = (await gjson(`/sites/${siteId}/drives?$select=id,name`)).value;
  inboxDriveId = (drives.find(d => d.name === CFG.inboxLibrary) || {}).id;
}
const PREFER = { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } };
async function findIndex(matter) {
  const j = await gjson(`/sites/${siteId}/lists/${listIds.index}/items?expand=fields&$filter=fields/Title eq '${odataStr(matter)}'&$top=1`, PREFER);
  return j && j.value[0] ? j.value[0].fields : null;
}
async function findLog(messageKey) {
  const j = await gjson(`/sites/${siteId}/lists/${listIds.log}/items?expand=fields&$filter=fields/MessageKey eq '${odataStr(messageKey)}'&$orderby=createdDateTime desc&$top=10`, PREFER);
  return j ? j.value.map(v => ({ ...v.fields, itemId: v.id, created: v.createdDateTime })) : [];
}
async function findRequests(messageKey) {
  const j = await gjson(`/sites/${siteId}/lists/${listIds.requests}/items?expand=fields&$filter=fields/MessageKey eq '${odataStr(messageKey)}'&$top=10`, PREFER);
  return j ? j.value.map(v => ({ ...v.fields, itemId: v.id, created: v.createdDateTime })) : [];
}

/* ---------- matter input ---------- */
function parseMatter(raw) {
  const m = String(raw || "").trim().replace(/^\[|\]$/g, "").match(/^(\d{6})([A-Za-z]{0,2})$/);
  return m ? { parent: m[1], tag: m[1] + m[2].toUpperCase() } : null;
}
async function checkMatter(inputId, resultId, buttonId) {
  const out = $(resultId), btn = $(buttonId);
  const m = parseMatter($(inputId).value);
  btn.disabled = true;
  if (!m) { show(out, "Enter 6 digits, optionally followed by up to 2 letters (e.g. 261760A).", "bad"); return null; }
  show(out, "Checking…", "muted");
  const idx = await findIndex(m.parent);
  if (idx && idx.FolderPath) {
    show(out, `✓ Matter ${esc(m.parent)} – <a href="${esc(idx.FolderPath)}" target="_blank" rel="noopener">open folder</a>`, "ok");
  } else {
    show(out, `Matter ${esc(m.parent)} has not been filed to before. Its folder will be located when the email is filed; if no folder exists it goes to review.`, "warn");
  }
  btn.disabled = false;
  return m;
}

/* ---------- compose ---------- */
function setSubject(s) {
  return new Promise((res, rej) => Office.context.mailbox.item.subject.setAsync(s, (r) =>
    r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error)));
}
function subjAsync() {
  return new Promise((res, rej) => Office.context.mailbox.item.subject.getAsync((r) =>
    r.status === Office.AsyncResultStatus.Succeeded ? res(r.value || "") : rej(r.error)));
}
async function addTag() {
  const m = parseMatter($("c-matter").value); if (!m) return;
  let s = await subjAsync();
  s = s.replace(TAG_RE, "").replace(/\s{2,}/g, " ").trim();
  const prefix = (s.match(/^((re|fw|fwd|aw)\s*:\s*)+/i) || [""])[0];
  s = prefix + `[${m.tag}] ` + s.slice(prefix.length);
  await setSubject(s.trim());
  show($("c-add-result"), `Subject updated: ${esc(s.trim())}`, "ok");
}
async function addMarker(marker) {
  let s = await subjAsync();
  if (s.toLowerCase().includes(marker.toLowerCase())) { show($("c-dnf-result"), `Already marked ${esc(marker)}.`, "muted"); return; }
  await setSubject((marker + " " + s).trim());
  show($("c-dnf-result"), `Marked ${esc(marker)} – this email will not be filed.`, "ok");
}

/* ---------- read: status ---------- */
let filedState = null;   // { matter, row, idx, saved:Set(attachmentId) } for the latest Filed row
const fmtWhen = (d) => new Date(d).toLocaleString("en-GB", { timeZone: CFG.timeZone, dateStyle: "medium", timeStyle: "short" });
async function loadStatus() {
  const item = Office.context.mailbox.item;
  const key = item.internetMessageId;
  const box = $("status");
  if (!key) { box.textContent = "No message ID available for this item."; return; }
  const [log, reqs] = await Promise.all([findLog(key), findRequests(key)]);
  const events = [
    ...log.map(l => ({ t: new Date(l.created).getTime(), kind: "log", l })),
    ...reqs.filter(r => r.Status === "Pending" || r.Status === "Failed").map(r => ({ t: new Date(r.created).getTime(), kind: "req", r })),
  ].sort((x, y) => y.t - x.t);
  const cats = item.categories ? await mbxAsync((cb) => item.categories.getAsync(cb)).catch(() => []) : [];
  const dnf = (cats || []).some(c => c.displayName === DNF_CATEGORY);
  let html = "";
  // most recent meaningful event: a later failure/review never hides a successful filing
  const lastFiled = log.filter(l => l.Status === "Filed").sort((x, y) => new Date(y.created) - new Date(x.created))[0];
  const lastRemoved = log.filter(l => l.Status === "Removed").sort((x, y) => new Date(y.created) - new Date(x.created))[0];
  let e = events[0];
  if (lastFiled && !(lastRemoved && new Date(lastRemoved.created) > new Date(lastFiled.created)) && e && e.kind === "log" && e.l.Status !== "Filed")
    e = { kind: "log", l: lastFiled };
  if (e && e.kind === "log") {
    const l = e.l, when = fmtWhen(l.created);
    if (l.Status === "Filed") html = `<div class="status-item ok">✓ Filed to <b>${esc(l.MatterNo)}</b> · ${esc(when)}${l.FolderPath ? ` · <a href="${esc(l.FolderPath)}" target="_blank" rel="noopener">open folder</a>` : ""}</div>`;
    else if (l.Status === "Review") html = `<div class="status-item warn">In review${l.MatterNo ? " (" + esc(l.MatterNo) + ")" : ""} · ${esc(when)}</div>`;
    else if (l.Status === "Removed") html = `<div class="status-item muted">Removed from ${esc(l.MatterNo)} · ${esc(when)}</div>`;
    else html = `<div class="status-item bad">Not filed – last attempt failed${l.MatterNo ? " (" + esc(l.MatterNo) + ")" : ""} · ${esc(when)}</div>`;
  } else if (e && e.kind === "req") {
    html = e.r.Status === "Pending" ? `<div class="status-item muted">⏳ Being filed to ${esc(e.r.MatterTag || e.r.MatterNo)} – usually within 5 minutes</div>`
                                    : `<div class="status-item bad">Filing to ${esc(e.r.MatterTag || e.r.MatterNo)} failed: ${esc(e.r.Result || "")}</div>`;
  }
  if (dnf) html = `<div class="status-item warn">⛔ Marked Do not file</div>` + html;
  box.className = ""; box.innerHTML = html || `<span class="muted">Not filed yet.</span>`;
  // which attachments are already saved?
  const latestFiled = lastFiled;
  const stillFiled = latestFiled && !(lastRemoved && new Date(lastRemoved.created) > new Date(lastFiled.created));
  filedState = null;
  if (stillFiled) {
    const idx = await findIndex(latestFiled.MatterNo).catch(() => null);
    const saved = new Set();
    if (idx && idx.DriveId && idx.AttachmentsFolderId) {
      const kids = await attachmentFiles(idx).catch(() => []);
      for (const r of log.filter(l => l.Status === "Filed" && l.MatterNo === latestFiled.MatterNo))
        for (const id of matchSavedAttachments(kids, r.created)) saved.add(id);
    }
    filedState = { matter: latestFiled.MatterNo, row: latestFiled, idx, saved };
    if (!$("r-matter").value) $("r-matter").value = latestFiled.MatterNo;
  }
  renderAttachmentList();
}
async function attachmentFiles(idx) {
  return ((await gjson(`/drives/${idx.DriveId}/items/${idx.AttachmentsFolderId}/children?$select=id,name,createdDateTime&$top=999`)) || { value: [] }).value;
}
// email attachments that have a matching file (same name, or dated / renamed copy) saved within 20 min of a filing
function matchSavedAttachments(kids, filedAt, alsoRecent) {
  const t0 = new Date(filedAt).getTime(), out = [];
  for (const a of Office.context.mailbox.item.attachments || []) {
    const n = (cleanName(a.name) || "attachment").toLowerCase(), i = n.lastIndexOf(".");
    const stem = i > 0 ? n.slice(0, i) : n, ext = i > 0 ? n.slice(i) : "";
    if (kids.some(k => { const kn = k.name.toLowerCase(); const kt = new Date(k.createdDateTime).getTime();
          return kn.startsWith(stem) && (!ext || kn.endsWith(ext)) && (Math.abs(kt - t0) < 20 * 60000 || (alsoRecent && kt >= alsoRecent)); }))
      out.push(a.id);
  }
  return out;
}

/* ---------- read: file this email ---------- */
function b64ToBytes(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
const mbxAsync = (fn) => new Promise((res, rej) => fn((r) => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)));
function cleanName(s) {
  return String(s || "").replace(TAG_RE, "").replace(/["*:<>?\/\\|]/g, "-").replace(/\s{2,}/g, " ").trim().slice(0, 76).trim();
}
function stamp(d) {
  const p = new Intl.DateTimeFormat("en-GB", { timeZone: CFG.timeZone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false })
    .formatToParts(d).reduce((a, x) => (a[x.type] = x.value, a), {});
  return `${p.year}-${p.month}-${p.day} ${p.hour}${p.minute}`;
}
// Upload bytes to "<base>:/<name>" where base is "/drives/{d}/items/{folderId}" or "/drives/{d}/root:/{path}" style prefix.
async function uploadTo(base, name, bytes) {
  const target = `${base}:/${encodeURIComponent(name)}`;
  if (bytes.length <= 4 * 1024 * 1024) {
    const r = await gfetch(`${target}:/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", headers: { "Content-Type": "application/octet-stream" }, body: bytes });
    if (!r.ok) { const e = new Error("Upload failed: " + r.status); e.status = r.status; throw e; }
    return r.json();
  }
  const s = await gjson(`${target}:/createUploadSession`, { method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }) });
  if (!s) { const e = new Error("Folder not found"); e.status = 404; throw e; }
  const CH = 320 * 1024 * 10; let last;
  for (let i = 0; i < bytes.length; i += CH) {
    const part = bytes.slice(i, i + CH);
    const r = await fetch(s.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${i}-${i + part.length - 1}/${bytes.length}` }, body: part });
    if (!r.ok) { const e = new Error("Upload failed: " + r.status); e.status = r.status; throw e; }
    last = r;
  }
  return last.json();
}
const inboxBase = (path) => `/drives/${inboxDriveId}/root:/${path.split("/").map(encodeURIComponent).join("/")}`;

function selectedAttachmentIds() {
  return [...document.querySelectorAll('#r-att-list input[type=checkbox]:checked')].map(c => c.value);
}
function renderAttachmentList() {
  const box = $("r-att-list"), item = Office.context.mailbox.item;
  const list = (item.attachments || []);
  const saved = filedState ? filedState.saved : new Set();
  if (!list.length) { box.innerHTML = '<span class="muted small">No attachments.</span>'; updateFileButton(); return; }
  box.innerHTML = list.map((a) => {
    const kb = a.size ? (a.size >= 1048576 ? (a.size / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(a.size / 1024)) + " KB") : "";
    if (saved.has(a.id)) return `<div class="att saved"><span class="tick">✓</span> ${esc(a.name)} <span class="muted small">${kb} · saved</span></div>`;
    if (a.attachmentType === "cloud") return `<div class="att disabled"><span class="tick muted">–</span> ${esc(a.name)} <span class="muted small">cloud link · not saved</span></div>`;
    if (a.isInline) return `<label class="chk att"><input type="checkbox" value="${esc(a.id)}"> ${esc(a.name)} <span class="muted small">${kb} · image in email body (kept inside the .eml)</span></label>`;
    return `<label class="chk att"><input type="checkbox" value="${esc(a.id)}" checked> ${esc(a.name)} <span class="muted small">${kb}</span></label>`;
  }).join("");
  box.querySelectorAll("input[type=checkbox]").forEach(c => c.onchange = updateFileButton);
  updateFileButton();
}
function updateFileButton() {
  const btn = $("r-file"), sel = selectedAttachmentIds().length;
  if (filedState) {
    btn.textContent = sel ? `Save ${sel} selected attachment${sel > 1 ? "s" : ""} to ${filedState.matter}` : "Already filed";
    btn.disabled = !sel;
  } else {
    btn.textContent = "File to matter";
    btn.disabled = !parseMatter($("r-matter").value);
  }
}
async function collectParts(item, wantedIds, out, attachmentsOnly) {
  show(out, attachmentsOnly ? "Reading attachments…" : "Reading email…", "muted");
  const eml = attachmentsOnly ? null : b64ToBytes(await mbxAsync((cb) => item.getAsFileAsync(cb)));
  const atts = []; let skipped = 0;
  {
    for (const a of item.attachments || []) {
      if (!wantedIds.includes(a.id)) { if (a.attachmentType === "cloud") skipped++; continue; }
      const c = await new Promise((res, rej) => item.getAttachmentContentAsync(a.id, (r) => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)));
      let bytes, name = cleanName(a.name) || "attachment";
      if (c.format === "base64") bytes = b64ToBytes(c.content);
      else if (c.format === "eml") { bytes = new TextEncoder().encode(c.content); if (!/\.eml$/i.test(name)) name += ".eml"; }
      else if (c.format === "iCalendar") { bytes = new TextEncoder().encode(c.content); if (!/\.ics$/i.test(name)) name += ".ics"; }
      else { skipped++; continue; }
      atts.push({ name, bytes });
    }
  }
  return { eml, atts, skipped };
}

// Instant path: matter folder already known (Folder Index) and the user has edit rights on it.
async function fileDirect(idx, m, meta, parts, out) {
  show(out, "Saving email to matter folder…", "muted");
  const saved = await uploadTo(`/drives/${idx.DriveId}/items/${idx.EmailsFolderId}`, meta.emlName, parts.eml);
  let n = 0;
  for (const a of parts.atts) {
    show(out, `Saving attachment ${++n} of ${parts.atts.length}…`, "muted");
    await uploadTo(`/drives/${idx.DriveId}/items/${idx.AttachmentsFolderId}`, meta.stamped(a.name), a.bytes);
  }
  const me = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  await gjson(`/sites/${siteId}/lists/${listIds.log}/items`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: {
    Title: meta.subject.slice(0, 255), MessageKey: meta.key, MatterNo: m.parent, Status: "Filed", MatchMethod: "Add-in",
    Mailbox: me, Direction: meta.from.toLowerCase() === me ? "Out" : "In", FromAddress: meta.from.toLowerCase(), SentUtc: meta.sentIso,
    FolderPath: idx.FolderPath || "", EmlUrl: saved.webUrl || "", AttachmentCount: parts.atts.length } }) });
  return saved;
}

// Queue path: the filing flow (full access) files it within ~5 minutes.
async function fileQueued(m, meta, parts, out) {
  const reqId = `${stamp(new Date()).replace(/[- ]/g, "")}-${Math.random().toString(36).slice(2, 8)}`;
  show(out, "Sending to filing queue…", "muted");
  await uploadTo(inboxBase(`${reqId}/email`), meta.emlName, parts.eml);
  for (const a of parts.atts) await uploadTo(inboxBase(`${reqId}/attachments`), a.name, a.bytes);
  await gjson(`/sites/${siteId}/lists/${listIds.requests}/items`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: {
    Title: reqId, MatterNo: m.parent, MatterTag: m.tag, MessageKey: meta.key, Subject: meta.subject.slice(0, 255),
    FromAddress: meta.from, SentUtc: meta.sentIso, RequestedBy: Office.context.mailbox.userProfile.emailAddress,
    EmlName: meta.emlName, AttachmentCount: parts.atts.length, Status: "Pending" } }) });
}

let pollTimer;
function pollStatus(key) {
  clearInterval(pollTimer); let n = 0;
  pollTimer = setInterval(async () => {
    n++;
    try {
      await loadStatus();
      const pending = (await findRequests(key)).some(r => r.Status === "Pending");
      if (!pending || n > 40) clearInterval(pollTimer);
    } catch (_) { clearInterval(pollTimer); }
  }, 15000);
}

async function fileEmail() {
  const out = $("r-file-result"), btn = $("r-file");
  const m = parseMatter($("r-matter").value); if (!m) return;
  const item = Office.context.mailbox.item;
  const key = item.internetMessageId;
  btn.disabled = true;
  // already filed: only save the extra attachments that were ticked
  if (filedState) {
    try {
      const ids = selectedAttachmentIds(); if (!ids.length) return;
      const idx = filedState.idx;
      if (!(idx && idx.DriveId && idx.AttachmentsFolderId)) throw new Error("The matter folder isn't known yet – try again in a few minutes.");
      const st = stamp(new Date(item.dateTimeCreated || Date.now()));
      const startedAt = Date.now() - 60000;
      const parts = await collectParts(item, ids, out, true);
      let n = 0;
      for (const a of parts.atts) {
        show(out, `Saving attachment ${++n} of ${parts.atts.length}…`, "muted");
        const i = a.name.lastIndexOf(".");
        const nm = i > 0 ? `${a.name.slice(0, i)} – ${st}${a.name.slice(i)}` : `${a.name} – ${st}`;
        await uploadTo(`/drives/${idx.DriveId}/items/${idx.AttachmentsFolderId}`, nm, a.bytes);
      }
      for (const id of ids) filedState.saved.add(id);
      renderAttachmentList();
      show(out, `✓ Saved ${parts.atts.length} attachment(s) to ${esc(filedState.matter)}.`, "ok");
    } catch (e) {
      show(out, (e.status === 401 || e.status === 403) ? "You don't have edit access to this matter folder. Please ask IT to save it." : "Failed: " + esc(e.message || e), "bad");
    } finally { updateFileButton(); }
    return;
  }
  try {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.14") || !item.getAsFileAsync)
      throw new Error("This Outlook version can't export the email. Please update Outlook (needs Mailbox 1.14).");
    const log = await findLog(key);
    if (log.some(l => l.Status === "Filed" && l.MatterNo === m.parent)) { show(out, `Already filed to ${esc(m.parent)}.`, "muted"); return; }
    const reqs = await findRequests(key);
    if (reqs.some(r => r.Status === "Pending" && r.MatterNo === m.parent)) { show(out, `Already queued for ${esc(m.parent)}.`, "muted"); return; }

    const received = new Date(item.dateTimeCreated || Date.now());
    const base = cleanName(item.subject) || "(No subject)";
    const st = stamp(received);
    const meta = {
      key, subject: item.subject || "", from: item.from ? item.from.emailAddress : "", sentIso: received.toISOString(),
      emlName: `${base} – ${st}.eml`,
      stamped: (n) => { const i = n.lastIndexOf("."); return i > 0 ? `${n.slice(0, i)} – ${st}${n.slice(i)}` : `${n} – ${st}`; },
    };
    const parts = await collectParts(item, selectedAttachmentIds(), out);
    const skippedNote = parts.skipped ? `, ${parts.skipped} cloud link(s) skipped` : "";

    const idx = await findIndex(m.parent);
    if (idx && idx.State === "Active" && idx.DriveId && idx.EmailsFolderId && idx.AttachmentsFolderId) {
      try {
        await fileDirect(idx, m, meta, parts, out);
        show(out, `✓ Filed to ${esc(m.tag)} now – ${parts.atts.length} attachment(s)${skippedNote}. <a href="${esc(idx.FolderPath)}" target="_blank" rel="noopener">Open folder</a>`, "ok");
        loadStatus().catch(fail);
        return;
      } catch (e) {
        if (e.status !== 401 && e.status !== 403 && e.status !== 404) throw e;   // no access → use the queue
      }
    }
    await fileQueued(m, meta, parts, out);
    show(out, `✓ Sent for filing to ${esc(m.tag)} – ${parts.atts.length} attachment(s)${skippedNote}. The status above updates automatically (usually within 5 minutes).`, "ok");
    loadStatus().catch(fail);
    pollStatus(key);
  } catch (e) {
    show(out, "Failed: " + esc(e.message || e), "bad");
  } finally { updateFileButton(); }
}


/* ---------- read: do not file / remove ---------- */
const DNF_CATEGORY = "Do not file / Personal";
async function ensureCategory() {
  const mc = Office.context.mailbox.masterCategories;
  let detail = "";
  if (mc) {
    const have = await mbxAsync((cb) => mc.getAsync(cb)).catch(() => null);
    if (have && have.some(c => c.displayName === DNF_CATEGORY)) return;
    try { await mbxAsync((cb) => mc.addAsync([{ displayName: DNF_CATEGORY, color: Office.MailboxEnums.CategoryColor.Preset0 }], cb)); return; }
    catch (e) { detail = (e && (e.message || e.name)) || String(e); }
  } else detail = "masterCategories not available";
  // fallback: create the category with Microsoft Graph (needs MailboxSettings.ReadWrite)
  try {
    const login = Office.context.mailbox.userProfile.emailAddress;
    const acct = pca.getAllAccounts().find(x => (x.username || "").toLowerCase() === login.toLowerCase());
    const req = { scopes: ["https://graph.microsoft.com/MailboxSettings.ReadWrite"], loginHint: login, account: acct };
    let tk;
    try { tk = (await pca.acquireTokenSilent(req)).accessToken; } catch (_) { tk = (await pca.acquireTokenPopup(req)).accessToken; }
    const list = await (await fetch(G + "/me/outlook/masterCategories", { headers: { Authorization: "Bearer " + tk } })).json();
    if ((list.value || []).some(c => c.displayName === DNF_CATEGORY)) return;
    const r = await fetch(G + "/me/outlook/masterCategories", { method: "POST", headers: { Authorization: "Bearer " + tk, "Content-Type": "application/json" },
      body: JSON.stringify({ displayName: DNF_CATEGORY, color: "preset0" }) });
    if (r.ok || r.status === 409) { await new Promise(res => setTimeout(res, 1500)); return; }
    detail += " | Graph " + r.status;
  } catch (e) { detail += " | Graph: " + (e.errorCode || e.message || e); }
  const err = new Error("NOCAT"); err.detail = detail; throw err;
}
const NOCAT_MSG = `Couldn't create the "${DNF_CATEGORY}" category in your Outlook.`;
async function graphItemFromUrl(url) {
  const id = "u!" + btoa(unescape(encodeURIComponent(url))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return gjson(`/shares/${id}/driveItem?$select=id,name,parentReference,createdDateTime`);
}
async function removeFiled(rows, out) {
  const item = Office.context.mailbox.item;
  let removed = 0, missing = 0;
  for (const l of rows) {
    const created = new Date(l.created).getTime();
    if (l.EmlUrl) {
      const di = await graphItemFromUrl(l.EmlUrl);
      if (di) { await gfetch(`/drives/${di.parentReference.driveId}/items/${di.id}`, { method: "DELETE" }); removed++; } else missing++;
    }
    // attachments saved for this email: same names (or dated/renamed variants), saved within 20 min of the filing
    const idx = await findIndex(l.MatterNo);
    if (idx && idx.DriveId && idx.AttachmentsFolderId) {
      const kids = (await gjson(`/drives/${idx.DriveId}/items/${idx.AttachmentsFolderId}/children?$select=id,name,createdDateTime&$top=999`) || { value: [] }).value;
      const stems = (item.attachments || []).map(a => { const n = cleanName(a.name) || "attachment"; const i = n.lastIndexOf("."); return [i > 0 ? n.slice(0, i) : n, i > 0 ? n.slice(i).toLowerCase() : ""]; });
      for (const k of kids) {
        if (Math.abs(new Date(k.createdDateTime).getTime() - created) > 20 * 60000) continue;
        const kn = k.name.toLowerCase();
        if (stems.some(([st, ext]) => kn.startsWith(st.toLowerCase()) && (!ext || kn.endsWith(ext)))) {
          await gfetch(`/drives/${idx.DriveId}/items/${k.id}`, { method: "DELETE" }); removed++;
        }
      }
    }
    await gjson(`/sites/${siteId}/lists/${listIds.log}/items/${l.itemId}/fields`, { method: "PATCH", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ Status: "Removed", Error: `Removed via add-in by ${Office.context.mailbox.userProfile.emailAddress}` }) });
  }
  return { removed, missing };
}
async function doNotFile() {
  const out = $("r-dnf-result"), item = Office.context.mailbox.item, key = item.internetMessageId;
  $("r-dnf").disabled = true;
  try {
    show(out, "Marking as Do not file…", "muted");
    let catErr = "";
    try { await ensureCategory(); } catch (e) { catErr = e.detail || e.message; }
    try { await mbxAsync((cb) => item.categories.addAsync([DNF_CATEGORY], cb)); }
    catch (e) { show(out, `${NOCAT_MSG}<div class="muted small">${esc(catErr || (e && e.message) || "")}</div>`, "bad"); return; }
    // cancel anything still queued
    for (const r of (await findRequests(key)).filter(r => r.Status === "Pending"))
      await gjson(`/sites/${siteId}/lists/${listIds.requests}/items/${r.itemId}/fields`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ Status: "Cancelled", Result: "Marked Do not file" }) });
    const filed = (await findLog(key)).filter(l => l.Status === "Filed");
    if (!filed.length) { show(out, "✓ Marked Do not file – this email will not be filed.", "ok"); loadStatus().catch(fail); return; }
    const matters = [...new Set(filed.map(f => f.MatterNo))].join(", ");
    show(out, `✓ Marked Do not file. This email is already filed to <b>${esc(matters)}</b>. Remove the saved copy from the matter folder?
      <div class="row" style="margin-top:6px"><button id="rm-yes" class="secondary">Yes, remove it</button><button id="rm-no" class="secondary">Keep it</button></div>
      <div class="muted small">Removed files go to the OneDrive recycle bin and can be restored.</div>`, "warn");
    $("rm-no").onclick = () => show(out, "✓ Marked Do not file. The copy already filed was kept.", "ok");
    $("rm-yes").onclick = async () => {
      $("rm-yes").disabled = $("rm-no").disabled = true;
      try {
        show(out, "Removing…", "muted");
        const r = await removeFiled(filed, out);
        show(out, `✓ Removed ${r.removed} file(s) from ${esc(matters)}${r.missing ? ` (${r.missing} already gone)` : ""}. Marked Do not file.`, "ok");
      } catch (e) {
        show(out, (e.status === 403 || e.status === 401) ? "You don't have edit access to that matter folder, so the copy couldn't be removed. Please ask IT to remove it." : "Remove failed: " + esc(e.message || e), "bad");
      }
      loadStatus().catch(fail);
    };
  } catch (e) { show(out, "Failed: " + esc(e.message || e), "bad"); }
  finally { $("r-dnf").disabled = false; }
}

/* ---------- start ---------- */
const IS_AUTH_POPUP = !!window.opener && /(code|error)=/.test(location.hash + location.search);
if (!IS_AUTH_POPUP) Office.onReady(async () => {
  try {
    $("who").textContent = Office.context.mailbox.userProfile.emailAddress;
    const diag = $("diag"); if (diag) diag.textContent = `${Office.context.diagnostics ? Office.context.diagnostics.platform + " " + Office.context.diagnostics.version : ""}`;
    setLoading("Signing in…");
    await initAuth();
    if (diag) diag.textContent += naa ? " · NAA" : " · popup sign-in";
    try { await token(false); }
    catch (e) {
      if (!e.needSignIn) throw e;
      $("loading").hidden = true;
      $("signin").hidden = false;
      $("signin-detail").textContent = String(e.message || "").replace("SIGNIN_NEEDED: ", "");
      await new Promise((res) => { $("signin-btn").onclick = async () => {
        $("signin-btn").disabled = true;
        try { await token(true); $("signin").hidden = true; res(); }
        catch (err) { $("signin-btn").disabled = false; $("signin-detail").textContent = "Sign-in failed: " + (err.errorCode || err.message || err); }
      }; });
    }
    setLoading("Connecting to Matter Filing site…");
    await withTimeout(initSite(), 20000, "Connecting to Matter Filing site");
    if (!listIds.index || !listIds.log || !listIds.requests || !inboxDriveId) throw new Error("Matter Filing site lists are missing or you don't have access to the Matter Filing site.");
    const compose = typeof Office.context.mailbox.item.subject === "object";
    $("loading").hidden = true;
    if (compose) {
      $("compose").hidden = false;
      $("c-check").onclick = () => checkMatter("c-matter", "c-check-result", "c-add").catch(fail);
      $("c-matter").onkeydown = (e) => { if (e.key === "Enter") $("c-check").click(); };
      $("c-add").onclick = () => addTag().catch(fail);
      $("c-conf").onclick = () => addMarker("[Confidential]").catch(fail);
      $("c-pers").onclick = () => addMarker("[Personal]").catch(fail);
      // prefill from an existing tag in the subject
      subjAsync().then(s => { const t = [...s.matchAll(TAG_RE)][0]; if (t) $("c-matter").value = t[1] + t[2]; }).catch(() => {});
    } else {
      $("read").hidden = false;
      const t = [...String(Office.context.mailbox.item.subject || "").matchAll(TAG_RE)][0];
      if (t) $("r-matter").value = t[1] + t[2];
      $("r-check").onclick = () => checkMatter("r-matter", "r-check-result", "r-file").catch(fail);
      $("r-matter").onkeydown = (e) => { if (e.key === "Enter") $("r-check").click(); };
      $("r-matter").oninput = updateFileButton;
      $("r-file").onclick = () => fileEmail();
      $("r-dnf").onclick = () => doNotFile();
      renderAttachmentList();
      loadStatus().catch(fail);
      if (Office.context.mailbox.addHandlerAsync && Office.EventType.ItemChanged)
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => location.reload());
    }
  } catch (e) { $("loading").hidden = true; fail(e); }
});
