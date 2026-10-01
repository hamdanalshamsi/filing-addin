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
  return j ? j.value.map(v => ({ ...v.fields, created: v.createdDateTime })) : [];
}
async function findRequests(messageKey) {
  const j = await gjson(`/sites/${siteId}/lists/${listIds.requests}/items?expand=fields&$filter=fields/MessageKey eq '${odataStr(messageKey)}'&$top=10`, PREFER);
  return j ? j.value.map(v => ({ ...v.fields, created: v.createdDateTime })) : [];
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
async function loadStatus() {
  const item = Office.context.mailbox.item;
  const key = item.internetMessageId;
  const box = $("status");
  if (!key) { box.textContent = "No message ID available for this item."; return; }
  const [log, reqs] = await Promise.all([findLog(key), findRequests(key)]);
  const rows = [];
  for (const l of log) {
    const when = new Date(l.created).toLocaleString("en-GB", { timeZone: CFG.timeZone, dateStyle: "medium", timeStyle: "short" });
    if (l.Status === "Filed") rows.push(`<div class="status-item ok">✓ Filed to <b>${esc(l.MatterNo)}</b> · ${esc(when)}${l.FolderPath ? ` · <a href="${esc(l.FolderPath)}" target="_blank" rel="noopener">open folder</a>` : ""}</div>`);
    else if (l.Status === "Review") rows.push(`<div class="status-item warn">In review${l.MatterNo ? " (" + esc(l.MatterNo) + ")" : ""} – ${esc(l.Error || "")} · ${esc(when)}</div>`);
    else rows.push(`<div class="status-item bad">${esc(l.Status)}${l.MatterNo ? " (" + esc(l.MatterNo) + ")" : ""} · ${esc(when)}</div>`);
  }
  for (const r of reqs.filter(r => r.Status === "Pending")) rows.push(`<div class="status-item muted">⏳ Queued for ${esc(r.MatterTag || r.MatterNo)} – will be filed within ~5 minutes</div>`);
  for (const r of reqs.filter(r => r.Status === "Failed")) rows.push(`<div class="status-item bad">Add-in request failed for ${esc(r.MatterTag || r.MatterNo)}: ${esc(r.Result || "")}</div>`);
  box.className = ""; box.innerHTML = rows.length ? rows.join("") : `<span class="muted">Not filed yet.</span>`;
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

async function collectParts(item, withAtt, out) {
  show(out, "Reading email…", "muted");
  const eml = b64ToBytes(await mbxAsync((cb) => item.getAsFileAsync(cb)));
  const atts = []; let skipped = 0;
  if (withAtt) {
    for (const a of item.attachments || []) {
      if (a.attachmentType === "cloud") { skipped++; continue; }
      if (a.isInline && a.size < CFG.inlineSkipBytes) continue;
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
    const parts = await collectParts(item, $("r-att").checked, out);
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
  } finally { btn.disabled = false; }
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
      $("r-file").onclick = () => fileEmail();
      loadStatus().catch(fail);
      if (Office.context.mailbox.addHandlerAsync && Office.EventType.ItemChanged)
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => location.reload());
    }
  } catch (e) { $("loading").hidden = true; fail(e); }
});
