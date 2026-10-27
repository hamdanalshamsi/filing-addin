/*
 * HAS Filing - Smart Alert on send (OnMessageSend)
 * Runs in classic Outlook (JavaScript-only runtime), new Outlook and Outlook on the web.
 * Kept to plain ES5-style callbacks so it runs in every Outlook runtime.
 *
 * Rule: if the email goes to at least one external recipient and the subject has
 * no matter tag ([261760] / [261760A]) and no [Confidential] / [Personal] marker,
 * show a prompt. The user can choose "Add to matter" (opens the HAS Filing pane)
 * or "Send anyway".
 */

var HAS_INTERNAL_DOMAINS = ["has.law", "alshamsilegal.com"];
var HAS_TAG_RE = /\[\s*\d{6}[A-Za-z]{0,2}\s*\]/;
var HAS_SKIP_RE = /\[\s*(confidential|personal)\s*\]/i;

function hasIsInternal(address) {
  var a = String(address || "").toLowerCase().trim();
  var at = a.lastIndexOf("@");
  if (at < 0) return true; // contact groups / unresolved names: not treated as external
  var domain = a.substring(at + 1);
  for (var i = 0; i < HAS_INTERNAL_DOMAINS.length; i++) {
    var d = HAS_INTERNAL_DOMAINS[i];
    if (domain === d || domain.slice(-(d.length + 1)) === "." + d) return true;
  }
  return false;
}

function hasGet(field, cb) {
  try {
    field.getAsync(function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
    });
  } catch (e) { cb(null); }
}

/* ---- conversation -> matter memory (same format as taskpane.js) ---- */
function hasImul(a, b) {
  if (Math.imul) return Math.imul(a, b);
  var ah = (a >>> 16) & 0xffff, al = a & 0xffff, bh = (b >>> 16) & 0xffff, bl = b & 0xffff;
  return ((al * bl) + (((ah * bl + al * bh) << 16) >>> 0) | 0);
}
function hasHex8(n) { var h = (n >>> 0).toString(16); while (h.length < 8) h = "0" + h; return h; }
function hasConvKey(id) {
  if (!id) return "";
  var a = 0x811c9dc5, b = 0x01000193 ^ 0x5bd1e995;
  for (var i = 0; i < id.length; i++) { var c = id.charCodeAt(i); a = hasImul(a ^ c, 16777619) >>> 0; b = hasImul(b ^ c, 2246822519) >>> 0; }
  return hasHex8(a) + hasHex8(b);
}
// stored value: "261760,261535" or "~261760" (leading ~ = court conversation)
function hasConvRaw(convId) {
  try {
    var m = Office.context.roamingSettings.get("hasConv") || {};
    var k = hasConvKey(convId);
    return k && m[k] ? String(m[k]) : "";
  } catch (e) { return ""; }
}
function hasConvCourt(convId) { return hasConvRaw(convId).charAt(0) === "~"; }
function hasConvTag(convId) { var v = hasConvRaw(convId); return v.charAt(0) === "~" ? "" : v; }
function hasCourtDomains() { try { return Office.context.roamingSettings.get("hasCourtDomains") || []; } catch (e) { return []; } }
function hasStripTags(subject) {
  return String(subject || "").replace(/\[\s*\d{6}[A-Za-z]{0,2}\s*\]/g, "").replace(/\s{2,}/g, " ").replace(/^\s+|\s+$/g, "");
}
function hasWithTag(subject, tag) {
  var s = String(subject || "").replace(/\[\s*\d{6}[A-Za-z]{0,2}\s*\]/g, "").replace(/\s{2,}/g, " ").replace(/^\s+|\s+$/g, "");
  var m = s.match(/^((re|fw|fwd|aw)\s*:\s*)+/i);
  var prefix = m ? m[0] : "";
  var tags = String(tag || "").split(","), out = "";
  for (var i = 0; i < tags.length && i < 5; i++) { var t = tags[i].replace(/^\s+|\s+$/g, ""); if (t) out += "[" + t + "] "; }
  return (prefix + out + s.slice(prefix.length)).replace(/\s+$/, "");
}
function hasConvId(item) { try { return item.conversationId || ""; } catch (e) { return ""; } }

/* ---- OnNewMessageCompose: reply / forward of a filed conversation gets its tag straight away ---- */
function onNewComposeHandler(event) {
  var done = false;
  function finish() { if (!done) { done = true; try { event.completed(); } catch (e) {} } }
  setTimeout(finish, 4000);
  try {
    var item = Office.context.mailbox.item;
    var convId = hasConvId(item);
    if (hasConvCourt(convId)) {
      hasGet(item.subject, function (subject) {
        subject = subject || "";
        if (!HAS_TAG_RE.test(subject)) { finish(); return; }
        try { item.subject.setAsync(hasStripTags(subject), function () { finish(); }); } catch (e) { finish(); }
      });
      return;
    }
    var tag = hasConvTag(convId);
    if (!tag) { finish(); return; }
    hasGet(item.subject, function (subject) {
      subject = subject || "";
      if (HAS_TAG_RE.test(subject) || HAS_SKIP_RE.test(subject)) { finish(); return; }
      try { item.subject.setAsync(hasWithTag(subject, tag), function () { finish(); }); }
      catch (e) { finish(); }
    });
  } catch (e) { finish(); }
}

var HAS_PROMPT = {
  allowEvent: false,
  errorMessage: "This email is not tagged to a matter. Do you want to add it to a matter? " +
    "Click \"Add to matter\" to add the matter tag (or mark it Confidential / Personal), then click Send again. " +
    "Or click \"Send anyway\" to send without a tag.",
  cancelLabel: "Add to matter",
  commandId: "hasComposeButton"
};

function onMessageSendHandler(event) {
  var done = false;
  function finish(opts) {
    if (done) return;
    done = true;
    try { event.completed(opts); }
    catch (e) { try { event.completed({ allowEvent: opts.allowEvent, errorMessage: opts.errorMessage }); } catch (e2) {} }
  }
  // never leave Outlook waiting: after 4 s decide with what we know
  setTimeout(function () { finish(HAS_PROMPT); }, 4000);
  try { hasCheck(event, finish); }
  catch (e) { console.log("[HAS] error", e); finish(HAS_PROMPT); }
}

function hasCheck(event, finish) {
  var item = Office.context.mailbox.item;

  hasGet(item.subject, function (subject) {
    subject = subject || "";
    var convId = hasConvId(item);
    var courtConv = hasConvCourt(convId);
    var courtDoms = hasCourtDomains();
    // court: remove any matter number and send without asking
    function sendAsCourt() {
      if (!HAS_TAG_RE.test(subject)) { finish({ allowEvent: true }); return; }
      try { item.subject.setAsync(hasStripTags(subject), function () { finish({ allowEvent: true }); }); }
      catch (e) { finish({ allowEvent: true }); }
    }
    if (courtConv) { sendAsCourt(); return; }
    if (courtDoms.length) {
      hasAnyRecipient(item, function (addr) { return hasDomainIn(addr, courtDoms); }, function (isCourt) {
        if (isCourt) sendAsCourt(); else hasCheckNormal(item, subject, finish);
      });
      return;
    }
    hasCheckNormal(item, subject, finish);
  });
}

function hasDomainIn(address, list) {
  var a = String(address || "").toLowerCase(), at = a.lastIndexOf("@");
  if (at < 0) return false;
  var d = a.substring(at + 1);
  for (var i = 0; i < list.length; i++) { var x = String(list[i]).toLowerCase(); if (d === x || d.slice(-(x.length + 1)) === "." + x) return true; }
  return false;
}
// cb(true) if any To/Cc/Bcc recipient matches test()
function hasAnyRecipient(item, test, cb) {
  var lists = [item.to, item.cc, item.bcc], pending = lists.length, hit = false;
  for (var i = 0; i < lists.length; i++) {
    (function (field) {
      if (!field) { if (--pending === 0) cb(hit); return; }
      hasGet(field, function (recips) {
        if (recips) for (var j = 0; j < recips.length; j++) if (test(recips[j].emailAddress)) { hit = true; break; }
        if (--pending === 0) cb(hit);
      });
    })(lists[i]);
  }
}

function hasCheckNormal(item, subject, finish) {
  (function () {
    if (HAS_TAG_RE.test(subject) || HAS_SKIP_RE.test(subject)) {
      finish({ allowEvent: true });
      return;
    }
    var convTag = hasConvTag(hasConvId(item));
    if (convTag) {
      try {
        item.subject.setAsync(hasWithTag(subject, convTag), function () { finish({ allowEvent: true }); });
      } catch (e) { finish({ allowEvent: true }); }
      return;
    }
    var lists = [item.to, item.cc, item.bcc];
    var pending = lists.length;
    var external = false;
    for (var i = 0; i < lists.length; i++) {
      (function (field) {
        if (!field) { if (--pending === 0) decide(); return; }
        hasGet(field, function (recips) {
          if (recips) {
            for (var j = 0; j < recips.length; j++) {
              if (!hasIsInternal(recips[j].emailAddress)) { external = true; break; }
            }
          }
          if (--pending === 0) decide();
        });
      })(lists[i]);
    }

    function decide() {
      if (!external) { finish({ allowEvent: true }); return; }
      finish(HAS_PROMPT);
    }
  })();
}

// Register the handler now and again once Office is ready (new Outlook / web need the second call)
function hasRegister() {
  try { Office.actions.associate("onMessageSendHandler", onMessageSendHandler); }
  catch (e) { console.log("[HAS] associate failed", e); }
  try { Office.actions.associate("onNewComposeHandler", onNewComposeHandler); }
  catch (e) { console.log("[HAS] associate failed", e); }
}
hasRegister();
try { if (Office.onReady) Office.onReady(function () { hasRegister(); }); } catch (e) {}
