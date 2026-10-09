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

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var done = false;
  function finish(opts) { if (!done) { done = true; event.completed(opts); } }

  hasGet(item.subject, function (subject) {
    subject = subject || "";
    if (HAS_TAG_RE.test(subject) || HAS_SKIP_RE.test(subject)) {
      finish({ allowEvent: true });
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
      finish({
        allowEvent: false,
        errorMessage:
          "This email is not tagged to a matter. Do you want to add it to a matter?\n\n" +
          "Click \"Add to matter\" to add the matter tag (or mark it Confidential / Personal), then click Send again. " +
          "Or click \"Send anyway\" to send without a tag.",
        cancelLabel: "Add to matter",
        commandId: "hasComposeButton"
      });
    }
  });
}

// Register the handler (required for every Outlook runtime)
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
