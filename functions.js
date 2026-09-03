// Edelburg signature add-in — runtime logic.
//
// checkSignature() is wired to OnNewMessageCompose (fires for new messages,
// replies, and forwards) via a manifest LaunchEvent, so it runs automatically
// whenever someone starts composing — no task pane or button involved.
//
// Loaded two ways (see manifest.xml <Runtimes>):
//   - Outlook on the web / new Outlook on Windows / Outlook on Mac load this
//     via functions.html.
//   - Classic Outlook on Windows loads this file directly (no HTML wrapper).
//
// Logs every step to the console (visible in the functions.html frame's
// devtools context, not the host page's) so a failure is diagnosable instead
// of silent — this add-in still doesn't surface anything to the end user on
// failure, since there's nothing actionable for them to do about a missing
// signature mid-compose.

console.log("[Edelburg Signature] functions.js loaded");

Office.onReady(function (info) {
  console.log("[Edelburg Signature] Office.onReady", info);
});

const SIGNATURE_BASE_URL = "https://s3-signatures.edelburg.net";

function getAliasFromEmail(email) {
  if (!email || email.indexOf("@") === -1) {
    return null;
  }
  return email.split("@")[0].toLowerCase();
}

// Everything we can cheaply learn about the environment before touching the
// body. Logged once per compose so a failure report carries the context
// Microsoft asks for (host, version, requirement set, item type, body format).
function logDiagnostics(callback) {
  const item = Office.context.mailbox.item;
  const diag = Office.context.mailbox.diagnostics || {};
  let mailbox110 = null;
  try {
    mailbox110 = Office.context.requirements.isSetSupported("Mailbox", "1.10");
  } catch (e) {
    mailbox110 = "isSetSupported threw: " + e;
  }
  console.log("[Edelburg Signature] diagnostics:", {
    hostName: diag.hostName,
    hostVersion: diag.hostVersion,
    OWAView: diag.OWAView,
    platform: Office.context.platform,
    mailbox110: mailbox110,
    itemType: item ? item.itemType : "(no item)",
    hasSetSignatureAsync: !!(item && item.body && typeof item.body.setSignatureAsync === "function"),
    hasPrependAsync: !!(item && item.body && typeof item.body.prependAsync === "function")
  });

  if (item && item.body && typeof item.body.getTypeAsync === "function") {
    item.body.getTypeAsync(function (result) {
      console.log("[Edelburg Signature] body.getTypeAsync:", result.status, result.value, result.error);
      callback();
    });
  } else {
    console.log("[Edelburg Signature] body.getTypeAsync not available");
    callback();
  }
}

// setSignatureAsync is the right API: it places the signature where Outlook
// would put its own, replaces an existing one instead of stacking, and
// doesn't dirty the form. In Outlook on the web it has been observed failing
// deterministically with a generic "Host Error" / code 5000 ("The operation
// is not supported") for this add-in, with every documented cause ruled out.
// So: try it a few times (it does have documented early-compose timing
// issues), and if it still fails, fall back to prependAsync — a plain body
// write that's been supported since Mailbox 1.1 and doesn't go through the
// signature subsystem at all. The fallback is slightly less polished (no
// replace-on-reinsert semantics; a blank line is prepended so the cursor has
// somewhere to go above the signature) but it gets a signature into the
// message, which is the point.
const SET_SIGNATURE_RETRY_DELAYS_MS = [300, 800, 1500];

function finish(eventObj, label) {
  console.log("[Edelburg Signature] done:", label);
  eventObj.completed();
}

function fallbackPrepend(html, eventObj) {
  console.log("[Edelburg Signature] falling back to body.prependAsync");
  Office.context.mailbox.item.body.prependAsync(
    "<div><br></div>" + html,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      console.log(
        "[Edelburg Signature] prependAsync result:",
        asyncResult.status,
        asyncResult.error
      );
      finish(eventObj, asyncResult.status === Office.AsyncResultStatus.Succeeded
        ? "signature inserted via prependAsync fallback"
        : "prependAsync fallback also failed");
    }
  );
}

function trySetSignature(html, eventObj, attempt) {
  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      console.log(
        "[Edelburg Signature] setSignatureAsync attempt " + attempt + " result:",
        asyncResult.status,
        asyncResult.error
      );
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        finish(eventObj, "signature set via setSignatureAsync");
        return;
      }
      if (attempt < SET_SIGNATURE_RETRY_DELAYS_MS.length) {
        setTimeout(function () {
          trySetSignature(html, eventObj, attempt + 1);
        }, SET_SIGNATURE_RETRY_DELAYS_MS[attempt]);
        return;
      }
      fallbackPrepend(html, eventObj);
    }
  );
}

function checkSignature(eventObj) {
  console.log("[Edelburg Signature] checkSignature fired");

  const email = Office.context.mailbox.userProfile.emailAddress;
  const alias = getAliasFromEmail(email);
  console.log("[Edelburg Signature] signed-in email:", email, "derived alias:", alias);

  if (!alias) {
    console.log("[Edelburg Signature] no usable alias, skipping");
    eventObj.completed();
    return;
  }

  const url = SIGNATURE_BASE_URL + "/" + encodeURIComponent(alias) + ".html";
  console.log("[Edelburg Signature] fetching", url);

  // no-store: the signature files change rarely, but when they do (a new
  // title, a fixed logo) a stale browser-cached copy is exactly the kind
  // of failure nobody notices until an email has gone out with it.
  fetch(url, { cache: "no-store" })
    .then(function (response) {
      console.log("[Edelburg Signature] fetch response status:", response.status);
      if (!response.ok) {
        throw new Error("HTTP " + response.status);
      }
      return response.text();
    })
    .then(function (html) {
      console.log("[Edelburg Signature] fetched HTML length:", html.length);
      // Outlook on the web refuses every body write whose HTML references
      // an SVG image, with nothing more specific than "Host Error" 5000
      // (Office.Body docs: "SVG files aren't supported in mail signatures";
      // OfficeDev/office-js#6020). That cost days to find once, so make it
      // loud if it ever regresses in the hosted signature files.
      if (/<img[^>]+src=["'][^"']*\.svg(\?[^"']*)?["']/i.test(html)) {
        console.error(
          "[Edelburg Signature] signature HTML references an .svg image — " +
          "Outlook on the web will reject the insert with Host Error 5000. " +
          "Serve a PNG/JPG instead."
        );
      }
      logDiagnostics(function () {
        trySetSignature(html, eventObj, 0);
      });
    })
    .catch(function (error) {
      // No hosted signature for this mailbox (shared/guest mailbox, or not
      // yet in employees.csv), or the fetch failed — leave the compose
      // window untouched rather than surfacing an error for something
      // that isn't actionable here. Still logged above for diagnosis.
      console.log("[Edelburg Signature] failed:", error);
      eventObj.completed();
    });
}

Office.actions.associate("checkSignature", checkSignature);
