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

// setSignatureAsync has documented timing issues: it can fail (observed as
// a generic "Host Error" / code 5000, "The operation is not supported")
// when called before the compose editor has fully finished initializing,
// which happens easily here since OnNewMessageCompose fires very early.
// Retry a few times with backoff before giving up.
const SET_SIGNATURE_RETRY_DELAYS_MS = [300, 800, 1500];

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
      if (asyncResult.status === Office.AsyncResultStatus.Failed && attempt < SET_SIGNATURE_RETRY_DELAYS_MS.length) {
        setTimeout(function () {
          trySetSignature(html, eventObj, attempt + 1);
        }, SET_SIGNATURE_RETRY_DELAYS_MS[attempt]);
        return;
      }
      eventObj.completed();
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

  fetch(url)
    .then(function (response) {
      console.log("[Edelburg Signature] fetch response status:", response.status);
      if (!response.ok) {
        throw new Error("HTTP " + response.status);
      }
      return response.text();
    })
    .then(function (html) {
      console.log("[Edelburg Signature] fetched HTML length:", html.length);
      trySetSignature(html, eventObj, 0);
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
