// Edelburg signature add-in — runtime logic.
//
// Shared by three load paths (see manifest.xml <Runtimes>):
//   - Outlook on the web / new Outlook on Windows / Outlook on Mac load this
//     via functions.html.
//   - Classic Outlook on Windows loads this file directly (no HTML wrapper).
//   - The "Insert Edelburg Signature" ribbon button's ExecuteFunction action
//     also runs through functions.html.
//
// checkSignature() is wired to OnNewMessageCompose (fires for new messages,
// replies, and forwards) via a manifest LaunchEvent, so it runs automatically
// whenever someone starts composing. insertSignatureCommand() backs the
// manual ribbon button, for the cases (a stale draft, a client where the
// launch event didn't fire) where auto-insert didn't happen or the user
// cleared the signature and wants it back.

Office.onReady();

const SIGNATURE_BASE_URL = "https://s3-signatures.edelburg.net";

function getAliasFromEmail(email) {
  if (!email || email.indexOf("@") === -1) {
    return null;
  }
  return email.split("@")[0].toLowerCase();
}

function insertSignature(done) {
  const email = Office.context.mailbox.userProfile.emailAddress;
  const alias = getAliasFromEmail(email);
  if (!alias) {
    done();
    return;
  }

  const url = SIGNATURE_BASE_URL + "/" + encodeURIComponent(alias) + ".html";

  fetch(url)
    .then(function (response) {
      if (!response.ok) {
        throw new Error("HTTP " + response.status);
      }
      return response.text();
    })
    .then(function (html) {
      Office.context.mailbox.item.body.setSignatureAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        function () {
          done();
        }
      );
    })
    .catch(function () {
      // No hosted signature for this mailbox (shared/guest mailbox, or not
      // yet in employees.csv) — leave the compose window untouched rather
      // than surfacing an error for something that isn't actionable here.
      done();
    });
}

function checkSignature(eventObj) {
  insertSignature(function () {
    eventObj.completed();
  });
}

function insertSignatureCommand(event) {
  insertSignature(function () {
    event.completed();
  });
}

Office.actions.associate("checkSignature", checkSignature);
Office.actions.associate("insertSignatureCommand", insertSignatureCommand);
