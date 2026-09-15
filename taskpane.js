// Edelburg signature add-in — task pane.
//
// Opened from the signature button on the compose ribbon. Lists every
// signature hosted for the signed-in mailbox, with a preview of each;
// clicking one inserts it with setSignatureAsync, which replaces the
// signature already in the message (normally the standard one that
// functions.js inserted on compose).
//
// The list comes from <domain>/<local>.json on R2, written by the
// signatures repo's Generate-Signatures.ps1, not from the manifest. So a
// new signature only needs new files in the bucket: no manifest change, no
// waiting for Microsoft 365 to roll it out.
//
// <domain>/<local>.json looks like this; "file" is relative to its folder:
//   { "signatures": [
//       { "key": "company-a", "label": "Company A", "file": "j.doe.html", "default": true },
//       { "key": "company-b", "label": "Company B", "file": "j.doe.company-b.html", "default": false } ] }
//
// The same page serves every tenant's add-in. Its manifest picks the
// language with ?lang=de or ?lang=fr on the task pane URL (default: de),
// and passes auto=0 when that add-in inserts nothing on compose (LAF), so
// the panel doesn't label a signature as the automatic one.

const SIGNATURE_BASE_URL = "https://s3-signatures.edelburg.net";

const STRINGS = {
  de: {
    pageTitle: "Signatur",
    heading: "Signatur wählen",
    hint: "Klicken Sie auf eine Signatur, um sie in diese E-Mail einzufügen. Sie ersetzt die vorhandene Signatur.",
    loading: "Signaturen werden geladen …",
    standard: "Standard",
    done: "✓ Eingefügt",
    inserting: "Wird eingefügt …",
    inserted: function (label) { return "Signatur „" + label + "“ eingefügt."; },
    insertFailed: function (message) { return "Die Signatur konnte nicht eingefügt werden: " + message; },
    unavailable: "Diese Signatur ist derzeit nicht verfügbar.",
    noMailbox: "Ihr Postfach konnte nicht erkannt werden.",
    none: function (email) { return "Für " + email + " sind keine Signaturen hinterlegt."; },
    onlyOne: "Für Ihr Postfach ist nur eine Signatur hinterlegt.",
    loadFailed: "Die Signaturen konnten nicht geladen werden. Bitte versuchen Sie es später erneut."
  },
  fr: {
    pageTitle: "Signature",
    heading: "Choisir une signature",
    hint: "Cliquez sur une signature pour l’insérer dans cet e-mail. Elle remplace la signature existante.",
    loading: "Chargement des signatures …",
    standard: "Par défaut",
    done: "✓ Insérée",
    inserting: "Insertion …",
    inserted: function (label) { return "Signature « " + label + " » insérée."; },
    insertFailed: function (message) { return "La signature n’a pas pu être insérée : " + message; },
    unavailable: "Cette signature n’est pas disponible pour le moment.",
    noMailbox: "Votre boîte aux lettres n’a pas pu être identifiée.",
    none: function (email) { return "Aucune signature n’est configurée pour " + email + "."; },
    onlyOne: "Une seule signature est configurée pour votre boîte aux lettres.",
    loadFailed: "Les signatures n’ont pas pu être chargées. Veuillez réessayer plus tard."
  }
};

const lang = (function () {
  const requested = new URLSearchParams(window.location.search).get("lang");
  return STRINGS[requested] ? requested : "de";
})();
const t = STRINGS[lang];
const autoInserts = new URLSearchParams(window.location.search).get("auto") !== "0";

const statusEl = document.getElementById("status");
const listEl = document.getElementById("list");

// Same as in functions.js: "<domain>/<local>", URL-encoded, or null.
function getMailboxPath(email) {
  const match = /^([^@\s\/\\]+)@([^@\s\/\\]+)$/.exec(email || "");
  if (!match) {
    return null;
  }
  return encodeURIComponent(match[2].toLowerCase()) + "/" + encodeURIComponent(match[1].toLowerCase());
}

// no-store, as in functions.js: a stale cached copy of a changed signature
// is exactly the failure nobody notices until an email has gone out.
function fetchHosted(path, as) {
  return fetch(SIGNATURE_BASE_URL + "/" + path, { cache: "no-store" })
    .then(function (response) {
      if (!response.ok) {
        const error = new Error("HTTP " + response.status + " for " + path);
        error.status = response.status;
        throw error;
      }
      return as === "json" ? response.json() : response.text();
    });
}

function setStatus(text, kind) {
  statusEl.textContent = text;
  statusEl.className = "status" + (kind ? " " + kind : "");
  statusEl.hidden = !text;
}

// Renders the signature in a shadow root, so the pane's CSS can't restyle
// it, then scales it to fit the pane: signatures are ~450px wide, the pane
// often ~320px.
function renderPreview(host, html) {
  const root = host.attachShadow({ mode: "open" });
  root.innerHTML =
    '<div style="display:inline-block;width:max-content;transform-origin:0 0;">' + html + "</div>";
  fitPreview(host);
}

function fitPreview(host) {
  const signature = host.shadowRoot && host.shadowRoot.firstElementChild;
  if (!signature || !host.clientWidth) {
    return;
  }
  signature.style.transform = "none";
  const scale = Math.min(1, host.clientWidth / signature.offsetWidth);
  signature.style.transform = "scale(" + scale + ")";
  host.style.height = Math.ceil(signature.offsetHeight * scale) + "px";
}

function insertSignature(entry, html, card) {
  const cards = listEl.querySelectorAll(".card");
  cards.forEach(function (c) {
    c.classList.remove("inserted");
    c.disabled = true;
  });
  setStatus(t.inserting);

  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      cards.forEach(function (c) {
        c.disabled = !c.dataset.html;
      });
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        card.classList.add("inserted");
        setStatus(t.inserted(entry.label), "ok");
      } else {
        console.error("[Edelburg Signature] setSignatureAsync failed:", result.error);
        setStatus(t.insertFailed(result.error.message), "error");
      }
    }
  );
}

function buildCard(entry, folder) {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "card";
  card.disabled = true;

  const head = document.createElement("span");
  head.className = "card-head";
  const label = document.createElement("span");
  label.className = "label";
  label.textContent = entry.label;
  head.appendChild(label);
  if (entry.default && autoInserts) {
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = t.standard;
    head.appendChild(badge);
  }
  const done = document.createElement("span");
  done.className = "done";
  done.textContent = t.done;
  head.appendChild(done);
  card.appendChild(head);

  const preview = document.createElement("span");
  preview.className = "preview";
  card.appendChild(preview);
  listEl.appendChild(card);

  fetchHosted(folder + "/" + encodeURIComponent(entry.file), "text")
    .then(function (html) {
      card.dataset.html = "1";
      renderPreview(preview, html);
      card.disabled = false;
      card.addEventListener("click", function () {
        insertSignature(entry, html, card);
      });
    })
    .catch(function (error) {
      console.error("[Edelburg Signature] could not load", entry.file, error);
      preview.className = "preview-missing";
      preview.textContent = t.unavailable;
    });
}

function init() {
  const email = Office.context.mailbox.userProfile.emailAddress;
  const path = getMailboxPath(email);
  if (!path) {
    setStatus(t.noMailbox, "error");
    return;
  }
  const folder = path.split("/")[0];

  fetchHosted(path + ".json", "json")
    .then(function (index) {
      const entries = (index && index.signatures) || [];
      if (entries.length === 0) {
        throw Object.assign(new Error("empty signature list"), { status: 404 });
      }
      setStatus("");
      entries.forEach(function (entry) {
        buildCard(entry, folder);
      });
      if (entries.length === 1) {
        setStatus(t.onlyOne);
      }
    })
    .catch(function (error) {
      console.error("[Edelburg Signature] could not load signature list for", path, error);
      setStatus(error.status === 404 ? t.none(email) : t.loadFailed, "error");
    });

  // The pane can be resized; keep previews fitted to its width.
  window.addEventListener("resize", function () {
    listEl.querySelectorAll(".preview").forEach(fitPreview);
  });
}

document.documentElement.lang = lang;
document.title = t.pageTitle;
document.getElementById("heading").textContent = t.heading;
document.getElementById("hint").textContent = t.hint;
setStatus(t.loading);

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});
