/**
 * SignatureForge - Outlook event handler
 *
 * Pipeline:
 *   1. Disable the user's client-side signature so we are the single source of truth.
 *   2. Read displayName and emailAddress from Office.context.mailbox.userProfile.
 *   3. Read title and phone from roaming settings.
 *   4. Fetch the git-managed default.html template.
 *   5. Substitute the fields and inject the result with setSignatureAsync.
 *
 * If anything fails, the event completes silently so sending mail is never blocked.
 */

const CONFIG_BASE_URL = "https://dchamorroarias2lsg.github.io/signatureforge";
const CACHE_TTL_MS = 60 * 1000;
const SETTINGS_KEYS = {
  title: "signatureforge.title",
  phone: "signatureforge.phone",
  websiteName: "signatureforge.websiteName",
  websiteUrl: "signatureforge.websiteUrl",
};

let _cache = { template: null, fetchedAt: 0 };

Office.onReady();

async function onNewMessageComposeHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const userProfile = Office.context.mailbox.userProfile || {};
    const displayName = (userProfile.displayName || "").trim();
    const emailAddress = (userProfile.emailAddress || "").trim().toLowerCase();

    if (!displayName && !emailAddress) {
      return event.completed();
    }

    await disableClientSignatureSafe();

    const templateHtml = await loadTemplate();
    const fields = loadFields();
    const rendered = renderTemplate(templateHtml, {
      displayName,
      firstName: getFirstName(displayName),
      email: emailAddress,
      title: fields.title,
      phone: fields.phone,
      websiteBlock: buildWebsiteBlock(fields.websiteName, fields.websiteUrl),
    });

    item.body.setSignatureAsync(
      rendered,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("SignatureForge setSignature failed:", result.error);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error("SignatureForge handler error:", err);
    event.completed();
  }
}

function disableClientSignatureSafe() {
  return new Promise((resolve) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item.disableClientSignatureAsync) {
        return resolve();
      }
      item.disableClientSignatureAsync(() => resolve());
    } catch (_) {
      resolve();
    }
  });
}

function loadFields() {
  const roaming = Office.context.roamingSettings;
  return {
    title: String(roaming.get(SETTINGS_KEYS.title) ?? ""),
    phone: String(roaming.get(SETTINGS_KEYS.phone) ?? ""),
    websiteName: String(roaming.get(SETTINGS_KEYS.websiteName) ?? ""),
    websiteUrl: String(roaming.get(SETTINGS_KEYS.websiteUrl) ?? ""),
  };
}

function getFirstName(displayName) {
  const name = String(displayName ?? "").trim();
  if (!name) return "";
  return name.split(/\s+/)[0];
}

function normalizeWebsiteUrl(url) {
  const value = String(url ?? "").trim();
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `https://${value}`;
}

function getWebsiteLabel(name, url) {
  const label = String(name ?? "").trim();
  if (label) return label;
  const normalized = normalizeWebsiteUrl(url);
  if (!normalized) return "";
  try {
    return new URL(normalized).hostname.replace(/^www\./i, "");
  } catch (_) {
    return normalized.replace(/^https?:\/\//i, "");
  }
}

function buildWebsiteBlock(name, url) {
  const label = getWebsiteLabel(name, url);
  if (!label) return "";
  const normalizedUrl = normalizeWebsiteUrl(url);
  if (!normalizedUrl) {
    return `
      <br>
      <span style="color:#252831;">
        ${escapeHtml(label)}
      </span>
    `;
  }
  return `
      <br>
      <a href="${escapeHtml(normalizedUrl)}"
         style="text-decoration:underline;">
        ${escapeHtml(label)}
      </a>
  `;
}

async function loadTemplate() {
  if (_cache.template && Date.now() - _cache.fetchedAt < CACHE_TTL_MS) {
    return _cache.template;
  }
  const url = `${CONFIG_BASE_URL}/default.html?t=${Date.now()}`;
  const resp = await fetch(url, { cache: "no-store" });
  if (!resp.ok) throw new Error(`default.html fetch failed: ${resp.status}`);
  const html = await resp.text();
  _cache.template = html;
  _cache.fetchedAt = Date.now();
  return html;
}

function renderTemplate(html, fields) {
  html = html.replace(/\{\{\{\s*([\w.-]+)\s*\}\}\}/g, (_, key) => {
    return fields[key] != null ? String(fields[key]) : "";
  });
  html = html.replace(/\{\{\s*([\w.-]+)\s*\}\}/g, (_, key) => {
    return fields[key] != null ? escapeHtml(String(fields[key])) : "";
  });
  return html;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
