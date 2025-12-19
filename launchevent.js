
// launchevent.js

// Associate function names in manifest with actual handlers
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);

const STYLE_ID = "comicsans-default-font";
const WRAPPER_ID = "cs-wrapper";
const CSS = `
  /* Applied once per item */
  <style id="${STYLE_ID}">
    body, div, p, span, td, li, a, table {
      font-family: "Comic Sans MS","Comic Sans",cursive !important;
    }
  </style>
`;

function onNewMessageCompose(event) {
  applyComicSans()
    .catch((e) => {
      // Optional: add a notification message here if you want
      console.warn("Comic Sans add-in:", e?.message || e);
    })
    .finally(() => event.completed());
}

async function applyComicSans() {
  await Office.onReady();

  const item = Office.context.mailbox.item;

  // Utility: promisify Office async APIs
  const getType = () =>
    new Promise((resolve, reject) => {
      if (!item.body || !item.body.getTypeAsync) { resolve(Office.CoercionType.Html); return; }
      item.body.getTypeAsync((r) =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)
      );
    });

  const getBody = (coercionType) =>
    new Promise((resolve, reject) => {
      item.body.getAsync(coercionType, (r) =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || "") : reject(r.error)
      );
    });

  const setBody = (html) =>
    new Promise((resolve, reject) => {
      item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (r) =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
      );
    });

  const encodeHtml = (t) =>
    (t || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  const type = await getType();

  if (type === Office.CoercionType.Html) {
    let html = await getBody(Office.CoercionType.Html);

    // Avoid duplicate injection if user opens multiple compose windows
    if (!new RegExp(`id=["']${STYLE_ID}["']`).test(html)) {
      const wrapped = `${CSS}<div id="${WRAPPER_ID}" style="font-family:'Comic Sans MS','Comic Sans',cursive">${html}</div>`;
      await setBody(wrapped);
    }
  } else {
    // Plain text compose → convert to HTML so we can apply font
    const text = await getBody(Office.CoercionType.Text);
    const htmlFromText = encodeHtml(text).replace(/\r?\n/g, "<br>");
    const html = `${CSS}<div id="${WRAPPER_ID}" style="font-family:'Comic Sans MS','Comic Sans',cursive">${htmlFromText}</div>`;
    await setBody(html);
  }
}
``
