// server.js
import express from "express";
import dotenv from "dotenv";
import axios from "axios";
import bodyParser from "body-parser";
import cors from "cors";
// ...
app.use(cors({ origin: "*"}));

dotenv.config();
const app = express();
app.use(bodyParser.json());

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,

  // Graph: lange Form der Site-ID ODER Pfad-Form wie
  // "rottachblechverarbeitung.sharepoint.com:/sites/Intranet-Rottach-Werke"
  SITE_ID,

  // Drive der Dokumentbibliothek (optional, für /public/list & /public/download)
  DRIVE_ID,

  // Für SharePoint-REST (Seiten-Inhalt):
  // z.B. SP_HOST="rottachblechverarbeitung.sharepoint.com"
  // und SP_SITE_PATH="/sites/Intranet-Rottach-Werke"
  SP_HOST,
  SP_SITE_PATH,

  // Push
  ONESIGNAL_APP_ID,
  ONESIGNAL_REST_KEY,

  // einfacher Schutz für /public/* Endpunkte
  PUBLIC_SHARED_KEY,

  PORT = 3000,
} = process.env;

// ---------------- Token-Helper ----------------
let _graphToken = null;
let _graphExp = 0;
let _spToken = null;
let _spExp = 0;

// OAuth v2 client_credentials für ein gewünschtes Resource "audience"
async function getTokenFor(resourceBaseUrl, cacheObj) {
  const now = Math.floor(Date.now() / 1000);
  if (cacheObj.token && cacheObj.exp - now > 60) return cacheObj.token;

  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: `${resourceBaseUrl}/.default`,
    grant_type: "client_credentials",
  });

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  cacheObj.token = data.access_token;
  cacheObj.exp = now + data.expires_in;
  return cacheObj.token;
}

async function getGraphToken() {
  return getTokenFor("https://graph.microsoft.com", {
    token: _graphToken,
    exp: _graphExp,
    set token(v) { _graphToken = v; },
    get token() { return _graphToken; },
    set exp(v) { _graphExp = v; },
    get exp() { return _graphExp; },
  });
}

async function getSharePointToken() {
  return getTokenFor(`https://${SP_HOST}`, {
    token: _spToken,
    exp: _spExp,
    set token(v) { _spToken = v; },
    get token() { return _spToken; },
    set exp(v) { _spExp = v; },
    get exp() { return _spExp; },
  });
}

// ---------------- Utils ----------------
function checkSharedKey(req, res) {
  if (!PUBLIC_SHARED_KEY) return true;
  if (req.query.key === PUBLIC_SHARED_KEY) return true;
  res.status(403).send("forbidden");
  return false;
}

function errOut(res, e, fallbackMsg) {
  const payload = e?.response?.data || { error: e.message || String(e) };
  console.error(payload);
  res.status(e?.response?.status || 500).json(payload ?? { error: fallbackMsg });
}

// ---------------- Files (optional) ----------------
app.get("/public/list", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  try {
    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${DRIVE_ID}/root/children?$top=200`;
    const { data } = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const simplified = (data.value || []).map((it) => ({
      id: it.id,
      name: it.name,
      isFolder: !!it.folder,
      size: it.size ?? 0,
      mimeType: it.file?.mimeType || null,
      lastModified: it.lastModifiedDateTime,
    }));
    res.json(simplified);
  } catch (e) {
    errOut(res, e, "error listing files");
  }
});

app.get("/public/download", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  const id = req.query.id;
  if (!id) return res.status(400).send("missing id");
  try {
    const token = await getGraphToken();
    const metaUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${id}?select=@microsoft.graph.downloadUrl,name,file`;
    const { data: meta } = await axios.get(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const dl = meta["@microsoft.graph.downloadUrl"];
    if (!dl) return res.status(404).send("no download url");

    const upstream = await axios.get(dl, { responseType: "stream" });
    res.setHeader("Content-Disposition", `inline; filename="${meta.name}"`);
    res.setHeader(
      "Content-Type",
      upstream.headers["content-type"] || "application/octet-stream"
    );
    upstream.data.pipe(res);
  } catch (e) {
    errOut(res, e, "error downloading file");
  }
});

// ---------------- Pages (Liste) via Graph ----------------
app.get("/public/pages", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  try {
    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/pages`;
    const { data } = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const pages = (data.value || []).map((p) => ({
      id: p.id,
      title: p.title,
      name: p.name,
      pageLayout: p.pageLayout,
      webUrl: p.webUrl,
    }));
    res.json(pages);
  } catch (e) {
    errOut(res, e, "error listing pages");
  }
});

// ---------------- Page (Inhalt) via SharePoint REST ----------------
app.get("/public/page", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  try {
    if (!SP_HOST || !SP_SITE_PATH) {
      return res.status(400).json({
        error:
          "SP_HOST and SP_SITE_PATH env vars are required, e.g. rottachblechverarbeitung.sharepoint.com and /sites/Intranet-Rottach-Werke",
      });
    }

    const name = req.query.name?.toString();
    const webUrl = req.query.webUrl?.toString();

    let serverRelativeUrl;
    if (name) {
      serverRelativeUrl = `${SP_SITE_PATH}/SitePages/${name}`;
    } else if (webUrl) {
      try {
        const u = new URL(webUrl);
        serverRelativeUrl = u.pathname; // beginnt mit /sites/...
      } catch {
        return res.status(400).json({ error: "invalid webUrl" });
      }
    } else {
      return res
        .status(400)
        .json({ error: "missing 'name' or 'webUrl' query param" });
    }

    const spToken = await getSharePointToken();
    const url = `https://${SP_HOST}${SP_SITE_PATH}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/$value`;
    const { data: html } = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${spToken}`,
        Accept: "text/html",
      },
      responseType: "text",
    });

    res.json({
      serverRelativeUrl,
      webUrl: `https://${SP_HOST}${serverRelativeUrl}`,
      html,
    });
  } catch (e) {
    errOut(res, e, "error loading page");
  }
});

// ---------------- Graph Webhook (Push) ----------------
app.all("/graph-webhook", async (req, res) => {
  const validationToken = req.query.validationToken || req.body?.validationToken;
  if (validationToken) {
    res.setHeader("Content-Type", "text/plain");
    return res.status(200).send(validationToken);
  }

  const notifications = req.body?.value;
  if (!notifications) return res.status(400).send("no notifications");

  for (const n of notifications) {
    try {
      await axios.post(
        "https://onesignal.com/api/v1/notifications",
        {
          app_id: ONESIGNAL_APP_ID,
          headings: { en: "Neue Info im Intranet" },
          contents: {
            en: `Änderung: ${n.changeType || "update"} in ${
              n.resource || "SharePoint"
            }`,
          },
          included_segments: ["Subscribed Users"],
        },
        {
          headers: {
            "Content-Type": "application/json; charset=utf-8",
            Authorization: `Basic ${ONESIGNAL_REST_KEY}`,
          },
        }
      );
    } catch (e) {
      console.error("OneSignal error", e?.response?.data || e.message);
    }
  }
  res.status(202).send("ok");
});

// ---------------- Health ----------------
app.get("/", (_req, res) => res.send("intranet-backend up"));

app.listen(PORT, () => console.log(`listening on :${PORT}`));
