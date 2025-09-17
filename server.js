// server.js
import express from "express";
import dotenv from "dotenv";
import axios from "axios";
import bodyParser from "body-parser";

dotenv.config();
const app = express();
app.use(bodyParser.json());

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SITE_ID,     // z.B. rottachblechverarbeitung.sharepoint.com:/sites/Intranet-Rottach-Werke  (oder die lange GUID-Form)
  DRIVE_ID,    // ID der Dokumentbibliothek
  ONESIGNAL_APP_ID,
  ONESIGNAL_REST_KEY,
  PUBLIC_SHARED_KEY, // optional: kleiner Schutz-Token für /public/* Endpoints
  PORT = 3000
} = process.env;

// ---- Helper: App-Token holen (client_credentials) ----
let _token = null;
let _exp = 0;
async function getAppToken() {
  const now = Math.floor(Date.now() / 1000);
  if (_token && _exp - now > 60) return _token;

  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  _token = data.access_token;
  _exp = now + data.expires_in;
  return _token;
}

// ---- Optional: ganz einfacher "shared key"-Check ----
function checkSharedKey(req, res) {
  if (!PUBLIC_SHARED_KEY) return true;
  if (req.query.key === PUBLIC_SHARED_KEY) return true;
  res.status(403).send("forbidden");
  return false;
}

// ---- Öffentliche Liste: Root der Bibliothek ----
app.get("/public/list", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  try {
    const token = await getAppToken();
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
    console.error(e?.response?.data || e.message);
    res.status(500).send("error listing files");
  }
});

// ---- Download-Proxy: streamed file to client ----
app.get("/public/download", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  const id = req.query.id;
  if (!id) return res.status(400).send("missing id");
  try {
    const token = await getAppToken();
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
    console.error(e?.response?.data || e.message);
    res.status(500).send("error downloading file");
  }
});

// ---- Site Pages listen (Titel, Name, URL) ----
app.get("/public/pages", async (req, res) => {
  if (!checkSharedKey(req, res)) return;
  try {
    const token = await getAppToken();
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
    console.error(e?.response?.data || e.message);
    res.status(500).send("error listing pages");
  }
});

// ---- Graph Webhook (Validation + Notifications) ----
app.all("/graph-webhook", async (req, res) => {
  // Validation: Graph sends ?validationToken=...
  const validationToken = req.query.validationToken || req.body?.validationToken;
  if (validationToken) {
    res.setHeader("Content-Type", "text/plain");
    return res.status(200).send(validationToken);
  }

  const notifications = req.body?.value;
  if (!notifications) return res.status(400).send("no notifications");

  // Send a simple broadcast via OneSignal
  for (const n of notifications) {
    try {
      await axios.post(
        "https://onesignal.com/api/v1/notifications",
        {
          app_id: ONESIGNAL_APP_ID,
          headings: { en: "Neue Info im Intranet" },
          contents: {
            en: `Änderung: ${n.changeType || "update"} in ${n.resource || "SharePoint"}`,
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

// ---- Health ----
app.get("/", (_req, res) => res.send("intranet-backend up"));

app.listen(PORT, () => console.log(`listening on :${PORT}`));

