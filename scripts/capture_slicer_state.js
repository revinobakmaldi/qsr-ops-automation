#!/usr/bin/env node
"use strict";

/**
 * Captures Power BI slicer states as a bookmark state string for use in ExportTo API.
 *
 * Reads a JSON config from --config <file> or stdin.
 * Writes JSON { reportKey: stateString, ... } to stdout.
 * Progress/errors go to stderr.
 *
 * Config format:
 * {
 *   "reports": [{
 *     "key": "operation",
 *     "workspaceId": "...",
 *     "reportId": "...",
 *     "slicers": [
 *       { "name": "Daily Slicer", "table": "Mapping_Date_RFID", "column": "DateOfBusiness", "value": "2026-05-25T00:00:00" },
 *       { "name": "Weekly Slicer", "table": "Map_Week_Only", "column": "WeekMonthYear", "value": "Week 4 May 2026" },
 *       { "name": "Monthly Slicer", "table": "Mapping_Date_RFID", "column": "MonthMap", "value": "2026-05-01T00:00:00" }
 *     ]
 *   }]
 * }
 */

const puppeteer = require("puppeteer");
const https = require("https");
const { URLSearchParams } = require("url");
const fs = require("fs");

// ---------- HTTP helpers (no axios dependency) ----------

function httpsRequest(method, url, body, headers = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const isJson = body && typeof body === "object";
    const payload = body ? (isJson ? JSON.stringify(body) : body) : null;
    const opts = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method,
      headers: {
        ...(payload ? {
          "Content-Type": isJson ? "application/json" : "application/x-www-form-urlencoded",
          "Content-Length": Buffer.byteLength(payload),
        } : {}),
        ...headers,
      },
    };
    const req = https.request(opts, (res) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => {
        const text = Buffer.concat(chunks).toString();
        try { resolve(JSON.parse(text)); } catch { resolve(text); }
      });
    });
    req.on("error", reject);
    if (payload) req.write(payload);
    req.end();
  });
}

const get = (url, headers) => httpsRequest("GET", url, null, headers);
const post = (url, body, headers) => httpsRequest("POST", url, body, headers);

// ---------- Azure AD + Power BI auth ----------

async function getAccessToken(tenantId, clientId, clientSecret) {
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://analysis.windows.net/powerbi/api/.default",
  });
  const data = await post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    params.toString()
  );
  if (!data.access_token)
    throw new Error(`Token request failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function getEmbedInfo(accessToken, workspaceId, reportId) {
  const auth = { Authorization: `Bearer ${accessToken}` };
  const base = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}`;
  const [reportData, tokenData] = await Promise.all([
    get(base, auth),
    post(`${base}/GenerateToken`, { accessLevel: "view" }, auth),
  ]);
  if (!tokenData.token)
    throw new Error(`GenerateToken failed: ${JSON.stringify(tokenData)}`);
  return { embedUrl: reportData.embedUrl, embedToken: tokenData.token };
}

// ---------- Puppeteer capture ----------

async function captureReportState(embedUrl, embedToken, reportId, slicers) {
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    // Forward browser console to stderr for debugging
    page.on("console", (msg) =>
      process.stderr.write(`[browser] ${msg.type()}: ${msg.text()}\n`)
    );
    page.on("pageerror", (err) =>
      process.stderr.write(`[browser error] ${err}\n`)
    );

    await page.setContent(
      `<!DOCTYPE html><html><head><meta charset="utf-8">
       <script src="https://cdn.jsdelivr.net/npm/powerbi-client@2.22.3/dist/powerbi.min.js"></script>
       </head><body>
       <div id="container" style="width:1280px;height:720px"></div>
       </body></html>`,
      { waitUntil: "domcontentloaded" }
    );

    // Wait for powerbi-client to finish loading
    await page.waitForFunction(() => typeof window.powerbi !== "undefined", { timeout: 30000 });
    process.stderr.write("powerbi-client loaded\n");

    const state = await page.evaluate(
      async (embedUrl, embedToken, reportId, slicers) => {
        const models = window["powerbi-client"].models;
        const report = window.powerbi.embed(
          document.getElementById("container"),
          {
            type: "report",
            id: reportId,
            embedUrl,
            accessToken: embedToken,
            tokenType: models.TokenType.Embed,
            settings: { filterPaneEnabled: false, navContentPaneEnabled: false },
          }
        );

        return new Promise((resolve, reject) => {
          const timer = setTimeout(
            () => reject("Timeout: report did not load within 120s"),
            120000
          );

          report.on("error", (e) => {
            clearTimeout(timer);
            reject(`Report error: ${JSON.stringify(e.detail)}`);
          });

          report.on("loaded", async () => {
            clearTimeout(timer);
            try {
              const pages = await report.getPages();

              // Track which slicer configs were successfully applied
              const applied = new Set();

              // Navigate to each page so its slicers become accessible
              for (const p of pages) {
                await report.setPage(p.name);
                await new Promise((r) => setTimeout(r, 1500));

                const visuals = await p.getVisuals();
                const slicerVisuals = visuals.filter((v) => v.type === "slicer");

                for (const cfg of slicers) {
                  for (const visual of slicerVisuals) {
                    try {
                      const slicerState = await visual.getSlicerState();
                      const targets = slicerState.targets || [];
                      const isMatch = targets.some(
                        (t) => t.table === cfg.table && t.column === cfg.column
                      );
                      if (!isMatch) continue;

                      await visual.setSlicerState({
                        filters: [
                          {
                            $schema: "http://powerbi.com/product/schema#basic",
                            target: { table: cfg.table, column: cfg.column },
                            operator: "In",
                            values: [cfg.value],
                            filterType: 1,
                          },
                        ],
                      });
                      applied.add(`${cfg.table}/${cfg.column}`);
                      console.log(`set ${cfg.table}/${cfg.column} = ${cfg.value} on ${p.displayName}`);
                    } catch (_) {}
                  }
                }
              }

              // Warn about any configs that were never matched
              for (const cfg of slicers) {
                const key = `${cfg.table}/${cfg.column}`;
                if (!applied.has(key)) {
                  console.log(`WARNING: no slicer found for ${key}`);
                }
              }

              // Wait for state to settle then capture
              await new Promise((r) => setTimeout(r, 3000));
              const bookmark = await report.bookmarksManager.capture({ allPages: true });
              resolve(bookmark.state);
            } catch (err) {
              reject(err.message || String(err));
            }
          });
        });
      },
      embedUrl,
      embedToken,
      reportId,
      slicers
    );

    return state;
  } finally {
    await browser.close();
  }
}

// ---------- Main ----------

async function main() {
  const argv = process.argv.slice(2);

  let config;
  const cfgIdx = argv.indexOf("--config");
  if (cfgIdx !== -1) {
    config = JSON.parse(fs.readFileSync(argv[cfgIdx + 1], "utf-8"));
  } else {
    const chunks = [];
    for await (const chunk of process.stdin) chunks.push(chunk);
    config = JSON.parse(Buffer.concat(chunks).toString());
  }

  const tenantId = process.env.POWERBI_TENANT_ID || process.env.AZURE_TENANT_ID;
  const clientId = process.env.POWERBI_CLIENT_ID || process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.POWERBI_CLIENT_SECRET || process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret)
    throw new Error(
      "Missing env vars: POWERBI_TENANT_ID, POWERBI_CLIENT_ID, POWERBI_CLIENT_SECRET"
    );

  const accessToken = await getAccessToken(tenantId, clientId, clientSecret);
  const results = {};

  for (const report of config.reports) {
    process.stderr.write(`Capturing state for ${report.key}...\n`);
    const embedInfo = await getEmbedInfo(accessToken, report.workspaceId, report.reportId);
    const state = await captureReportState(
      embedInfo.embedUrl,
      embedInfo.embedToken,
      report.reportId,
      report.slicers
    );
    results[report.key] = state;
    process.stderr.write(`Done: ${report.key} (state length: ${state ? state.length : 0})\n`);
  }

  process.stdout.write(JSON.stringify(results) + "\n");
}

main().catch((err) => {
  process.stderr.write(`Fatal: ${err}\n`);
  process.exit(1);
});
