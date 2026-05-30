#!/usr/bin/env node
"use strict";

/**
 * Generates Power BI report PDFs via Puppeteer (headless Chrome).
 * Bypasses ExportTo API limitations (unsupported visuals, slicer interactions).
 *
 * Input (JSON to --config file or stdin):
 * {
 *   "workspaceId": "...",
 *   "reportId": "...",
 *   "canvasWidth": 1280,   // optional, default 1280
 *   "canvasHeight": 720,   // optional, default 720
 *   "slicers": [           // date slicers applied to all jobs
 *     { "table": "...", "column": "...", "value": "..." }
 *   ],
 *   "jobs": [
 *     {
 *       "exportKey": "operation.regional.regional_2",
 *       "pageName": "636418bdf7359888f0ea",
 *       "outputFile": "/tmp/op_reg2.pdf",
 *       "filters": [       // value filters (regional/area/store)
 *         { "table": "Mapping_Resto_RFID", "column": "RegionOperational", "values": ["REGIONAL 2"] }
 *       ]
 *     }
 *   ]
 * }
 *
 * Output (JSON to stdout):
 * [{ "exportKey": "...", "success": true, "outputFile": "..." }, ...]
 */

const puppeteer = require("puppeteer");
const https = require("https");
const { URLSearchParams } = require("url");
const fs = require("fs");

// ---------- HTTP helpers ----------

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

const get = (url, h) => httpsRequest("GET", url, null, h);
const post = (url, b, h) => httpsRequest("POST", url, b, h);

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
  if (!data.access_token) throw new Error(`Token failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function getEmbedInfo(accessToken, workspaceId, reportId) {
  const auth = { Authorization: `Bearer ${accessToken}` };
  const base = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}`;
  const [reportData, tokenData] = await Promise.all([
    get(base, auth),
    post(`${base}/GenerateToken`, { accessLevel: "view" }, auth),
  ]);
  if (!tokenData.token) throw new Error(`GenerateToken failed: ${JSON.stringify(tokenData)}`);
  return { embedUrl: reportData.embedUrl, embedToken: tokenData.token };
}

// ---------- PDF generation ----------

async function exportJobs(embedInfo, reportId, slicers, jobs, canvasWidth, canvasHeight, renderWait) {
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
  });

  const results = [];

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: canvasWidth, height: canvasHeight });

    page.on("console", (msg) => {
      if (msg.type() !== "warning" && !msg.text().includes("Slow network")) {
        process.stderr.write(`[browser] ${msg.type()}: ${msg.text()}\n`);
      }
    });

    const embedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/powerbi-client@2.22.3/dist/powerbi.min.js"></script>
<style>*{margin:0;padding:0;overflow:hidden}body,html{width:${canvasWidth}px;height:${canvasHeight}px}</style>
</head><body>
<div id="container" style="width:${canvasWidth}px;height:${canvasHeight}px"></div>
</body></html>`;

    await page.setContent(embedHtml, { waitUntil: "domcontentloaded" });
    await page.waitForFunction(() => typeof window.powerbi !== "undefined", { timeout: 30000 });

    // Embed report and wait for loaded
    const loaded = await page.evaluate(
      async (embedUrl, embedToken, reportId) => {
        const models = window["powerbi-client"].models;
        const report = window.powerbi.embed(document.getElementById("container"), {
          type: "report",
          id: reportId,
          embedUrl,
          accessToken: embedToken,
          tokenType: models.TokenType.Embed,
          settings: { filterPaneEnabled: false, navContentPaneEnabled: false },
        });
        window._report = report;
        window._rendered = false;
        report.on("rendered", () => { window._rendered = true; });

        return new Promise((resolve, reject) => {
          const t = setTimeout(() => reject("report load timeout"), 120000);
          report.on("error", (e) => { clearTimeout(t); reject(JSON.stringify(e.detail)); });
          report.on("loaded", () => { clearTimeout(t); resolve(true); });
        });
      },
      embedInfo.embedUrl, embedInfo.embedToken, reportId
    );

    if (!loaded) throw new Error("Report failed to load");
    process.stderr.write("Report loaded\n");

    // Zoom the report so its native page width fills the container width exactly.
    // Power BI defaults to "Fit to Page" which leaves side whitespace when the
    // report's aspect ratio doesn't match the container's. setZoom overrides this.
    const fitResult = await page.evaluate(async (containerWidth) => {
      try {
        const pages = await window._report.getPages();
        const active = pages.find(p => p.isActive) || pages[0];
        const ds = active && active.defaultSize;
        // Widescreen (type 0) = 1280×720, Standard (type 1) = 960×720; custom has explicit w/h
        const nativeW = (ds && ds.width) || (ds && ds.type === 0 ? 1280 : ds && ds.type === 1 ? 960 : null);
        const nativeH = (ds && ds.height) || (ds && ds.type === 0 ? 720 : ds && ds.type === 1 ? 720 : null);
        if (!nativeW) return { zoom: 1, ds };
        const zoom = containerWidth / nativeW;
        await window._report.setZoom(zoom);
        return { zoom, nativeW, nativeH, scaledH: nativeH ? Math.ceil(nativeH * zoom) : null, ds };
      } catch (e) {
        return { zoom: 1, error: e.message };
      }
    }, canvasWidth);
    process.stderr.write(`Page fit: ${JSON.stringify(fitResult)}\n`);
    if (fitResult.zoom !== 1) await new Promise((r) => setTimeout(r, 2000));

    // Collect per-page heights — each page has its own defaultSize.height.
    // We need this to set the correct viewport+PDF height per job; using a
    // single height for all pages causes Power BI to vertically center shorter
    // pages, adding whitespace at the top.
    const allPages = await page.evaluate(() =>
      window._report.getPages().then(ps => ps.map(p => ({
        name: p.name, displayName: p.displayName, defaultSize: p.defaultSize,
      })))
    );
    const pageHeights = {};
    for (const p of allPages) {
      const h = p.defaultSize && p.defaultSize.height ? p.defaultSize.height : null;
      if (h) {
        pageHeights[p.name] = Math.ceil(h * fitResult.zoom);
        pageHeights[p.displayName] = Math.ceil(h * fitResult.zoom);
      }
    }

    // Expand viewport to the tallest page upfront so mid-session we only shrink
    const maxH = Math.max(canvasHeight, ...Object.values(pageHeights));
    await page.setViewport({ width: canvasWidth, height: maxH });
    await page.evaluate((h) => {
      document.documentElement.style.height = h + "px";
      document.body.style.height = h + "px";
      const container = document.getElementById("container");
      if (container) container.style.height = h + "px";
    }, maxH);
    await new Promise((r) => setTimeout(r, 3000));
    process.stderr.write(`Viewport set to ${canvasWidth}×${maxH}\n`);

    let currentViewportH = maxH;

    // Set date slicers on all pages
    if (slicers.length > 0) {
      for (const p of allPages) {
        await page.evaluate(async (pageName) => window._report.setPage(pageName), p.name);
        await new Promise((r) => setTimeout(r, 1200));

        for (const cfg of slicers) {
          await page.evaluate(async (pageName, cfg) => {
            const pages = await window._report.getPages();
            const pg = pages.find(x => x.name === pageName);
            if (!pg) return;
            const visuals = await pg.getVisuals();
            for (const v of visuals.filter(x => x.type === "slicer")) {
              try {
                const st = await v.getSlicerState();
                if ((st.targets || []).some(t => t.table === cfg.table && t.column === cfg.column)) {
                  await v.setSlicerState({ filters: [{ $schema: "http://powerbi.com/product/schema#basic", target: { table: cfg.table, column: cfg.column }, operator: "In", values: [cfg.value], filterType: 1 }] });
                }
              } catch (_) {}
            }
          }, p.name, cfg);
        }
        process.stderr.write(`Slicers set on ${p.displayName}\n`);
      }
    }

    // Process each job
    for (const job of jobs) {
      try {
        process.stderr.write(`Generating ${job.exportKey}...\n`);

        // Apply value filters
        await page.evaluate(async (filters) => {
          const filterObjs = filters.map(f => ({
            $schema: "http://powerbi.com/product/schema#basic",
            target: { table: f.table, column: f.column },
            operator: "In",
            values: f.values,
            filterType: 1,
          }));
          await window._report.setFilters(filterObjs);
        }, job.filters || []);

        // Navigate to the export page. If already on the target page, navigate
        // away first — setActive() on the current page does not fire a rendered
        // event, so we'd fall through to the 30s timeout before continuing.
        await page.evaluate(async (pageName) => {
          const pages = await window._report.getPages();
          const target = pages.find(p => p.name === pageName || p.displayName === pageName);
          if (!target) throw new Error(`Page not found: ${pageName}`);
          const current = pages.find(p => p.isActive);
          if (current && current.name === target.name) {
            const other = pages.find(p => p.name !== target.name);
            if (other) await other.setActive();
            await new Promise((r) => setTimeout(r, 800));
          }
          window._rendered = false;
          await target.setActive();
        }, job.pageName);

        // Wait for rendered event or timeout
        await page.waitForFunction(() => window._rendered === true, { timeout: 30000 })
          .catch(() => {});

        // Debounce-wait: hold until the rendered event has not fired for renderWait ms.
        // Data-bound visuals trigger a second rendered event after their queries finish,
        // so a fixed wait risks capturing charts that are still loading. This approach
        // resets the timer each time rendered fires, resolving only when rendering is stable.
        await page.evaluate(async (debounceMs) => {
          await new Promise((resolve) => {
            let timer = setTimeout(resolve, debounceMs);
            const bump = () => { clearTimeout(timer); timer = setTimeout(resolve, debounceMs); };
            window._report.on("rendered", bump);
            // Hard cap at 3× debounce so a perpetually re-rendering report doesn't block forever
            setTimeout(resolve, debounceMs * 3);
          });
        }, renderWait);

        // Resize viewport+container to this page's exact height so Power BI
        // fills the container perfectly — a shorter page (e.g. Store = 3700px)
        // in a taller container (e.g. 4280px from Area) gets vertically centred,
        // creating whitespace at the top.
        const jobPdfH = pageHeights[job.pageName] || maxH;
        if (jobPdfH !== currentViewportH) {
          await page.setViewport({ width: canvasWidth, height: jobPdfH });
          await page.evaluate((h) => {
            document.documentElement.style.height = h + "px";
            document.body.style.height = h + "px";
            const container = document.getElementById("container");
            if (container) container.style.height = h + "px";
          }, jobPdfH);
          await new Promise((r) => setTimeout(r, 1500));
          currentViewportH = jobPdfH;
        }

        // Generate PDF
        const pdfBuffer = await page.pdf({
          width: `${canvasWidth}px`,
          height: `${jobPdfH}px`,
          printBackground: true,
          margin: { top: "0px", right: "0px", bottom: "0px", left: "0px" },
        });

        fs.writeFileSync(job.outputFile, pdfBuffer);
        process.stderr.write(`Done ${job.exportKey}: ${pdfBuffer.length} bytes\n`);
        const ok = { exportKey: job.exportKey, success: true, outputFile: job.outputFile };
        results.push(ok);
        process.stdout.write(JSON.stringify(ok) + "\n");
      } catch (err) {
        const msg = err && err.message ? err.message : JSON.stringify(err);
        process.stderr.write(`Failed ${job.exportKey}: ${msg}\n`);
        const fail = { exportKey: job.exportKey, success: false, error: msg };
        results.push(fail);
        process.stdout.write(JSON.stringify(fail) + "\n");
      }
    }
  } finally {
    await browser.close();
  }

  return results;
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
    throw new Error("Missing POWERBI_TENANT_ID, POWERBI_CLIENT_ID, POWERBI_CLIENT_SECRET env vars");

  const accessToken = await getAccessToken(tenantId, clientId, clientSecret);
  const embedInfo = await getEmbedInfo(accessToken, config.workspaceId, config.reportId);
  process.stderr.write(`Embed token acquired for ${config.reportId}\n`);

  const results = await exportJobs(
    embedInfo,
    config.reportId,
    config.slicers || [],
    config.jobs,
    config.canvasWidth || 1280,
    config.canvasHeight || 720,
    config.renderWait != null ? config.renderWait : 4000
  );

  // Results already streamed to stdout per-job; nothing left to write.
}

main().catch((err) => {
  process.stderr.write(`Fatal: ${err}\n`);
  process.exit(1);
});
