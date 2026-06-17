// ==UserScript==
// @name         WikiData QIDS Capture Panel v8
// @namespace    myNamespace
// @version      2.8.0
// @match        https://query.wikidata.org/*
// @grant        GM_xmlhttpRequest
// @grant        GM_download
// @grant        GM_getResourceText
// @resource     PANEL_HTML https://raw.githubusercontent.com/a-mccutchan87/wdqs-tampermonkey-panel/main/wdqs-panel.html
// @resource     PANEL_CSS  https://raw.githubusercontent.com/a-mccutchan87/wdqs-tampermonkey-panel/main/wdqs-panel.css
// @require https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

/* ---------------------------
   SECTION 1: Config
---------------------------- */
const CFG = {
    // panel + storage
    panelId: "wdqs-capture-panel",
    storageKey: "wdqs_capture_store_v3",
    uiKey: "wdqs_capture_ui_v3",

    // UI behavior
    maxStatusMs: 3500,
    previewMaxChars: 9000,

    // WDQS scan safety (if you later add generic QID extraction)
    maxBindingsToScan: 25000,

    // Wikidata API constraints
    wbgetentitiesBatch: 50,
    imageProp: "P18",
    descLanguage: "en",

    // export defaults
    defaultBaseName: "wdqs_capture",
    xlsx: {
        sheetNames: {
            nodes: "nodes",
            properties: "properties",
            links: "links",
            images: "images"
        }
    }
};
const DBG = {
  on: true,
  ns: "[WDQS-CAP]",
  log(...a) { if (this.on) console.log(this.ns, ...a); },
  warn(...a) { if (this.on) console.warn(this.ns, ...a); },
  err(...a) { if (this.on) console.error(this.ns, ...a); },
};
const TRACE = {
    enabled: true,                 // master switch
    prefix: "[WDQS-CAP]",

    log(...a) {
        if (!this.enabled) return;
        console.log(this.prefix, ...a);
    },
    warn(...a) {
        if (!this.enabled) return;
        console.warn(this.prefix, ...a);
    },
    error(...a) {
        if (!this.enabled) return;
        console.error(this.prefix, ...a);
    },

    group(label) {
        if (!this.enabled) return;
        console.group(this.prefix, label);
    },
    groupEnd() {
        if (!this.enabled) return;
        console.groupEnd();
    }
};
/* ---------------------------
   SECTION 2: Utilities
---------------------------- */
const nowISO = () => new Date().toISOString();

function safeJsonParse(str) {
    try { return JSON.parse(str); } catch { return null; }
}

function uniq(arr) {
    return [...new Set((arr || []).filter(Boolean))];
}

function chunk(arr, n) {
    const out = [];
    for (let i = 0; i < (arr || []).length; i += n) out.push(arr.slice(i, i + n));
    return out;
}

function csvEscape(v) {
    const s = String(v ?? "");
    const t = s.replace(/"/g, '""');
    return /[",\n\r]/.test(t) ? `"${t}"` : t;
}

function toCsv(headers, rows) {
    const out = [];
    out.push(headers.map(csvEscape).join(","));
    for (const r of rows) out.push(headers.map(h => csvEscape(r?.[h])).join(","));
    return out.join("\n");
}

function jsonCell(obj) {
    // single-line JSON for CSV/XLSX cells
    try { return JSON.stringify(obj); } catch { return ""; }
}

function gmGetText(url) {
    return new Promise((resolve) => {
        if (typeof GM_xmlhttpRequest !== "function") return resolve(null);

        GM_xmlhttpRequest({
            method: "GET",
            url,
            onload: (res) => resolve(res?.responseText || null),
            onerror: () => resolve(null),
            ontimeout: () => resolve(null),
        });
    });
}

function injectScriptText(code) {
    const s = document.createElement("script");
    s.textContent = code;
    (document.head || document.documentElement).appendChild(s);
    s.remove();
}

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch {
        try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand("copy");
            ta.remove();
            return !!ok;
        } catch {
            return false;
        }
    }
}

function clampSheetName(name) {
    return String(name || "Sheet")
        .replace(/[:\\\/\?\*\[\]]/g, " ")
        .slice(0, 31)
        .trim() || "Sheet";
}

function canExportXlsx() {
    return typeof window.XLSX !== "undefined" && !!window.XLSX?.utils;
}

function rowsToSheet(headers, rows) {
    const aoa = [headers];
    for (const r of rows) aoa.push(headers.map(h => r?.[h] ?? ""));
    return window.XLSX.utils.aoa_to_sheet(aoa);
}

/* ---------------------------
   Store helpers (read/write + migration)
---------------------------- */
function readUIState() {
    const raw = localStorage.getItem(CFG.uiKey);
    const parsed = raw ? safeJsonParse(raw) : null;
    return {
        collapsed: !!parsed?.collapsed,
        tall: !!parsed?.tall,
        height: parsed?.height || 360,
        treeSummaryOpen: !!parsed?.treeSummaryOpen
    };
}

function writeUIState(next) {
    localStorage.setItem(CFG.uiKey, JSON.stringify(next));
}

function readStore() {
    const raw = localStorage.getItem(CFG.storageKey);
    const parsed = raw ? safeJsonParse(raw) : null;

    if (parsed && typeof parsed === "object") {
        // meta
        if (!parsed.meta) {
            parsed.meta = {
                name: "",
                createdAt: nowISO(),
                updatedAt: nowISO(),
                source: "query.wikidata.org",
                version: CFG?.version || "2.5"
            };
        }
        if (!parsed.options) parsed.options = { storeFullResponses: false };

        // migrate old key "nodes" -> "targets"
        if (!Array.isArray(parsed.targets)) {
            if (Array.isArray(parsed.nodes)) {
                parsed.targets = parsed.nodes;
                delete parsed.nodes;
            } else {
                parsed.targets = [];
            }
        }

        if (!("lastPreview" in parsed)) parsed.lastPreview = null;

        // legacy keys kept for compatibility
        if (!parsed.qidIndex) parsed.qidIndex = {};
        if (!Array.isArray(parsed.captures)) parsed.captures = [];

        // images cache
        if (!parsed.images) parsed.images = { byQid: {}, byLink: {}, nextImgN: 1, updatedAt: null };
        if (!parsed.images.byQid) parsed.images.byQid = {};
        if (!parsed.images.byLink) parsed.images.byLink = {};
        if (!parsed.images.nextImgN) parsed.images.nextImgN = 1;

        // descriptions cache
        if (!parsed.descriptions) parsed.descriptions = { byQid: {}, updatedAt: null, language: CFG.descLanguage };
        if (!parsed.descriptions.byQid) parsed.descriptions.byQid = {};
        if (!parsed.descriptions.language) parsed.descriptions.language = CFG.descLanguage;

        return parsed;
    }

    // default store
    return {
        meta: {
            name: "",
            createdAt: nowISO(),
            updatedAt: nowISO(),
            source: "query.wikidata.org",
            version: "2.5"
        },
        options: { storeFullResponses: false },
        // legacy
        qidIndex: {},
        captures: [],
        // main
        targets: [],
        // UI
        lastPreview: null,
        // caches
        images: { byQid: {}, byLink: {}, nextImgN: 1, updatedAt: null },
        descriptions: { byQid: {}, updatedAt: null, language: CFG.descLanguage }
    };
}

function writeStore(store) {
    if (!store.meta) store.meta = {};
    store.meta.updatedAt = nowISO();
    localStorage.setItem(CFG.storageKey, JSON.stringify(store, null, 2));
}

/* ---------------------------
   Status + Preview helpers
---------------------------- */
function setStatus(text, kind = "info") {
    const el = document.querySelector(`#${CFG.panelId} .wdqs-cap-status`);
    if (!el) return;
    el.textContent = text;
    el.dataset.kind = kind;

    clearTimeout(setStatus._t);
    setStatus._t = setTimeout(() => {
        el.textContent = "Idle";
        el.dataset.kind = "idle";
    }, CFG.maxStatusMs);
}

function setPreview(obj) {
    const pre = document.querySelector(`#${CFG.panelId} .wdqs-cap-preview`);
    if (!pre) return;

    let s = "";
    try { s = JSON.stringify(obj, null, 2); } catch { s = String(obj); }

    if (s.length > CFG.previewMaxChars) s = s.slice(0, CFG.previewMaxChars) + "\n…(truncated)…";
    pre.textContent = s;
}

function humanCount(n) {
    return (n ?? 0).toLocaleString();
}

function setActionsDisabled(disabled) {
    const panel = document.getElementById(CFG.panelId);
    if (!panel) return;
    const els = panel.querySelectorAll("button, input, label.wdqs-btn, label.wdqs-btn-up");
    for (const el of els) {
        if (el.tagName === "LABEL") {
            el.style.pointerEvents = disabled ? "none" : "";
            el.style.opacity = disabled ? "0.6" : "";
        } else {
            el.disabled = !!disabled;
        }
    }
}

/* ---------------------------
   WDQS response helpers
---------------------------- */
function isSparqlUrl(url) {
    return typeof url === "string" && url.includes("/sparql");
}

function extractUrl(resource) {
    if (typeof resource === "string") return resource;
    if (resource && typeof resource === "object" && typeof resource.url === "string") return resource.url;
    return "";
}

function isWdqsJsonResponse(resp) {
    try {
        const ct = resp.headers?.get?.("content-type") || "";
        return ct.includes("application/sparql-results+json") || ct.includes("application/json");
    } catch {
        return false;
    }
}

/* ---------------------------
   Wikidata/Commons helpers
---------------------------- */
function commonsFilePathUrl(fileName) {
    if (!fileName) return "";
    return `https://commons.wikimedia.org/wiki/Special:FilePath/${encodeURIComponent(String(fileName).replace(/ /g, "_"))}`;
}

function isEntityUri(v) {
    return typeof v === "string" && (
        v.startsWith("http://www.wikidata.org/entity/") ||
        v.startsWith("https://www.wikidata.org/entity/")
    );
}

function qidFromEntityUri(v) {
    if (!isEntityUri(v)) return null;
    const last = (v.split("/").pop() || "").trim();
    return /^Q\d+$/.test(last) ? last : null;
}

/* Optional: generic scan if you ever want it for non-summary queries */
function extractQidsFromWdqsResults(data) {
    const qids = new Set();
    const bindings = data?.results?.bindings;
    if (!Array.isArray(bindings)) return qids;

    const limit = Math.min(bindings.length, CFG.maxBindingsToScan);
    for (let i = 0; i < limit; i++) {
        const row = bindings[i];
        if (!row || typeof row !== "object") continue;
        for (const k of Object.keys(row)) {
            const value = row[k]?.value;
            const qid = qidFromEntityUri(value);
            if (qid) qids.add(qid);
        }
    }
    return qids;
}

/* ---------------------------
   Store traversal helpers
---------------------------- */
function collectAllQidsFromStore(store) {
    const qids = new Set();
    const items = Array.isArray(store?.targets) ? store.targets : [];

    for (const s of items) {
        const sid = s?.sourceID || "";
        if (/^Q\d+$/.test(sid)) qids.add(sid);

        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                if (/^Q\d+$/.test(tid)) qids.add(tid);
            }
        }
    }

    return [...qids];
}

function refreshCounts(store) {
    const panel = document.getElementById(CFG.panelId);
    if (!panel) return;

    const total = Array.isArray(store.targets) ? store.targets.length : 0;

    const uniqQids = new Set();
    const items = Array.isArray(store.targets) ? store.targets : [];
    for (const s of items) {
        if (s?.sourceID) uniqQids.add(s.sourceID);
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) if (n?.targetID) uniqQids.add(n.targetID);
        }
    }

    panel.querySelector(".wdqs-cap-captures").textContent = humanCount(total);
    panel.querySelector(".wdqs-cap-qids").textContent = humanCount(uniqQids.size);
}

/* ---------------------------
   SECTION 3: Conversion helpers
---------------------------- */


/* ---- parsing helpers ---- */
function looksLikePropertySummary(headVars) {
    const v = headVars || [];
    return v.includes("p") && v.includes("pl_") && v.includes("count") && v.includes("ol_");
}

function propIdFromUri(uri) {
    if (typeof uri !== "string") return "";
    const m = uri.match(/\/prop\/direct\/(P\d+)$/);
    return m ? m[1] : "";
}

function splitOlLiteral(s) {
    if (typeof s !== "string") return [];
    return s.split(",").map(x => x.trim()).filter(Boolean);
}

function decodeQueryFromUrl(url) {
    try {
        const u = new URL(url, location.origin);
        return u.searchParams.get("query") || "";
    } catch {
        return "";
    }
}

function extractSubjectQidFromQueryText(q) {
    // looks for <https://www.wikidata.org/entity/Q####> in the query
    const m = String(q || "").match(/<https?:\/\/www\.wikidata\.org\/entity\/(Q\d+)>/);
    return m ? m[1] : "";
}

/* ---- Wikidata lookups for conversion ---- */
async function qidFromLabel(label, language = "en") {
    // best-effort: 1st hit from wbsearchentities
    const url =
        "https://www.wikidata.org/w/api.php" +
        `?action=wbsearchentities&search=${encodeURIComponent(label)}` +
        `&language=${encodeURIComponent(language)}` +
        "&format=json&limit=1&origin=*";

    const r = await fetch(url);
    const j = await r.json();
    const hit = j?.search?.[0];
    return hit?.id || "";
}

async function enrichNodeMeta(qid) {
    // label + enwiki sitelink for the *source* node
    const url =
        "https://www.wikidata.org/w/api.php" +
        "?action=wbgetentities" +
        `&ids=${encodeURIComponent(qid)}` +
        "&props=labels|sitelinks" +
        "&languages=en" +
        "&sitefilter=enwiki" +
        "&format=json&origin=*";

    const r = await fetch(url);
    const j = await r.json();
    const ent = j?.entities?.[qid];

    const label = ent?.labels?.en?.value || "";
    const title = ent?.sitelinks?.enwiki?.title || "";

    return {
        label,
        wikipediaLink: title ? `https://en.wikipedia.org/wiki/${encodeURIComponent(title.replace(/ /g, "_"))}` : ""
    };
}

/* ---- main converter: property-summary JSON -> store.targets node ---- */
async function convertWdqsPropertySummaryToTargets({
    sourceLabel,
    sourceID,
    wdqsJson,
    onProgress
}) {
    const bindings = wdqsJson?.results?.bindings || [];
    const targets = [];

    // cache label->qid lookups during this conversion
    const labelToQid = new Map();

    // progress accounting: sum label count in ol_ plus a few steps
    let totalLabels = 0;
    for (const row of bindings) totalLabels += splitOlLiteral(row?.ol_?.value || "").length;
    const totalSteps = Math.max(1, totalLabels + 2);
    let done = 0;

    const bump = (msg) => {
        done++;
        const pct = Math.round((done / totalSteps) * 100);
        if (typeof onProgress === "function") onProgress(pct, msg);
    };

    if (typeof onProgress === "function") onProgress(1, `Resolving properties/targets for ${sourceID}…`);

    for (let i = 0; i < bindings.length; i++) {
        const row = bindings[i];

        const propertyID = propIdFromUri(row?.p?.value || "");
        const propertyLabel = row?.pl_?.value || "";
        const nodeCount = row?.count?.value || "";

        const labels = splitOlLiteral(row?.ol_?.value || "");
        const targetNodes = [];

        for (let li = 0; li < labels.length; li++) {
            const lbl = labels[li];

            if (!labelToQid.has(lbl)) {
                if (typeof onProgress === "function") {
                    onProgress(
                        Math.round((done / totalSteps) * 100),
                        `Finding QID for “${lbl}”… (${done}/${totalSteps})`
                    );
                }
                labelToQid.set(lbl, await qidFromLabel(lbl, "en"));
            }

            targetNodes.push({
                targetLabel: lbl,
                targetID: labelToQid.get(lbl) || ""
            });

            bump(`Resolved “${lbl}” (${done}/${totalSteps})`);
        }

        targets.push({
            propertyID,
            propertyLabel,
            nodeCount,
            targetNodes
        });

        if (typeof onProgress === "function") {
            onProgress(
                Math.round((done / totalSteps) * 100),
                `Built ${i + 1}/${bindings.length} property buckets…`
            );
        }
    }

    if (typeof onProgress === "function") onProgress(95, `Fetching label/wiki links for ${sourceID}…`);

    // refine the source label + wikipedia link
    let meta = { label: sourceLabel || "", wikipediaLink: "" };
    try {
        meta = await enrichNodeMeta(sourceID);
    } catch {
        // best-effort; keep what we have
    }

    if (typeof onProgress === "function") onProgress(100, `Done building node for ${sourceID}.`);

    return {
        sourceLabel: meta.label || sourceLabel || "",
        sourceID,
        sourceWikiDataLink: `https://www.wikidata.org/wiki/${sourceID}`,
        sourceWikipediaLink: meta.wikipediaLink || "",
        targets
    };
}

/* ---------------------------
   Descriptions cache (EXPORT-TIME)
---------------------------- */

function normalizeDescriptionsStore(store) {
    logDesc("normalizeDescriptionsStore");

    if (!store.descriptions) {
        store.descriptions = {
            byQid: {},
            updatedAt: null,
            language: CFG.descLanguage
        };
    }

    if (!store.descriptions.byQid) store.descriptions.byQid = {};
    if (!store.descriptions.language) store.descriptions.language = CFG.descLanguage;

    return store;
}

async function fetchDescriptionsForQids(qids, onProgress) {
    logDesc("fetchDescriptionsForQids", qids.length);

    const out = {};
    const clean = uniq(qids).filter(q => /^Q\d+$/.test(q));
    const batches = chunk(clean, CFG.wbgetentitiesBatch);
    const lang = CFG.descLanguage || "en";

    for (let bi = 0; bi < batches.length; bi++) {
        const ids = batches[bi];

        const url =
            "https://www.wikidata.org/w/api.php" +
            "?action=wbgetentities" +
            `&ids=${encodeURIComponent(ids.join("|"))}` +
            "&props=descriptions" +
            `&languages=${encodeURIComponent(lang)}` +
            "&format=json&origin=*";

        const r = await fetch(url);
        const j = await r.json();

        for (const qid of ids) {
            const ent = j?.entities?.[qid];
            out[qid] = ent?.descriptions?.[lang]?.value || "";
        }

        if (typeof onProgress === "function") {
            const pct = Math.round(((bi + 1) / batches.length) * 100);
            onProgress(pct, `Fetching descriptions… ${bi + 1}/${batches.length} batches`);
        }
    }

    return out;
}

async function ensureDescriptionsCachedForNodesExport(store, onProgress) {
    logDesc("ensureDescriptionsCachedForNodesExport");

    store = normalizeDescriptionsStore(store);

    const qids = collectAllQidsFromStore(store);
    const missing = qids.filter(q => !store.descriptions.byQid[q]);

    if (!missing.length) {
        logDesc("descriptions already cached");
        if (typeof onProgress === "function") onProgress(100, "Descriptions already cached.");
        return store;
    }

    const fetched = await fetchDescriptionsForQids(missing, onProgress);

    for (const [qid, desc] of Object.entries(fetched)) {
        store.descriptions.byQid[qid] = desc || "";
    }

    store.descriptions.updatedAt = nowISO();
    return store;
}

function attachDescriptionsToUniqueNodes(uniqueNodes, store) {
    logDesc("attachDescriptionsToUniqueNodes", uniqueNodes.length);

    const byQid = store?.descriptions?.byQid || {};

    for (const n of uniqueNodes) {
        n.description = byQid[n.qid] || "";
    }

    return uniqueNodes;
}


/* ---------------------------
   SECTION 4: Image enrichment (export-time)
---------------------------- */

/* ---- commons helper (Special:FilePath) ---- */
function commonsFilePathUrl(fileName) {
    if (!fileName) return "";
    return `https://commons.wikimedia.org/wiki/Special:FilePath/${encodeURIComponent(String(fileName).replace(/ /g, "_"))}`;
}

/* ---- store normalization ---- */
function normalizeImagesStore(store) {
    if (!store.images) store.images = { byQid: {}, byLink: {}, nextImgN: 1, updatedAt: null };
    if (!store.images.byQid) store.images.byQid = {};
    if (!store.images.byLink) store.images.byLink = {};
    if (!store.images.nextImgN) store.images.nextImgN = 1;
    return store;
}

/* ---- assign ids ---- */
function assignImageId(store, imageLink) {
    store = normalizeImagesStore(store);
    if (!imageLink) return "";
    if (store.images.byLink[imageLink]) return store.images.byLink[imageLink];

    const id = `img_${store.images.nextImgN++}`;
    store.images.byLink[imageLink] = id;
    return id;
}

/* ---- fetch P18 images for QIDs (batched wbgetentities claims) ---- */

async function fetchImagesForQids(qids, onProgress) {
    const out = {};
    const clean = uniq(qids).filter(q => /^Q\d+$/.test(q));
    const batches = chunk(clean, CFG.wbgetentitiesBatch);

    for (let bi = 0; bi < batches.length; bi++) {
        const ids = batches[bi];

        const url =
            "https://www.wikidata.org/w/api.php" +
            "?action=wbgetentities" +
            `&ids=${encodeURIComponent(ids.join("|"))}` +
            "&props=claims" +
            "&format=json&origin=*";

        const r = await fetch(url);
        const j = await r.json();

        for (const qid of ids) {
            const ent = j?.entities?.[qid];
            const p18 = ent?.claims?.[CFG.imageProp];
            const links = [];

            if (Array.isArray(p18)) {
                for (const cl of p18) {
                    const fileName = cl?.mainsnak?.datavalue?.value;
                    if (fileName) links.push(commonsFilePathUrl(fileName));
                }
            }

            out[qid] = uniq(links);
        }

        if (typeof onProgress === "function") {
            const pct = Math.round(((bi + 1) / batches.length) * 100);
            onProgress(pct, `Fetching images… ${bi + 1}/${batches.length} batches`);
        }
    }

    return out;
}

/* ---- ensure images cached in store for qids ---- */

async function ensureImagesCachedForExport(store, qids, onProgress) {
    store = normalizeImagesStore(store);

    const missing = [];
    for (const q of qids) {
        if (!store.images.byQid[q]) missing.push(q);
    }

    if (!missing.length) {
        if (typeof onProgress === "function") onProgress(100, "Images already cached.");
        return store;
    }

    const fetched = await fetchImagesForQids(missing, onProgress);

    for (const [qid, links] of Object.entries(fetched)) {
        const arr = Array.isArray(links) ? links : [];
        store.images.byQid[qid] = uniq(arr);

        // assign ids for each unique link
        for (const link of store.images.byQid[qid]) assignImageId(store, link);
    }

    store.images.updatedAt = nowISO();
    return store;
}

/* ---- attach image columns to node rows (for nodes CSV/XLSX) ---- */

function attachImagesToUniqueNodes(uniqueNodes, store, includeImageIds) {
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    for (const n of uniqueNodes) {
        const links = Array.isArray(byQid[n.qid]) ? byQid[n.qid] : [];

        // match your node CSV example: quoted links separated by comma+space
        n.images = links.map(u => `"${u}"`).join(", ");

        if (includeImageIds) {
            const ids = links.map(u => byLink[u]).filter(Boolean);
            n.imageIds = ids.join(", ");
        } else {
            n.imageIds = "";
        }
    }
    return uniqueNodes;
}

/* ---- (optional) enrich the nested targets JSON (sourceImages/targetImages) ---- */
function enrichTargetsJsonWithImages(targets, store) {
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    const imageObjsForQid = (qid) => {
        const links = Array.isArray(byQid[qid]) ? byQid[qid] : [];
        return links
            .map(link => ({ imageId: byLink[link] || "", imageLink: link }))
            .filter(x => x.imageLink);
    };

    for (const s of targets) {
        const sid = s?.sourceID || "";
        s.sourceImages = imageObjsForQid(sid);

        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                n.targetImages = imageObjsForQid(tid);
            }
        }
    }
}


function buildUniqueImagesRowsUnique(store) {
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    // imageId -> { imageId, imageLink, qids:Set }
    const imgMap = new Map();

    for (const [qid, links] of Object.entries(byQid)) {
        if (!/^Q\d+$/.test(qid)) continue;

        const arr = Array.isArray(links) ? links : [];
        for (const link of arr) {
            if (!link) continue;

            const imageId = byLink[link] || "";
            if (!imageId) continue;

            if (!imgMap.has(imageId)) {
                imgMap.set(imageId, { imageId, imageLink: link, qids: new Set() });
            }
            imgMap.get(imageId).qids.add(qid);
        }
    }

    const rows = [];
    for (const v of imgMap.values()) {
        const qidsArr = [...v.qids].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
        const qidsStr = qidsArr.join(", ");

        rows.push({
            imageId: v.imageId,
            qid: qidsStr,
            imageLink: v.imageLink,
            JSON: jsonCell({ imageId: v.imageId, qids: qidsArr, imageLink: v.imageLink })
        });
    }

    rows.sort((a, b) =>
        String(a.imageId || "").localeCompare(String(b.imageId || ""), undefined, { numeric: true })
    );

    return rows;
}

/* ---------------------------
   SECTION 5: Unique node + property summaries (for CSV exports)
---------------------------- */

/* ---- properties ---- */
function buildUniqueProperties(store) {
    const seen = new Map(); // pid -> {propertyID, propertyLabel, wikidataLink}
    const items = Array.isArray(store?.targets) ? store.targets : [];

    for (const s of items) {
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const pid = b?.propertyID || "";
            if (!/^P\d+$/.test(pid)) continue;

            if (!seen.has(pid)) {
                const label = b?.propertyLabel || "";
                seen.set(pid, {
                    propertyID: pid,
                    propertyLabel: label,
                    wikidataLink: `https://www.wikidata.org/wiki/Property:${pid}`
                });
            } else {
                const cur = seen.get(pid);
                if (!cur.propertyLabel && b?.propertyLabel) cur.propertyLabel = b.propertyLabel;
            }
        }
    }

    const arr = [...seen.values()];
    arr.sort((a, b) => a.propertyID.localeCompare(b.propertyID, undefined, { numeric: true }));
    return arr;
}

/* ---- nodes (unique sources + targets) ---- */
function buildUniqueNodes(store) {
    const map = new Map(); // qid -> row
    const items = Array.isArray(store?.targets) ? store.targets : [];

    for (const s of items) {
        // include source node itself
        const sid = s?.sourceID || "";
        if (/^Q\d+$/.test(sid)) {
            if (!map.has(sid)) {
                map.set(sid, {
                    qid: sid,
                    label: s?.sourceLabel || "",
                    wikidataLink: s?.sourceWikiDataLink || `https://www.wikidata.org/wiki/${sid}`,
                    wikipediaLink: s?.sourceWikipediaLink || "",
                    kind: "source",
                    // filled later during export-prep
                    description: "",
                    images: "",
                    imageIds: ""
                });
            } else {
                const cur = map.get(sid);
                if (!cur.label && s?.sourceLabel) cur.label = s.sourceLabel;
                if (!cur.wikipediaLink && s?.sourceWikipediaLink) cur.wikipediaLink = s.sourceWikipediaLink;
            }
        }

        // include all target nodes across buckets
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                if (!/^Q\d+$/.test(tid)) continue;

                if (!map.has(tid)) {
                    map.set(tid, {
                        qid: tid,
                        label: n?.targetLabel || "",
                        wikidataLink: `https://www.wikidata.org/wiki/${tid}`,
                        wikipediaLink: "",
                        kind: "target",
                        // filled later during export-prep
                        description: "",
                        images: "",
                        imageIds: ""
                    });
                } else {
                    const cur = map.get(tid);
                    if (!cur.label && n?.targetLabel) cur.label = n.targetLabel;
                }
            }
        }
    }

    const out = [...map.values()];
    out.sort((a, b) => a.qid.localeCompare(b.qid, undefined, { numeric: true }));
    return out;
}


function buildLinksRows(store, includeImageIds) {
    const rows = [];
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    const items = Array.isArray(store?.targets) ? store.targets : [];
    for (const s of items) {
        const sid = s?.sourceID || "";
        const sLabel = s?.sourceLabel || "";
        const sWd = s?.sourceWikiDataLink || (sid ? `https://www.wikidata.org/wiki/${sid}` : "");
        const sWp = s?.sourceWikipediaLink || "";

        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const pid = b?.propertyID || "";
            const pLabel = b?.propertyLabel || "";
            const nodeCount = b?.nodeCount ?? "";

            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                const tLabel = n?.targetLabel || "";

                const base = {
                    sourceID: sid,
                    sourceLabel: sLabel,
                    propertyID: pid,
                    propertyLabel: pLabel,
                    nodeCount: nodeCount,
                    targetID: tid,
                    targetLabel: tLabel,
                    sourceWikiDataLink: sWd,
                    sourceWikipediaLink: sWp
                };

                // optional image id columns for link rows
                if (includeImageIds) {
                    const sImgIds = sid && Array.isArray(byQid[sid])
                        ? byQid[sid].map(u => byLink[u]).filter(Boolean).join(", ")
                        : "";
                    const tImgIds = tid && Array.isArray(byQid[tid])
                        ? byQid[tid].map(u => byLink[u]).filter(Boolean).join(", ")
                        : "";
                    base.sourceImageIds = sImgIds;
                    base.targetImageIds = tImgIds;
                }

                // JSON col per your example
                base.JSON = jsonCell({ source: sid, propertyLabel: pLabel, target: tid });

                rows.push(base);
            }
        }
    }

    rows.sort((a, b) => {
        const ak = `${a.sourceID}|${a.propertyID}|${a.targetID}`;
        const bk = `${b.sourceID}|${b.propertyID}|${b.targetID}`;
        return ak.localeCompare(bk, undefined, { numeric: true, sensitivity: "base" });
    });

    return rows;
}

/* ---- PROPERTIES rows ---- index id labels JSON */

function buildPropertiesRows(store) {
    const props = buildUniqueProperties(store);
    return props.map((p, i) => ({
        index: i + 1,
        id: p.propertyID,
        labels: p.propertyLabel,
        JSON: jsonCell({ pid: p.propertyID, label: p.propertyLabel })
    }));
}

/* ---- NODES rows ---- */

function buildNodesRows(store, includeImages) {
    const nodes = buildUniqueNodes(store);

    // images are attached later during export-prep,
    // but if caller already ensured images cache, we can attach now.
    if (includeImages) attachImagesToUniqueNodes(nodes, store, true);

    // ensure fields exist + JSON column (refresh after descriptions/images are attached)
    for (const n of nodes) {
        if (typeof n.description !== "string") n.description = "";
        if (!("images" in n)) n.images = "";
        if (!("imageIds" in n)) n.imageIds = "";

        const imageIdsArr = String(n.imageIds || "").split(",").map(x => x.trim()).filter(Boolean);

        n.JSON = jsonCell({
            qid: n.qid,
            label: n.label,
            description: n.description || "",
            wikidataLink: n.wikidataLink,
            wikipediaLink: n.wikipediaLink,
            kind: n.kind,
            imageIds: imageIdsArr
        });
    }

    return nodes;
}

/* ---- Headers helpers ---- */
function getLinksHeaders(includeImageIds) {
    return includeImageIds
        ? [
            "sourceID", "sourceLabel", "propertyID", "propertyLabel", "nodeCount",
            "targetID", "targetLabel",
            "sourceWikiDataLink", "sourceWikipediaLink",
            "sourceImageIds", "targetImageIds",
            "JSON"
        ]
        : [
            "sourceID", "sourceLabel", "propertyID", "propertyLabel", "nodeCount",
            "targetID", "targetLabel",
            "sourceWikiDataLink", "sourceWikipediaLink",
            "JSON"
        ];
}

function getPropertiesHeaders() {
    return ["index", "id", "labels", "JSON"];
}

function getNodesHeaders(includeImages) {
    return includeImages
        ? ["qid", "label", "description", "wikidataLink", "wikipediaLink", "kind", "images", "imageIds", "JSON"]
        : ["qid", "label", "description", "wikidataLink", "wikipediaLink", "kind", "JSON"];
}

function getImagesHeaders() {
    return ["imageId", "qid", "imageLink", "JSON"];
}

/* ---------------------------
   SECTION 6: Core: Merge capture
---------------------------- */

/* ---- store schema helpers ---- */

function ensureStoreShape(store) {
    // minimal defensive normalization (readStore already does more)
    if (!store.meta) store.meta = { name: "", createdAt: nowISO(), updatedAt: nowISO(), source: "query.wikidata.org", version: "2.4" };
    if (!store.options) store.options = { storeFullResponses: false };
    if (!Array.isArray(store.targets)) store.targets = [];
    if (!Array.isArray(store.captures)) store.captures = [];
    if (!store.qidIndex || typeof store.qidIndex !== "object") store.qidIndex = {};
    if (!("lastPreview" in store)) store.lastPreview = null;
    return store;
}

function makeTripleKey(sourceID, propertyID, targetID) {
    return `${sourceID || ""}|${propertyID || ""}|${targetID || ""}`;
}

/* ---- find/create helpers for nested structure ---- */

function getOrCreateSourceNode(targetsArr, srcObj) {
    const sid = srcObj?.sourceID || "";
    if (!sid) return null;

    let src = targetsArr.find(x => x?.sourceID === sid);
    if (!src) {
        src = {
            sourceLabel: srcObj?.sourceLabel || "",
            sourceID: sid,
            sourceWikiDataLink: srcObj?.sourceWikiDataLink || `https://www.wikidata.org/wiki/${sid}`,
            sourceWikipediaLink: srcObj?.sourceWikipediaLink || "",
            targets: []
        };
        targetsArr.push(src);
    } else {
        // backfill label/links if missing
        if (!src.sourceLabel && srcObj?.sourceLabel) src.sourceLabel = srcObj.sourceLabel;
        if (!src.sourceWikiDataLink && srcObj?.sourceWikiDataLink) src.sourceWikiDataLink = srcObj.sourceWikiDataLink;
        if (!src.sourceWikipediaLink && srcObj?.sourceWikipediaLink) src.sourceWikipediaLink = srcObj.sourceWikipediaLink;
        if (!Array.isArray(src.targets)) src.targets = [];
    }
    return src;
}

function getOrCreatePropertyBucket(srcNode, bucketObj) {
    const pid = bucketObj?.propertyID || "";
    if (!pid) return null;

    let buck = (srcNode.targets || []).find(x => x?.propertyID === pid);
    if (!buck) {
        buck = {
            propertyID: pid,
            propertyLabel: bucketObj?.propertyLabel || "",
            nodeCount: bucketObj?.nodeCount ?? "",
            targetNodes: []
        };
        srcNode.targets.push(buck);
    } else {
        if (!buck.propertyLabel && bucketObj?.propertyLabel) buck.propertyLabel = bucketObj.propertyLabel;
        // keep nodeCount if empty
        if ((buck.nodeCount === "" || buck.nodeCount == null) && (bucketObj?.nodeCount != null)) buck.nodeCount = bucketObj.nodeCount;
        if (!Array.isArray(buck.targetNodes)) buck.targetNodes = [];
    }
    return buck;
}

/* ---- build a dedupe index from current store.targets ---- */
function buildDedupeIndexFromTargets(targetsArr) {
    const index = new Set();

    for (const s of (targetsArr || [])) {
        const sid = s?.sourceID || "";
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const pid = b?.propertyID || "";
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                if (!sid || !pid || !tid) continue;
                index.add(makeTripleKey(sid, pid, tid));
            }
        }
    }
    return index;
}

/* ---- merge a converted "property summary" capture into store.targets ---- */

function mergeConvertedTargetsIntoStore(store, converted, onProgress) {
    store = ensureStoreShape(store);

    const targetsArr = store.targets;
    const index = buildDedupeIndexFromTargets(targetsArr);

    // converted: { sourceID, sourceLabel, sourceWikiDataLink, sourceWikipediaLink, targets:[{propertyID,..., targetNodes:[]}] }
    const src = getOrCreateSourceNode(targetsArr, converted);
    if (!src) return store;

    const buckets = Array.isArray(converted?.targets) ? converted.targets : [];
    const total = Math.max(1, buckets.length);
    let done = 0;

    for (const b of buckets) {
        const buck = getOrCreatePropertyBucket(src, b);
        if (!buck) continue;

        const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
        for (const n of nodes) {
            const tid = n?.targetID || "";
            const pid = buck.propertyID;
            const sid = src.sourceID;

            if (!sid || !pid || !tid) continue;

            const key = makeTripleKey(sid, pid, tid);
            if (index.has(key)) continue;

            buck.targetNodes.push({
                targetLabel: n?.targetLabel || "",
                targetID: tid
            });

            index.add(key);
        }

        done++;
        if (typeof onProgress === "function") {
            const pct = Math.round((done / total) * 100);
            onProgress(pct, `Merged ${done}/${total} property buckets…`);
        }
    }

    return store;
}

/* ---- create a capture record (optional raw response storage) ---- */

function makeCaptureRecord({ url, queryText, kind, sourceID, note }) {
    return {
        capturedAt: nowISO(),
        url: url || "",
        kind: kind || "unknown",
        sourceID: sourceID || "",
        query: queryText || "",
        note: note || ""
    };
}

/* ---- main entry: merge a WDQS JSON response (and optionally raw response) ---- */

async function mergeWdqsJsonIntoStore({ url, data, storeFull, onProgress }) {
    let store = readStore();
    store = ensureStoreShape(store);

    const queryText = decodeQueryFromUrl(url);
    const headVars = data?.head?.vars || [];
    const isSummary = looksLikePropertySummary(headVars);

    // Always record a capture entry
    const inferredSourceID = extractSubjectQidFromQueryText(queryText);
    const cap = makeCaptureRecord({
        url,
        queryText,
        kind: isSummary ? "property-summary" : "generic",
        sourceID: inferredSourceID,
        note: isSummary ? "Converted property-summary into targets." : "Captured generic WDQS JSON."
    });

    // Optional: store raw response (big)
    if (storeFull) {
        // keep raw captures lightweight-ish: only last N? (optional). Here we just push.
        store.captures.push({
            ...cap,
            response: data
        });
    } else {
        store.captures.push(cap);
    }

    // Always keep lastPreview usable
    store.lastPreview = {
        capturedAt: cap.capturedAt,
        kind: cap.kind,
        url: cap.url,
        sourceID: cap.sourceID,
        headVars: Array.isArray(headVars) ? headVars : [],
        bindingsCount: Array.isArray(data?.results?.bindings) ? data.results.bindings.length : 0
    };

    // Also keep a fast qidIndex of what we saw (generic or summary)
    // This is legacy-ish but useful for debugging and quick counts.
    try {
        const qids = extractQidsFromWdqsResults(data);
        for (const q of qids) store.qidIndex[q] = true;
    } catch { }

    if (!isSummary) {
        // For non-summary queries, we do NOT convert into store.targets
        writeStore(store);
        return store;
    }

    // Summary: Convert then merge into nested store.targets
    const subjectQid = inferredSourceID || "";
    if (!/^Q\d+$/.test(subjectQid)) {
        // We can’t merge without a sourceID. Still keep capture.
        writeStore(store);
        return store;
    }

    if (typeof onProgress === "function") onProgress(5, `Converting property summary for ${subjectQid}…`);

    const converted = await convertWdqsPropertySummaryToTargets({
        sourceLabel: "",           // will be filled by enrichNodeMeta() inside converter
        sourceID: subjectQid,
        wdqsJson: data,
        onProgress
    });

    if (typeof onProgress === "function") onProgress(70, "Merging into store…");

    store = mergeConvertedTargetsIntoStore(store, converted, onProgress);

    // Persist
    writeStore(store);

    if (typeof onProgress === "function") onProgress(100, "Merged capture.");
    return store;
}

/* ---------------------------
   SECTION 7: UI Panel (resizable + collapse + preview)
---------------------------- */


function injectPanelResourceCss() {
    if (injectPanelResourceCss.done) return;

    if (typeof GM_getResourceText !== "function") {
        throw new Error("GM_getResourceText is unavailable. Add @grant GM_getResourceText and @resource PANEL_CSS.");
    }

    const css = GM_getResourceText("PANEL_CSS") || "";
    const style = document.createElement("style");
    style.dataset.wdqsCaptureResource = "panel-css";
    style.textContent = css;
    document.documentElement.appendChild(style);

    injectPanelResourceCss.done = true;
}

function createPanelFromHtmlResource() {
    if (typeof GM_getResourceText !== "function") {
        throw new Error("GM_getResourceText is unavailable. Add @grant GM_getResourceText and @resource PANEL_HTML.");
    }

    const html = GM_getResourceText("PANEL_HTML") || "";
    const template = document.createElement("template");
    template.innerHTML = html.trim();

    const panel = template.content.firstElementChild;
    if (!panel) throw new Error("PANEL_HTML did not contain a root panel element.");

    if (!panel.id) panel.id = CFG.panelId;
    return panel;
}

function ensurePanel() {
    if (document.getElementById(CFG.panelId)) {
        return {
            panel: document.getElementById(CFG.panelId),
            api: getPanelApi()
        };
    }

    const ui = readUIState();
    const store = readStore();

    injectPanelResourceCss();

    const panel = createPanelFromHtmlResource();
    panel.id = CFG.panelId;
    panel.classList.toggle("is-collapsed", ui.collapsed);
    panel.classList.toggle("is-tall", ui.tall);

    if (ui.tall) {
        panel.style.height = `calc(100vh - 28px)`;
    } else {
        panel.style.height = `${ui.height}px`;
    }

    document.body.appendChild(panel);

    // ---- initialize form values (NO EVENT WIRING HERE) ----
    const nameInput = panel.querySelector(".wdqs-cap-name");
    if (nameInput) nameInput.value = store?.meta?.name || "";

    const fullChk = panel.querySelector(".wdqs-cap-full");
    if (fullChk) fullChk.checked = !!store?.options?.storeFullResponses;

    // init image checkboxes (UI only)
    const imgChk = panel.querySelector(".wdqs-cap-img");
    const imgUniqChk = panel.querySelector(".wdqs-cap-img-uniq");

    const treeToggle = panel.querySelector(".wdqs-cap-tree-on");
    const treeWrap = panel.querySelector(".wdqs-cap-tree-wrap");
    if (imgChk) imgChk.checked = false;
    if (imgUniqChk) {
        imgUniqChk.checked = false;
        imgUniqChk.disabled = !(imgChk && imgChk.checked);
    }

    // ---- apply initial mode (collapsed/tall/normal) ----
    function applyPanelModeFromUIState() {
        const s = readUIState();

        panel.classList.toggle("is-collapsed", !!s.collapsed);
        panel.classList.toggle("is-tall", !!s.tall);

        // IMPORTANT: inline height must not fight collapsed
        if (s.collapsed) {
            panel.style.height = "";
            panel.style.resize = "none";
        } else if (s.tall) {
            panel.style.height = `calc(100vh - 28px)`;
            panel.style.resize = "none";
        } else {
            panel.style.height = `${s.height || 360}px`;
            panel.style.resize = "vertical";
        }

        // keep header reachable
        panel.scrollTop = 0;

        // icons
        const toggleIcon = panel.querySelector(".wdqs-cap-toggle-ic");
        if (toggleIcon) toggleIcon.textContent = s.collapsed ? "▸" : "▾";

        const tallIcon = panel.querySelector(".wdqs-cap-tall-ic");
        if (tallIcon) tallIcon.textContent = s.tall ? "⤡" : "⤢";
    }

    applyPanelModeFromUIState();

    // initial counts + preview
    refreshCounts(store);
    setPreview(store.lastPreview || {});

    return { panel, api: getPanelApi() };
}

/* ---- small accessors so other sections don’t querySelector everything repeatedly ---- */

function getPanelApi() {
    const panel = document.getElementById(CFG.panelId);
    if (!panel) return null;

    return {
        panel,

        // inputs
        nameInput: panel.querySelector(".wdqs-cap-name"),
        fullChk: panel.querySelector(".wdqs-cap-full"),
        imgChk: panel.querySelector(".wdqs-cap-img"),
        imgUniqChk: panel.querySelector(".wdqs-cap-img-uniq"),

        // buttons
        btnCopy: panel.querySelector(".wdqs-btn-copy"),
        btnDl: panel.querySelector(".wdqs-btn-dl"),
        btnCsvNodes: panel.querySelector(".wdqs-btn-csv-nodes"),
        btnCsvProps: panel.querySelector(".wdqs-btn-csv-props"),
        btnCsvLinks: panel.querySelector(".wdqs-btn-csv-links"),
        btnXlsx: panel.querySelector(".wdqs-btn-xlsx"),
        btnClear: panel.querySelector(".wdqs-btn-clear"),
        fileInput: panel.querySelector(".wdqs-file"),

        // areas
        preview: panel.querySelector(".wdqs-cap-preview"),
        status: panel.querySelector(".wdqs-cap-status"),

        // progress
        capProgressWrap: panel.querySelector(".wdqs-cap-capprogress"),
        expProgressWrap: panel.querySelector(".wdqs-cap-progress"),

        // tree summary
        treeToggle: panel.querySelector(".wdqs-cap-tree-on"),
        treeWrap: panel.querySelector(".wdqs-cap-tree-wrap"),
        treeBox: panel.querySelector(".wdqs-cap-tree-box")
    };
}

/* ---------------------------
   SECTION 8: Helpers for export actions
---------------------------- */

/* ---------------------------
   Capture-time progress helpers
---------------------------- */

function getCaptureProgressEls() {
    const wrap = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogress`);
    const bar = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogressbar > span`);
    const txt = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogresstxt`);
    return { wrap, bar, txt };
}

function setCaptureProgress(visible, pct = 0, text = "") {
    const { wrap, bar, txt } = getCaptureProgressEls();
    if (!wrap || !bar || !txt) return;

    wrap.classList.toggle("is-slot", !visible);
    wrap.classList.toggle("is-active", visible);

    bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
    txt.textContent = text || "";
}

/* --------------------------------
   Progress helpers (capture-time)
--------------------------------- */

function getCaptureProgressEls() {
    const wrap = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogress`);
    const bar = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogressbar > span`);
    const txt = document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogresstxt`);
    return { wrap, bar, txt };
}

function setCaptureProgress(visible, pct = 0, text = "") {
    const { wrap, bar, txt } = getCaptureProgressEls();
    if (!wrap || !bar || !txt) return;

    wrap.classList.toggle("is-slot", !visible);
    wrap.classList.toggle("is-active", visible);

    bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
    txt.textContent = text || "";
}

/**
 * Your capture pipeline calls makeProgressReporter(...).
 * Define it (and keep the name) so existing logic doesn’t break.
 */

function makeProgressReporter(labelPrefix = "Working…") {
    let lastPaint = 0;
    return function report(pct, msg) {
        const now = performance.now();
        if (now - lastPaint < 60 && pct < 100) return;
        lastPaint = now;
        setCaptureProgress(true, pct, msg ? msg : labelPrefix);
    };
}

/* --------------------------------
   Progress helpers (export-time)
--------------------------------- */

function getExportProgressEls() {
    const wrap = document.querySelector(`#${CFG.panelId} .wdqs-cap-progress`);
    const bar = document.querySelector(`#${CFG.panelId} .wdqs-cap-progressbar > span`);
    const txt = document.querySelector(`#${CFG.panelId} .wdqs-cap-progresstxt`);
    return { wrap, bar, txt };
}

function setExportProgress(visible, pct = 0, text = "") {
    const { wrap, bar, txt } = getExportProgressEls();
    if (!wrap || !bar || !txt) return;

    wrap.classList.toggle("is-slot", !visible);
    wrap.classList.toggle("is-active", visible);

    bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
    txt.textContent = text || "";
}

function makeCaptureProgressReporter(labelPrefix = "Working…") {
    let lastPaint = 0;
    return function report(pct, msg) {
        const now = performance.now();
        if (now - lastPaint < 60 && pct < 100) return;
        lastPaint = now;
        setCaptureProgress(true, pct, msg ? msg : labelPrefix);
    };
}

/* --------------------------------
   Enable/disable panel actions
--------------------------------- */

function setActionsDisabled(disabled) {
    const panel = document.getElementById(CFG.panelId);
    if (!panel) return;
    const els = panel.querySelectorAll("button, input, label.wdqs-btn, label.wdqs-btn-up");
    for (const el of els) {
        if (el.tagName === "LABEL") {
            el.style.pointerEvents = disabled ? "none" : "";
            el.style.opacity = disabled ? "0.6" : "";
        } else {
            el.disabled = !!disabled;
        }
    }
}

/* --------------------------------
   Clipboard helper
--------------------------------- */

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch {
        try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand("copy");
            ta.remove();
            return !!ok;
        } catch {
            return false;
        }
    }
}

/* --------------------------------
   File download helpers
--------------------------------- */

function safeFileName(name) {
    return String(name || "download")
        .replace(/[\\/:*?"<>|]+/g, "_")
        .replace(/\s+/g, " ")
        .trim() || "download";
}

function downloadBlob(filename, blob) {
    const cleanName = safeFileName(filename);
    const url = URL.createObjectURL(blob);

    try {
        // Use the normal browser download flow for generated Blob files.
        // GM_download is better for remote URLs, but it can silently fail with blob: URLs
        // depending on the userscript manager/browser permission setup.
        const a = document.createElement("a");
        a.href = url;
        a.download = cleanName;
        a.rel = "noopener";
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        a.remove();

        setTimeout(() => URL.revokeObjectURL(url), 15000);
        return true;
    } catch (e) {
        URL.revokeObjectURL(url);
        console.error("[WDQS-CAP] Download failed:", e);
        return false;
    }
}

function downloadText(filename, text, mime = "text/plain;charset=utf-8") {
    return downloadBlob(filename, new Blob([text], { type: mime }));
}

/* --------------------------------
   CSV helpers
--------------------------------- */

function csvEscape(v) {
    const s = String(v ?? "");
    const t = s.replace(/"/g, '""');
    return /[",\n\r]/.test(t) ? `"${t}"` : t;
}

function toCsv(headers, rows) {
    const out = [];
    out.push(headers.map(csvEscape).join(","));
    for (const r of rows) out.push(headers.map(h => csvEscape(r?.[h])).join(","));
    return out.join("\n");
}

function jsonCell(obj) {
    try { return JSON.stringify(obj); } catch { return ""; }
}

/* --------------------------------
   XLSX helpers
--------------------------------- */

function canExportXlsx() {
    return typeof window.XLSX !== "undefined"
        && typeof window.XLSX.utils === "object";
}

function clampSheetName(name) {
    return String(name || "Sheet")
        .replace(/[:\\\/\?\*\[\]]/g, " ")
        .slice(0, 31)
        .trim() || "Sheet";
}

function rowsToSheet(headers, rows) {
    const aoa = [headers];
    for (const r of rows) aoa.push(headers.map(h => r?.[h] ?? ""));
    return window.XLSX.utils.aoa_to_sheet(aoa);
}

/* --------------------------------
   Store prep for exports
   - Nodes exports ALWAYS fetch descriptions
   - Optional images fetch (P18) when requested
--------------------------------- */

async function prepareExportStore({ wantImages, onProgress }) {
    let s = readStore();

    // ALWAYS cache descriptions for Nodes CSV/XLSX
    try {
        if (typeof onProgress === "function") onProgress(2, "Preparing export store…");
        s = normalizeDescriptionsStore(s);
        s = await ensureDescriptionsCachedForNodesExport(s, onProgress);
    } catch (e) {
        console.warn("Descriptions fetch failed:", e);
        // continue; descriptions will be blank
    }

    // Optional: cache images for export
    if (wantImages) {
        try {
            const qids = collectAllQidsFromStore(s);
            s = normalizeImagesStore(s);
            s = await ensureImagesCachedForExport(s, qids, onProgress);
        } catch (e) {
            console.warn("Images fetch failed:", e);
            // continue; image columns blank
        }
    }

    writeStore(s); // persist caches
    return s;
}

/* --------------------------------
   Export row builders (Links/Props/Nodes)
   NOTE: These use helpers from SECTION 5:
   - buildUniqueProperties(store)
   - buildUniqueNodes(store)
   And cache attachment helpers from SECTION 4/Descriptions section:
   - attachImagesToUniqueNodes(...)
   - attachDescriptionsToUniqueNodes(...)
--------------------------------- */

// Headers
function getLinksHeaders(includeImageIds) {
    return includeImageIds
        ? [
            "sourceID", "sourceLabel", "propertyID", "propertyLabel", "nodeCount",
            "targetID", "targetLabel",
            "sourceWikiDataLink", "sourceWikipediaLink",
            "sourceImageIds", "targetImageIds",
            "JSON"
        ]
        : [
            "sourceID", "sourceLabel", "propertyID", "propertyLabel", "nodeCount",
            "targetID", "targetLabel",
            "sourceWikiDataLink", "sourceWikipediaLink",
            "JSON"
        ];
}

function getPropertiesHeaders() {
    return ["index", "id", "labels", "JSON"];
}

function getNodesHeaders(includeImages) {
    return includeImages
        ? ["qid", "label", "description", "wikidataLink", "wikipediaLink", "kind", "images", "imageIds", "JSON"]
        : ["qid", "label", "description", "wikidataLink", "wikipediaLink", "kind", "JSON"];
}

function getImagesHeaders() {
    return ["imageId", "qid", "imageLink", "JSON"];
}

// Links rows
function buildLinksRows(store, includeImageIds) {
    const rows = [];
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    const items = Array.isArray(store.targets) ? store.targets : [];
    for (const s of items) {
        const sid = s?.sourceID || "";
        const sLabel = s?.sourceLabel || "";
        const sWd = s?.sourceWikiDataLink || (sid ? `https://www.wikidata.org/wiki/${sid}` : "");
        const sWp = s?.sourceWikipediaLink || "";

        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const pid = b?.propertyID || "";
            const pLabel = b?.propertyLabel || "";
            const nodeCount = b?.nodeCount ?? "";

            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                const tLabel = n?.targetLabel || "";

                const base = {
                    sourceID: sid,
                    sourceLabel: sLabel,
                    propertyID: pid,
                    propertyLabel: pLabel,
                    nodeCount: nodeCount,
                    targetID: tid,
                    targetLabel: tLabel,
                    sourceWikiDataLink: sWd,
                    sourceWikipediaLink: sWp
                };

                if (includeImageIds) {
                    const sImgIds = sid && byQid[sid]
                        ? (byQid[sid].map(u => byLink[u]).filter(Boolean).join(", "))
                        : "";
                    const tImgIds = tid && byQid[tid]
                        ? (byQid[tid].map(u => byLink[u]).filter(Boolean).join(", "))
                        : "";
                    base.sourceImageIds = sImgIds;
                    base.targetImageIds = tImgIds;
                }

                base.JSON = jsonCell({ source: sid, propertyLabel: pLabel, target: tid });
                rows.push(base);
            }
        }
    }

    rows.sort((a, b) => {
        const ak = `${a.sourceID}|${a.propertyID}|${a.targetID}`;
        const bk = `${b.sourceID}|${b.propertyID}|${b.targetID}`;
        return ak.localeCompare(bk, undefined, { numeric: true, sensitivity: "base" });
    });

    return rows;
}

// Properties rows
function buildPropertiesRows(store) {
    const props = buildUniqueProperties(store);
    return props.map((p, i) => ({
        index: i + 1,
        id: p.propertyID,
        labels: p.propertyLabel,
        JSON: jsonCell({ pid: p.propertyID, label: p.propertyLabel })
    }));
}

// Nodes rows
function buildNodesRows(store, includeImages) {
    let nodes = buildUniqueNodes(store);

    // ALWAYS include descriptions in nodes exports
    attachDescriptionsToUniqueNodes(nodes, store);

    if (includeImages) attachImagesToUniqueNodes(nodes, store, true);

    for (const n of nodes) {
        if (typeof n.description !== "string") n.description = "";
        if (!("images" in n)) n.images = "";
        if (!("imageIds" in n)) n.imageIds = "";

        const imageIdsArr = String(n.imageIds || "")
            .split(",")
            .map(x => x.trim())
            .filter(Boolean);

        n.JSON = jsonCell({
            qid: n.qid,
            label: n.label,
            description: n.description || "",
            wikidataLink: n.wikidataLink,
            wikipediaLink: n.wikipediaLink,
            kind: n.kind,
            imageIds: imageIdsArr
        });
    }

    return nodes;
}

/* --------------------------------
   UNIQUE IMAGES ROWS
--------------------------------- */

function buildUniqueImagesRowsUnique(store) {
    const byQid = store?.images?.byQid || {};
    const byLink = store?.images?.byLink || {};

    // imageId -> { imageId, imageLink, qids:Set() }
    const imgMap = new Map();

    for (const [qid, links] of Object.entries(byQid)) {
        if (!/^Q\d+$/.test(qid)) continue;

        const arr = Array.isArray(links) ? links : [];
        for (const link of arr) {
            if (!link) continue;
            const imageId = byLink[link] || "";
            if (!imageId) continue;

            if (!imgMap.has(imageId)) {
                imgMap.set(imageId, { imageId, imageLink: link, qids: new Set() });
            }
            imgMap.get(imageId).qids.add(qid);
        }
    }

    const rows = [];
    for (const v of imgMap.values()) {
        const qids = [...v.qids].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
        const qidCell = qids.join(", ");
        rows.push({
            imageId: v.imageId,
            qid: qidCell,
            imageLink: v.imageLink,
            JSON: jsonCell({ imageId: v.imageId, qids, imageLink: v.imageLink })
        });
    }

    rows.sort((a, b) => String(a.imageId || "").localeCompare(String(b.imageId || ""), undefined, { numeric: true }));
    return rows;
}

/* --------------------------------
   High-level exporters used by UI
--------------------------------- */

async function exportJsonStore() {
    const s = readStore();
    const base = (s?.meta?.name || "wdqs_capture").trim() || "wdqs_capture";
    downloadText(`${base}.json`, JSON.stringify(s, null, 2));
}

async function exportCsvProperties() {
    const s = readStore();
    const base = (s?.meta?.name || "wdqs_capture").trim() || "wdqs_capture";
    const rows = buildPropertiesRows(s);
    const csv = toCsv(getPropertiesHeaders(), rows);
    downloadText(`${base}__properties.csv`, "\uFEFF" + csv);
}

async function exportCsvLinks({ wantImages, onProgress }) {
    const report = typeof onProgress === "function" ? onProgress : () => { };
    const s = wantImages ? await prepareExportStore({ wantImages: true, onProgress: report }) : readStore();

    const base = (s?.meta?.name || "wdqs_capture").trim() || "wdqs_capture";
    const rows = buildLinksRows(s, wantImages);
    const csv = toCsv(getLinksHeaders(wantImages), rows);
    downloadText(`${base}__links.csv`, "\uFEFF" + csv);
}

async function exportCsvNodes({ wantImages, wantUniqueImagesCsv, onProgress }) {
    const report = typeof onProgress === "function" ? onProgress : () => { };
    const s = await prepareExportStore({ wantImages: !!wantImages, onProgress: report });

    const base = (s?.meta?.name || "wdqs_capture").trim() || "wdqs_capture";

    const nodes = buildNodesRows(s, !!wantImages);
    const csv = toCsv(getNodesHeaders(!!wantImages), nodes);
    downloadText(`${base}__nodes.csv`, "\uFEFF" + csv);

    if (wantImages && wantUniqueImagesCsv) {
        const imgRows = buildUniqueImagesRowsUnique(s);
        const imgCsv = toCsv(getImagesHeaders(), imgRows);
        downloadText(`${base}__images.csv`, "\uFEFF" + imgCsv);
    }
}

async function exportXlsxSheets({ wantImages, wantUniqueImagesSheet, onProgress }) {

    if (!canExportXlsx()) {
        throw new Error("XLSX library not available");
    }

    const report = typeof onProgress === "function" ? onProgress : () => { };
    const store = await prepareExportStore({
        wantImages: !!wantImages,
        onProgress: report
    });

    const base =
        (store?.meta?.name || "wdqs_capture").trim() || "wdqs_capture";

    const propsRows = buildPropertiesRows(store);
    const linksRows = buildLinksRows(store, !!wantImages);
    const nodesRows = buildNodesRows(store, !!wantImages);
    const imgRows =
        (wantImages && wantUniqueImagesSheet)
            ? buildUniqueImagesRowsUnique(store)
            : [];

    const wb = window.XLSX.utils.book_new();

    const wsNodes = rowsToSheet(getNodesHeaders(!!wantImages), nodesRows);
    window.XLSX.utils.book_append_sheet(wb, wsNodes, clampSheetName("nodes"));

    const wsProps = rowsToSheet(getPropertiesHeaders(), propsRows);
    window.XLSX.utils.book_append_sheet(wb, wsProps, clampSheetName("properties"));

    const wsLinks = rowsToSheet(getLinksHeaders(!!wantImages), linksRows);
    window.XLSX.utils.book_append_sheet(wb, wsLinks, clampSheetName("links"));

    if (wantImages && wantUniqueImagesSheet) {
        const wsImgs = rowsToSheet(getImagesHeaders(), imgRows);
        window.XLSX.utils.book_append_sheet(wb, wsImgs, clampSheetName("images"));
    }

    const out = window.XLSX.write(wb, {
        bookType: "xlsx",
        type: "array"
    });

    downloadBlob(
        `${base}.xlsx`,
        new Blob([out], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        })
    );
}

/* =================================
   Export progress helpers (RESTORED)
================================== */

function getExportProgressEls() {
    const wrap = document.querySelector(`#${CFG.panelId} .wdqs-cap-progress`);
    const bar = document.querySelector(`#${CFG.panelId} .wdqs-cap-progressbar > span`);
    const txt = document.querySelector(`#${CFG.panelId} .wdqs-cap-progresstxt`);
    return { wrap, bar, txt };
}

function setExportProgress(visible, pct = 0, text = "") {
    const { wrap, bar, txt } = getExportProgressEls();
    if (!wrap || !bar || !txt) {
        DBG.warn("Export progress elements missing");
        return;
    }

    wrap.classList.toggle("is-slot", !visible);
    bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
    txt.textContent = text || "";
}

function makeExportProgressReporter(labelPrefix = "Exporting…") {
    DBG.log("makeExportProgressReporter created:", labelPrefix);

    let lastPaint = 0;
    return function report(pct, msg) {
        const now = performance.now();
        if (now - lastPaint < 60 && pct < 100) return;
        lastPaint = now;

        DBG.log("EXPORT PROGRESS", pct, msg);
        setExportProgress(true, pct, msg || labelPrefix);
    };
}

function resetStoreToEmptyButKeepMeta(prev) {
    const createdAt = prev?.meta?.createdAt || nowISO();
    const name = prev?.meta?.name || "";

    return {
        meta: {
            name,
            createdAt,
            updatedAt: nowISO(),
            source: "query.wikidata.org",
            version: prev?.meta?.version || "2.5",
        },
        options: {
            storeFullResponses: !!prev?.options?.storeFullResponses
        },

        // main dataset
        targets: [],

        // legacy / logs
        qidIndex: {},
        captures: [],

        // UI preview
        lastPreview: null,

        // caches (optional: clear them so next export re-fetches fresh)
        images: { byQid: {}, byLink: {}, nextImgN: 1, updatedAt: null },
        descriptions: { byQid: {}, updatedAt: null, language: CFG.descLanguage }
    };
}

/* ---------------------------
   SECTION 9: Wire up UI behaviors
---------------------------- */


/* ---------------------------
   SECTION 9A: Tree summary viewer helpers
---------------------------- */

function escHtml(v) {
    return String(v ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function labelForStoreQid(store, qid) {
    const items = Array.isArray(store?.targets) ? store.targets : [];
    for (const s of items) {
        if (s?.sourceID === qid) return s?.sourceLabel || qid;
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            const hit = nodes.find(n => n?.targetID === qid);
            if (hit) return hit?.targetLabel || qid;
        }
    }
    return qid;
}

function buildSourceMap(store) {
    const map = new Map();
    const items = Array.isArray(store?.targets) ? store.targets : [];
    for (const s of items) {
        if (s?.sourceID) map.set(s.sourceID, s);
    }
    return map;
}

function isQidAcquiredAnywhere(store, qid) {
    if (!/^Q\d+$/.test(qid || "")) return false;
    const items = Array.isArray(store?.targets) ? store.targets : [];
    for (const s of items) {
        if (s?.sourceID === qid) return true;
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            if (nodes.some(n => n?.targetID === qid)) return true;
        }
    }
    return false;
}

function extractClickedQid(evt) {
    const path = typeof evt.composedPath === "function" ? evt.composedPath() : [];
    for (const el of path) {
        if (!el || el === document || el === window) continue;
        const href = el.getAttribute?.("href") || "";
        const title = el.getAttribute?.("title") || "";
        const dataQid = el.getAttribute?.("data-qid") || el.getAttribute?.("data-entity-id") || "";
        const text = el.textContent || "";
        const combined = [href, title, dataQid, text].join(" ");
        const m = combined.match(/\bQ\d+\b/);
        if (m) return m[0];
    }
    return "";
}

function findIncomingAcquiredLinks(store, qid, sourceMap) {
    const incoming = [];
    const items = Array.isArray(store?.targets) ? store.targets : [];
    for (const s of items) {
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                if (n?.targetID !== qid) continue;
                incoming.push({
                    sourceID: s.sourceID || "",
                    sourceLabel: s.sourceLabel || s.sourceID || "",
                    propertyID: b.propertyID || "",
                    propertyLabel: b.propertyLabel || b.propertyID || "",
                    isSourceAcquired: sourceMap.has(s.sourceID || "")
                });
            }
        }
    }
    return incoming;
}

function renderAcquiredNodeTree(sourceNode, sourceMap, depth = 0, seen = new Set()) {
    const qid = sourceNode?.sourceID || "";
    const label = sourceNode?.sourceLabel || qid;
    const line = `<div><span class="wdqs-cap-tree-node">${escHtml(label)}</span> <span class="wdqs-cap-tree-qid">(${escHtml(qid)})</span> <span class="wdqs-cap-tree-acquired">acquired</span></div>`;

    if (!qid || seen.has(qid)) {
        return `${line}<ul><li><span class="wdqs-cap-tree-muted">Cycle already shown.</span></li></ul>`;
    }

    if (depth >= 8) {
        return `${line}<ul><li><span class="wdqs-cap-tree-muted">Depth limit reached.</span></li></ul>`;
    }

    const nextSeen = new Set(seen);
    nextSeen.add(qid);

    const childRows = [];
    const buckets = Array.isArray(sourceNode?.targets) ? sourceNode.targets : [];
    for (const b of buckets) {
        const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
        for (const n of nodes) {
            const tid = n?.targetID || "";
            if (!sourceMap.has(tid)) continue;

            const childSource = sourceMap.get(tid);
            const targetLabel = n?.targetLabel || childSource?.sourceLabel || tid;
            const prop = b?.propertyLabel || b?.propertyID || "property";
            const childHtml = renderAcquiredNodeTree(childSource, sourceMap, depth + 1, nextSeen);
            childRows.push(`<li><span class="wdqs-cap-tree-prop">${escHtml(prop)}</span>: <span>${escHtml(targetLabel)}</span> <span class="wdqs-cap-tree-qid">(${escHtml(tid)})</span>${childHtml ? `<ul><li>${childHtml}</li></ul>` : ""}</li>`);
        }
    }

    if (!childRows.length) {
        return `${line}<ul><li><span class="wdqs-cap-tree-muted">No linked child nodes have been acquired yet.</span></li></ul>`;
    }

    return `${line}<ul>${childRows.join("")}</ul>`;
}

function renderTreeSummaryForQid(qid) {
    const store = readStore();
    const sourceMap = buildSourceMap(store);
    const treeBox = document.querySelector(`#${CFG.panelId} .wdqs-cap-tree-box`);
    if (!treeBox) return;

    if (!/^Q\d+$/.test(qid || "")) {
        treeBox.innerHTML = `<span class="wdqs-cap-tree-empty">No QID detected from the clicked node.</span>`;
        return;
    }

    const label = labelForStoreQid(store, qid);
    const acquiredSource = sourceMap.get(qid);
    const acquiredAnywhere = isQidAcquiredAnywhere(store, qid);
    const incoming = findIncomingAcquiredLinks(store, qid, sourceMap);

    if (!acquiredAnywhere) {
        treeBox.innerHTML = `<div><span class="wdqs-cap-tree-node">${escHtml(label)}</span> <span class="wdqs-cap-tree-qid">(${escHtml(qid)})</span></div><div class="wdqs-cap-tree-empty">This node is not in the captured JSON yet.</div>`;
        setStatus(`${qid} is not acquired yet.`, "info");
        return;
    }

    let html = "";
    if (acquiredSource) {
        html += renderAcquiredNodeTree(acquiredSource, sourceMap);
    } else {
        html += `<div><span class="wdqs-cap-tree-node">${escHtml(label)}</span> <span class="wdqs-cap-tree-qid">(${escHtml(qid)})</span> <span class="wdqs-cap-tree-acquired">acquired as target</span></div>`;
        html += `<ul><li><span class="wdqs-cap-tree-muted">This node is present as a target, but its own child summary has not been acquired yet.</span></li></ul>`;
    }

    if (incoming.length) {
        html += `<div style="margin-top:8px;"><span class="wdqs-cap-tree-muted">Incoming acquired links</span></div><ul>`;
        html += incoming.map(x => `<li><span class="wdqs-cap-tree-prop">${escHtml(x.propertyLabel)}</span>: <span>${escHtml(x.sourceLabel)}</span> <span class="wdqs-cap-tree-qid">(${escHtml(x.sourceID)})</span></li>`).join("");
        html += `</ul>`;
    }

    treeBox.innerHTML = html;
    setStatus(`Tree summary shown for ${qid}.`, "ok");
}

function wirePanelBehaviors(panel) {
    if (!panel) return;

    const store = readStore();

    const toggleBtn = panel.querySelector(".wdqs-cap-toggle");
    const toggleIcon = panel.querySelector(".wdqs-cap-toggle-ic");

    const tallBtn = panel.querySelector(".wdqs-cap-tall");
    const tallIcon = panel.querySelector(".wdqs-cap-tall-ic");

    const nameInput = panel.querySelector(".wdqs-cap-name");
    const fullChk = panel.querySelector(".wdqs-cap-full");

    const imgChk = panel.querySelector(".wdqs-cap-img");
    const imgUniqChk = panel.querySelector(".wdqs-cap-img-uniq");

    const treeToggle = panel.querySelector(".wdqs-cap-tree-on");
    const treeWrap = panel.querySelector(".wdqs-cap-tree-wrap");

    // --- mode applier (single source of truth) ---
    function applyMode(state) {
        // state: { collapsed, tall, height }
        panel.classList.toggle("is-collapsed", !!state.collapsed);
        panel.classList.toggle("is-tall", !!state.tall);

        if (state.collapsed) {
            panel.style.height = "";         // let CSS .is-collapsed win
            panel.style.resize = "none";
        } else if (state.tall) {
            panel.style.height = `calc(100vh - 28px)`;
            panel.style.resize = "none";
        } else {
            panel.style.height = `${state.height || 360}px`;
            panel.style.resize = "vertical";
        }

        // keep header reachable; feels like "can't collapse" if header is scrolled away
        panel.scrollTop = 0;

        if (toggleIcon) toggleIcon.textContent = state.collapsed ? "▸" : "▾";
        if (tallIcon) tallIcon.textContent = state.tall ? "⤡" : "⤢";
    }

    // --- init ---
    let ui = readUIState();
    applyMode(ui);

    // --- collapse toggle ---
    if (toggleBtn) {
        toggleBtn.addEventListener("click", () => {
            ui = readUIState();

            const nextCollapsed = !ui.collapsed;

            // when collapsing: keep tall flag as-is in storage? NO.
            // collapsed should override visuals, but we also want a sane restore:
            // keep tall flag, but collapsing always clears inline height anyway.
            const next = { ...ui, collapsed: nextCollapsed };

            writeUIState(next);
            applyMode(next);
        });
    }

    // --- tall toggle ---
    if (tallBtn) {
        tallBtn.addEventListener("click", () => {
            ui = readUIState();

            // If currently collapsed, uncollapse first
            const nextTall = !ui.tall;
            const next = {
                ...ui,
                tall: nextTall,
                collapsed: false // tall implies expanded
            };

            writeUIState(next);
            applyMode(next);
        });
    }

    // --- persist height ONLY in normal mode ---
    let ro;
    try {
        ro = new ResizeObserver(() => {
            const s = readUIState();
            if (s.collapsed) return;
            if (s.tall) return;

            const h = Math.round(panel.getBoundingClientRect().height);
            if (h > 80 && h < 2000) {
                writeUIState({ ...s, height: h });
            }
        });
        ro.observe(panel);
    } catch (e) {
        console.warn("ResizeObserver unavailable:", e);
    }


    // --- tree summary toggle + WDQS node click viewer ---
    if (treeToggle && treeWrap) {
        treeToggle.checked = !!ui.treeSummaryOpen;
        treeWrap.classList.toggle("is-hidden", !treeToggle.checked);

        treeToggle.addEventListener("change", () => {
            const cur = readUIState();
            const next = { ...cur, treeSummaryOpen: !!treeToggle.checked };
            writeUIState(next);
            treeWrap.classList.toggle("is-hidden", !treeToggle.checked);
            setStatus(treeToggle.checked ? "Tree summary viewer ON. Click a Q-node." : "Tree summary viewer OFF.", "info");
        });
    }

    if (!window.__wdqsCapTreeClickInstalled) {
        window.__wdqsCapTreeClickInstalled = true;
        document.addEventListener("click", (evt) => {
            const panelEl = document.getElementById(CFG.panelId);
            if (!panelEl) return;
            if (panelEl.contains(evt.target)) return;

            const on = panelEl.querySelector(".wdqs-cap-tree-on")?.checked;
            if (!on) return;

            const qid = extractClickedQid(evt);
            if (!qid) return;

            panelEl.querySelector(".wdqs-cap-tree-wrap")?.classList.remove("is-hidden");
            renderTreeSummaryForQid(qid);
        }, true);
    }

    // --- store bindings ---
    if (nameInput) {
        nameInput.value = store?.meta?.name || "";
        nameInput.addEventListener("input", () => {
            const s = readStore();
            s.meta = s.meta || {};
            s.meta.name = nameInput.value.trim();
            writeStore(s);
            setStatus("Name saved.", "info");
        });
    }

    if (fullChk) {
        fullChk.checked = !!store?.options?.storeFullResponses;
        fullChk.addEventListener("change", () => {
            const s = readStore();
            s.options = s.options || {};
            s.options.storeFullResponses = !!fullChk.checked;
            writeStore(s);
            setStatus(fullChk.checked ? "Full response storage ON." : "Full response storage OFF.", "info");
        });
    }

    // --- export-time image checkboxes ---
    if (imgChk && imgUniqChk) {
        imgChk.checked = false;
        imgUniqChk.checked = false;
        imgUniqChk.disabled = !imgChk.checked;

        imgChk.addEventListener("change", () => {
            imgUniqChk.disabled = !imgChk.checked;
            if (!imgChk.checked) imgUniqChk.checked = false;
            setStatus(imgChk.checked ? "Export-time image fetch ON." : "Export-time image fetch OFF.", "info");
        });

        imgUniqChk.addEventListener("change", () => {
            if (imgUniqChk.checked && !imgChk.checked) imgUniqChk.checked = false;
            setStatus(imgUniqChk.checked ? "Unique images CSV will be included." : "Unique images CSV disabled.", "info");
        });
    }

    // --- export / copy buttons ---
    const btnCopy = panel.querySelector(".wdqs-btn-copy");
    const btnDl = panel.querySelector(".wdqs-btn-dl");
    const btnCsvNodes = panel.querySelector(".wdqs-btn-csv-nodes");
    const btnCsvProps = panel.querySelector(".wdqs-btn-csv-props");
    const btnCsvLinks = panel.querySelector(".wdqs-btn-csv-links");
    const btnXlsx = panel.querySelector(".wdqs-btn-xlsx");

    async function runPanelAction(label, fn, doneMessage) {
        const report = makeExportProgressReporter(label);
        try {
            setActionsDisabled(true);
            setExportProgress(true, 1, label);
            await fn(report);
            setExportProgress(true, 100, doneMessage || "Done.");
            setStatus(doneMessage || "Done.", "ok");
        } catch (e) {
            console.error("[WDQS-CAP]", label, e);
            setStatus(e?.message || "Action failed.", "warn");
        } finally {
            setActionsDisabled(false);
            setTimeout(() => setExportProgress(false), 700);
        }
    }

    if (btnCopy) {
        btnCopy.addEventListener("click", async () => {
            const s = readStore();
            const ok = await copyText(JSON.stringify(s, null, 2));
            setStatus(ok ? "Copied JSON." : "Copy failed.", ok ? "ok" : "warn");
        });
    }

    if (btnDl) {
        btnDl.addEventListener("click", () => {
            runPanelAction("Downloading JSON…", async () => {
                await exportJsonStore();
            }, "JSON download started.");
        });
    }

    if (btnCsvProps) {
        btnCsvProps.addEventListener("click", () => {
            runPanelAction("Downloading properties CSV…", async () => {
                await exportCsvProperties();
            }, "Properties CSV download started.");
        });
    }

    if (btnCsvLinks) {
        btnCsvLinks.addEventListener("click", () => {
            runPanelAction("Downloading links CSV…", async (report) => {
                await exportCsvLinks({
                    wantImages: !!imgChk?.checked,
                    onProgress: report
                });
            }, "Links CSV download started.");
        });
    }

    if (btnCsvNodes) {
        btnCsvNodes.addEventListener("click", () => {
            runPanelAction("Downloading nodes CSV…", async (report) => {
                await exportCsvNodes({
                    wantImages: !!imgChk?.checked,
                    wantUniqueImagesCsv: !!imgUniqChk?.checked,
                    onProgress: report
                });
            }, "Nodes CSV download started.");
        });
    }

    if (btnXlsx) {
        btnXlsx.addEventListener("click", () => {
            runPanelAction("Downloading XLSX…", async (report) => {
                await exportXlsxSheets({
                    wantImages: !!imgChk?.checked,
                    wantUniqueImagesSheet: !!imgUniqChk?.checked,
                    onProgress: report
                });
            }, "XLSX download started.");
        });
    }

    const btnClear = panel.querySelector(".wdqs-btn-clear");

    if (btnClear) {
        btnClear.addEventListener("click", () => {
            // optional confirm (recommended)
            const ok = confirm("Clear all captured data (targets, captures, qidIndex) and reset the dataset?");
            if (!ok) return;

            const prev = readStore();
            const next = resetStoreToEmptyButKeepMeta(prev);

            writeStore(next);

            refreshCounts(next);
            setPreview({ clearedAt: nowISO(), message: "Dataset cleared." });
            setStatus("Cleared dataset.", "ok");

            console.log("[WDQS-CAP][CLEAR] store reset:", {
                targets: next.targets.length,
                captures: next.captures.length,
                qidIndex: Object.keys(next.qidIndex).length
            });
        });
    }


    // ---- initial render ----
    refreshCounts(store);
    setPreview(store.lastPreview || {});
}

/* ---------------------------
   SECTION 10: Interceptors (fetch + XHR) + queue
---------------------------- */

const CaptureQueue = (() => {
    let q = [];
    let running = false;

    async function drain() {
        if (running) return;
        running = true;

        while (q.length) {
            const job = q.shift();
            try {
                await job();
            } catch (e) {
                console.error("Capture job failed:", e);
            }
        }
        running = false;
    }

    function enqueue(fn) {
        q.push(fn);
        drain();
    }

    return { enqueue };
})();

/* ---------------------------
   Capture router: decide how to process a WDQS JSON response
---------------------------- */
function makeCaptureJob({ url, queryText, data }) {
    return async () => {
        const store = readStore();

        // If the response is property-summary shaped, convert it into a single "source node"
        const headVars = data?.head?.vars || [];
        const isSummary = looksLikePropertySummary(headVars);

        // Try to extract the subject QID from the query (used for summary conversion)
        const subjectQid = extractSubjectQidFromQueryText(queryText);

        // Always store lastPreview for UI
        store.lastPreview = {
            capturedAt: nowISO(),
            kind: isSummary ? "propertySummary" : "qids",
            url,
            subjectQid: subjectQid || "",
            headVars: Array.isArray(headVars) ? headVars : [],
            sample: (() => {
                try {
                    // keep preview light: just show first 1 binding if present
                    const b = data?.results?.bindings;
                    return Array.isArray(b) && b.length ? b[0] : null;
                } catch {
                    return null;
                }
            })()
        };

        // Full-response storage (optional)
        if (store?.options?.storeFullResponses) {
            store.captures = Array.isArray(store.captures) ? store.captures : [];
            store.captures.push({
                capturedAt: nowISO(),
                url,
                query: queryText || "",
                headVars,
                data
            });

            // avoid unbounded growth (lightly)
            if (store.captures.length > 25) store.captures.splice(0, store.captures.length - 25);
        }

        writeStore(store);
        setPreview(store.lastPreview);

        // Main routes
        if (isSummary && subjectQid) {
            // Convert to your "targets" node structure, then merge
            const report = makeCaptureProgressReporter("Building node from summary…");
            setCaptureProgress(true, 1, "Converting property summary…");

            try {
                const nodeObj = await convertWdqsPropertySummaryToTargets({
                    sourceLabel: "",
                    sourceID: subjectQid,
                    wdqsJson: data,
                    onProgress: report
                });

                // merge into store with dedupe
                const next = mergeTargetNodeIntoStore(store, nodeObj);

                writeStore(next);
                refreshCounts(next);
                setPreview(next.lastPreview || store.lastPreview || {});
                setStatus(`Captured summary for ${subjectQid}.`, "ok");
            } catch (e) {
                console.error(e);
                setStatus("Capture failed (summary conversion).", "warn");
            } finally {
                setCaptureProgress(false);
                document.querySelector(`#${CFG.panelId} .wdqs-cap-capprogress`)?.classList.add("is-slot");
            }
            return;
        }

        // Otherwise: treat as general WDQS results → extract QIDs
        const qids = extractQidsFromWdqsResults(data);
        const qidArr = [...qids];

        if (!qidArr.length) {
            // still a valid capture; just nothing extractable
            setStatus("Captured response (no QIDs found).", "info");
            refreshCounts(store);
            return;
        }

        // Merge QIDs into legacy indices if you still want that compatibility
        // (targets structure is for property summary; qidIndex/captures are legacy)
        store.qidIndex = store.qidIndex || {};
        for (const q of qidArr) store.qidIndex[q] = true;

        writeStore(store);
        refreshCounts(store);
        setStatus(`Captured ${humanCount(qidArr.length)} QIDs.`, "ok");
    };
}

/* ---------------------------
   Store merge for "targets" node object (dedupe by sourceID|propertyID|targetID)
   - This is SECTION 6 logic in a reusable helper
---------------------------- */

function mergeTargetNodeIntoStore(store, incomingNodeObj) {
    const cur = store || readStore();
    const merged = Array.isArray(cur.targets) ? [...cur.targets] : [];
    const index = new Set();

    // build existing index
    for (const s of merged) {
        const sid = s?.sourceID || "";
        const buckets = Array.isArray(s?.targets) ? s.targets : [];
        for (const b of buckets) {
            const pid = b?.propertyID || "";
            const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
            for (const n of nodes) {
                const tid = n?.targetID || "";
                index.add(`${sid}|${pid}|${tid}`);
            }
        }
    }

    // incomingNodeObj is a single "source node"
    const sid = incomingNodeObj?.sourceID || "";
    if (!sid) return cur;

    let src = merged.find(x => x?.sourceID === sid);
    if (!src) {
        src = {
            sourceLabel: incomingNodeObj?.sourceLabel || "",
            sourceID: sid,
            sourceWikiDataLink: incomingNodeObj?.sourceWikiDataLink || `https://www.wikidata.org/wiki/${sid}`,
            sourceWikipediaLink: incomingNodeObj?.sourceWikipediaLink || "",
            targets: []
        };
        merged.push(src);
    } else {
        if (!src.sourceLabel && incomingNodeObj?.sourceLabel) src.sourceLabel = incomingNodeObj.sourceLabel;
        if (!src.sourceWikipediaLink && incomingNodeObj?.sourceWikipediaLink) src.sourceWikipediaLink = incomingNodeObj.sourceWikipediaLink;
        if (!src.sourceWikiDataLink && incomingNodeObj?.sourceWikiDataLink) src.sourceWikiDataLink = incomingNodeObj.sourceWikiDataLink;
    }

    const incBuckets = Array.isArray(incomingNodeObj?.targets) ? incomingNodeObj.targets : [];
    for (const b of incBuckets) {
        const pid = b?.propertyID || "";
        if (!pid) continue;

        let buck = (src.targets || []).find(x => x?.propertyID === pid);
        if (!buck) {
            buck = {
                propertyID: pid,
                propertyLabel: b?.propertyLabel || "",
                nodeCount: b?.nodeCount ?? "",
                targetNodes: []
            };
            src.targets.push(buck);
        } else {
            if (!buck.propertyLabel && b?.propertyLabel) buck.propertyLabel = b.propertyLabel;
            if (!buck.nodeCount && b?.nodeCount) buck.nodeCount = b.nodeCount;
        }

        const nodes = Array.isArray(b?.targetNodes) ? b.targetNodes : [];
        for (const n of nodes) {
            const tid = n?.targetID || "";
            const key = `${sid}|${pid}|${tid}`;
            if (!sid || !pid || !tid) continue;

            if (!index.has(key)) {
                buck.targetNodes.push({
                    targetLabel: n?.targetLabel || "",
                    targetID: tid
                });
                index.add(key);
            }
        }
    }

    cur.targets = merged;
    return cur;
}

/* ---------------------------
   fetch() interceptor
---------------------------- */

function installFetchInterceptor() {
    if (!window.fetch) return;

    const origFetch = window.fetch.bind(window);

    window.fetch = async function interceptedFetch(resource, init) {
        const url = extractUrl(resource);
        const isSparql = isSparqlUrl(url);

        // Perform the real fetch first
        const resp = await origFetch(resource, init);

        // Only inspect WDQS sparql JSON responses
        if (!isSparql) return resp;

        // Clone so we don't consume the body
        try {
            const clone = resp.clone();
            if (!isWdqsJsonResponse(clone)) return resp;

            // Try JSON parse
            const data = await clone.json();

            // Extract query text from URL
            const queryText = decodeQueryFromUrl(url);

            // Enqueue capture job (serialized)
            CaptureQueue.enqueue(makeCaptureJob({ url, queryText, data }));
        } catch (e) {
            // ignore parsing errors
        }

        return resp;
    };
}

/* ---------------------------
   XHR interceptor
---------------------------- */

function installXhrInterceptor() {
    const XHROpen = XMLHttpRequest.prototype.open;
    const XHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
        this.__wdqs_url = url;
        this.__wdqs_method = method;
        return XHROpen.call(this, method, url, ...rest);
    };

    XMLHttpRequest.prototype.send = function (body) {
        const xhr = this;

        function onReady() {
            try {
                if (xhr.readyState !== 4) return;

                const url = xhr.__wdqs_url || "";
                if (!isSparqlUrl(url)) return;

                const ct = xhr.getResponseHeader("content-type") || "";
                if (!ct.includes("application/sparql-results+json") && !ct.includes("application/json")) return;

                // Parse JSON safely
                const data = safeJsonParse(xhr.responseText);
                if (!data) return;

                const queryText = decodeQueryFromUrl(url);
                CaptureQueue.enqueue(makeCaptureJob({ url, queryText, data }));
            } catch (e) {
                // ignore
            }
        }

        try {
            xhr.addEventListener("readystatechange", onReady);
        } catch (e) { }

        return XHRSend.call(xhr, body);
    };
}

/* ---------------------------
   Boot: install interceptors + panel
---------------------------- */

function bootWdqsCapture() {
    try {
        ensurePanel(); // SECTION 7
        const panel = document.getElementById(CFG.panelId);
        if (panel) wirePanelBehaviors(panel); // SECTION 9

        installFetchInterceptor();
        installXhrInterceptor();

        setStatus("Interceptor ready.", "ok");
    } catch (e) {
        console.error(e);
    }
}

// Start
bootWdqsCapture();

/* ======================================================
   CENTRAL DEBUG / TRACE SYSTEM (single-point logging)
====================================================== */



function traceFn(fn, name = fn.name || "anonymous") {
    return function tracedFunction(...args) {
        TRACE.group(`CALL ${name}`);
        TRACE.log("args:", args);

        try {
            const result = fn.apply(this, args);

            // handle async functions
            if (result && typeof result.then === "function") {
                return result
                    .then(res => {
                        TRACE.log("return (async):", res);
                        TRACE.groupEnd();
                        return res;
                    })
                    .catch(err => {
                        TRACE.error("error (async):", err);
                        TRACE.groupEnd();
                        throw err;
                    });
            }

            TRACE.log("return:", result);
            TRACE.groupEnd();
            return result;

        } catch (err) {
            TRACE.error("error:", err);
            TRACE.groupEnd();
            throw err;
        }
    };
}

function logDesc(...args) {
    console.log("[WDQS-CAP][DESC]", ...args);
}

