/* SnipCompta — complément Excel d'extraction et de pointage de pièces justificatives.
 * Visionneuse PDF/image, Text / Sum / Validation / Exception / Table Snip, OCR local,
 * recherche multi-documents, rapprochement automatique, liens cellule <-> preuve
 * enregistrés dans le classeur.
 */
"use strict";
(() => {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  const LIB = (path) => new URL("lib/" + path, location.href).href;
  const PDF_WORKER = LIB("pdfjs/pdf.worker.min.js");
  const NUMFMT = "#,##0.00";
  const NAME_PREFIX = "_snip_";
  const DATA_SHEET = "_SnipData";
  const DOCS_SHEET = "_SnipDocs";
  const CHUNK = 30000;          // caractères base64 par cellule (limite Excel 32 767)
  const MAX_STORE_MB = 15;

  const STYLE = {
    text:   { fill: "#DDEBFF", font: "#1F4E9A", bold: false, color: "#2F6FD6", label: "T", name: "Texte" },
    sum:    { fill: "#FFF1C2", font: "#7A5600", bold: false, color: "#E39B00", label: "Σ", name: "Somme" },
    valid:  { fill: "#E3F4E6", font: "#1E7B34", bold: true,  color: "#22A04B", label: "✓", name: "Validation" },
    except: { fill: "#FBE4E2", font: "#B3261E", bold: true,  color: "#D93025", label: "✗", name: "Exception" },
    table:  { fill: "#EFE8FB", font: "#4B2A8A", bold: false, color: "#7B4FD1", label: "▦", name: "Tableau" },
    comment:{ fill: "#FFE8D1", font: "#8A4B00", bold: false, color: "#F28C28", label: "✎", name: "Commentaire" },
  };

  const S = {
    inOffice: false,
    docs: new Map(),      // id -> doc
    cur: null, page: 1, zoom: 1, scale: 1, mode: "nav",
    snips: new Map(),     // id -> snip
    sumParts: [], table: null, commentDraft: null, selSnip: null, flash: null, hits: [],
    opts: { store: true, ocr: true, advance: true, mergeRows: true },
    renderToken: 0, renderTask: null, ocrWorker: null,
  };
  let xl = null;

  /* ====================================================================== */
  /* Utilitaires                                                            */
  /* ====================================================================== */
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const norm = (s) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const round2 = (v) => Math.round(v * 100) / 100;
  const fmtNum = (v) => v.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtVal = (v) => (typeof v === "number" ? fmtNum(v) : String(v));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function status(msg, kind = "") {
    const el = $("#status");
    el.textContent = msg;
    el.className = kind;
  }

  function colLetters(c) {
    let s = ""; c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function lettersToCol(L) { let n = 0; for (const ch of L.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
  const a1 = (r, c) => colLetters(c) + (r + 1);
  const rangeA1 = (r, c, h, w) => (h === 1 && w === 1 ? a1(r, c) : a1(r, c) + ":" + a1(r + h - 1, c + w - 1));

  /** "'Feuil 1'!$B$3:$D$5" -> {sheet, r1, c1, r2, c2, address} */
  function parseRef(ref) {
    if (!ref) return null;
    ref = String(ref).replace(/^=/, "");
    const i = ref.lastIndexOf("!");
    if (i < 0) return null;
    let sheet = ref.slice(0, i);
    const rng = ref.slice(i + 1).replace(/\$/g, "");
    if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    const m = rng.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!m) return null;
    const c1 = lettersToCol(m[1]), r1 = +m[2] - 1;
    const c2 = m[3] ? lettersToCol(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
    return { sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2), address: rng };
  }

  function hashBytes(b) {
    let h1 = 0x811c9dc5, h2 = 0x9747b28c ^ b.length;
    for (let i = 0; i < b.length; i++) {
      h1 = Math.imul(h1 ^ b[i], 16777619);
      h2 = Math.imul(h2 + b[i], 0x5bd1e995); h2 ^= h2 >>> 15;
    }
    return "d" + (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36) + b.length.toString(36);
  }
  function toB64(bytes) {
    let s = ""; const CH = 0x8000;
    for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(s);
  }
  function fromB64(b64) {
    const s = atob(b64); const out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
  }

  /* ---------------- Montants (formats français et anglais) -------------- */
  function parseAmount(raw) {
    if (raw == null) return null;
    let s = String(raw).replace(/[   ']/g, " ").replace(/[€$£]|\bEUR\b|\bUSD\b|\bCHF\b/gi, "").trim();
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
    if (/^[-−–]/.test(s)) { neg = true; s = s.replace(/^[-−–]\s*/, ""); }
    else if (/^\+/.test(s)) s = s.slice(1).trim();
    if (/[-−]$/.test(s)) { neg = true; s = s.slice(0, -1).trim(); }
    if (/\s*(CR|DB|DR)$/i.test(s)) { if (/(DB|DR)$/i.test(s)) neg = true; s = s.replace(/\s*(CR|DB|DR)$/i, ""); }
    if (!/^\d[\d ., ]*$/.test(s) || !/\d$/.test(s)) return null;
    if (s.includes(" ")) {
      if (!/^\d{1,3}( \d{3})+([.,]\d+)?$/.test(s)) return null;
      s = s.replace(/ /g, "");
    }
    const lastDot = s.lastIndexOf("."), lastCom = s.lastIndexOf(",");
    let intPart = s, decPart = "";
    if (lastDot >= 0 && lastCom >= 0) {
      const d = Math.max(lastDot, lastCom);
      const thou = s.slice(0, d), sepT = d === lastDot ? "," : ".";
      if (thou.includes(d === lastDot ? "." : ",")) return null;
      if (!thou.split(sepT).slice(1).every((g) => g.length === 3)) return null;
      intPart = thou.split(sepT).join(""); decPart = s.slice(d + 1);
    } else if (lastDot >= 0 || lastCom >= 0) {
      const sep = lastCom >= 0 ? "," : ".";
      const parts = s.split(sep);
      if (parts.length > 2) {
        if (!parts.slice(1).every((p) => p.length === 3)) return null;
        intPart = parts.join("");
      } else if (sep === "." && parts[1].length === 3 && parts[0] !== "0") {
        intPart = parts.join("");      // 1.500 = mille cinq cents (usage européen)
      } else { intPart = parts[0]; decPart = parts[1]; }
    }
    if (!/^\d+$/.test(intPart) || (decPart && !/^\d+$/.test(decPart))) return null;
    const v = parseFloat(intPart + (decPart ? "." + decPart : ""));
    if (!isFinite(v)) return null;
    return neg ? -v : v;
  }

  const AMT_RE = /(?<![\w.,\/:])[-−+(]?\s?(?:\d{1,3}(?:[   .]\d{3})+(?:[.,]\d{1,2})?|\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:[.,]\d{1,2})?)\)?(?:\s?[-−](?!\d))?(?![\w\/:]|[.,]\d)/g;
  /** Tous les montants d'une ligne de texte, avec leur position. */
  function findAmounts(text) {
    const out = []; let m; AMT_RE.lastIndex = 0;
    while ((m = AMT_RE.exec(text))) {
      const raw = m[0].trim();
      const v = parseAmount(raw);
      if (v == null) continue;
      const start = m.index + (m[0].length - m[0].trimStart().length);
      out.push({ value: v, start, end: start + raw.length, hasDec: /[.,]\d{1,2}\)?\s?[-−]?$/.test(raw) && !/^\d{1,3}([ .]\d{3})+$/.test(raw.replace(/[-−()+\s]$/g, "")) });
    }
    return out;
  }

  /** Texte extrait -> {v: valeur pour Excel, f: format de nombre}. */
  function typed(text) {
    const t = String(text ?? "").trim();
    if (!t) return { v: "", f: "General" };
    if (/^0\d/.test(t) && /^\d+$/.test(t)) return { v: t, f: "@" };          // référence à zéros initiaux
    const v = parseAmount(t);
    if (v == null) return { v: t, f: "General" };
    const hasDec = /[.,]\d{1,2}\)?\s?[-−]?$/.test(t.replace(/\s?(€|EUR)$/i, ""));
    return { v, f: hasDec || !Number.isInteger(v) ? NUMFMT : "General" };
  }

  /* ====================================================================== */
  /* Texte des pages : couche texte PDF ou OCR                              */
  /* ====================================================================== */
  function wordsFromTextContent(tc, vp) {
    const out = [];
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const t = pdfjsLib.Util.transform(vp.transform, it.transform);
      const h = Math.hypot(t[2], t[3]) || Math.abs(it.height) || 8;
      const x = t[4], base = t[5];
      const width = (it.width || h * it.str.length * 0.5) * vp.scale;
      const cw = width / it.str.length;
      const re = /\S+/g; let m;
      while ((m = re.exec(it.str))) {
        out.push({ str: m[0], x0: x + m.index * cw, x1: x + (m.index + m[0].length) * cw, y0: base - h * 0.82, y1: base + h * 0.22 });
      }
    }
    return out;
  }

  /** Regroupe des mots en lignes et calcule le texte de chaque ligne. */
  function groupLines(words) {
    const ws = words.map((w) => ({ ...w, yc: (w.y0 + w.y1) / 2, h: Math.max(1, w.y1 - w.y0) }))
      .sort((a, b) => a.yc - b.yc || a.x0 - b.x0);
    const lines = [];
    for (const w of ws) {
      const L = lines[lines.length - 1];
      if (L && Math.abs(w.yc - L.yc) < 0.5 * Math.min(w.h, L.h)) {
        L.words.push(w);
        L.yc = (L.yc * (L.words.length - 1) + w.yc) / L.words.length;
      } else lines.push({ words: [w], yc: w.yc, h: w.h });
    }
    for (const L of lines) {
      L.words.sort((a, b) => a.x0 - b.x0);
      let text = "", prev = null;
      for (const w of L.words) {
        if (prev && w.x0 - prev.x1 > 0.15 * L.h) text += " ";
        w.start = text.length; text += w.str; w.end = text.length; prev = w;
      }
      L.text = text;
      L.x0 = Math.min(...L.words.map((w) => w.x0)); L.x1 = Math.max(...L.words.map((w) => w.x1));
      L.y0 = Math.min(...L.words.map((w) => w.y0)); L.y1 = Math.max(...L.words.map((w) => w.y1));
    }
    return lines;
  }

  const joinWords = (ws) => (ws.length ? groupLines(ws).map((l) => l.text).join(" ").trim() : "");
  const inRect = (w, r) => { const cx = (w.x0 + w.x1) / 2, cy = (w.y0 + w.y1) / 2; return cx >= r.x0 && cx <= r.x1 && cy >= r.y0 && cy <= r.y1; };
  const wordsInRect = (words, r) => words.filter((w) => inRect(w, r));
  function bboxOf(items, pad = 0) {
    if (!items.length) return null;
    return { x0: Math.min(...items.map((w) => w.x0)) - pad, y0: Math.min(...items.map((w) => w.y0)) - pad,
             x1: Math.max(...items.map((w) => w.x1)) + pad, y1: Math.max(...items.map((w) => w.y1)) + pad };
  }
  function spanBBox(L, start, end) {
    const ws = L.words.filter((w) => w.end > start && w.start < end);
    return bboxOf(ws.length ? ws : L.words, 1.5);
  }

  async function getOcr() {
    if (typeof Tesseract === "undefined") throw new Error("Le moteur OCR n'a pas pu être chargé (fichiers lib/ manquants ?)");
    if (!S.ocrWorker) {
      status("Chargement du moteur OCR (la première fois, ~10-20 s)…", "busy");
      S.ocrWorker = Tesseract.createWorker("fra+eng", 1, {
        workerPath: LIB("tesseract/worker.min.js"),
        corePath: LIB("tesseract/core"),
        langPath: LIB("tesseract/lang"),
      });
      S.ocrWorker.catch(() => { S.ocrWorker = null; });
    }
    return S.ocrWorker;
  }

  async function ocrPage(doc, p, pg) {
    if (pg.ocrPromise) return pg.ocrPromise;
    pg.ocrPromise = (async () => {
      const worker = await getOcr();
      status(`Reconnaissance du texte (OCR) — ${doc.name}, page ${p}…`, "busy");
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      let f;
      if (doc.kind === "pdf") {
        const page = await doc.pdf.getPage(p);
        f = clamp(2500 / pg.w, 1, 4);
        const vp = page.getViewport({ scale: f });
        canvas.width = Math.ceil(vp.width); canvas.height = Math.ceil(vp.height);
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
      } else {
        f = pg.w < 1400 ? 2 : pg.w > 3500 ? 3500 / pg.w : 1;
        canvas.width = Math.ceil(pg.w * f); canvas.height = Math.ceil(pg.h * f);
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(doc.img, 0, 0, canvas.width, canvas.height);
      }
      const { data } = await worker.recognize(canvas);
      pg.words = (data.words || []).filter((w) => w.text && w.text.trim() && w.confidence > 20)
        .map((w) => ({ str: w.text.trim(), x0: w.bbox.x0 / f, y0: w.bbox.y0 / f, x1: w.bbox.x1 / f, y1: w.bbox.y1 / f }));
      pg.lines = groupLines(pg.words);
      pg.ocrDone = true;
      status(`OCR terminé : ${pg.words.length} mots reconnus (${doc.name}, p. ${p}).`, "ok");
    })();
    try { await pg.ocrPromise; } catch (e) { pg.ocrPromise = null; throw e; }
  }

  /** Infos et texte d'une page. allowOcr : lance l'OCR si la page est un scan. */
  async function pageText(doc, p, { allowOcr = S.opts.ocr } = {}) {
    await openDocData(doc);
    let pg = doc.pages[p];
    if (!pg) {
      pg = doc.pages[p] = {};
      if (doc.kind === "pdf") {
        const page = await doc.pdf.getPage(p);
        const vp = page.getViewport({ scale: 1 });
        pg.w = vp.width; pg.h = vp.height;
        const tc = await page.getTextContent();
        pg.words = wordsFromTextContent(tc, vp);
        pg.needsOcr = pg.words.reduce((a, w) => a + w.str.length, 0) < 15;
      } else {
        pg.w = doc.img.naturalWidth; pg.h = doc.img.naturalHeight; pg.words = []; pg.needsOcr = true;
      }
      pg.lines = groupLines(pg.words);
    }
    if (pg.needsOcr && !pg.ocrDone && allowOcr) await ocrPage(doc, p, pg);
    return pg;
  }

  /* ====================================================================== */
  /* Documents                                                              */
  /* ====================================================================== */
  function kindOf(name, mime) {
    if (/pdf/i.test(mime) || /\.pdf$/i.test(name)) return "pdf";
    return "image";
  }

  async function openDocData(doc) {
    if (doc.ready) return doc.ready;
    doc.ready = (async () => {
      if (!doc.bytes) {
        status(`Chargement de « ${doc.name} » depuis le classeur…`, "busy");
        doc.bytes = await xl.loadDocBytes(doc.id);
        if (!doc.bytes) throw new Error(`Le document « ${doc.name} » n'est pas disponible : réimportez-le.`);
      }
      if (doc.kind === "pdf") {
        if (typeof pdfjsLib === "undefined") throw new Error("La visionneuse PDF n'a pas pu être chargée (fichiers lib/ manquants ?)");
        doc.pdf = await pdfjsLib.getDocument({ data: doc.bytes.slice() }).promise;
        doc.numPages = doc.pdf.numPages;
      } else {
        const url = URL.createObjectURL(new Blob([doc.bytes], { type: doc.mime || "image/png" }));
        doc.img = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error("Image illisible")); im.src = url; });
        doc.numPages = 1;
      }
    })();
    try { await doc.ready; } catch (e) { doc.ready = null; throw e; }
  }

  async function importFiles(files) {
    let last = null;
    for (const f of files) {
      try {
        const bytes = new Uint8Array(await f.arrayBuffer());
        const id = hashBytes(bytes);
        let doc = S.docs.get(id);
        if (!doc) {
          doc = { id, name: f.name, mime: f.type || (/\.pdf$/i.test(f.name) ? "application/pdf" : "image/png"), bytes, pages: {}, stored: false };
          doc.kind = kindOf(doc.name, doc.mime);
          S.docs.set(id, doc);
          await openDocData(doc);
          if (S.opts.store && S.inOffice) {
            if (bytes.length > MAX_STORE_MB * 1048576) status(`« ${f.name} » dépasse ${MAX_STORE_MB} Mo : non enregistré dans le classeur.`, "warn");
            else { await xl.saveDoc(doc, (pct) => status(`Enregistrement de « ${f.name} » dans le classeur… ${pct} %`, "busy")); doc.stored = true; }
          }
        } else if (!doc.bytes) { doc.bytes = bytes; }
        last = doc;
        status(`« ${doc.name} » importé (${doc.numPages || "?"} page${doc.numPages > 1 ? "s" : ""}).`, "ok");
      } catch (e) { console.error(e); status(`Erreur à l'import de « ${f.name} » : ${e.message}`, "err"); }
    }
    refreshDocUI();
    if (last) { await openDoc(last.id, 1); switchTab("viewer"); }
  }

  async function openDoc(id, page = 1) {
    const doc = S.docs.get(id);
    if (!doc) return;
    if (S.cur !== id) { S.table = null; }
    S.cur = id; S.page = page;
    $("#docSelect").value = id;
    try { await openDocData(doc); } catch (e) { status(e.message, "err"); return; }
    S.page = clamp(page, 1, doc.numPages);
    await render();
  }

  async function deleteDoc(id) {
    const doc = S.docs.get(id);
    if (!doc) return;
    try { if (S.inOffice && doc.stored) await xl.deleteDoc(id); } catch (e) { status("Suppression impossible : " + e.message, "err"); return; }
    S.docs.delete(id);
    if (S.cur === id) { S.cur = null; const first = S.docs.keys().next().value; if (first) await openDoc(first); else render(); }
    refreshDocUI();
    status(`Document « ${doc.name} » retiré. Les snips liés restent dans Excel.`, "ok");
  }

  function refreshDocUI() {
    const sel = $("#docSelect"), ms = $("#matchDoc");
    const docs = [...S.docs.values()];
    sel.innerHTML = docs.length ? docs.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("") : `<option value="">— aucun document —</option>`;
    if (S.cur) sel.value = S.cur;
    ms.innerHTML = `<option value="*">Tous les documents</option>` + docs.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("");
    const ul = $("#docList");
    ul.innerHTML = docs.length ? "" : `<li class="hint">Aucun document pour l'instant.</li>`;
    for (const d of docs) {
      const li = document.createElement("li");
      const nSn = [...S.snips.values()].filter((s) => s.docId === d.id).length;
      li.innerHTML = `<div class="grow click"><div class="ttl">${esc(d.name)}</div><div class="sub">${d.numPages ? d.numPages + " p. · " : ""}${nSn} snip(s) · ${d.stored ? "enregistré dans le classeur" : "session uniquement"}</div></div>
        <button data-open>Ouvrir</button><button class="danger" data-del>Retirer</button>`;
      li.querySelector("[data-open]").onclick = () => { openDoc(d.id); switchTab("viewer"); };
      li.querySelector(".grow").onclick = () => { openDoc(d.id); switchTab("viewer"); };
      const del = li.querySelector("[data-del]");
      del.onclick = () => {
        if (del.dataset.confirm) { deleteDoc(d.id); return; }
        del.dataset.confirm = "1"; del.textContent = "Confirmer ?";
        setTimeout(() => { del.dataset.confirm = ""; del.textContent = "Retirer"; }, 3000);
      };
      ul.appendChild(li);
    }
  }

  /* ====================================================================== */
  /* Rendu de la page + calque                                              */
  /* ====================================================================== */
  const curDoc = () => S.docs.get(S.cur);

  async function render() {
    const doc = curDoc();
    const token = ++S.renderToken;
    const wrap = $("#pageWrap"), empty = $("#emptyState");
    if (!doc) { wrap.hidden = true; empty.hidden = false; $("#pageInfo").textContent = "–"; return; }
    wrap.hidden = false; empty.hidden = true;
    try {
      const pg = await pageText(doc, S.page, { allowOcr: false });
      if (token !== S.renderToken) return;
      const avail = Math.max(120, $("#viewer").clientWidth - 14);
      const scale = Math.max(0.15, (avail / pg.w) * S.zoom);
      S.scale = scale;
      const dpr = window.devicePixelRatio || 1;
      const canvas = $("#pageCanvas");
      const W = Math.round(pg.w * scale), H = Math.round(pg.h * scale);
      wrap.style.width = W + "px"; wrap.style.height = H + "px";
      canvas.style.width = W + "px"; canvas.style.height = H + "px";
      canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr);
      const ctx = canvas.getContext("2d");
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
      $("#pageInfo").textContent = `${S.page} / ${doc.numPages}`;
      drawOverlay();
      if (doc.kind === "pdf") {
        if (S.renderTask) { try { S.renderTask.cancel(); } catch (_) { /* ignore */ } }
        const page = await doc.pdf.getPage(S.page);
        if (token !== S.renderToken) return;
        const vp = page.getViewport({ scale: scale * dpr });
        S.renderTask = page.render({ canvasContext: ctx, viewport: vp });
        try { await S.renderTask.promise; } catch (e) { if (e && e.name === "RenderingCancelledException") return; throw e; }
      } else {
        ctx.drawImage(doc.img, 0, 0, canvas.width, canvas.height);
      }
      if (token !== S.renderToken) return;
      // OCR en arrière-plan si la page est un scan
      if (pg.needsOcr && !pg.ocrDone && S.opts.ocr) {
        pageText(doc, S.page, { allowOcr: true }).catch((e) => status("OCR impossible : " + e.message, "err"));
      }
    } catch (e) { console.error(e); status("Affichage impossible : " + e.message, "err"); }
  }

  function box(ov, r, cls, color, label, title) {
    const sc = S.scale;
    const d = document.createElement("div");
    d.className = "ovbox " + cls;
    d.style.left = r.x0 * sc + "px"; d.style.top = r.y0 * sc + "px";
    d.style.width = Math.max(2, (r.x1 - r.x0) * sc) + "px"; d.style.height = Math.max(2, (r.y1 - r.y0) * sc) + "px";
    if (color) { d.style.borderColor = color; d.style.background = color + "1f"; }
    if (title) d.title = title;
    if (label) { const b = document.createElement("span"); b.className = "badge"; b.style.background = color || "#555"; b.textContent = label; d.appendChild(b); }
    ov.appendChild(d);
    return d;
  }

  function drawOverlay() {
    const ov = $("#overlay");
    ov.innerHTML = "";
    ov.classList.toggle("drawing", S.mode !== "nav");
    const doc = curDoc(); if (!doc) return;
    const p = S.page, sc = S.scale;
    for (const sn of S.snips.values()) {
      for (const r of sn.rects) {
        if ((r.docId || sn.docId) !== doc.id || r.page !== p) continue;
        const st = STYLE[sn.type];
        const showBadge = sn.type !== "table" || /_0$/.test(sn.id) || S.selSnip === sn.id;
        box(ov, r, "snip" + (S.selSnip === sn.id ? " sel" : ""), st.color, showBadge ? st.label : "", `${st.name} → ${sn.sheet}!${sn.address}`);
      }
    }
    for (const part of S.sumParts) {
      if (part.docId === doc.id && part.page === p) box(ov, part.rect, "part", STYLE.sum.color, fmtNum(part.total));
    }
    for (const h of S.hits) if (h.docId === doc.id && h.page === p) box(ov, h.rect, "hit");
    if (S.flash && S.flash.docId === doc.id && S.flash.page === p) box(ov, S.flash.rect, "flash");
    const cd = S.commentDraft;
    if (cd && cd.docId === doc.id && cd.page === p) box(ov, cd.rect, "part", STYLE.comment.color, "✎");
    const t = S.table;
    if (t && t.docId === doc.id && t.page === p) {
      const r = t.rect;
      const tb = document.createElement("div");
      tb.className = "tablebox";
      tb.style.left = r.x0 * sc + "px"; tb.style.top = r.y0 * sc + "px";
      tb.style.width = (r.x1 - r.x0) * sc + "px"; tb.style.height = (r.y1 - r.y0) * sc + "px";
      for (const y of t.rowYs || []) {
        const rl = document.createElement("div"); rl.className = "rowline"; rl.style.top = (y - r.y0) * sc + "px"; tb.appendChild(rl);
      }
      t.seps.forEach((x, i) => {
        const s = document.createElement("div");
        s.className = "sep"; s.dataset.i = i; s.style.left = (x - r.x0) * sc + "px";
        s.title = "Glisser pour déplacer, double-clic pour supprimer";
        tb.appendChild(s);
      });
      ov.appendChild(tb);
    }
  }

  /* ---------------- Interaction souris / tactile ------------------------ */
  let drag = null;
  function evPt(e) {
    const r = $("#overlay").getBoundingClientRect();
    return { x: (e.clientX - r.left) / S.scale, y: (e.clientY - r.top) / S.scale };
  }
  function bindOverlay() {
    const ov = $("#overlay");
    ov.addEventListener("pointerdown", (e) => {
      if (e.button !== 0 || !curDoc()) return;
      if (e.target.classList.contains("sep")) {
        drag = { kind: "sep", idx: +e.target.dataset.i, el: e.target };
        ov.setPointerCapture(e.pointerId); e.preventDefault(); return;
      }
      const pt = evPt(e);
      drag = { kind: "rect", x: pt.x, y: pt.y, moved: false, el: null };
      if (S.mode !== "nav") { ov.setPointerCapture(e.pointerId); e.preventDefault(); }
    });
    ov.addEventListener("pointermove", (e) => {
      if (!drag) return;
      const pt = evPt(e);
      if (drag.kind === "sep") {
        const t = S.table; if (!t) return;
        t.seps[drag.idx] = clamp(pt.x, t.rect.x0 + 2, t.rect.x1 - 2);
        drag.el.style.left = (t.seps[drag.idx] - t.rect.x0) * S.scale + "px";
        throttledPreview();
        return;
      }
      if (S.mode === "nav") return;
      if (!drag.moved && Math.hypot(pt.x - drag.x, pt.y - drag.y) * S.scale > 4) drag.moved = true;
      if (drag.moved) {
        const r = { x0: Math.min(drag.x, pt.x), y0: Math.min(drag.y, pt.y), x1: Math.max(drag.x, pt.x), y1: Math.max(drag.y, pt.y) };
        if (!drag.el) drag.el = box(ov, r, "drag");
        drag.el.style.left = r.x0 * S.scale + "px"; drag.el.style.top = r.y0 * S.scale + "px";
        drag.el.style.width = (r.x1 - r.x0) * S.scale + "px"; drag.el.style.height = (r.y1 - r.y0) * S.scale + "px";
      }
    });
    const end = (e) => {
      if (!drag) return;
      const d = drag; drag = null;
      if (d.kind === "sep") { S.table.seps.sort((a, b) => a - b); drawOverlay(); updateTablePreview(); return; }
      const pt = evPt(e);
      if (d.el) d.el.remove();
      if (!d.moved || S.mode === "nav") { handleClick(pt); return; }
      const r = { x0: Math.min(d.x, pt.x), y0: Math.min(d.y, pt.y), x1: Math.max(d.x, pt.x), y1: Math.max(d.y, pt.y) };
      handleRect(r).catch((err) => { console.error(err); status("Erreur : " + err.message, "err"); });
    };
    ov.addEventListener("pointerup", end);
    ov.addEventListener("pointercancel", () => { if (drag && drag.el) drag.el.remove(); drag = null; });
    ov.addEventListener("dblclick", (e) => {
      if (e.target.classList.contains("sep") && S.table) {
        S.table.seps.splice(+e.target.dataset.i, 1); drawOverlay(); updateTablePreview();
      }
    });
  }

  function handleClick(pt) {
    const doc = curDoc(); if (!doc) return;
    const t = S.table;
    if (t && t.docId === doc.id && t.page === S.page && pt.x > t.rect.x0 && pt.x < t.rect.x1 && pt.y > t.rect.y0 && pt.y < t.rect.y1) {
      if (!t.seps.some((x) => Math.abs(x - pt.x) * S.scale < 6)) { t.seps.push(pt.x); t.seps.sort((a, b) => a - b); drawOverlay(); updateTablePreview(); }
      return;
    }
    let hit = null;
    for (const sn of S.snips.values()) {
      for (const r of sn.rects) {
        if ((r.docId || sn.docId) === doc.id && r.page === S.page && pt.x >= r.x0 && pt.x <= r.x1 && pt.y >= r.y0 && pt.y <= r.y1) {
          if (!hit || sn.created > hit.created) hit = sn;
        }
      }
    }
    if (hit) {
      S.selSnip = hit.id; drawOverlay();
      xl.selectAddr(hit.sheet, hit.address).catch((e) => status("Cellule introuvable : " + e.message, "warn"));
      status(`${STYLE[hit.type].name} : ${hit.sheet}!${hit.address} = ${fmtVal(hit.value)}`);
    } else if (S.selSnip) { S.selSnip = null; drawOverlay(); }
  }

  async function handleRect(r) {
    switch (S.mode) {
      case "text": return doTextSnip(r);
      case "sum": return addSumPart(r);
      case "valid": case "except": return doMarkSnip(S.mode, r);
      case "table": return startTable(r);
      case "comment": return startComment(r);
    }
  }

  /* ====================================================================== */
  /* Les snips                                                              */
  /* ====================================================================== */
  async function placeSnip(type, value, rects, fmt) {
    const doc = curDoc();
    const ac = await xl.activeCell();
    const id = uid();
    const address = await xl.write(ac.sheet, ac.row, ac.col, [[value]], type, fmt ? [[fmt]] : null);
    const sn = { id, group: id, type, sheet: ac.sheet, address, docId: rects[0].docId || doc.id, docName: doc.name, rects, value, created: Date.now() };
    S.snips.set(id, sn);
    await xl.addLinks([sn]);
    await persistSnips();
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + 1, ac.col);
    S.selSnip = id;
    drawOverlay(); refreshSnipList();
    status(`${STYLE[type].name} → ${ac.sheet}!${address} : ${fmtVal(value)}`, "ok");
  }

  async function doTextSnip(r) {
    const doc = curDoc(), p = S.page;
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    const text = joinWords(ws);
    if (!text) return status(pg.needsOcr && !S.opts.ocr ? "Page scannée : activez l'OCR dans l'onglet Documents." : "Aucun texte dans la zone sélectionnée.", "warn");
    const tv = typed(text);
    await placeSnip("text", tv.v, [{ docId: doc.id, page: p, ...bboxOf(ws, 2) }], tv.f);
  }

  async function doMarkSnip(type, r) {
    const doc = curDoc(), p = S.page;
    let rect = r;
    try {
      const pg = await pageText(doc, p, { allowOcr: true });
      const ws = wordsInRect(pg.words, r);
      if (ws.length) rect = bboxOf(ws, 2);
    } catch (_) { /* sans texte : on garde la zone dessinée */ }
    await placeSnip(type, type === "valid" ? "✓" : "✗", [{ docId: doc.id, page: p, ...rect }]);
  }

  /* ---------------- Sum Snip ------------------------------------------- */
  async function addSumPart(r) {
    const doc = curDoc(), p = S.page;
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    const lines = groupLines(ws);
    let amts = [];
    for (const L of lines) for (const a of findAmounts(L.text)) amts.push({ ...a, bbox: spanBBox(L, a.start, a.end) });
    if (amts.some((a) => a.hasDec)) amts = amts.filter((a) => a.hasDec);
    if (!amts.length) return status("Aucun montant trouvé dans cette zone.", "warn");
    const total = round2(amts.reduce((s, a) => s + a.value, 0));
    S.sumParts.push({ docId: doc.id, page: p, rect: bboxOf(ws, 2), values: amts.map((a) => a.value), total });
    drawOverlay(); showSumPanel();
  }

  function showSumPanel() {
    const panel = $("#ctxPanel");
    if (!S.sumParts.length) { panel.hidden = true; return; }
    const total = round2(S.sumParts.reduce((s, p) => s + p.total, 0));
    panel.hidden = false;
    panel.innerHTML = `<div><strong>Somme en cours</strong> <span class="hint">(dessinez d'autres zones, même sur d'autres pages)</span></div>
      <ul class="parts list">${S.sumParts.map((p, i) => `<li><span>p.${p.page} · ${p.values.length} montant(s)</span><span>${fmtNum(p.total)} <button data-rm="${i}" title="Retirer">×</button></span></li>`).join("")}</ul>
      <div class="total">Total : ${fmtNum(total)}</div>
      <div class="row"><button class="primary" id="sumInsert">Insérer le total (Entrée)</button><button id="sumClear">Vider</button></div>`;
    $$("[data-rm]", panel).forEach((b) => (b.onclick = () => { S.sumParts.splice(+b.dataset.rm, 1); drawOverlay(); showSumPanel(); }));
    $("#sumInsert").onclick = () => insertSum().catch((e) => status("Erreur : " + e.message, "err"));
    $("#sumClear").onclick = () => { S.sumParts = []; drawOverlay(); showSumPanel(); };
  }

  async function insertSum() {
    if (!S.sumParts.length) return;
    const total = round2(S.sumParts.reduce((s, p) => s + p.total, 0));
    const rects = S.sumParts.map((p) => ({ docId: p.docId, page: p.page, ...p.rect }));
    S.sumParts = [];
    showSumPanel();
    await placeSnip("sum", total, rects);
  }

  /* ---------------- Table Snip ----------------------------------------- */
  function autoColumns(words, r) {
    const n = Math.max(1, Math.ceil(r.x1 - r.x0));
    const cov = new Uint16Array(n);
    for (const w of words) {
      const a = clamp(Math.floor(w.x0 - r.x0), 0, n - 1), b = clamp(Math.ceil(w.x1 - r.x0), 0, n - 1);
      for (let i = a; i <= b; i++) cov[i]++;
    }
    const nLines = groupLines(words).length;
    const tol = Math.floor(nLines * 0.06);
    let first = 0, last = n - 1;
    while (first < n && cov[first] <= tol) first++;
    while (last > 0 && cov[last] <= tol) last--;
    const seps = [];
    let i = first;
    while (i <= last) {
      if (cov[i] <= tol) {
        let j = i; while (j <= last && cov[j] <= tol) j++;
        if (j - i >= 6) seps.push(r.x0 + (i + j) / 2);
        i = j;
      } else i++;
    }
    return seps;
  }

  function extractTable(pg, rect, seps, merge) {
    const words = wordsInRect(pg.words, rect);
    const lines = groupLines(words);
    const cols = [rect.x0, ...[...seps].sort((a, b) => a - b), rect.x1];
    const nc = cols.length - 1;
    const rows = [];
    for (const L of lines) {
      const buckets = Array.from({ length: nc }, () => []);
      for (const w of L.words) {
        const xc = (w.x0 + w.x1) / 2;
        let k = cols.findIndex((x, idx) => idx > 0 && xc < x) - 1;
        if (k < 0) k = nc - 1;
        buckets[clamp(k, 0, nc - 1)].push(w);
      }
      const cells = buckets.map(joinWords);
      const row = { cells, y0: L.y0, y1: L.y1, x0: rect.x0, x1: rect.x1 };
      const prev = rows[rows.length - 1];
      const isCont = merge && prev && nc > 1 && !cells[0] && !cells.some((c) => c && parseAmount(c) != null);
      if (isCont) {
        cells.forEach((c, i) => { if (c) prev.cells[i] = prev.cells[i] ? prev.cells[i] + " " + c : c; });
        prev.y1 = L.y1;
      } else rows.push(row);
    }
    return rows;
  }

  async function startTable(r) {
    const doc = curDoc(), p = S.page;
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    if (!ws.length) return status("Aucun texte dans la zone : élargissez la sélection ou activez l'OCR.", "warn");
    S.table = { docId: doc.id, page: p, rect: r, seps: autoColumns(ws, r), from: p, to: p };
    drawOverlay(); updateTablePreview();
    status("Tableau détecté : ajustez les colonnes si besoin, puis « Insérer dans Excel ».", "ok");
  }

  let previewTimer = null;
  const throttledPreview = () => { clearTimeout(previewTimer); previewTimer = setTimeout(updateTablePreview, 120); };

  function updateTablePreview() {
    const t = S.table, panel = $("#ctxPanel");
    if (!t) { if (!S.sumParts.length) panel.hidden = true; return; }
    const doc = S.docs.get(t.docId);
    const pg = doc.pages[t.page];
    const rows = extractTable(pg, t.rect, t.seps, S.opts.mergeRows);
    t.rowYs = rows.slice(1).map((r) => r.y0 - 1);
    const show = rows.slice(0, 8);
    const nc = t.seps.length + 1;
    panel.hidden = false;
    panel.innerHTML = `<div><strong>Tableau</strong> — ${rows.length} ligne(s) × ${nc} colonne(s) sur cette page
        <div class="hint">Cliquez dans le cadre violet pour ajouter une colonne, glissez un trait pour le déplacer, double-cliquez pour le supprimer.</div></div>
      <table>${show.map((r) => `<tr>${r.cells.map((c) => `<td class="${parseAmount(c) != null ? "num" : ""}">${esc(c)}</td>`).join("")}</tr>`).join("")}</table>
      ${rows.length > show.length ? `<div class="hint">… ${rows.length - show.length} autre(s) ligne(s)</div>` : ""}
      <div class="row">Pages <input type="number" id="tFrom" min="1" max="${doc.numPages}" value="${t.from}"> à <input type="number" id="tTo" min="1" max="${doc.numPages}" value="${t.to}">
        <span class="hint">(même zone sur chaque page)</span></div>
      <label class="chk"><input type="checkbox" id="tMerge" ${S.opts.mergeRows ? "checked" : ""}> Fusionner les lignes de suite (1<sup>re</sup> colonne vide, sans montant)</label>
      <div class="row"><button class="primary" id="tInsert">Insérer dans Excel (Entrée)</button><button id="tAuto">Colonnes auto</button><button id="tCancel">Annuler</button></div>`;
    $("#tFrom").onchange = (e) => { t.from = clamp(+e.target.value || 1, 1, doc.numPages); };
    $("#tTo").onchange = (e) => { t.to = clamp(+e.target.value || t.page, 1, doc.numPages); };
    $("#tMerge").onchange = (e) => { S.opts.mergeRows = e.target.checked; saveOpts(); updateTablePreview(); drawOverlay(); };
    $("#tInsert").onclick = () => insertTable().catch((e) => { console.error(e); status("Erreur : " + e.message, "err"); });
    $("#tAuto").onclick = () => { t.seps = autoColumns(wordsInRect(pg.words, t.rect), t.rect); drawOverlay(); updateTablePreview(); };
    $("#tCancel").onclick = () => { S.table = null; drawOverlay(); updateTablePreview(); };
    drawOverlay();
  }

  async function insertTable() {
    const t = S.table; if (!t) return;
    const doc = S.docs.get(t.docId);
    const from = Math.min(t.from, t.to), to = Math.max(t.from, t.to);
    const all = [];
    for (let p = from; p <= to; p++) {
      status(`Extraction du tableau — page ${p}/${to}…`, "busy");
      const pg = await pageText(doc, p, { allowOcr: true });
      for (const r of extractTable(pg, t.rect, t.seps, S.opts.mergeRows)) all.push({ ...r, page: p });
    }
    if (!all.length) return status("Aucune ligne extraite.", "warn");
    const nc = t.seps.length + 1;
    const tv = all.map((r) => r.cells.map(typed));
    const values = tv.map((r) => r.map((c) => c.v)), fmts = tv.map((r) => r.map((c) => c.f));
    const ac = await xl.activeCell();
    await xl.write(ac.sheet, ac.row, ac.col, values, "table", fmts);
    const group = uid();
    const made = all.map((r, i) => ({
      id: group + "_" + i, group, type: "table", sheet: ac.sheet, address: rangeA1(ac.row + i, ac.col, 1, nc),
      docId: doc.id, docName: doc.name, rects: [{ docId: doc.id, page: r.page, x0: r.x0, y0: r.y0 - 1, x1: r.x1, y1: r.y1 + 1 }],
      value: r.cells.filter(Boolean).join(" | "), created: Date.now() + i / 1000,
    }));
    for (const sn of made) S.snips.set(sn.id, sn);
    await xl.addLinks(made);
    await persistSnips();
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + all.length, ac.col);
    S.table = null; updateTablePreview(); drawOverlay(); refreshSnipList();
    status(`Tableau inséré : ${all.length} ligne(s) × ${nc} colonne(s) en ${ac.sheet}!${rangeA1(ac.row, ac.col, all.length, nc)}.`, "ok");
  }

  /* ---------------- Comment Snip --------------------------------------- */
  async function startComment(r) {
    const doc = curDoc(), p = S.page;
    let rect = r, found = "";
    try {
      const pg = await pageText(doc, p, { allowOcr: true });
      const ws = wordsInRect(pg.words, r);
      if (ws.length) { rect = bboxOf(ws, 2); found = joinWords(ws); }
    } catch (_) { /* zone sans texte */ }
    S.commentDraft = { docId: doc.id, page: p, rect, found };
    drawOverlay();
    const panel = $("#ctxPanel");
    panel.hidden = false;
    panel.innerHTML = `<div><strong>Commentaire</strong> <span class="hint">— ajouté en commentaire Excel sur la cellule active, avec lien vers cette zone</span></div>
      ${found ? `<div class="hint">Zone : « ${esc(found.slice(0, 120))} »</div>` : ""}
      <textarea id="cText" rows="3" style="width:100%;font:inherit;border:1px solid var(--line);border-radius:6px;padding:4px" placeholder="Ex. : montant différent de la facture, à voir avec le client"></textarea>
      <div class="row"><button class="primary" id="cAdd">Ajouter le commentaire</button><button id="cCancel">Annuler</button></div>`;
    const ta = $("#cText"); setTimeout(() => ta.focus(), 0);
    ta.onkeydown = (e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) $("#cAdd").click(); };
    $("#cCancel").onclick = () => { S.commentDraft = null; panel.hidden = true; drawOverlay(); };
    $("#cAdd").onclick = () => insertComment(ta.value).catch((e) => status("Commentaire impossible : " + e.message, "err"));
  }

  async function insertComment(text) {
    const cd = S.commentDraft; if (!cd) return;
    text = String(text || "").trim() || "Voir pièce";
    const doc = S.docs.get(cd.docId);
    const ac = await xl.activeCell();
    const address = a1(ac.row, ac.col);
    await xl.addComment(ac.sheet, ac.row, ac.col, `${text}\n— ${doc.name}, p. ${cd.page}`);
    await xl.format([{ sheet: ac.sheet, address, type: "comment" }]);
    const id = uid();
    const sn = { id, group: id, auto: true, type: "comment", sheet: ac.sheet, address, docId: doc.id, docName: doc.name,
      rects: [{ docId: doc.id, page: cd.page, ...cd.rect }], value: text, created: Date.now() };
    S.snips.set(id, sn);
    await xl.addLinks([sn]);
    await persistSnips();
    S.commentDraft = null; $("#ctxPanel").hidden = true;
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + 1, ac.col);
    drawOverlay(); refreshSnipList();
    status(`Commentaire ajouté en ${ac.sheet}!${address}.`, "ok");
  }

  /* ---------------- Trouver les sommes (Find all sums) ----------------- */
  /** Suites de montants consécutifs, alignés en colonne, dont le total vaut target. */
  function sumRuns(pg, target, tol) {
    const items = [];
    for (const L of pg.lines) for (const a of findAmounts(L.text)) if (a.hasDec) items.push({ v: Math.abs(a.value), bbox: spanBBox(L, a.start, a.end) });
    const cols = [];
    for (const it of items) {
      let c = cols.find((c) => Math.abs(c.x1 - it.bbox.x1) < 6);
      if (!c) { c = { x1: it.bbox.x1, items: [] }; cols.push(c); }
      c.items.push(it);
    }
    const runs = [];
    for (const c of cols) {
      const arr = c.items.sort((a, b) => a.bbox.y0 - b.bbox.y0);
      for (let i = 0; i < arr.length; i++) {
        let sum = 0;
        for (let j = i; j < arr.length; j++) {
          sum += arr[j].v;
          if (j > i && Math.abs(sum - target) <= tol) runs.push(arr.slice(i, j + 1));
          if (sum - target > tol) break;
        }
      }
    }
    return runs.sort((a, b) => a.length - b.length);
  }

  async function findSums() {
    const ul = $("#matchResults"), info = $("#matchInfo");
    ul.innerHTML = "";
    if (!S.docs.size) { info.textContent = "Importez d'abord les pièces justificatives."; return; }
    const sel = await xl.selected();
    const tol = Math.abs(parseAmount($("#matchTol").value) ?? 0.01) + 1e-9;
    const scope = $("#matchDoc").value;
    const docs = scope === "*" ? [...S.docs.values()] : [S.docs.get(scope)].filter(Boolean);
    const pages = await allPages(docs, true, (d, p) => status(`Lecture de « ${d.name} », p. ${p}…`, "busy"));
    const created = [], formats = [], results = [];
    const group = uid();
    for (let i = 0; i < sel.values.length; i++) {
      for (let j = 0; j < sel.values[i].length; j++) {
        const raw = sel.values[i][j];
        const num = typeof raw === "number" ? raw : parseAmount(raw);
        if (num == null || num === 0) continue;
        const addr = a1(sel.row + i, sel.col + j);
        let found = null;
        for (const { doc, p, pg } of pages) {
          const runs = sumRuns(pg, Math.abs(num), tol);
          if (runs.length) { found = { doc, p, run: runs[0] }; break; }
        }
        if (found) {
          const sn = { id: group + "_" + created.length, group, auto: true, type: "sum", sheet: sel.sheet, address: addr, docId: found.doc.id, docName: found.doc.name,
            rects: found.run.map((it) => ({ docId: found.doc.id, page: found.p, ...it.bbox })), value: raw, created: Date.now() + created.length / 1000 };
          created.push(sn); formats.push({ sheet: sel.sheet, address: addr, type: "sum" });
          results.push({ ok: true, addr, raw, sn, n: found.run.length });
        } else {
          formats.push({ sheet: sel.sheet, address: addr, type: "except" });
          results.push({ ok: false, addr, raw });
        }
      }
    }
    if (!results.length) { info.textContent = "Sélectionnez dans Excel des cellules contenant des totaux."; status("Aucun montant dans la sélection.", "warn"); return; }
    await xl.format(formats);
    for (const sn of created) S.snips.set(sn.id, sn);
    if (created.length) { await xl.addLinks(created); await persistSnips(); }
    const ok = results.filter((r) => r.ok).length;
    info.innerHTML = `Trouver les sommes : <span class="ok">${ok} total(aux) retrouvé(s)</span> · <span class="ko">${results.length - ok} non trouvé(s)</span>.`;
    for (const r of results) {
      const li = document.createElement("li");
      li.className = "click";
      li.innerHTML = `<span class="chip" style="background:${r.ok ? STYLE.sum.color : STYLE.except.color}">${r.ok ? "Σ" : "✗"}</span>
        <div class="grow"><div class="ttl">${esc(sel.sheet)}!${r.addr} — ${esc(fmtVal(r.raw))}</div>
        <div class="sub">${r.ok ? `somme de ${r.n} montants · ${esc(r.sn.docName)} · page ${r.sn.rects[0].page}` : "aucune suite de montants ne donne ce total"}</div></div>`;
      li.onclick = r.ok ? () => goToSnip(r.sn) : () => xl.selectAddr(sel.sheet, r.addr).catch(() => {});
      ul.appendChild(li);
    }
    refreshSnipList(); drawOverlay();
    status(`Trouver les sommes : ${ok}/${results.length} total(aux) justifié(s).`, "ok");
  }

  /* ---------------- Export de la piste d'audit -------------------------- */
  async function exportSnips() {
    const all = [...S.snips.values()].sort((a, b) => a.created - b.created);
    if (!all.length) return status("Aucun snip à exporter.", "warn");
    const rows = all.map((sn) => ({ type: STYLE[sn.type].name, sheet: sn.sheet, address: sn.address, value: sn.value, doc: sn.docName, page: sn.rects[0].page, date: new Date(sn.created).toLocaleString("fr-FR") }));
    const name = await xl.exportAudit(rows);
    status(`Piste d'audit exportée dans la feuille « ${name} » (${rows.length} ligne(s)).`, "ok");
  }

  async function deleteSnipsInSelection() {
    const sel = await xl.selected();
    const r2 = sel.row + sel.values.length - 1, c2 = sel.col + (sel.values[0] || []).length - 1;
    const hit = [...S.snips.values()].filter((sn) => {
      const ref = parseRef(`'${sn.sheet.replace(/'/g, "''")}'!${sn.address}`);
      return ref && ref.sheet === sel.sheet && ref.r1 <= r2 && ref.r2 >= sel.row && ref.c1 <= c2 && ref.c2 >= sel.col;
    });
    if (!hit.length) return status("Aucun snip dans la sélection Excel.", "warn");
    await deleteSnips(hit.map((s) => s.id));
    status(`${hit.length} snip(s) supprimé(s) dans la sélection (les valeurs sont conservées).`, "ok");
  }

  async function deleteAllSnips() {
    const ids = [...S.snips.keys()];
    if (!ids.length) return status("Aucun snip dans ce classeur.", "warn");
    await deleteSnips(ids);
    status(`${ids.length} snip(s) supprimé(s) (les valeurs sont conservées).`, "ok");
  }

  async function ocrCurrentDoc() {
    const doc = curDoc();
    if (!doc) return status("Ouvrez d'abord un document.", "warn");
    await openDocData(doc);
    let done = 0;
    for (let p = 1; p <= doc.numPages; p++) {
      const pg = await pageText(doc, p, { allowOcr: false });
      if (pg.needsOcr && !pg.ocrDone) { await ocrPage(doc, p, pg); done++; }
    }
    status(done ? `OCR terminé sur ${done} page(s) de « ${doc.name} ».` : `« ${doc.name} » contient déjà du texte lisible : pas besoin d'OCR.`, "ok");
  }

  async function toggleStore() {
    S.opts.store = !S.opts.store;
    $("#optStore").checked = S.opts.store; saveOpts();
    if (S.opts.store && S.inOffice) {
      for (const d of S.docs.values()) {
        if (d.stored || !d.bytes || d.bytes.length > MAX_STORE_MB * 1048576) continue;
        await xl.saveDoc(d, (pct) => status(`Enregistrement de « ${d.name} »… ${pct} %`, "busy")); d.stored = true;
      }
      refreshDocUI();
    }
    status(S.opts.store ? "Les PDF sont maintenant enregistrés dans le classeur." : "Les nouveaux PDF ne seront plus enregistrés dans le classeur (liens conservés).", "ok");
  }

  /* ---------------- Suppression / annulation --------------------------- */
  async function persistSnips() {
    try { await xl.saveSnips([...S.snips.values()]); }
    catch (e) { console.error(e); status("Snips non enregistrés dans le classeur : " + e.message, "warn"); }
  }

  async function deleteSnips(ids, { clear = false } = {}) {
    const list = ids.map((id) => S.snips.get(id)).filter(Boolean);
    if (!list.length) return;
    await xl.removeLinks(list, { clear });
    for (const sn of list) S.snips.delete(sn.id);
    if (list.some((s) => s.id === S.selSnip)) S.selSnip = null;
    await persistSnips();
    drawOverlay(); refreshSnipList();
  }

  async function undoLast() {
    const all = [...S.snips.values()];
    if (!all.length) return status("Rien à annuler.", "warn");
    const last = all.reduce((a, b) => (b.created > a.created ? b : a));
    const grp = all.filter((s) => s.group === last.group);
    await deleteSnips(grp.map((s) => s.id), { clear: !last.auto });
    status(`Annulé : ${STYLE[last.type].name} (${grp.length > 1 ? grp.length + " lignes" : last.sheet + "!" + last.address}).`, "ok");
  }

  /* ---------------- Navigation cellule <-> preuve ---------------------- */
  async function goToLocation(docId, page, rect) {
    if (!S.docs.has(docId)) return false;
    switchTab("viewer");
    S.flash = rect ? { docId, page, rect } : null;
    await openDoc(docId, page);
    if (rect) {
      const v = $("#viewer");
      v.scrollTop = Math.max(0, rect.y0 * S.scale - 60);
      v.scrollLeft = Math.max(0, rect.x0 * S.scale - 30);
      setTimeout(() => { if (S.flash && S.flash.rect === rect) { S.flash = null; drawOverlay(); } }, 1700);
    }
    return true;
  }

  async function goToSnip(sn) {
    const r = sn.rects[0];
    S.selSnip = sn.id;
    const ok = await goToLocation(r.docId || sn.docId, r.page, r);
    if (!ok) status(`Document « ${sn.docName} » absent : réimportez-le pour voir la preuve.`, "warn");
    else status(`${STYLE[sn.type].name} : ${sn.sheet}!${sn.address} = ${fmtVal(sn.value)} — ${sn.docName}, p. ${r.page}`);
  }

  let selTimer = null;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(async () => {
      try {
        const info = await xl.selectionInfo();
        // met à jour les adresses (lignes insérées, cellules déplacées…)
        for (const [id, ref] of Object.entries(info.links)) {
          const sn = S.snips.get(id);
          if (sn && ref) { sn.sheet = ref.sheet; sn.address = ref.address; }
        }
        let hit = null;
        for (const [id, ref] of Object.entries(info.links)) {
          if (!ref || ref.sheet !== info.sheet) continue;
          if (info.row >= ref.r1 && info.row <= ref.r2 && info.col >= ref.c1 && info.col <= ref.c2) {
            const sn = S.snips.get(id);
            if (sn && (!hit || sn.created > hit.created)) hit = sn;
          }
        }
        if (hit && hit.id !== S.selSnip) await goToSnip(hit);
      } catch (e) { console.warn(e); }
    }, 180);
  }

  /* ---------------- Liste des snips ------------------------------------ */
  function refreshSnipList() {
    const ul = $("#snipList");
    const all = [...S.snips.values()].sort((a, b) => b.created - a.created);
    $("#snipCount").textContent = all.length ? `${all.length} snip(s) dans ce classeur. Cliquez pour voir la preuve et la cellule.` : "Aucun snip pour l'instant.";
    ul.innerHTML = "";
    for (const sn of all.slice(0, 500)) {
      const st = STYLE[sn.type];
      const li = document.createElement("li");
      li.innerHTML = `<span class="chip" style="background:${st.color}">${st.label}</span>
        <div class="grow click"><div class="ttl">${esc(sn.sheet)}!${esc(sn.address)} — ${esc(fmtVal(sn.value))}</div>
        <div class="sub">${esc(sn.docName)} · p. ${sn.rects[0].page}${sn.auto ? " · rapprochement" : ""}</div></div>
        <button class="danger" title="Supprimer le lien (la valeur reste dans la cellule)">×</button>`;
      li.querySelector(".grow").onclick = async () => { await goToSnip(sn); xl.selectAddr(sn.sheet, sn.address).catch(() => {}); };
      li.querySelector("button").onclick = () => deleteSnips([sn.id]).then(() => status("Lien supprimé (la valeur est conservée).", "ok"));
      ul.appendChild(li);
    }
    if (all.length > 500) ul.insertAdjacentHTML("beforeend", `<li class="hint">… ${all.length - 500} de plus</li>`);
    refreshDocUISoft();
  }
  function refreshDocUISoft() { if ($("#tab-docs").classList.contains("active")) refreshDocUI(); }

  /* ====================================================================== */
  /* Recherche multi-documents                                              */
  /* ====================================================================== */
  async function allPages(docs, allowOcr, onProgress) {
    const out = [];
    let i = 0;
    for (const doc of docs) {
      try { await openDocData(doc); } catch (e) { status(e.message, "warn"); continue; }
      for (let p = 1; p <= doc.numPages; p++) {
        onProgress && onProgress(doc, p, ++i);
        const pg = await pageText(doc, p, { allowOcr });
        out.push({ doc, p, pg });
      }
    }
    return out;
  }

  async function doSearch(q) {
    q = q.trim();
    const ul = $("#searchResults"), info = $("#searchInfo");
    ul.innerHTML = ""; S.hits = [];
    if (!q) return;
    if (!S.docs.size) { info.textContent = "Aucun document à parcourir."; return; }
    const qn = norm(q), qAmt = parseAmount(q);
    const allowOcr = $("#searchOcr").checked;
    status("Recherche…", "busy");
    const pages = await allPages([...S.docs.values()], allowOcr, (d, p) => status(`Recherche dans « ${d.name} », p. ${p}…`, "busy"));
    const res = []; let skipped = 0;
    for (const { doc, p, pg } of pages) {
      if (pg.needsOcr && !pg.ocrDone) skipped++;
      for (const L of pg.lines) {
        const ln = norm(L.text);
        const found = [];
        let k = ln.indexOf(qn);
        while (k >= 0 && qn) { found.push([k, k + qn.length]); k = ln.indexOf(qn, k + 1); }
        if (qAmt != null) for (const a of findAmounts(L.text)) if (Math.abs(Math.abs(a.value) - Math.abs(qAmt)) < 0.005 && !found.some(([s, e]) => s < a.end && e > a.start)) found.push([a.start, a.end]);
        for (const [s, e] of found) res.push({ doc, p, L, s, e, rect: spanBBox(L, s, e) });
      }
      if (res.length > 300) break;
    }
    info.textContent = `${res.length} résultat(s)${skipped ? ` · ${skipped} page(s) scannée(s) non lue(s) (cochez l'option OCR)` : ""}.`;
    S.hits = res.map((r) => ({ docId: r.doc.id, page: r.p, rect: r.rect }));
    for (const r of res.slice(0, 300)) {
      const li = document.createElement("li");
      li.className = "click";
      const t = r.L.text;
      li.innerHTML = `<div class="grow"><div class="ttl">${esc(t.slice(Math.max(0, r.s - 40), r.s))}<mark>${esc(t.slice(r.s, r.e))}</mark>${esc(t.slice(r.e, r.e + 50))}</div>
        <div class="sub">${esc(r.doc.name)} · page ${r.p}</div></div>`;
      li.onclick = () => goToLocation(r.doc.id, r.p, r.rect);
      ul.appendChild(li);
    }
    drawOverlay();
    status(`Recherche terminée : ${res.length} résultat(s).`, "ok");
  }

  /* ====================================================================== */
  /* Rapprochement automatique (Document Matching)                          */
  /* ====================================================================== */
  async function runMatch() {
    const btn = $("#matchRun"); btn.disabled = true;
    const ul = $("#matchResults"), info = $("#matchInfo");
    ul.innerHTML = "";
    try {
      if (!S.docs.size) { info.textContent = "Importez d'abord les pièces justificatives."; return; }
      const sel = await xl.selected();
      const scope = $("#matchDoc").value;
      const docs = scope === "*" ? [...S.docs.values()] : [S.docs.get(scope)].filter(Boolean);
      const tol = Math.abs(parseAmount($("#matchTol").value) ?? 0.01) + 1e-9;
      const ignoreSign = $("#matchAbs").checked, once = $("#matchOnce").checked;
      const pages = await allPages(docs, true, (d, p) => status(`Lecture de « ${d.name} », p. ${p}…`, "busy"));
      // index des montants et des lignes
      const amts = [], lines = [];
      for (const { doc, p, pg } of pages) {
        pg.lines.forEach((L, li) => {
          lines.push({ doc, p, L, key: `${doc.id}:${p}:${li}`, n: norm(L.text) });
          for (const a of findAmounts(L.text)) amts.push({ doc, p, L, a, key: `${doc.id}:${p}:${li}` });
        });
      }
      const used = new Set();
      const created = [], formats = [], results = [];
      const group = uid();
      for (let i = 0; i < sel.values.length; i++) {
        for (let j = 0; j < sel.values[i].length; j++) {
          const raw = sel.values[i][j];
          if (raw === "" || raw == null) continue;
          const row = sel.row + i, col = sel.col + j, addr = a1(row, col);
          let num = typeof raw === "number" ? raw : parseAmount(raw);
          let best = null;
          if (num != null) {
            for (const c of amts) {
              if (once && used.has(c.key)) continue;
              const d = ignoreSign ? Math.abs(Math.abs(c.a.value) - Math.abs(num)) : Math.abs(c.a.value - num);
              if (d <= tol && (!best || d < best.d)) { best = { d, doc: c.doc, p: c.p, rect: spanBBox(c.L, c.a.start, c.a.end), key: c.key }; if (d < 1e-9) break; }
            }
          }
          if (!best && typeof raw === "string" && raw.trim().length >= 2) {
            const qn = norm(raw.trim());
            for (const c of lines) {
              if (once && used.has(c.key)) continue;
              const k = c.n.indexOf(qn);
              if (k >= 0) { best = { doc: c.doc, p: c.p, rect: spanBBox(c.L, k, k + qn.length), key: c.key }; break; }
            }
          }
          if (best) {
            if (once) used.add(best.key);
            const sn = { id: group + "_" + created.length, group, auto: true, type: "valid", sheet: sel.sheet, address: addr, docId: best.doc.id, docName: best.doc.name,
              rects: [{ docId: best.doc.id, page: best.p, ...best.rect }], value: raw, created: Date.now() + created.length / 1000 };
            created.push(sn); formats.push({ sheet: sel.sheet, address: addr, type: "valid" });
            results.push({ ok: true, addr, raw, sn });
          } else {
            formats.push({ sheet: sel.sheet, address: addr, type: "except" });
            results.push({ ok: false, addr, raw });
          }
        }
      }
      if (!results.length) { info.textContent = "La sélection Excel est vide."; status("Sélectionnez des cellules à rapprocher.", "warn"); return; }
      await xl.format(formats);
      for (const sn of created) S.snips.set(sn.id, sn);
      if (created.length) { await xl.addLinks(created); await persistSnips(); }
      const ok = results.filter((r) => r.ok).length;
      info.innerHTML = `<span class="ok">${ok} trouvé(s)</span> · <span class="ko">${results.length - ok} non trouvé(s)</span> sur ${results.length} cellule(s).`;
      for (const r of results) {
        const li = document.createElement("li");
        li.className = r.ok ? "click" : "";
        li.innerHTML = `<span class="chip" style="background:${r.ok ? STYLE.valid.color : STYLE.except.color}">${r.ok ? "✓" : "✗"}</span>
          <div class="grow"><div class="ttl">${esc(sel.sheet)}!${r.addr} — ${esc(fmtVal(r.raw))}</div>
          <div class="sub">${r.ok ? esc(r.sn.docName) + " · page " + r.sn.rects[0].page : "introuvable dans les documents"}</div></div>`;
        if (r.ok) li.onclick = () => goToSnip(r.sn);
        else li.onclick = () => xl.selectAddr(sel.sheet, r.addr).catch(() => {});
        ul.appendChild(li);
      }
      refreshSnipList(); drawOverlay();
      status(`Rapprochement terminé : ${ok}/${results.length} justifié(s).`, "ok");
    } catch (e) { console.error(e); status("Rapprochement impossible : " + e.message, "err"); }
    finally { btn.disabled = false; }
  }

  /* ====================================================================== */
  /* Adaptateur Excel (Office.js)                                           */
  /* ====================================================================== */
  function cellOut(v) {
    if (typeof v === "number") return v;
    const s = String(v ?? "");
    return /^[=+\-@]/.test(s) && parseAmount(s) == null ? "'" + s : s;
  }
  function applyStyle(rg, type) {
    const st = STYLE[type];
    rg.format.fill.color = st.fill;
    rg.format.font.color = st.font;
    rg.format.font.bold = st.bold;
  }
  async function ensureSheet(ctx, name) {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    await ctx.sync();
    if (ws.isNullObject) {
      const active = ctx.workbook.worksheets.getActiveWorksheet();
      ws = ctx.workbook.worksheets.add(name);
      ws.visibility = "VeryHidden";
      active.activate();
      await ctx.sync();
    }
    return ws;
  }

  const OfficeXL = {
    async activeCell() {
      return Excel.run(async (ctx) => {
        const c = ctx.workbook.getActiveCell();
        c.load("rowIndex,columnIndex");
        const ws = c.worksheet; ws.load("name");
        await ctx.sync();
        if (ws.name === DATA_SHEET || ws.name === DOCS_SHEET) throw new Error("feuille technique sélectionnée");
        return { sheet: ws.name, row: c.rowIndex, col: c.columnIndex };
      });
    },
    async write(sheet, row, col, values, type, fmts) {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        const h = values.length, w = Math.max(...values.map((r) => r.length));
        const vals = values.map((r) => { const a = r.map(cellOut); while (a.length < w) a.push(""); return a; });
        const rg = ws.getRangeByIndexes(row, col, h, w);
        rg.numberFormat = vals.map((r, i) => r.map((v, j) => (fmts && fmts[i] && fmts[i][j]) || (typeof v === "number" ? NUMFMT : "General")));
        rg.values = vals;
        applyStyle(rg, type);
        await ctx.sync();
        return rangeA1(row, col, h, w);
      });
    },
    async select(sheet, row, col) {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        ws.getRangeByIndexes(row, col, 1, 1).select();
        await ctx.sync();
      });
    },
    async selectAddr(sheet, address) {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        ws.activate(); ws.getRange(address).select();
        await ctx.sync();
      });
    },
    async selected() {
      return Excel.run(async (ctx) => {
        const rg = ctx.workbook.getSelectedRange();
        rg.load("values,rowIndex,columnIndex,rowCount,columnCount");
        const ws = rg.worksheet; ws.load("name");
        await ctx.sync();
        if (rg.rowCount * rg.columnCount > 20000) throw new Error("sélection trop grande (20 000 cellules max)");
        return { sheet: ws.name, row: rg.rowIndex, col: rg.columnIndex, values: rg.values };
      });
    },
    async format(list) {
      return Excel.run(async (ctx) => {
        for (const f of list) applyStyle(ctx.workbook.worksheets.getItem(f.sheet).getRange(f.address), f.type);
        await ctx.sync();
      });
    },
    async addLinks(snips) {
      const add = (withVisible) => Excel.run(async (ctx) => {
        for (const sn of snips) {
          const rg = ctx.workbook.worksheets.getItem(sn.sheet).getRange(sn.address);
          const nm = ctx.workbook.names.add(NAME_PREFIX + sn.id, rg, "SnipCompta");
          if (withVisible) nm.visible = false;
        }
        await ctx.sync();
      });
      try { await add(true); } catch (e) { await add(false); }
    },
    async removeLinks(snips, { clear }) {
      return Excel.run(async (ctx) => {
        const names = snips.map((sn) => ctx.workbook.names.getItemOrNullObject(NAME_PREFIX + sn.id));
        names.forEach((n) => n.load("name"));
        await ctx.sync();
        names.forEach((nm, i) => {
          const sn = snips[i];
          let rg = null;
          if (!nm.isNullObject) rg = nm.getRange();
          else { try { rg = ctx.workbook.worksheets.getItem(sn.sheet).getRange(sn.address); } catch (_) { rg = null; } }
          if (rg) {
            rg.format.fill.clear();
            rg.format.font.color = "#000000";
            rg.format.font.bold = false;
            if (clear) rg.clear("Contents");
          }
          if (!nm.isNullObject) nm.delete();
        });
        await ctx.sync();
      }).then(() => OfficeXL._removeComments(snips.filter((s) => s.type === "comment")));
    },
    async _removeComments(list) {
      if (!list.length || !Office.context.requirements.isSetSupported("ExcelApi", "1.10")) return;
      try {
        await Excel.run(async (ctx) => {
          for (const sn of list) {
            const c = ctx.workbook.comments.getItemByCell(ctx.workbook.worksheets.getItem(sn.sheet).getRange(sn.address));
            c.delete();
          }
          await ctx.sync();
        });
      } catch (_) { /* commentaire déjà supprimé à la main */ }
    },
    async addComment(sheet, row, col, text) {
      if (!Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
        // Excel trop ancien : le texte est écrit dans la cellule
        return OfficeXL.write(sheet, row, col, [[text]], "comment");
      }
      return Excel.run(async (ctx) => {
        const rg = ctx.workbook.worksheets.getItem(sheet).getRangeByIndexes(row, col, 1, 1);
        const existing = ctx.workbook.comments.getItemByCell(rg);
        existing.load("content");
        try { await ctx.sync(); existing.content = existing.content + "\n\n" + text; }
        catch (_) { ctx.workbook.comments.add(rg, text); }
        await ctx.sync();
      });
    },
    async exportAudit(rows) {
      return Excel.run(async (ctx) => {
        const wss = ctx.workbook.worksheets; wss.load("items/name");
        await ctx.sync();
        const names = new Set(wss.items.map((w) => w.name));
        let name = "Piste d'audit", k = 2;
        while (names.has(name)) name = `Piste d'audit (${k++})`;
        const ws = wss.add(name);
        const q = (sh) => "'" + sh.replace(/'/g, "''") + "'";
        const data = [["Type", "Feuille", "Cellule", "Valeur", "Document", "Page", "Créé le"],
          ...rows.map((r) => [r.type, r.sheet, `=HYPERLINK("#${q(r.sheet).replace(/"/g, '""')}!${r.address.split(":")[0]}","${r.address}")`,
            typeof r.value === "number" ? r.value : cellOut(r.value), cellOut(r.doc), r.page, r.date])];
        const rg = ws.getRangeByIndexes(0, 0, data.length, 7);
        rg.formulas = data;
        const head = ws.getRange("A1:G1");
        head.format.font.bold = true; head.format.fill.color = "#DDEBFF";
        ws.freezePanes.freezeRows(1);
        rg.format.autofitColumns();
        ws.activate();
        await ctx.sync();
        return name;
      });
    },
    async selectionInfo() {
      return Excel.run(async (ctx) => {
        const c = ctx.workbook.getActiveCell();
        c.load("rowIndex,columnIndex");
        const ws = c.worksheet; ws.load("name");
        const names = ctx.workbook.names; names.load("items/name,items/value");
        await ctx.sync();
        const links = {};
        for (const n of names.items) if (n.name.startsWith(NAME_PREFIX)) links[n.name.slice(NAME_PREFIX.length)] = parseRef(n.value);
        return { sheet: ws.name, row: c.rowIndex, col: c.columnIndex, links };
      });
    },
    async saveSnips(arr) {
      return Excel.run(async (ctx) => {
        const ws = await ensureSheet(ctx, DATA_SHEET);
        ws.getRange().clear();
        const rows = [["id", "json"], ...arr.map((s) => [s.id, JSON.stringify(s)])];
        for (let i = 0; i < rows.length; i += 500) {
          const part = rows.slice(i, i + 500);
          const rg = ws.getRangeByIndexes(i, 0, part.length, 2);
          rg.numberFormat = part.map(() => ["@", "@"]);
          rg.values = part;
          await ctx.sync();
        }
      });
    },
    async loadSnips() {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(DATA_SHEET);
        await ctx.sync();
        if (ws.isNullObject) return [];
        const ur = ws.getUsedRangeOrNullObject(true); ur.load("values");
        const names = ctx.workbook.names; names.load("items/name,items/value");
        await ctx.sync();
        if (ur.isNullObject) return [];
        const links = {};
        for (const n of names.items) if (n.name.startsWith(NAME_PREFIX)) links[n.name.slice(NAME_PREFIX.length)] = parseRef(n.value);
        const out = [];
        for (const [id, json] of ur.values.slice(1)) {
          try {
            const sn = JSON.parse(json);
            const ref = links[id];
            if (ref) { sn.sheet = ref.sheet; sn.address = ref.address; }
            out.push(sn);
          } catch (_) { /* ligne illisible */ }
        }
        return out;
      });
    },
    async saveDoc(doc, onProgress) {
      const b64 = toB64(doc.bytes);
      const chunks = [];
      for (let i = 0; i < b64.length; i += CHUNK) chunks.push("~" + b64.slice(i, i + CHUNK));
      return Excel.run(async (ctx) => {
        const ws = await ensureSheet(ctx, DOCS_SHEET);
        const ur = ws.getUsedRangeOrNullObject(true); ur.load("rowIndex,rowCount");
        await ctx.sync();
        const start = ur.isNullObject ? 0 : ur.rowIndex + ur.rowCount;
        const B = 25;
        for (let i = 0; i < chunks.length; i += B) {
          const part = chunks.slice(i, i + B).map((c, k) => [doc.id, "~" + doc.name, doc.mime, i + k, chunks.length, c]);
          const rg = ws.getRangeByIndexes(start + i, 0, part.length, 6);
          rg.numberFormat = part.map(() => ["@", "@", "@", "0", "0", "@"]);
          rg.values = part;
          await ctx.sync();
          onProgress && onProgress(Math.round(((i + part.length) / chunks.length) * 100));
        }
      });
    },
    async listDocs() {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(DOCS_SHEET);
        await ctx.sync();
        if (ws.isNullObject) return [];
        const ur = ws.getUsedRangeOrNullObject(true); ur.load("rowIndex,rowCount");
        await ctx.sync();
        if (ur.isNullObject) return [];
        const rg = ws.getRangeByIndexes(0, 0, ur.rowIndex + ur.rowCount, 5); rg.load("values");
        await ctx.sync();
        const seen = new Map();
        for (const [id, name, mime] of rg.values) if (id && !seen.has(id)) seen.set(id, { id, name: String(name).replace(/^~/, ""), mime });
        return [...seen.values()];
      });
    },
    async _rowsOf(ctx, ws, id) {
      const ur = ws.getUsedRangeOrNullObject(true); ur.load("rowIndex,rowCount");
      await ctx.sync();
      if (ur.isNullObject) return null;
      const col = ws.getRangeByIndexes(0, 0, ur.rowIndex + ur.rowCount, 1); col.load("values");
      await ctx.sync();
      let r1 = -1, r2 = -1;
      col.values.forEach(([v], i) => { if (v === id) { if (r1 < 0) r1 = i; r2 = i; } });
      return r1 < 0 ? null : { r1, r2 };
    },
    async loadDocBytes(id) {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(DOCS_SHEET);
        await ctx.sync();
        if (ws.isNullObject) return null;
        const rows = await OfficeXL._rowsOf(ctx, ws, id);
        if (!rows) return null;
        const rg = ws.getRangeByIndexes(rows.r1, 3, rows.r2 - rows.r1 + 1, 3); rg.load("values");
        await ctx.sync();
        const parts = rg.values.slice().sort((a, b) => a[0] - b[0]).map((r) => String(r[2]).replace(/^~/, ""));
        return fromB64(parts.join(""));
      });
    },
    async deleteDoc(id) {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(DOCS_SHEET);
        await ctx.sync();
        if (ws.isNullObject) return;
        const rows = await OfficeXL._rowsOf(ctx, ws, id);
        if (!rows) return;
        ws.getRangeByIndexes(rows.r1, 0, rows.r2 - rows.r1 + 1, 6).delete("Up");
        await ctx.sync();
      });
    },
    onSelection(cb) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, cb);
    },
  };

  /* ====================================================================== */
  /* Adaptateur démo (hors d'Excel) : mini-feuille en mémoire               */
  /* ====================================================================== */
  const DemoXL = (() => {
    const SHEET = "Feuil1";
    const cells = new Map();
    let act = { r: 0, c: 0 }, selEnd = null, selCb = null;
    const key = (r, c) => r + "," + c;
    function grid() {
      let R = 12, C = 6;
      for (const k of cells.keys()) { const [r, c] = k.split(",").map(Number); R = Math.max(R, r + 3); C = Math.max(C, c + 2); }
      R = Math.max(R, act.r + 3);
      const r1 = Math.min(act.r, selEnd ? selEnd.r : act.r), r2 = Math.max(act.r, selEnd ? selEnd.r : act.r);
      const c1 = Math.min(act.c, selEnd ? selEnd.c : act.c), c2 = Math.max(act.c, selEnd ? selEnd.c : act.c);
      let h = "<table><tr><th></th>" + Array.from({ length: C }, (_, c) => `<th>${colLetters(c)}</th>`).join("") + "</tr>";
      for (let r = 0; r < R; r++) {
        h += `<tr><th>${r + 1}</th>`;
        for (let c = 0; c < C; c++) {
          const v = cells.get(key(r, c));
          const st = v && v.type ? STYLE[v.type] : null;
          const cls = [r === act.r && c === act.c ? "act" : "", r >= r1 && r <= r2 && c >= c1 && c <= c2 && (r !== act.r || c !== act.c) ? "insel" : "", v && typeof v.v === "number" ? "num" : ""].join(" ");
          const style = st ? `background:${st.fill};color:${st.font};${st.bold ? "font-weight:700" : ""}` : "";
          h += `<td data-r="${r}" data-c="${c}" class="${cls}" style="${style}">${v ? esc(typeof v.v === "number" ? fmtNum(v.v) : v.v) : ""}</td>`;
        }
        h += "</tr>";
      }
      $("#demoGrid").innerHTML = h + "</table>";
    }
    function fire() { if (selCb) selCb(); }
    function bind() {
      $("#demoGrid").addEventListener("click", (e) => {
        const td = e.target.closest("td"); if (!td) return;
        const r = +td.dataset.r, c = +td.dataset.c;
        if (e.shiftKey) selEnd = { r, c }; else { act = { r, c }; selEnd = null; }
        grid(); fire();
      });
      grid();
    }
    const range = () => {
      const e = selEnd || act;
      return { r1: Math.min(act.r, e.r), r2: Math.max(act.r, e.r), c1: Math.min(act.c, e.c), c2: Math.max(act.c, e.c) };
    };
    return {
      bind,
      _cells: cells,
      async activeCell() { return { sheet: SHEET, row: act.r, col: act.c }; },
      async write(sheet, row, col, values, type) {
        const w = Math.max(...values.map((r) => r.length));
        values.forEach((r, i) => { for (let j = 0; j < w; j++) cells.set(key(row + i, col + j), { v: r[j] ?? "", type }); });
        grid();
        return rangeA1(row, col, values.length, w);
      },
      async select(sheet, row, col) { act = { r: row, c: col }; selEnd = null; grid(); },
      async selectAddr(sheet, address) {
        const ref = parseRef(SHEET + "!" + address);
        act = { r: ref.r1, c: ref.c1 }; selEnd = ref.r2 !== ref.r1 || ref.c2 !== ref.c1 ? { r: ref.r2, c: ref.c2 } : null; grid();
      },
      async selected() {
        const g = range();
        const values = [];
        for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const v = cells.get(key(r, c)); row.push(v ? v.v : ""); } values.push(row); }
        return { sheet: SHEET, row: g.r1, col: g.c1, values };
      },
      async format(list) {
        for (const f of list) { const ref = parseRef(SHEET + "!" + f.address); for (let r = ref.r1; r <= ref.r2; r++) for (let c = ref.c1; c <= ref.c2; c++) { const v = cells.get(key(r, c)) || { v: "" }; v.type = f.type; cells.set(key(r, c), v); } }
        grid();
      },
      async addLinks() {},
      async removeLinks(snips, { clear }) {
        for (const sn of snips) {
          const ref = parseRef(SHEET + "!" + sn.address);
          for (let r = ref.r1; r <= ref.r2; r++) for (let c = ref.c1; c <= ref.c2; c++) {
            if (clear) cells.delete(key(r, c)); else { const v = cells.get(key(r, c)); if (v) v.type = null; }
          }
        }
        grid();
      },
      async selectionInfo() {
        const links = {};
        for (const sn of S.snips.values()) links[sn.id] = parseRef(sn.sheet + "!" + sn.address);
        return { sheet: SHEET, row: act.r, col: act.c, links };
      },
      async addComment(sheet, row, col, text) { const v = cells.get(key(row, col)) || { v: "" }; v.note = text; cells.set(key(row, col), v); grid(); },
      async exportAudit(rows) { this._audit = rows; return "Piste d'audit"; },
      async saveSnips() {}, async loadSnips() { return []; },
      async saveDoc() {}, async listDocs() { return []; }, async loadDocBytes() { return null; }, async deleteDoc() {},
      onSelection(cb) { selCb = cb; },
    };
  })();

  /* ====================================================================== */
  /* Interface                                                              */
  /* ====================================================================== */
  function switchTab(name) {
    $$(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
    $$(".tab").forEach((t) => t.classList.toggle("active", t.id === "tab-" + name));
    if (name === "docs") refreshDocUI();
    if (name === "snips") refreshSnipList();
    if (name === "viewer") setTimeout(render, 0);
  }

  function setMode(m) {
    S.mode = m;
    $$(".modes button").forEach((b) => b.classList.toggle("active", b.dataset.mode === m));
    if (m !== "table" && S.table) { S.table = null; updateTablePreview(); }
    if (m !== "comment" && S.commentDraft) { S.commentDraft = null; $("#ctxPanel").hidden = true; }
    drawOverlay();
    const tips = {
      nav: "Navigation : cliquez sur une zone colorée pour aller à sa cellule Excel.",
      text: "Text Snip : sélectionnez une cellule Excel, puis entourez le texte ou le montant dans le document.",
      sum: "Sum Snip : entourez un ou plusieurs montants (plusieurs zones possibles), puis « Insérer le total ».",
      valid: "Validation Snip : entourez l'élément vérifié — un ✓ vert est inscrit dans la cellule active.",
      except: "Exception Snip : entourez l'élément en anomalie — un ✗ rouge est inscrit dans la cellule active.",
      table: "Table Snip : entourez tout le tableau (sans le titre), ajustez les colonnes puis insérez.",
      comment: "Commentaire : entourez la zone concernée, tapez votre remarque — elle devient un commentaire Excel lié à la pièce.",
    };
    status(tips[m]);
  }

  function loadOpts() {
    try { Object.assign(S.opts, JSON.parse(localStorage.getItem("snipcompta.opts") || "{}")); } catch (_) { /* stockage indisponible */ }
    $("#optStore").checked = S.opts.store; $("#optOcr").checked = S.opts.ocr; $("#optAdvance").checked = S.opts.advance;
  }
  function saveOpts() { try { localStorage.setItem("snipcompta.opts", JSON.stringify(S.opts)); } catch (_) { /* ignore */ } }

  function bindUI() {
    $$(".tabs button").forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
    $$(".modes button").forEach((b) => (b.onclick = () => setMode(b.dataset.mode)));
    $$("[data-action=import]").forEach((b) => (b.onclick = () => $("#fileInput").click()));
    $("#fileInput").onchange = (e) => { const fs = [...e.target.files]; e.target.value = ""; importFiles(fs); };
    $("#docSelect").onchange = (e) => e.target.value && openDoc(e.target.value, 1);
    $("#prevPage").onclick = () => { const d = curDoc(); if (d && S.page > 1) { S.page--; render(); } };
    $("#nextPage").onclick = () => { const d = curDoc(); if (d && S.page < d.numPages) { S.page++; render(); } };
    $("#zoomIn").onclick = () => { S.zoom = Math.min(5, S.zoom * 1.25); render(); };
    $("#zoomOut").onclick = () => { S.zoom = Math.max(0.3, S.zoom / 1.25); render(); };
    $("#undoBtn").onclick = $("#undoBtn2").onclick = () => undoLast().catch((e) => status("Annulation impossible : " + e.message, "err"));
    $("#refreshSnips").onclick = async () => { try { await reloadSnips(); status("Snips actualisés.", "ok"); } catch (e) { status(e.message, "err"); } };
    $("#searchForm").onsubmit = (e) => { e.preventDefault(); doSearch($("#searchInput").value).catch((err) => status("Recherche impossible : " + err.message, "err")); };
    $("#matchRun").onclick = runMatch;
    $("#sumsRun").onclick = () => findSums().catch((e) => status("Recherche des sommes impossible : " + e.message, "err"));
    $("#exportBtn").onclick = () => exportSnips().catch((e) => status("Export impossible : " + e.message, "err"));
    $("#delSelBtn").onclick = () => deleteSnipsInSelection().catch((e) => status("Suppression impossible : " + e.message, "err"));
    $("#optStore").onchange = (e) => { S.opts.store = e.target.checked; saveOpts(); };
    $("#optOcr").onchange = (e) => { S.opts.ocr = e.target.checked; saveOpts(); };
    $("#optAdvance").onchange = (e) => { S.opts.advance = e.target.checked; saveOpts(); };

    document.addEventListener("keydown", (e) => {
      if (e.target.matches("input, select, textarea")) return;
      if (e.key === "Enter") {
        if (S.table) { e.preventDefault(); insertTable().catch((err) => status("Erreur : " + err.message, "err")); }
        else if (S.sumParts.length) { e.preventDefault(); insertSum().catch((err) => status("Erreur : " + err.message, "err")); }
      } else if (e.key === "Escape") {
        S.table = null; S.sumParts = []; S.commentDraft = null; $("#ctxPanel").hidden = true; drawOverlay();
      } else if (/^[1-7]$/.test(e.key) && !e.ctrlKey && !e.altKey && !e.metaKey) {
        setMode(["nav", "text", "sum", "valid", "except", "table", "comment"][+e.key - 1]);
      } else if (e.key === "PageDown") $("#nextPage").click();
      else if (e.key === "PageUp") $("#prevPage").click();
    });

    const v = $("#viewer");
    v.addEventListener("dragover", (e) => { e.preventDefault(); v.classList.add("drop"); });
    v.addEventListener("dragleave", () => v.classList.remove("drop"));
    v.addEventListener("drop", (e) => { e.preventDefault(); v.classList.remove("drop"); if (e.dataTransfer.files.length) importFiles([...e.dataTransfer.files]); });
    let rt = null;
    window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(render, 150); });
    bindOverlay();
  }

  async function reloadSnips() {
    const list = await xl.loadSnips();
    S.snips.clear();
    for (const sn of list) S.snips.set(sn.id, sn);
    refreshSnipList(); drawOverlay();
  }

  /* ====================================================================== */
  /* Boutons du ruban (onglet SNIPCOMPTA)                                   */
  /* ====================================================================== */
  function registerRibbon() {
    if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) return;
    const show = () => { try { if (Office.addin && Office.addin.showAsTaskpane) Office.addin.showAsTaskpane().catch(() => {}); } catch (_) { /* ancien Excel */ } };
    const act = (name, fn) => Office.actions.associate(name, (event) => {
      show();
      Promise.resolve().then(fn).catch((e) => { console.error(e); status("Erreur : " + e.message, "err"); });
      if (event && event.completed) event.completed();
    });
    const mode = (m) => () => { switchTab("viewer"); setMode(m); };
    act("openViewer", () => switchTab("viewer"));
    act("importDocs", () => {
      switchTab("docs");
      try { $("#fileInput").click(); } catch (_) { /* bloqué : l'utilisateur cliquera */ }
      status("Cliquez sur « Importer des documents » pour choisir vos PDF.");
    });
    act("organizeDocs", () => switchTab("docs"));
    act("modeText", mode("text"));
    act("modeValid", mode("valid"));
    act("modeExcept", mode("except"));
    act("modeSum", mode("sum"));
    act("modeTable", mode("table"));
    act("modeComment", mode("comment"));
    act("undoLast", () => undoLast());
    act("delSelSnips", () => deleteSnipsInSelection());
    act("delAllSnips", () => { switchTab("snips"); return deleteAllSnips(); });
    act("docMatching", () => { switchTab("match"); return runMatch(); });
    act("findSums", () => { switchTab("match"); return findSums(); });
    act("toggleStore", () => toggleStore());
    act("exportSnips", () => exportSnips());
    act("removeDoc", () => { switchTab("docs"); status("Cliquez sur « Retirer » à côté du document à enlever du classeur."); });
    act("ocrDoc", () => { switchTab("viewer"); return ocrCurrentDoc(); });
    act("openSettings", () => { switchTab("docs"); $("#optStore").scrollIntoView(); });
    act("openHelp", () => { switchTab("docs"); $("#helpBox").scrollIntoView(); });
  }

  let started = false;
  async function init(inOffice) {
    if (started) return; started = true;
    S.inOffice = inOffice;
    xl = inOffice ? OfficeXL : DemoXL;
    if (typeof pdfjsLib !== "undefined") pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
    loadOpts();
    bindUI();
    if (!inOffice) { $("#demoBar").hidden = false; DemoXL.bind(); }
    xl.onSelection(onSelectionChanged);
    if (inOffice) { try { registerRibbon(); } catch (e) { console.warn("Ruban :", e); } }
    refreshDocUI();
    try {
      if (inOffice) {
        status("Lecture des snips et documents du classeur…", "busy");
        const docs = await xl.listDocs();
        for (const d of docs) S.docs.set(d.id, { ...d, kind: kindOf(d.name, d.mime), bytes: null, pages: {}, stored: true });
        await reloadSnips();
        refreshDocUI();
        if (docs.length) await openDoc(docs[0].id);
        status(docs.length || S.snips.size ? `Prêt — ${docs.length} document(s), ${S.snips.size} snip(s) dans ce classeur.` : "Prêt. Importez un document pour commencer.", "ok");
      } else {
        status("Mode démo : ouvert hors d'Excel. Les valeurs vont dans la mini-feuille en bas.", "warn");
      }
    } catch (e) { console.error(e); status("Initialisation : " + e.message, "err"); }
    render();
  }

  // exposé pour les tests automatisés
  window.SnipCompta = { S, registerRibbon, findSums, exportSnips, deleteSnipsInSelection, deleteAllSnips, insertComment, sumRuns, typed, parseAmount, findAmounts, groupLines, autoColumns, extractTable, wordsInRect, DemoXL, setMode, handleRect, insertSum, insertTable, importFiles, doSearch, runMatch, pageText };

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady((info) => init(!!info && info.host === Office.HostType.Excel));
  } else {
    document.addEventListener("DOMContentLoaded", () => init(false));
    if (document.readyState !== "loading") init(false);
  }
})();
