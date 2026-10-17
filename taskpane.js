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
    opts: { store: true, ocr: true, advance: true, mergeRows: true, pasteImport: true, autoSearch: true, sortFollow: true, winOnImport: true, theme: "auto", cropMargins: true },
    hitIdx: 0, hitQuery: "", op: null, prog: null, unsaved: false, userActed: false, lastPointer: null, zooms: {},
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
  const BUILD = "20261017-1512"; // remplacé à la publication (stamp.py)
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  let statusTimer = null;
  /**
   * Message à l'utilisateur. kind : "" | ok | warn | err | busy (barre de progression annulable) | saving (progression sans annulation).
   * actions : boutons de solution [{label, fn}].
   */
  function status(msg, kind = "", actions) {
    if (kind === "busy" && S.op && S.op.cancelled) throw cancelErr();
    const el = $("#status");
    if (!el) return;
    clearTimeout(statusTimer);
    if (kind === "busy" || kind === "saving") { showProgress(msg, kind === "busy" && !!S.op); return; }
    hideProgress();
    if (!msg) { el.hidden = true; return; }
    el.innerHTML = "";
    const sp = document.createElement("span"); sp.className = "smsg"; sp.textContent = msg; el.appendChild(sp);
    for (const a of actions || []) {
      const b = document.createElement("button"); b.className = "sact"; b.textContent = a.label;
      b.onclick = () => { el.hidden = true; Promise.resolve().then(a.fn).catch((e) => reportError(e)); };
      el.appendChild(b);
    }
    const x = document.createElement("button"); x.className = "sclose"; x.title = "Fermer"; x.textContent = "×"; x.onclick = () => { el.hidden = true; };
    el.appendChild(x);
    el.className = kind;
    el.hidden = false;
    const ms = (actions && actions.length ? 16000 : 0) + (kind === "err" ? 10000 : kind === "warn" ? 7000 : 4200);
    statusTimer = setTimeout(() => { el.hidden = true; }, ms);
    el.onmouseenter = () => clearTimeout(statusTimer);
    el.onmouseleave = () => { statusTimer = setTimeout(() => { el.hidden = true; }, 3000); };
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
      const sx = Math.hypot(t[0], t[1]);
      if (sx && (Math.abs(t[1]) > 0.02 * sx || t[0] < 0)) {
        // texte tourné (page pivotée, texte vertical ou à l'envers) : boîte englobante des 4 coins
        const dx = t[0] / sx, dy = t[1] / sx, ux = t[2] / h, uy = t[3] / h;
        while ((m = re.exec(it.str))) {
          const s0 = m.index * cw, s1 = (m.index + m[0].length) * cw, xs = [], ys = [];
          for (const s of [s0, s1]) for (const k of [0.82, -0.22]) { xs.push(x + dx * s + ux * h * k); ys.push(base + dy * s + uy * h * k); }
          out.push({ str: m[0], x0: Math.min(...xs), x1: Math.max(...xs), y0: Math.min(...ys), y1: Math.max(...ys) });
        }
        continue;
      }
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
  /* Sélection au caractère près : seul ce qui est dans la zone est pris */
  const charWidths = new Map(); let cwCtx = null;
  function charW(ch) {
    let v = charWidths.get(ch);
    if (v == null) {
      if (!cwCtx) { cwCtx = document.createElement("canvas").getContext("2d"); cwCtx.font = "100px Arial, Helvetica, sans-serif"; }
      v = Math.max(8, cwCtx.measureText(ch).width); charWidths.set(ch, v);
    }
    return v;
  }
  /** Partie d'un mot couverte par la zone (un caractère compte s'il est couvert au moins au tiers). */
  function clipWord(w, r) {
    const cy = (w.y0 + w.y1) / 2;
    if (cy < r.y0 || cy > r.y1 || w.x1 <= r.x0 || w.x0 >= r.x1) return null;
    if (w.x0 >= r.x0 - 0.5 && w.x1 <= r.x1 + 0.5) return w;
    const ch = [...w.str], ws = ch.map(charW), k = (w.x1 - w.x0) / ws.reduce((a, b) => a + b, 0);
    let x = w.x0, i0 = -1, i1 = -1, a0 = 0, a1 = 0;
    ch.forEach((c, i) => {
      const cw = ws[i] * k, ov = Math.min(x + cw, r.x1) - Math.max(x, r.x0);
      if (ov >= cw / 3) { if (i0 < 0) { i0 = i; a0 = x; } i1 = i; a1 = x + cw; }
      x += cw;
    });
    if (i0 < 0) return null;
    return { ...w, str: ch.slice(i0, i1 + 1).join(""), x0: a0, x1: a1 };
  }
  const clipWords = (words, r) => words.map((w) => clipWord(w, r)).filter(Boolean);
  /** Cadre serré autour du texte retenu, sans déborder de la zone dessinée. */
  function tightBox(ws, r) {
    const b = bboxOf(ws, 1.5);
    return { x0: Math.max(b.x0, r.x0 - 1), y0: Math.max(b.y0, r.y0 - 1), x1: Math.min(b.x1, r.x1 + 1), y1: Math.min(b.y1, r.y1 + 1) };
  }
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
    const kind = S.opts.ocrBest ? "best" : "std";
    if (S.ocrWorker && S.ocrKind !== kind) {           // moteur changé dans les paramètres
      const old = S.ocrWorker; S.ocrWorker = null;
      old.then((w) => w.terminate()).catch(() => {});
    }
    if (!S.ocrWorker) {
      S.ocrKind = kind;
      status(kind === "best" ? "Chargement du moteur OCR précis (la première fois, ~20-40 s)…" : "Chargement du moteur OCR (la première fois, ~10-20 s)…", "saving");
      S.ocrWorker = Tesseract.createWorker("fra+eng", 1, {
        workerPath: LIB("tesseract/worker.min.js"),
        corePath: LIB("tesseract/core"),
        langPath: LIB(kind === "best" ? "tesseract/lang-best" : "tesseract/lang"),
        cachePath: "snipcompta-" + kind,                  // les deux moteurs ne se mélangent pas dans le cache
        logger: (m) => { try { if (S.ocrMsg && m && m.status === "recognizing text") status(`${S.ocrMsg} ${Math.round((m.progress || 0) * 100)} %`, S.op ? "busy" : "saving"); } catch (_) { /* annulation : gérée par onCancel */ } },
      });
      S.ocrWorker.catch(() => { S.ocrWorker = null; });
    }
    return S.ocrWorker;
  }

  async function ocrPage(doc, p, pg) {
    if (pg.ocrPromise) return pg.ocrPromise;
    pg.ocrPromise = (async () => {
      const worker = await getOcr();
      S.ocrMsg = `Lecture du scan (OCR) — ${doc.name}, page ${p}…`;
      status(S.ocrMsg, "busy");
      const skew = await pageSkew(doc, p, pg);
      const f = doc.kind === "pdf" ? clamp(2500 / pg.w, 1, 4) : clamp(2500 / pg.w, 0.6, 3);
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(pg.w * f); canvas.height = Math.ceil(pg.h * f);
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      await paintPage(ctx, doc, p, f, skew);
      // copie nettoyée, uniquement pour la lecture : la pièce affichée ne change pas
      if (S.opts.ocrClean !== false) { await sleep(0); cleanForOcr(ctx, canvas.width, canvas.height); }
      let data;
      try { ({ data } = await Promise.race([worker.recognize(canvas), onCancel()])); }
      finally { S.ocrMsg = null; }
      pg.words = (data.words || []).filter((w) => w.text && w.text.trim() && w.confidence > 20)
        .map((w) => ({ str: w.text.trim(), x0: w.bbox.x0 / f, y0: w.bbox.y0 / f, x1: w.bbox.x1 / f, y1: w.bbox.y1 / f }));
      pg.lines = groupLines(pg.words);
      pg.ocrDone = true;
      if (doc.crop === null) doc.crop = undefined;
      status(`OCR terminé : ${pg.words.length} mots reconnus (${doc.name}, p. ${p}).`, "ok");
    })();
    try { await pg.ocrPromise; } catch (e) { pg.ocrPromise = null; throw e; }
  }

  /* ---------------- Scans : redressement + nettoyage avant lecture -------- */
  /** Dessine la page entière dans ctx (taille du canvas = page × sc), redressée de -skew degrés. */
  async function paintPage(ctx, doc, p, sc, skew) {
    const W = ctx.canvas.width, H = ctx.canvas.height;
    ctx.save(); ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, H); ctx.restore();
    let m = null;
    if (skew) {
      const th = (-skew * Math.PI) / 180, c = Math.cos(th), s = Math.sin(th), cx = W / 2, cy = H / 2;
      m = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy];
    }
    if (doc.kind === "pdf") {
      const page = await doc.pdf.getPage(p);
      const params = { canvasContext: ctx, viewport: page.getViewport({ scale: sc }) };
      if (m) params.transform = m;
      await page.render(params).promise;
    } else {
      ctx.save(); if (m) ctx.setTransform(...m); ctx.drawImage(doc.img, 0, 0, W, H); ctx.restore();
    }
  }

  function grayOf(ctx, W, H) {
    const d = ctx.getImageData(0, 0, W, H).data, n = W * H, g = new Uint8ClampedArray(n);
    for (let i = 0, j = 0; i < n; i++, j += 4) g[i] = (d[j] * 77 + d[j + 1] * 150 + d[j + 2] * 29) >> 8;
    return g;
  }
  /** Fond ramené au blanc : ombres de photo, papier gris ou jauni, éclairage inégal. */
  function flattenBg(g, W, H) {
    const B = Math.max(12, Math.round(Math.max(W, H) / 64));
    const gw = Math.ceil(W / B), gh = Math.ceil(H / B), bg = new Float32Array(gw * gh), hist = new Uint32Array(256);
    for (let by = 0; by < gh; by++) for (let bx = 0; bx < gw; bx++) {
      hist.fill(0); let n = 0;
      const ye = Math.min(H, (by + 1) * B), xe = Math.min(W, (bx + 1) * B);
      for (let y = by * B; y < ye; y += 2) { const o = y * W; for (let x = bx * B; x < xe; x += 2) { hist[g[o + x]]++; n++; } }
      let k = Math.max(1, Math.round(n * 0.1)), v = 255;          // couleur du papier = 90e centile du bloc
      while (v > 0) { k -= hist[v]; if (k <= 0) break; v--; }
      bg[by * gw + bx] = v;
    }
    const tmp = new Float32Array(bg.length);
    const pass = (src, dst, max) => {
      for (let y = 0; y < gh; y++) for (let x = 0; x < gw; x++) {
        let acc = 0, c = 0;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = y + dy, xx = x + dx; if (yy < 0 || xx < 0 || yy >= gh || xx >= gw) continue;
          const v = src[yy * gw + xx]; if (max) { if (v > acc) acc = v; } else { acc += v; c++; }
        }
        dst[y * gw + x] = max ? acc : acc / c;
      }
    };
    pass(bg, tmp, true); pass(tmp, bg, false); pass(bg, tmp, false);   // blocs pleins d'encre ignorés, puis lissage
    const X0 = new Int32Array(W), X1 = new Int32Array(W), TX = new Float32Array(W);
    for (let x = 0; x < W; x++) { const fx = clamp((x + 0.5) / B - 0.5, 0, gw - 1); X0[x] = Math.floor(fx); X1[x] = Math.min(gw - 1, X0[x] + 1); TX[x] = fx - X0[x]; }
    const out = new Uint8ClampedArray(W * H);
    for (let y = 0; y < H; y++) {
      const fy = clamp((y + 0.5) / B - 0.5, 0, gh - 1), r0 = Math.floor(fy) * gw, r1 = Math.min(gh - 1, Math.floor(fy) + 1) * gw, ty = fy - Math.floor(fy), o = y * W;
      for (let x = 0; x < W; x++) {
        const tx = TX[x], a = tmp[r0 + X0[x]] * (1 - tx) + tmp[r0 + X1[x]] * tx, b = tmp[r1 + X0[x]] * (1 - tx) + tmp[r1 + X1[x]] * tx;
        const bgv = a * (1 - ty) + b * ty;
        out[o + x] = (g[o + x] * 255) / (bgv > 24 ? bgv : 24);
      }
    }
    return out;
  }
  /** Léger flou [1 2 1] : enlève le grain (photo, scan bruité) sans empâter le texte. */
  function soften(g, W, H) {
    const t = new Uint8ClampedArray(g.length);
    for (let y = 0; y < H; y++) { const o = y * W; t[o] = g[o]; t[o + W - 1] = g[o + W - 1]; for (let x = 1; x < W - 1; x++) t[o + x] = (g[o + x - 1] + 2 * g[o + x] + g[o + x + 1] + 2) >> 2; }
    for (let x = 0; x < W; x++) { g[x] = t[x]; g[(H - 1) * W + x] = t[(H - 1) * W + x]; }
    for (let y = 1; y < H - 1; y++) { const o = y * W; for (let x = 0; x < W; x++) g[o + x] = (t[o + x - W] + 2 * t[o + x] + t[o + x + W] + 2) >> 2; }
    return g;
  }
  /** Contraste : l'encre la plus foncée devient noire, le papier (et son grain) blanc. */
  function stretchGray(g) {
    const hist = new Uint32Array(256); for (let i = 0; i < g.length; i++) hist[g[i]]++;
    const n = g.length;
    const pct = (q) => { let v = 0, acc = 0; while (v < 255 && (acc += hist[v]) < n * q) v++; return v; };
    let lo = pct(0.004);
    const med = Math.max(pct(0.5), 170);
    // le grain du papier descend sous la médiane : on blanchit tout ce qui ressemble au papier
    const tune = S.ocrTune || {};
    let hi = clamp(pct(tune.hiQ ?? 0.15), Math.round((lo + med) / 2) + 12, med);
    lo = Math.min(lo, hi - 50);
    const lut = new Uint8ClampedArray(256);
    for (let v = 0; v < 256; v++) { const t = clamp((v - lo) / (hi - lo), 0, 1); lut[v] = 255 * Math.pow(t, tune.gamma ?? 1); }
    for (let i = 0; i < n; i++) g[i] = lut[g[i]];
    return g;
  }
  /** Inclinaison des lignes de texte (degrés, + = la ligne descend vers la droite). */
  function detectSkew(g, W, H) {
    const pts = [];
    for (let y = 0; y < H; y++) { const o = y * W; for (let x = 0; x < W; x++) if (g[o + x] < 110) pts.push(x, y); }
    const n = pts.length / 2; if (n < 300) return 0;
    const stride = 2 * Math.max(1, Math.ceil(n / 50000)), off = Math.ceil(W * Math.tan((7 * Math.PI) / 180)) + 2;
    const bins = new Float64Array(H + 2 * off + 2);
    const score = (deg) => {
      const t = Math.tan((deg * Math.PI) / 180); bins.fill(0);
      for (let i = 0; i < pts.length; i += stride) bins[Math.round(pts[i + 1] - pts[i] * t) + off]++;
      let s = 0; for (let i = 0; i < bins.length; i++) s += bins[i] * bins[i];
      return s;
    };
    let best = 0, bs = score(0);
    const s0 = bs;
    for (let d = -6; d <= 6.001; d += 0.5) { const s = score(d); if (s > bs) { bs = s; best = d; } }
    const c = best;
    for (let d = c - 0.5; d <= c + 0.501; d += 0.05) { const s = score(d); if (s > bs) { bs = s; best = d; } }
    if (Math.abs(best) < 0.2 || bs < s0 * 1.04) return 0;
    return Math.round(best * 100) / 100;
  }
  /** Inclinaison d'une page scannée (calculée une fois, sur une petite image). */
  function pageSkew(doc, p, pg) {
    if (pg.skew != null) return Promise.resolve(pg.skew);
    if (!pg.needsOcr || S.opts.ocrClean === false) return Promise.resolve(0);
    if (!pg.skewP) {
      pg.skewP = (async () => {
        const sc = 1000 / pg.w, W = Math.round(pg.w * sc), H = Math.round(pg.h * sc);
        const c = document.createElement("canvas"); c.width = W; c.height = H;
        const ctx = c.getContext("2d", { willReadFrequently: true });
        await paintPage(ctx, doc, p, sc, 0);
        return detectSkew(stretchGray(flattenBg(grayOf(ctx, W, H), W, H)), W, H);
      })().then((v) => (pg.skew = v), (e) => { console.warn("inclinaison", e); return (pg.skew = 0); });
    }
    return pg.skewP;
  }
  /** Remplace l'image du canvas par sa version nettoyée (niveaux de gris, fond blanc, contraste). */
  /** Grain de l'image (0 = net, > 0,06 = scan/photo bruité), mesuré sur le papier. */
  function grainOf(g, W, H) {
    const hist = new Uint32Array(256); let n = 0;
    for (let y = 2; y < H - 2; y += 3) { const o = y * W; for (let x = 2; x < W - 2; x += 3) {
      const c = g[o + x], m = (g[o + x - 2] + g[o + x + 2] + g[o + x - 2 * W] + g[o + x + 2 * W]) >> 2;
      if (c < 90 || m < 90) continue;                                // encre : ignorée
      hist[Math.min(255, Math.abs(c - m) * 4)]++; n++;
    } }
    if (!n) return 0;
    let v = 0, acc = 0; while (v < 255 && (acc += hist[v]) < n * 0.5) v++;
    return v / 4 / 255;
  }
  function cleanForOcr(ctx, W, H) {
    const tune = S.ocrTune || {};
    let g = grayOf(ctx, W, H);
    const grain = grainOf(g, W, H);
    S.lastGrain = grain;
    if (tune.soften === true || (tune.soften !== false && grain > (tune.grainMin ?? 0.05))) g = soften(g, W, H);
    g = stretchGray(flattenBg(g, W, H));
    const im = ctx.createImageData(W, H), d = im.data;
    for (let i = 0, j = 0; i < g.length; i++, j += 4) { d[j] = d[j + 1] = d[j + 2] = g[i]; d[j + 3] = 255; }
    ctx.putImageData(im, 0, 0);
  }

  /** Infos et texte d'une page. allowOcr : lance l'OCR si la page est un scan. */
  async function pageText(doc, p, { allowOcr = S.opts.ocr } = {}) {
    await openDocData(doc);
    let pg = doc.pages[p];
    if (!pg) {
      pg = doc.pages[p] = {};
      pg.loading = (async () => {
        if (doc.kind === "pdf") {
          const page = await doc.pdf.getPage(p);
          const vp = page.getViewport({ scale: 1 });
          const tc = await page.getTextContent();
          const words = wordsFromTextContent(tc, vp);
          pg.w = vp.width; pg.h = vp.height; pg.words = words;
          pg.needsOcr = words.reduce((a, w) => a + w.str.length, 0) < 15;
        } else {
          pg.w = doc.img.naturalWidth; pg.h = doc.img.naturalHeight; pg.words = []; pg.needsOcr = true;
        }
        pg.lines = groupLines(pg.words);
      })();
      try { await pg.loading; } catch (e) { if (doc.pages[p] === pg) delete doc.pages[p]; throw e; }
      finally { pg.loading = null; }
    } else if (pg.loading) await pg.loading;
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
        status(`Chargement de « ${doc.name} » depuis le classeur…`, "saving");
        doc.bytes = await xl.loadDocBytes(doc.id);
        if (!doc.bytes) throw new Error(`Le document « ${doc.name} » n'est pas disponible : réimportez-le.`);
      }
      if (doc.kind === "pdf") {
        if (typeof pdfjsLib === "undefined") throw new Error("La visionneuse PDF n'a pas pu être chargée (fichiers lib/ manquants ?)");
        doc.pdf = await openPdfWithPassword(doc);
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
    const versions = [];
    files = await expandFiles(files);
    let k = 0;
    for (const f of files) {
      k++;
      if (files.length > 1) status(`Import de « ${f.name} » (${k}/${files.length})…`, "busy");
      try {
        const bytes = new Uint8Array(await f.arrayBuffer());
        const id = hashBytes(bytes);
        let doc = S.docs.get(id);
        if (!doc) {
          const prev = [...S.docs.values()].find((d) => d.name === f.name && d.id !== id);
          doc = { id, name: f.name, mime: f.type || (/\.pdf$/i.test(f.name) ? "application/pdf" : "image/png"), bytes, pages: {}, stored: false, folder: S.importFolder || "" };
          if (prev) versions.push([prev, doc]);
          doc.kind = kindOf(doc.name, doc.mime);
          S.docs.set(id, doc);
          await openDocData(doc);
          if (S.opts.store && S.inOffice) {
            if (bytes.length > MAX_STORE_MB * 1048576) status(`« ${f.name} » dépasse ${MAX_STORE_MB} Mo : non enregistré dans le classeur.`, "warn");
            else { await xl.saveDoc(doc, (pct) => status(`Enregistrement de « ${f.name} » dans le classeur… ${pct} %`, "saving")); doc.stored = true; }
          }
        } else if (!doc.bytes) { doc.bytes = bytes; }
        last = doc;
        status(`« ${doc.name} » importé (${doc.numPages || "?"} page${doc.numPages > 1 ? "s" : ""}).`, "ok");
      } catch (e) {
        if (isCancel(e)) throw e;
        console.error(e);
        reportError(e, files.length === 1 ? () => runAction(() => importFiles([f])) : null, `Import de « ${f.name} » impossible`);
      }
    }
    refreshDocUI();
    if (last) { await openDoc(last.id, 1); switchTab("viewer"); }
    if (versions.length) offerVersionTransfer(versions[0][0], versions[0][1]);
    if (last && S.inOffice && !S.isWindow && S.opts.winOnImport) { if (S.win) winSend({ event: "reloadDocs" }); else openWindow(); }
  }

  async function openDoc(id, page = 1) {
    const doc = S.docs.get(id);
    if (!doc) return;
    if (S.win) winSend({ event: "open", doc: id, page });
    if (S.cur !== id) { S.table = null; }
    S.cur = id; S.page = page;
    $("#docSelect").value = id;
    try { await openDocData(doc); } catch (e) { status(e.message, "err"); return; }
    S.page = clamp(page, 1, doc.numPages);
    updateSide();
    await render();
    scrollToPage(S.page);
    if (S.hitQuery && !S.hits.some((h) => h.docId === id)) cellHighlight(S.hitQuery);
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
    sel.innerHTML = docs.length ? docs.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("") : `<option value="">Aucun document</option>`;
    document.body.classList.toggle("nodoc", !docs.length);
    if (S.cur) sel.value = S.cur;
    const opts = docs.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("");
    ms.innerHTML = `<option value="*">Tous les documents</option>` + opts;
    for (const id of ["#cmpA", "#cmpB"]) { const e = $(id); if (e) { const v = e.value; e.innerHTML = opts; if (v && S.docs.has(v)) e.value = v; } }
    renderDocList();
    renderAutoDocs();
  }

  /* ====================================================================== */
  /* Rendu de la page + calque                                              */
  /* ====================================================================== */
  const curDoc = () => S.docs.get(S.cur);

  /* ---------------- Visionneuse : toutes les pages empilées --------------- */
  let pageObserver = null, thumbObserver = null;
  function loadZooms() { try { return JSON.parse(localStorage.getItem("snipcompta.zoom") || "{}"); } catch (_) { return {}; } }
  S.zooms = loadZooms();
  const zoomFor = (doc) => (doc && S.zooms[doc.id] != null ? S.zooms[doc.id] : 1);
  function setZoom(doc, z) {
    if (!doc) return;
    S.zooms[doc.id] = z === "page" ? "page" : clamp(z, 0.3, 6);
    try { localStorage.setItem("snipcompta.zoom", JSON.stringify(S.zooms)); } catch (_) { /* ignore */ }
    render();
  }
  async function pageDims(doc) {
    if (doc.dims && doc.dims.length === doc.numPages) return doc.dims;
    const dims = [];
    if (doc.kind === "pdf") {
      for (let p = 1; p <= doc.numPages; p++) { const vp = (await doc.pdf.getPage(p)).getViewport({ scale: 1 }); dims.push({ w: vp.width, h: vp.height }); }
    } else dims.push({ w: doc.img.naturalWidth, h: doc.img.naturalHeight });
    doc.dims = dims;
    return dims;
  }
  const pageEl = (p) => $(`#pages .pg[data-p="${p}"]`);
  const curOv = () => $(`#pages .ov[data-p="${S.page}"]`);

  async function render() {
    const doc = curDoc();
    const token = ++S.renderToken;
    const pagesEl = $("#pages"), empty = $("#emptyState");
    if (!doc) { pagesEl.innerHTML = ""; pagesEl.dataset.key = ""; pagesEl.hidden = true; empty.hidden = false; $("#pageInfo").textContent = "–"; renderThumbs(null); return; }
    pagesEl.hidden = false; empty.hidden = true;
    try {
      await openDocData(doc);
      const dims = await pageDims(doc);
      if (token !== S.renderToken) return;
      const th = $("#thumbs"); if (th) th.hidden = doc.numPages < 2;
      const v = $("#viewer");
      if (!v.clientWidth) return;   // onglet masqué : on dessinera en y revenant
      const maxW = Math.max(...dims.map((d) => d.w)), maxH = Math.max(...dims.map((d) => d.h));
      const padX = v.clientWidth < 560 ? 12 : 28;
      const availW = Math.max(120, v.clientWidth - padX), availH = Math.max(120, v.clientHeight - 28);
      const z = zoomFor(doc);
      // marges blanches ignorées : on cadre sur la zone où il y a du texte
      const crop = z !== "page" && S.opts.cropMargins !== false ? await cropOf(doc, maxW) : null;
      if (token !== S.renderToken) return;
      const cw = crop ? crop.x1 - crop.x0 : maxW, fitW = availW / cw;
      const scale = Math.max(0.1, z === "page" ? Math.min(availW / maxW, availH / maxH) : fitW * z);
      S.cropX = crop ? crop.x0 : 0;
      const key = doc.id + "|" + scale.toFixed(4) + "|" + doc.numPages + "|" + (doc.rev || 0) + "|" + (crop ? crop.x0.toFixed(1) + "-" + crop.x1.toFixed(1) : "");
      if (pagesEl.dataset.key !== key) {
        const keepPage = S.page, el0 = pageEl(keepPage);
        const frac = el0 ? (v.scrollTop - el0.offsetTop) / Math.max(1, el0.offsetHeight) : 0;
        S.scale = scale;
        pagesEl.dataset.key = key; pagesEl.innerHTML = "";
        if (pageObserver) pageObserver.disconnect();
        pageObserver = new IntersectionObserver((entries) => entries.forEach((en) => { if (en.isIntersecting) renderPage(doc, +en.target.dataset.p, en.target, key); }), { root: v, rootMargin: "700px 0px" });
        dims.forEach((d, i) => {
          const p = i + 1;
          const pg = document.createElement("div");
          pg.className = "pg"; pg.dataset.p = p;
          const fw = Math.round(d.w * scale), fh = Math.round(d.h * scale), off = Math.round(S.cropX * scale);
          pg.dataset.fw = fw; pg.dataset.fh = fh;
          pg.style.width = (crop ? Math.max(40, Math.min(Math.round(cw * scale), fw - off)) : fw) + "px"; pg.style.height = fh + "px";
          const pos = `left:${-off}px;width:${fw}px;height:${fh}px`;
          pg.innerHTML = `<canvas style="${pos}"></canvas><div class="ov" data-p="${p}" style="${pos}"></div>`;
          pagesEl.appendChild(pg); pageObserver.observe(pg);
        });
        const el1 = pageEl(clamp(keepPage, 1, doc.numPages));
        if (el1) v.scrollTop = el1.offsetTop + Math.max(0, frac) * el1.offsetHeight;
        renderThumbs(doc);
      }
      S.scale = scale;
      $("#zoomPct").textContent = Math.round((scale / (96 / 72)) * 100) + " %";
      updatePageInfo();
      drawOverlay();
    } catch (e) { console.error(e); status("Affichage impossible : " + e.message, "err"); }
  }
  async function renderPage(doc, p, el, key) {
    if (el.dataset.done === key) return;
    el.dataset.done = key;
    try {
      const canvas = el.querySelector("canvas");
      const W = +el.dataset.fw, H = +el.dataset.fh, dpr = window.devicePixelRatio || 1;
      const pg = await pageText(doc, p, { allowOcr: false });
      const skew = await pageSkew(doc, p, pg);          // scan de travers : affiché redressé
      if (el.dataset.done !== key) return;
      canvas.style.width = W + "px"; canvas.style.height = H + "px";
      canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr);
      await paintPage(canvas.getContext("2d"), doc, p, (W / doc.dims[p - 1].w) * dpr, skew);
      if (pg.needsOcr && !pg.ocrDone && S.opts.ocr) runAction(() => pageText(doc, p, { allowOcr: true }), { label: "OCR impossible" });
    } catch (e) { if (!e || e.name !== "RenderingCancelledException") console.warn("page", p, e); el.dataset.done = ""; }
  }
  function updatePageInfo() {
    const doc = curDoc();
    $("#pageInfo").textContent = doc ? `${S.page} / ${doc.numPages}` : "–";
    $$("#thumbs .th").forEach((t) => t.classList.toggle("on", +t.dataset.p === S.page));
  }
  function trackScroll() {
    const v = $("#viewer"), doc = curDoc(); if (!doc) return;
    const y = v.scrollTop + v.clientHeight * 0.35;
    let cur = 1;
    for (const el of $$("#pages .pg")) { if (el.offsetTop <= y) cur = +el.dataset.p; else break; }
    if (cur !== S.page) { S.page = cur; updatePageInfo(); const th = $(`#thumbs .th[data-p="${cur}"]`); if (th) th.scrollIntoView({ block: "nearest" }); }
  }
  function scrollToPage(p, rect) {
    const v = $("#viewer"), el = pageEl(p); if (!el) return;
    S.page = p; updatePageInfo();
    v.scrollTop = el.offsetTop - 10 + (rect ? Math.max(0, rect.y0 * S.scale - 70) : 0);
    if (rect) v.scrollLeft = Math.max(0, el.offsetLeft + (rect.x0 - (S.cropX || 0)) * S.scale - 40);
  }
  /** Zone utile des pages (là où il y a du texte), pour ne pas afficher les grandes marges blanches. */
  async function cropOf(doc, maxW) {
    if (doc.crop !== undefined && doc.cropRev === (doc.rev || 0)) return doc.crop;
    let x0 = Infinity, x1 = -Infinity, n = 0;
    for (let p = 1; p <= Math.min(doc.numPages, 6); p++) {
      let pg;
      try { pg = await pageText(doc, p, { allowOcr: false }); } catch (_) { continue; }
      if (!pg.words || pg.words.length < 5) continue;
      for (const w of pg.words) { if (w.x0 < x0) x0 = w.x0; if (w.x1 > x1) x1 = w.x1; }
      n++;
    }
    let crop = null;
    if (n && x1 > x0) {
      crop = { x0: Math.max(0, x0 - 20), x1: Math.min(maxW, x1 + 40) };   // à droite : place pour les étiquettes
      if (crop.x1 - crop.x0 > maxW * 0.9) crop = null;     // presque pleine largeur : inutile
    }
    doc.crop = crop; doc.cropRev = doc.rev || 0;
    return crop;
  }
  function renderThumbs(doc) {
    const box = $("#thumbs"); if (!box) return;
    if (thumbObserver) thumbObserver.disconnect();
    box.innerHTML = "";
    if (!doc || doc.numPages < 2) { box.hidden = true; return; }
    thumbObserver = new IntersectionObserver((entries) => entries.forEach(async (en) => {
      if (!en.isIntersecting || en.target.dataset.done) return;
      en.target.dataset.done = "1";
      const p = +en.target.dataset.p, c = en.target.querySelector("canvas"), d = doc.dims[p - 1], w = 84, s = w / d.w;
      c.width = Math.round(w * 2); c.height = Math.round(d.h * s * 2); c.style.width = w + "px"; c.style.height = Math.round(d.h * s) + "px";
      try { const pg = await pageText(doc, p, { allowOcr: false }); await paintPage(c.getContext("2d"), doc, p, s * 2, await pageSkew(doc, p, pg)); }
      catch (_) { en.target.dataset.done = ""; }
    }), { root: box, rootMargin: "300px 0px" });
    for (let p = 1; p <= doc.numPages; p++) {
      const t = document.createElement("button");
      t.className = "th" + (p === S.page ? " on" : ""); t.dataset.p = p; t.title = "Page " + p;
      t.innerHTML = `<canvas></canvas><span>${p}</span>`;
      t.onclick = () => scrollToPage(p);
      box.appendChild(t); thumbObserver.observe(t);
    }
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
    const pagesEl = $("#pages");
    pagesEl.classList.toggle("drawing", S.mode !== "nav");
    pagesEl.classList.toggle("noting", S.mode === "note" || (S.mode === "markup" && ["text", "check", "cross"].includes(S.markTool)));
    const doc = curDoc();
    for (const ov of $$("#pages .ov")) { ov.innerHTML = ""; if (doc) drawPageOverlay(ov, doc, +ov.dataset.p); }
    syncSideList();
  }
  function drawPageOverlay(ov, doc, p) {
    const sc = S.scale;
    for (const sn of S.snips.values()) {
      for (const [ri, r] of sn.rects.entries()) {
        if ((r.docId || sn.docId) !== doc.id || r.page !== p) continue;
        const st = STYLE[sn.type] || STYLE.text;
        const showBadge = sn.type !== "table" || /_0$/.test(sn.id) || S.selSnip === sn.id;
        const cell = String(sn.address || "").split(":")[0];
        const el = box(ov, r, "snip t-" + sn.type + (S.selSnip === sn.id ? " sel" : ""), st.color, "", `${st.name} → ${sn.sheet}!${sn.address}`);
        if (showBadge) {
          const tag = document.createElement("span");
          tag.className = "tag"; tag.style.background = st.color;
          tag.innerHTML = `<i>${esc(st.label)}</i><b>${esc(cell)}</b>`;
          el.appendChild(tag);
        }
        if (S.selSnip === sn.id && RESIZABLE.has(sn.type)) {
          el.classList.add("resizable");
          for (const h of ["n", "s", "e", "w", "nw", "ne", "sw", "se"]) {
            const k = document.createElement("i");
            k.className = "hdl h-" + h; k.dataset.h = h; k.dataset.sid = sn.id; k.dataset.ri = ri;
            k.title = "Glisser pour agrandir ou réduire la zone";
            el.appendChild(k);
          }
        }
      }
    }
    for (const part of S.sumParts) if (part.docId === doc.id && part.page === p) box(ov, part.rect, "part", STYLE.sum.color, fmtNum(part.total));
    S.hits.forEach((h, i) => { if (h.docId === doc.id && h.page === p) box(ov, h.rect, "hit" + (S.hitQuery && i === S.hitIdx ? " cur" : "")); });
    if (S.flash && S.flash.docId === doc.id && S.flash.page === p) box(ov, S.flash.rect, "flash");
    const cd = S.commentDraft;
    if (cd && cd.docId === doc.id && cd.page === p) box(ov, cd.rect, "part", STYLE.comment.color, "✎");
    drawExtras(ov, doc, p);
    const t = S.table;
    if (t && t.docId === doc.id && t.page === p) {
      const r = t.rect;
      const tb = document.createElement("div");
      tb.className = "tablebox";
      tb.style.left = r.x0 * sc + "px"; tb.style.top = r.y0 * sc + "px";
      tb.style.width = (r.x1 - r.x0) * sc + "px"; tb.style.height = (r.y1 - r.y0) * sc + "px";
      for (const y of t.rowYs || []) { const rl = document.createElement("div"); rl.className = "rowline"; rl.style.top = (y - r.y0) * sc + "px"; tb.appendChild(rl); }
      for (const row of t.rows || []) {
        if (!t.excl || !t.excl.has(rowKey(p, row))) continue;
        const rx = document.createElement("div"); rx.className = "rowx";
        rx.style.top = (row.y0 - 1 - r.y0) * sc + "px"; rx.style.height = (row.y1 - row.y0 + 2) * sc + "px";
        tb.appendChild(rx);
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
  function evPt(e, ov) {
    ov = ov || (drag && drag.ov) || (e.target && e.target.closest && e.target.closest(".ov")) || curOv();
    const r = ov.getBoundingClientRect();
    return { x: (e.clientX - r.left) / S.scale, y: (e.clientY - r.top) / S.scale };
  }
  function bindOverlay() {
    const host = $("#pages");
    host.addEventListener("pointerdown", (e) => {
      const ov = e.target.closest(".ov");
      if (e.button !== 0 || !curDoc() || !ov) return;
      if (S.page !== +ov.dataset.p) { S.page = +ov.dataset.p; updatePageInfo(); }
      S.lastPointer = { x: e.clientX, y: e.clientY };
      if (e.target.classList.contains("hdl")) {
        const sn = S.snips.get(e.target.dataset.sid); if (!sn) return;
        const ri = +e.target.dataset.ri;
        drag = { kind: "resize", sn, ri, h: e.target.dataset.h, r0: { ...sn.rects[ri] }, r: null, el: e.target.parentElement, ov };
        hideHover(); ov.setPointerCapture(e.pointerId); e.preventDefault(); e.stopPropagation(); return;
      }
      if (e.target.classList.contains("sep")) {
        drag = { kind: "sep", idx: +e.target.dataset.i, el: e.target, ov };
        ov.setPointerCapture(e.pointerId); e.preventDefault(); return;
      }
      const pt = evPt(e, ov);
      drag = { kind: "rect", x: pt.x, y: pt.y, moved: false, el: null, ov };
      hideHover();
      if (S.mode !== "nav") { ov.setPointerCapture(e.pointerId); e.preventDefault(); }
    });
    host.addEventListener("pointermove", (e) => {
      if (!drag) { hoverAt(e); return; }
      const pt = evPt(e);
      if (drag.kind === "resize") {
        const r = drag.r = resizedRect(drag.r0, drag.h, pt), el = drag.el, sc = S.scale;
        el.style.left = r.x0 * sc + "px"; el.style.top = r.y0 * sc + "px";
        el.style.width = (r.x1 - r.x0) * sc + "px"; el.style.height = (r.y1 - r.y0) * sc + "px";
        liveResize(drag.sn, drag.ri, r, e);
        return;
      }
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
        if (!drag.el) drag.el = box(drag.ov, r, "drag");
        drag.el.style.left = r.x0 * S.scale + "px"; drag.el.style.top = r.y0 * S.scale + "px";
        drag.el.style.width = (r.x1 - r.x0) * S.scale + "px"; drag.el.style.height = (r.y1 - r.y0) * S.scale + "px";
        liveValue(r, e);
      }
    });
    host.addEventListener("pointerleave", () => { if (!drag) hideHover(); });
    const end = (e) => {
      if (!drag) return;
      const d = drag; drag = null;
      hideLive();
      S.lastPointer = { x: e.clientX, y: e.clientY };
      if (d.kind === "sep") { S.table.seps.sort((a, b) => a - b); drawOverlay(); updateTablePreview(); return; }
      if (d.kind === "resize") { if (d.r) runAction(() => resizeSnip(d.sn, d.ri, d.r), { label: "Ajustement impossible" }); else drawOverlay(); return; }
      const pt = evPt(e, d.ov);
      if (d.el) d.el.remove();
      if (!d.moved || S.mode === "nav") { handleClick(pt, e); return; }
      const r = { x0: Math.min(d.x, pt.x), y0: Math.min(d.y, pt.y), x1: Math.max(d.x, pt.x), y1: Math.max(d.y, pt.y) };
      runAction(() => handleRect(r, { ctrl: e.ctrlKey, from: { x: d.x, y: d.y }, to: pt }));
    };
    host.addEventListener("pointerup", end);
    host.addEventListener("pointercancel", () => { if (drag && drag.kind === "resize") drawOverlay(); else if (drag && drag.el) drag.el.remove(); drag = null; hideLive(); });
    host.addEventListener("dblclick", (e) => {
      if (e.target.classList.contains("sep") && S.table) { S.table.seps.splice(+e.target.dataset.i, 1); drawOverlay(); updateTablePreview(); }
    });
    const v = $("#viewer");
    let st = null;
    v.addEventListener("scroll", () => { if (st) return; st = requestAnimationFrame(() => { st = null; trackScroll(); }); });
    v.addEventListener("wheel", (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const doc = curDoc(); if (!doc) return;
      let z = zoomFor(doc); if (z === "page") z = 1;
      clearTimeout(v._zt);
      S._zoomPending = clamp((S._zoomPending || z) * (e.deltaY < 0 ? 1.12 : 1 / 1.12), 0.3, 6);
      v._zt = setTimeout(() => { const nz = S._zoomPending; S._zoomPending = null; setZoom(doc, nz); }, 70);
    }, { passive: false });
  }

  function handleClick(pt, ev) {
    const doc = curDoc(); if (!doc) return;
    if (extraClick(pt)) return;
    const t = S.table;
    if (t && t.docId === doc.id && t.page === S.page && pt.x > t.rect.x0 && pt.x < t.rect.x1 && pt.y > t.rect.y0 && pt.y < t.rect.y1) {
      if (!t.seps.some((x) => Math.abs(x - pt.x) * S.scale < 6)) { t.seps.push(pt.x); t.seps.sort((a, b) => a - b); drawOverlay(); updateTablePreview(); }
      return;
    }
    if (CLICK_MODES.has(S.mode)) { queueClickSnip(pt, ev); return; }
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

  async function handleRect(r, opts = {}) {
    if (opts.ctrl && ["text", "sum", "valid", "except"].includes(S.mode)) return attachEvidence(r);
    if (await extraRect(r, opts)) return;
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
    const sn = { id, group: id, type, sheet: ac.sheet, address, docId: rects[0].docId || doc.id, docName: doc.name, rects, value, cellVals: [value], created: Date.now() };
    S.snips.set(id, sn);
    await xl.addLinks([sn]);
    await persistSnips();
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + 1, ac.col);
    S.selSnip = id;
    drawOverlay(); refreshSnipList();
    confirmSnip(fmtVal(value), address, type);
  }

  async function doTextSnip(r, target) {
    const doc = curDoc(), p = S.page;
    if (target) { const tv = typed(target.text); return placeSnip("text", tv.v, [{ docId: doc.id, page: p, ...target.rect }], tv.f); }
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = clipWords(pg.words, r);
    const text = joinWords(ws);
    if (!text) return status(pg.needsOcr && !S.opts.ocr ? "Page scannée : activez l'OCR dans les réglages." : "Aucun texte dans la zone sélectionnée.", "warn");
    const tv = typed(text);
    await placeSnip("text", tv.v, [{ docId: doc.id, page: p, ...tightBox(ws, r) }], tv.f);
  }

  async function doMarkSnip(type, r) {
    const doc = curDoc(), p = S.page;
    let rect = r;
    try {
      const pg = await pageText(doc, p, { allowOcr: true });
      const ws = wordsInRect(pg.words, r);
      if (ws.length) rect = bboxOf(ws, 2);
    } catch (_) { /* sans texte : on garde la zone dessinée */ }
    const rects = [{ docId: doc.id, page: p, ...rect }];
    const ac = await xl.activeCell();
    const cur = await xl.readCell(ac.sheet, ac.row, ac.col);
    if (cur !== "" && cur != null) return linkExistingCell(type, ac, rects, cur);
    await placeSnip(type, type === "valid" ? "✓" : "✗", rects);
  }

  /* ---------------- Ajuster la zone d'un snip (poignées) ----------------- */
  const RESIZABLE = new Set(["text", "sum", "valid", "except", "comment"]);
  function resizedRect(r0, h, pt) {
    const r = { x0: r0.x0, y0: r0.y0, x1: r0.x1, y1: r0.y1 }, m = 3;
    if (h.includes("w")) r.x0 = Math.min(pt.x, r0.x1 - m);
    if (h.includes("e")) r.x1 = Math.max(pt.x, r0.x0 + m);
    if (h.includes("n")) r.y0 = Math.min(pt.y, r0.y1 - m);
    if (h.includes("s")) r.y1 = Math.max(pt.y, r0.y0 + m);
    return r;
  }
  function parseA1(a) {
    const m = /^\$?([A-Z]+)\$?(\d+)/i.exec(String(a || "")); if (!m) return null;
    let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64;
    return { row: +m[2] - 1, col: c - 1 };
  }
  /** Nouvelle valeur d'un snip si sa zone n° ri devient r (null = valeur inchangée). */
  async function valueAfterResize(sn, ri, r, { live = false } = {}) {
    const rect = sn.rects[ri], doc = S.docs.get(rect.docId || sn.docId); if (!doc) return null;
    const textOf = async (rc) => { const pg = live ? doc.pages[rc.page] : await pageText(doc, rc.page, { allowOcr: true }); return pg && pg.words ? pg.words : []; };
    if (sn.type === "text" && ri === 0) {
      const ws = clipWords(await textOf(rect), r), text = joinWords(ws);
      if (!text) return { empty: true };
      const tv = typed(text);
      return { value: tv.v, fmt: tv.f, box: tightBox(ws, r) };
    }
    if (sn.type === "sum") {
      let total = 0, n = 0;
      for (const [i, rc] of sn.rects.entries()) {
        const z = i === ri ? { ...rc, ...r } : rc, d = S.docs.get(z.docId || sn.docId); if (!d) continue;
        const pg = live ? d.pages[z.page] : await pageText(d, z.page, { allowOcr: true });
        const a = amountsIn(wordsInRect((pg && pg.words) || [], z)); n += a.length;
        total += a.reduce((s, x) => s + x.value, 0);
      }
      return { value: round2(total), fmt: NUMFMT, count: n };
    }
    return null;
  }
  let liveTok = 0;
  async function liveResize(sn, ri, r, e) {
    const tok = ++liveTok, v = await valueAfterResize(sn, ri, r, { live: true });
    if (tok !== liveTok || !drag) return;
    const el = $("#live"); if (!el) return;
    if (!v) { el.hidden = true; return; }
    el.textContent = v.empty ? "(aucun texte)" : (sn.type === "sum" ? "Σ " : "") + fmtVal(v.value);
    el.hidden = false; placeFloat(el, e.clientX + 16, e.clientY + 18);
  }
  async function resizeSnip(sn, ri, r) {
    hideLive(); const cf = $("#confirm"); if (cf) cf.hidden = true;
    if (!S.snips.has(sn.id) || !sn.rects[ri]) return drawOverlay();
    const v = await valueAfterResize(sn, ri, r);
    if (v && v.empty) { drawOverlay(); return status("Aucun texte dans la nouvelle zone : le snip n'a pas changé.", "warn"); }
    sn.rects[ri] = { ...sn.rects[ri], ...r, ...(v && v.box ? v.box : {}) };
    let msg = `Zone ajustée (${sn.sheet}!${sn.address}).`;
    if (v && v.value !== sn.value) {
      const rc = parseA1(sn.address);
      if (rc) {
        await xl.write(sn.sheet, rc.row, rc.col, [[v.value]], sn.type, v.fmt ? [[v.fmt]] : null);
        sn.value = v.value; sn.cellVals = [v.value];
        msg = `Zone ajustée : ${sn.sheet}!${sn.address} = ${fmtVal(v.value)}`;
      }
    }
    S.selSnip = sn.id;
    await persistSnips();
    drawOverlay(); refreshSnipList();
    status(msg, "ok");
  }

  /* ---------------- Sum Snip ------------------------------------------- */
  async function addSumPart(r) {
    const doc = curDoc(), p = S.page;
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    const amts = amountsIn(ws);
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
    $("#sumInsert").onclick = () => runAction(() => insertSum());
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
    // écart minimal entre deux colonnes, proportionnel à la taille du texte
    // (une image scannée compte en pixels, pas en points : « 1 590,07 » ne doit pas être coupé)
    const hs = words.map((w) => w.y1 - w.y0).sort((a, b) => a - b);
    const minGap = Math.max(6, 0.62 * (hs.length ? hs[hs.length >> 1] : 9));
    let first = 0, last = n - 1;
    while (first < n && cov[first] <= tol) first++;
    while (last > 0 && cov[last] <= tol) last--;
    const seps = [];
    let i = first;
    while (i <= last) {
      if (cov[i] <= tol) {
        let j = i; while (j <= last && cov[j] <= tol) j++;
        if (j - i >= minGap) seps.push(r.x0 + (i + j) / 2);
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

  async function startTable(r, page) {
    const doc = curDoc(), p = page || S.page;
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    if (!ws.length) return status("Aucun texte dans la zone : élargissez la sélection ou activez l'OCR.", "warn");
    S.table = { docId: doc.id, page: p, rect: r, seps: autoColumns(ws, r), from: p, to: p, excl: new Set() };
    drawOverlay(); updateTablePreview();
  }

  let previewTimer = null;
  const throttledPreview = () => { clearTimeout(previewTimer); previewTimer = setTimeout(updateTablePreview, 120); };

  const rowKey = (page, r) => page + ":" + Math.round(r.y0);
  function updateTablePreview() {
    const t = S.table, panel = $("#ctxPanel");
    if (!t) { if (!S.sumParts.length) panel.hidden = true; return; }
    t.excl = t.excl || new Set();
    const doc = S.docs.get(t.docId);
    const pg = doc.pages[t.page];
    const rows = extractTable(pg, t.rect, t.seps, S.opts.mergeRows);
    t.rows = rows;
    t.rowYs = rows.slice(1).map((r) => r.y0 - 1);
    const nc = t.seps.length + 1;
    const kept = rows.filter((r) => !t.excl.has(rowKey(t.page, r))).length;
    const multi = doc.numPages > 1;
    panel.hidden = false;
    panel.innerHTML = `<div class="tp-head"><strong>Tableau</strong> <span class="hint">${kept} ligne(s) retenue(s) × ${nc} colonne(s)</span></div>
      <div class="hint">Cliquez dans le cadre pour ajouter une colonne · glissez un trait pour le déplacer · <b>×</b> retire une ligne · <b>⇹</b> fusionne deux colonnes.</div>
      <div class="tp-wrap"><table class="tp">
        <thead><tr><th></th>${Array.from({ length: nc }, (_, i) => `<th><span>Col. ${i + 1}</span>${i < nc - 1 ? `<button class="tp-merge" data-merge="${i}" title="Fusionner les colonnes ${i + 1} et ${i + 2}">⇹</button>` : ""}</th>`).join("")}</tr></thead>
        <tbody>${rows.map((r, i) => {
          const off = t.excl.has(rowKey(t.page, r));
          return `<tr class="${off ? "off" : ""}"><td class="tp-x"><button data-row="${i}" title="${off ? "Remettre cette ligne" : "Retirer cette ligne"}">${off ? "↺" : "×"}</button></td>${r.cells.map((c) => `<td class="${parseAmount(c) != null ? "num" : ""}">${esc(c)}</td>`).join("")}</tr>`;
        }).join("")}</tbody></table></div>
      ${multi ? `<div class="row">Pages <input type="number" id="tFrom" min="1" max="${doc.numPages}" value="${t.from}"> à <input type="number" id="tTo" min="1" max="${doc.numPages}" value="${t.to}">
        <span class="hint">(même zone sur chaque page)</span></div>` : ""}
      <label class="chk"><input type="checkbox" id="tMerge" ${S.opts.mergeRows ? "checked" : ""}> Recoller les libellés sur plusieurs lignes</label>
      <div class="row"><button class="primary" id="tInsert">Insérer dans Excel (Entrée)</button><button id="tAuto">Colonnes auto</button><button id="tCancel">Annuler</button></div>`;
    if (multi) {
      $("#tFrom").onchange = (e) => { t.from = clamp(+e.target.value || 1, 1, doc.numPages); };
      $("#tTo").onchange = (e) => { t.to = clamp(+e.target.value || t.page, 1, doc.numPages); };
    }
    $$("[data-row]", panel).forEach((b) => (b.onclick = () => { const k = rowKey(t.page, rows[+b.dataset.row]); if (t.excl.has(k)) t.excl.delete(k); else t.excl.add(k); updateTablePreview(); }));
    $$("[data-merge]", panel).forEach((b) => (b.onclick = () => { t.seps.sort((a, b2) => a - b2).splice(+b.dataset.merge, 1); updateTablePreview(); }));
    $("#tMerge").onchange = (e) => { S.opts.mergeRows = e.target.checked; saveOpts(); updateTablePreview(); };
    $("#tInsert").onclick = () => runAction(insertTable, { label: "Insertion du tableau impossible" });
    $("#tAuto").onclick = () => { t.seps = autoColumns(wordsInRect(pg.words, t.rect), t.rect); updateTablePreview(); };
    $("#tCancel").onclick = () => { S.table = null; drawOverlay(); updateTablePreview(); };
    drawOverlay();
  }

  async function insertTable() {
    const t = S.table; if (!t) return;
    const doc = S.docs.get(t.docId);
    const from = Math.min(t.from, t.to), to = Math.max(t.from, t.to);
    const all = [];
    for (let p = from; p <= to; p++) {
      if (to > from) status(`Extraction du tableau — page ${p - from + 1}/${to - from + 1}…`, "busy");
      const pg = await pageText(doc, p, { allowOcr: true });
      for (const r of extractTable(pg, t.rect, t.seps, S.opts.mergeRows)) if (!(t.excl && t.excl.has(rowKey(p, r)))) all.push({ ...r, page: p });
    }
    if (!all.length) return status("Aucune ligne à insérer.", "warn");
    const nc = t.seps.length + 1;
    const tv = all.map((r) => r.cells.map(typed));
    const values = tv.map((r) => r.map((c) => c.v)), fmts = tv.map((r) => r.map((c) => c.f));
    const ac = await xl.activeCell();
    await xl.write(ac.sheet, ac.row, ac.col, values, "table", fmts);
    const group = uid();
    const made = all.map((r, i) => ({
      id: group + "_" + i, group, type: "table", sheet: ac.sheet, address: rangeA1(ac.row + i, ac.col, 1, nc),
      docId: doc.id, docName: doc.name, rects: [{ docId: doc.id, page: r.page, x0: r.x0, y0: r.y0 - 1, x1: r.x1, y1: r.y1 + 1 }],
      value: r.cells.filter(Boolean).join(" | "), cellVals: values[i], created: Date.now() + i / 1000,
    }));
    for (const sn of made) S.snips.set(sn.id, sn);
    await xl.addLinks(made);
    await persistSnips();
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + all.length, ac.col);
    S.table = null; S.selSnip = made[0].id; updateTablePreview(); drawOverlay(); refreshSnipList();
    const addr = rangeA1(ac.row, ac.col, all.length, nc);
    status(`Tableau inséré : ${all.length} ligne(s) × ${nc} colonne(s) en ${ac.sheet}!${addr}.`, "ok");
    confirmSnip(`${all.length} ligne(s)`, addr, "table");
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
    $("#cAdd").onclick = () => runAction(() => insertComment(ta.value), { label: "Commentaire impossible" });
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
        await xl.saveDoc(d, (pct) => status(`Enregistrement de « ${d.name} »… ${pct} %`, "saving")); d.stored = true;
      }
      refreshDocUI();
    }
    status(S.opts.store ? "Les PDF sont maintenant enregistrés dans le classeur." : "Les nouveaux PDF ne seront plus enregistrés dans le classeur (liens conservés).", "ok");
  }

  /* ---------------- Suppression / annulation --------------------------- */

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
    scrollToPage(page, rect);
    if (rect) {
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
    if (S.win) { winSend({ event: "selection" }); return; }
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
        if (hit) { if (S.hitQuery) clearHits(); if (hit.id !== S.selSnip) await goToSnip(hit); }
        else if (S.opts.autoSearch) { if (info.text) autoSearchCell(info.text); else if (S.hitQuery) clearHits(); }
      } catch (e) { console.warn(e); }
    }, 180);
  }

  /* ---------------- Liste des snips ------------------------------------ */
  function refreshSnipList() {
    S.snipRev = (S.snipRev || 0) + 1; syncSideList();
    const ul = $("#snipList");
    const all = filteredSnips().sort((a, b) => b.created - a.created);
    $("#snipCount").textContent = all.length ? `${all.length} snip(s) dans ce classeur. Cliquez pour voir la preuve et la cellule.` : "Aucun snip pour l'instant.";
    ul.innerHTML = "";
    for (const sn of all.slice(0, 500)) {
      const st = STYLE[sn.type];
      const li = document.createElement("li");
      li.innerHTML = `<span class="chip" style="background:${st.color}">${st.label}</span>
        <div class="grow click"><div class="ttl">${esc(sn.sheet)}!${esc(sn.address)} — ${esc(fmtVal(sn.value))}</div>
        <div class="sub">${esc(sn.docName)} · p. ${sn.rects[0].page}${sn.auto ? " · rapprochement" : ""}</div></div>
        <button class="danger" title="Supprimer le lien (la valeur reste dans la cellule)">×</button>`;
      li.querySelector(".grow").onclick = async () => { S.selSnip = sn.id; await goToSnip(sn); xl.selectAddr(sn.sheet, sn.address).catch(() => {}); };
      li.querySelector("button").onclick = () => deleteSnips([sn.id]).then(() => status("Lien supprimé (la valeur est conservée).", "ok"));
      ul.appendChild(li);
    }
    if (all.length > 500) ul.insertAdjacentHTML("beforeend", `<li class="hint">… ${all.length - 500} de plus</li>`);
    renderAnnotList();
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
      const total = docs.reduce((a, d) => a + (d.numPages || 1), 0);
      for (let p = 1; p <= doc.numPages; p++) {
        S.prog = { i: ++i, total };
        try { onProgress && onProgress(doc, p, i); } finally { S.prog = null; }
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
    const terms = searchTerms(q), qAmt = isDateLike(q) ? null : parseAmount(q);
    const allowOcr = $("#searchOcr").checked;
    status("Recherche…", "busy");
    const pages = await allPages([...S.docs.values()], allowOcr, (d, p) => status(`Recherche dans « ${d.name} », p. ${p}…`, "busy"));
    const res = []; let skipped = 0;
    for (const { doc, p, pg } of pages) {
      if (pg.needsOcr && !pg.ocrDone) skipped++;
      for (const L of pg.lines) {
        for (const [s, e] of findInLine(L, terms, qAmt)) res.push({ doc, p, L, s, e, rect: spanBBox(L, s, e) });
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
        <div class="sub">${esc(r.doc.name)} · page ${r.p}</div></div><button title="Search & Snip : inscrire ce résultat dans la cellule Excel active">Snip</button>`;
      li.querySelector(".grow").onclick = () => goToLocation(r.doc.id, r.p, r.rect);
      li.querySelector("button").onclick = (ev) => { ev.stopPropagation(); runAction(() => snipSearchResult(r), { label: "Snip impossible" }); };
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
        rg.load("values,text,rowIndex,columnIndex,rowCount,columnCount");
        const ws = rg.worksheet; ws.load("name");
        await ctx.sync();
        if (rg.rowCount * rg.columnCount > 20000) throw new Error("sélection trop grande (20 000 cellules max)");
        return { sheet: ws.name, row: rg.rowIndex, col: rg.columnIndex, values: rg.values, texts: rg.text };
      });
    },
    async readCell(sheet, row, col) {
      return Excel.run(async (ctx) => {
        const rg = ctx.workbook.worksheets.getItem(sheet).getRangeByIndexes(row, col, 1, 1);
        rg.load("values");
        await ctx.sync();
        return rg.values[0][0];
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
        c.load("rowIndex,columnIndex,text");
        const ws = c.worksheet; ws.load("name");
        const names = ctx.workbook.names; names.load("items/name,items/value");
        await ctx.sync();
        const links = {};
        for (const n of names.items) if (n.name.startsWith(NAME_PREFIX)) links[n.name.slice(NAME_PREFIX.length)] = parseRef(n.value);
        return { sheet: ws.name, row: c.rowIndex, col: c.columnIndex, links, text: String((c.text && c.text[0] && c.text[0][0]) || "") };
      });
    },
    async saveSnips(arr, meta) {
      return Excel.run(async (ctx) => {
        const ws = await ensureSheet(ctx, DATA_SHEET);
        ws.getRange().clear();
        const mj = JSON.stringify(meta || {});
        const metaRows = [];
        for (let i = 0, k = 0; i < Math.max(1, mj.length); i += CHUNK, k++) metaRows.push(["__meta__" + k, mj.slice(i, i + CHUNK)]);
        const rows = [["id", "json"], ...metaRows, ...arr.map((s) => [s.id, JSON.stringify(s)])];
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
        if (ws.isNullObject) return { list: [], meta: null };
        const ur = ws.getUsedRangeOrNullObject(true); ur.load("values");
        const names = ctx.workbook.names; names.load("items/name,items/value");
        await ctx.sync();
        if (ur.isNullObject) return { list: [], meta: null };
        const links = {};
        for (const n of names.items) if (n.name.startsWith(NAME_PREFIX)) links[n.name.slice(NAME_PREFIX.length)] = parseRef(n.value);
        const out = [];
        let metaStr = "";
        for (const [id, json] of ur.values.slice(1)) {
          if (String(id).startsWith("__meta__")) { metaStr += String(json); continue; }
          try {
            const sn = JSON.parse(json);
            const ref = links[id];
            if (ref) { sn.sheet = ref.sheet; sn.address = ref.address; }
            out.push(sn);
          } catch (_) { /* ligne illisible */ }
        }
        let meta = null;
        try { meta = metaStr ? JSON.parse(metaStr) : null; } catch (_) { meta = null; }
        return { list: out, meta };
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
        return { sheet: SHEET, row: g.r1, col: g.c1, values, texts: values.map((r) => r.map((v) => (typeof v === "number" ? fmtNum(v) : String(v)))) };
      },
      async readCell(sheet, row, col) { const v = cells.get(key(row, col)); return v ? v.v : ""; },
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
        const v = cells.get(key(act.r, act.c));
        return { sheet: SHEET, row: act.r, col: act.c, links, text: v ? (typeof v.v === "number" ? fmtNum(v.v) : String(v.v)) : "" };
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
    const menu = $("#mainMenu"); if (menu) menu.hidden = true;
    $$(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
    $$(".tab").forEach((t) => t.classList.toggle("active", t.id === "tab-" + name));
    if (name === "docs") refreshDocUI();
    if (name === "snips") refreshSnipList();
    if (name === "viewer") setTimeout(render, 0);
    if (S.padPanel) S.padPanel();
    const cf = $("#confirm"); if (cf && name !== "viewer") cf.hidden = true;
  }

  function setMode(m) {
    S.mode = m;
    $$(".modes button").forEach((b) => b.classList.toggle("active", b.dataset.mode === m));
    if (m !== "table" && S.table) { S.table = null; updateTablePreview(); }
    if (m !== "comment" && S.commentDraft) { S.commentDraft = null; $("#ctxPanel").hidden = true; }
    if (m !== "markup" && m !== "note" && S.panelKind === "markup") { $("#ctxPanel").hidden = true; S.panelKind = null; }
    drawOverlay();
    const tips = {
      text: "Texte : cliquez un montant ou un mot (double-clic pour la ligne entière), ou entourez une zone.",
      sum: "Somme : cliquez ou entourez les montants à additionner, puis « Insérer le total ».",
      valid: "Validation : cliquez l'élément vérifié ; un ✓ est inscrit dans la cellule active.",
      except: "Exception : cliquez l'élément en anomalie ; un ✗ est inscrit dans la cellule active.",
      table: "Tableau : cliquez dans le tableau, vérifiez l'aperçu, puis insérez.",
      comment: "Commentaire : entourez la zone concernée, tapez votre remarque — elle devient un commentaire Excel lié à la pièce.",
      markup: "Annotation : choisissez un outil (surligner, texte, flèche, coche, croix, caviarder) puis dessinez sur la pièce.",
      note: "Commentaire sur la pièce : cliquez à l'endroit voulu du document et tapez votre remarque.",
      formfield: "Extraction : entourez la VALEUR à extraire (ex. le numéro de facture), pas son libellé.",
      formkey: "Extraction : entourez le LIBELLÉ qui sert de repère (ex. « Facture n° »).",
    };
    // chaque conseil n'apparaît que les deux premières fois
    let seen = {};
    try { seen = JSON.parse(localStorage.getItem("snipcompta.tips") || "{}"); } catch (_) { /* ignore */ }
    if (tips[m] && (seen[m] || 0) < 2) {
      status(tips[m]); seen[m] = (seen[m] || 0) + 1;
      try { localStorage.setItem("snipcompta.tips", JSON.stringify(seen)); } catch (_) { /* ignore */ }
    } else if (m === "nav") status("");
    modeExtras(m);
    const chip = $("#modeChip");
    if (chip) {
      const names = { text: "Texte", sum: "Somme", valid: "Validation", except: "Exception", table: "Tableau", comment: "Commentaire", markup: "Annotation", note: "Bulle", formfield: "Extraction : la valeur", formkey: "Extraction : le repère" };
      chip.hidden = m === "nav";
      chip.textContent = (names[m] || m) + " ✕";
      chip.title = "Outil actif — cliquez (ou Échap) pour revenir à la navigation";
    }
  }

  function loadOpts() {
    try { Object.assign(S.opts, JSON.parse(localStorage.getItem("snipcompta.opts") || "{}")); } catch (_) { /* stockage indisponible */ }
    if ((S.opts.v || 0) < 5) { S.opts.autoSearch = true; S.opts.v = 5; saveOpts(); }   // v5 : surlignage activé par défaut
    $("#optStore").checked = S.opts.store; $("#optOcr").checked = S.opts.ocr; $("#optAdvance").checked = S.opts.advance;
    for (const k of ["pasteImport", "autoSearch", "sortFollow", "winOnImport", "cropMargins", "ocrBest"]) { const el = $("#opt_" + k); if (el) { el.checked = S.opts[k] !== false && !!S.opts[k]; el.onchange = () => { S.opts[k] = el.checked; saveOpts(); if (k === "cropMargins") render(); if (k === "ocrBest") rereadScans(); }; } }
  }
  /** Moteur OCR changé : les scans déjà lus seront relus avec le nouveau moteur. */
  function rereadScans() {
    let n = 0;
    for (const d of S.docs.values()) for (const pg of Object.values(d.pages || {})) {
      if (pg && pg.needsOcr && (pg.ocrDone || pg.ocrPromise)) { pg.ocrDone = false; pg.ocrPromise = null; pg.words = []; pg.lines = []; n++; }
    }
    status(S.opts.ocrBest ? "Lecture précise activée" + (n ? " : les scans ouverts vont être relus." : ".") : "Lecture standard rétablie" + (n ? " : les scans ouverts vont être relus." : "."), "ok");
    if (n) { const pe = $("#pages"); if (pe) pe.dataset.key = ""; render(); }
  }
  function saveOpts() { try { localStorage.setItem("snipcompta.opts", JSON.stringify(S.opts)); } catch (_) { /* ignore */ } }

  function bindUI() {
    $$(".tabs button").forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
    $$(".modes button").forEach((b) => (b.onclick = () => setMode(b.dataset.mode)));
    $$("[data-action=import]").forEach((b) => (b.onclick = () => $("#fileInput").click()));
    $("#fileInput").onchange = (e) => { const fs = [...e.target.files]; e.target.value = ""; hideImport(); runAction(() => importFiles(fs), { label: "Import impossible" }); };
    $("#docSelect").onchange = (e) => e.target.value && openDoc(e.target.value, 1);
    $("#prevPage").onclick = () => { const d = curDoc(); if (d && S.page > 1) scrollToPage(S.page - 1); };
    $("#nextPage").onclick = () => { const d = curDoc(); if (d && S.page < d.numPages) scrollToPage(S.page + 1); };
    const zstep = (f) => { const d = curDoc(); if (!d) return; let z = zoomFor(d); if (z === "page") z = 1; setZoom(d, z * f); };
    $("#zoomIn").onclick = () => zstep(1.2);
    $("#zoomOut").onclick = () => zstep(1 / 1.2);
    $("#zoomFitW").onclick = () => setZoom(curDoc(), 1);
    $("#zoomFitP").onclick = () => setZoom(curDoc(), "page");
    const rot = (q, all) => { const d = curDoc(); if (!d) return status("Ouvrez d'abord une pièce.", "warn"); runAction(() => rotateDoc(d, all ? "all" : S.page, q), { label: "Rotation impossible" }); };
    $("#rotR").onclick = (e) => rot(1, e.shiftKey); $("#rotL").onclick = (e) => rot(3, e.shiftKey);
    const mclose = () => { const m = $("#mainMenu"); if (m) m.hidden = true; };
    $("#mRotR").onclick = () => { mclose(); rot(1); }; $("#mRotL").onclick = () => { mclose(); rot(3); };
    $("#mRotAll").onclick = () => { mclose(); rot(1, true); };
    $("#undoBtn").onclick = $("#undoBtn2").onclick = () => runAction(() => undoLast(), { label: "Annulation impossible" });
    $("#refreshSnips").onclick = async () => { try { await reloadSnips(); status("Snips actualisés.", "ok"); } catch (e) { status(e.message, "err"); } };
    $("#searchForm").onsubmit = (e) => { e.preventDefault(); runAction(() => doSearch($("#searchInput").value), { label: "Recherche impossible" }); };
    $("#matchRun").onclick = runMatch;
    $("#sumsRun").onclick = () => runAction(() => findSums(), { label: "Recherche des sommes impossible" });
    $("#exportBtn").onclick = () => runAction(() => exportSnips(), { label: "Export impossible" });
    $("#delSelBtn").onclick = () => runAction(() => deleteSnipsInSelection(), { label: "Suppression impossible" });
    $("#optStore").onchange = (e) => { S.opts.store = e.target.checked; saveOpts(); };
    $("#optOcr").onchange = (e) => { S.opts.ocr = e.target.checked; saveOpts(); };
    $("#optAdvance").onchange = (e) => { S.opts.advance = e.target.checked; saveOpts(); };
    $("#optTheme").value = S.opts.theme || "auto";
    $("#optTheme").onchange = (e) => { S.opts.theme = e.target.value; saveOpts(); applyTheme(); };
    $("#progress .pcancel").onclick = requestCancel;

    document.addEventListener("keydown", (e) => {
      if (e.target.matches("input, select, textarea")) return;
      if (e.key === "Enter") {
        if (S.table) { e.preventDefault(); runAction(() => insertTable()); }
        else if (S.sumParts.length) { e.preventDefault(); runAction(() => insertSum()); }
      } else if (e.key === "Escape") {
        if (S.op) requestCancel();
        if (S.hitQuery) clearHits();
        hideHover(); hideLive(); $("#confirm").hidden = true;
        S.table = null; S.sumParts = []; S.commentDraft = null; S.selMarkup = null; $("#ctxPanel").hidden = true; hideMenu(); drawOverlay();
      } else if (e.ctrlKey && !e.altKey && ({ ",": 1, ".": 1, "[": 1, "=": 1, "m": 1, "M": 1, "z": 1, "Z": 1 })[e.key]) {
        e.preventDefault();
        const map = { ",": "valid", ".": "except", "[": "sum", "=": "text", m: "table", M: "table" };
        if (e.key === "z" || e.key === "Z") runAction(() => undoLast(), { label: "Annulation impossible" });
        else { switchTab("viewer"); setMode(map[e.key]); }
      } else if (e.key === "F3" && S.hits.length) {
        e.preventDefault(); gotoHit(S.hitIdx + (e.shiftKey ? -1 : 1));
      } else if (e.ctrlKey && e.key === "0") {
        e.preventDefault(); setZoom(curDoc(), 1);
      } else if (e.key === "Delete" && S.selMarkup) {
        deleteMarkup(S.selMarkup);
      } else if ((e.key === "ArrowRight" || e.key === "ArrowLeft") && !e.ctrlKey && $("#tab-viewer").classList.contains("active")) {
        (e.key === "ArrowRight" ? $("#nextPage") : $("#prevPage")).click();
      } else if (/^[1-9]$/.test(e.key) && !e.ctrlKey && !e.altKey && !e.metaKey) {
        if (+e.key > 7) { setMode(+e.key === 8 ? "markup" : "note"); return; }
        setMode(["nav", "text", "sum", "valid", "except", "table", "comment"][+e.key - 1]);
      } else if (e.key === "PageDown") $("#nextPage").click();
      else if (e.key === "PageUp") $("#prevPage").click();
    });

    const v = $("#viewer");
    v.addEventListener("dragover", (e) => { e.preventDefault(); v.classList.add("drop"); });
    v.addEventListener("dragleave", () => v.classList.remove("drop"));
    v.addEventListener("drop", (e) => { e.preventDefault(); v.classList.remove("drop"); if (e.dataTransfer.files.length) { const fs = [...e.dataTransfer.files]; runAction(() => importFiles(fs), { label: "Import impossible" }); } });
    let rt = null;
    window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(render, 150); });
    if (typeof ResizeObserver !== "undefined") {
      // le panneau flottant (somme, tableau…) ne doit jamais cacher le bas du document
      const cp = $("#ctxPanel");
      const pad = () => {
        const h = cp.hidden || !$("#tab-viewer").classList.contains("active") ? 0 : cp.offsetHeight;
        $("#viewer").style.paddingBottom = h ? h + 24 + "px" : "";
        document.body.style.setProperty("--panel-h", h ? h + 10 + "px" : "0px");
      };
      S.padPanel = pad;
      new ResizeObserver(pad).observe(cp);
      new MutationObserver(pad).observe(cp, { attributes: true, attributeFilter: ["hidden"] });
      let lastW = 0;
      new ResizeObserver(() => { const w = $("#viewer").clientWidth; if (w && Math.abs(w - lastW) > 2) { lastW = w; clearTimeout(rt); rt = setTimeout(render, 120); } }).observe($("#viewer"));
    }
    bindOverlay();
  }

  async function reloadSnips() {
    const res = await xl.loadSnips();
    const list = Array.isArray(res) ? res : res.list;
    if (res && res.meta) { S.meta = normalizeMeta(res.meta); applyDocMeta(); }
    S.snips.clear();
    for (const sn of list) S.snips.set(sn.id, sn);
    refreshSnipList(); drawOverlay();
  }

  /* ====================================================================== */
  /* V2 — organiseur, annotations, automatisations                          */
  /* ====================================================================== */
  function normalizeMeta(m) {
    m = m && typeof m === "object" ? m : {};
    return {
      markups: Array.isArray(m.markups) ? m.markups : [],
      notes: Array.isArray(m.notes) ? m.notes : [],
      docMeta: m.docMeta && typeof m.docMeta === "object" ? m.docMeta : {},
      folders: Array.isArray(m.folders) ? m.folders : [],
    };
  }
  S.meta = normalizeMeta(null);
  Object.assign(S, { markTool: "highlight", markColor: "#FFD400", selMarkup: null, selNote: null, diffMarks: [], sumMarks: [],
    importFolder: "", panelKind: null, snipFilter: "all", selBox: null, cmp: null,
    form: { mode: "doc", refDoc: null, fields: [], pendingKeyFor: null, rows: null },
    rm: { header: true, once: true, cols: [] } });
  STYLE.head = { fill: "#F2F4F7", font: "#1D2433", bold: true, color: "#667085", label: "", name: "En-tête" };
  const SNIP_TYPES = ["text", "sum", "valid", "except", "table", "comment"];
  const STYLE_DEFAULT = JSON.parse(JSON.stringify(STYLE));
  const libPdf = () => { if (typeof PDFLib === "undefined") throw new Error("module d'édition PDF absent (dossier lib/pdflib)"); return PDFLib; };
  let persistTimer = null;
  const persistSoon = () => { clearTimeout(persistTimer); persistTimer = setTimeout(persistSnips, 600); };
  const normRect = (m) => ({ x0: Math.min(m.x0, m.x1), y0: Math.min(m.y0, m.y1), x1: Math.max(m.x0, m.x1), y1: Math.max(m.y0, m.y1) });
  const inR = (pt, r, pad = 0) => pt.x >= r.x0 - pad && pt.x <= r.x1 + pad && pt.y >= r.y0 - pad && pt.y <= r.y1 + pad;
  const refOf = (sn) => parseRef(`'${String(sn.sheet).replace(/'/g, "''")}'!${sn.address}`);

  function applyDocMeta() {
    for (const [id, m] of Object.entries(S.meta.docMeta)) {
      const d = S.docs.get(id); if (!d) continue;
      if (m.name) d.name = m.name;
      if (m.folder != null) d.folder = m.folder;
    }
  }
  function setDocMeta(doc) { S.meta.docMeta[doc.id] = { name: doc.name, folder: doc.folder || "" }; }

  function openPanel(kind, html) {
    const panel = $("#ctxPanel");
    panel.hidden = false; S.panelKind = kind; panel.innerHTML = html;
    return panel;
  }
  function closePanel(kind) { if (!kind || S.panelKind === kind) { $("#ctxPanel").hidden = true; S.panelKind = null; } }

  /* ---------------- Styles personnalisables ----------------------------- */
  function loadStyles() {
    try {
      const s = JSON.parse(localStorage.getItem("snipcompta.styles") || "{}");
      for (const k of Object.keys(s)) if (STYLE[k] && k !== "head") Object.assign(STYLE[k], { fill: s[k].fill, font: s[k].font });
    } catch (_) { /* stockage indisponible */ }
  }
  function saveStyles() {
    const o = {};
    for (const k of SNIP_TYPES) o[k] = { fill: STYLE[k].fill, font: STYLE[k].font };
    try { localStorage.setItem("snipcompta.styles", JSON.stringify(o)); } catch (_) { /* ignore */ }
  }
  function renderStyleBox() {
    const box = $("#styleBox"); if (!box) return;
    box.innerHTML = SNIP_TYPES.map((k) => { const st = STYLE[k]; return `<div class="row"><span class="chip" style="background:${st.color}">${st.label}</span><span style="min-width:84px">${st.name}</span>
      fond <input type="color" data-k="${k}" data-p="fill" value="${st.fill.toLowerCase()}"> texte <input type="color" data-k="${k}" data-p="font" value="${st.font.toLowerCase()}"></div>`; }).join("")
      + `<div class="row"><button id="styleApply">Appliquer aux snips existants</button><button id="styleReset">Couleurs d'origine</button></div>`;
    $$("input[type=color]", box).forEach((inp) => (inp.onchange = () => { STYLE[inp.dataset.k][inp.dataset.p] = inp.value.toUpperCase(); saveStyles(); }));
    $("#styleApply").onclick = async () => {
      const list = [...S.snips.values()].filter((s) => s.address).map((s) => ({ sheet: s.sheet, address: s.address, type: s.type }));
      try { if (list.length) await xl.format(list); status(`Couleurs appliquées à ${list.length} snip(s).`, "ok"); } catch (e) { status("Mise en forme impossible : " + e.message, "err"); }
    };
    $("#styleReset").onclick = () => { for (const k of SNIP_TYPES) Object.assign(STYLE[k], { fill: STYLE_DEFAULT[k].fill, font: STYLE_DEFAULT[k].font }); saveStyles(); renderStyleBox(); };
  }

  /* ---------------- Import élargi : zip, dossiers, texte, mots de passe --- */
  const SUPPORTED_RE = /\.(pdf|png|jpe?g|jfif|bmp|gif|webp)$/i;
  const TEXT_RE = /\.(txt|csv|tsv|xml|json|log|md|html?)$/i;
  function mimeOf(n) {
    return /\.pdf$/i.test(n) ? "application/pdf" : /\.png$/i.test(n) ? "image/png" : /\.(jpe?g|jfif)$/i.test(n) ? "image/jpeg"
      : /\.gif$/i.test(n) ? "image/gif" : /\.bmp$/i.test(n) ? "image/bmp" : /\.webp$/i.test(n) ? "image/webp" : "text/plain";
  }
  async function expandFiles(files) {
    const out = [];
    for (const f of files) {
      try {
        if (/\.zip$/i.test(f.name)) {
          if (typeof JSZip === "undefined") throw new Error("module zip absent (dossier lib/jszip)");
          const zip = await JSZip.loadAsync(await f.arrayBuffer());
          const entries = Object.values(zip.files).filter((e) => !e.dir && !/(^|\/)(__MACOSX|\.)/.test(e.name));
          for (const e of entries) {
            const base = e.name.split("/").pop();
            if (SUPPORTED_RE.test(base) || TEXT_RE.test(base)) out.push(...await expandFiles([new File([await e.async("uint8array")], base, { type: mimeOf(base) })]));
          }
          status(`Archive « ${f.name} » ouverte (${entries.length} fichier(s)).`, "ok");
        } else if (TEXT_RE.test(f.name)) {
          let text = await f.text();
          if (/\.html?$/i.test(f.name)) { const d = new DOMParser().parseFromString(text, "text/html"); text = d.body ? d.body.textContent : text; }
          out.push(new File([await textToPdf(text, f.name)], f.name.replace(/\.[^.]+$/, "") + ".pdf", { type: "application/pdf" }));
        } else if (SUPPORTED_RE.test(f.name) || /^image\//.test(f.type) || /pdf/.test(f.type)) out.push(f);
        else status(`« ${f.name} » : format non pris en charge. Enregistrez-le d'abord en PDF (Fichier → Enregistrer sous → PDF).`, "warn");
      } catch (e) { status(`Import de « ${f.name} » impossible : ${e.message}`, "err"); }
    }
    return out;
  }
  const winAnsi = (s) => String(s).replace(/\t/g, "    ").replace(/[^\x20-\x7E\xA0-\xFF€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/g, "?");
  async function textToPdf(text, title) {
    const { PDFDocument, StandardFonts, rgb } = libPdf();
    const pdf = await PDFDocument.create();
    const font = await pdf.embedFont(StandardFonts.Courier);
    const size = 8.5, lh = 11, W = 595.28, H = 841.89, m = 36;
    const maxChars = Math.floor((W - 2 * m) / (size * 0.6));
    const lines = [];
    for (const raw of String(text).replace(/\r/g, "").split("\n")) {
      let l = winAnsi(raw);
      if (!l.length) { lines.push(""); continue; }
      while (l.length > maxChars) { lines.push(l.slice(0, maxChars)); l = l.slice(maxChars); }
      lines.push(l);
    }
    const perPage = Math.floor((H - 2 * m - 16) / lh);
    for (let i = 0; i < Math.max(1, lines.length); i += perPage) {
      const page = pdf.addPage([W, H]);
      page.drawText(winAnsi(title), { x: m, y: H - m, size: 9, font, color: rgb(0.4, 0.4, 0.4) });
      lines.slice(i, i + perPage).forEach((l, k) => page.drawText(l, { x: m, y: H - m - 16 - (k + 1) * lh, size, font }));
    }
    return pdf.save();
  }
  async function openPdfWithPassword(doc) {
    for (let attempt = 0; attempt < 6; attempt++) {
      try { return await pdfjsLib.getDocument({ data: doc.bytes.slice(), password: doc.password || undefined }).promise; }
      catch (e) {
        if (!e || e.name !== "PasswordException") throw e;
        doc.password = await askPassword(doc.name, attempt > 0);
        if (doc.password == null) throw new Error(`« ${doc.name} » est protégé par un mot de passe.`);
      }
    }
    throw new Error("Mot de passe incorrect.");
  }
  function askPassword(name, wrong) {
    return new Promise((resolve) => {
      switchTab("viewer");
      openPanel("pw", `<div><strong>Document protégé</strong> — « ${esc(name)} »${wrong ? ' <span class="ko">mot de passe incorrect</span>' : ""}</div>
        <div class="row"><input type="password" id="pwIn" placeholder="Mot de passe du PDF" style="flex:1"><button class="primary" id="pwOk">Ouvrir</button><button id="pwNo">Annuler</button></div>
        <div class="hint">Le mot de passe n'est pas enregistré : il sera redemandé à la réouverture du classeur.</div>`);
      const done = (v) => { closePanel("pw"); resolve(v); };
      $("#pwOk").onclick = () => done($("#pwIn").value);
      $("#pwIn").onkeydown = (e) => { if (e.key === "Enter") done($("#pwIn").value); };
      $("#pwNo").onclick = () => done(null);
      setTimeout(() => $("#pwIn").focus(), 0);
    });
  }

  /* ---------------- Organiseur de documents ------------------------------ */
  function renderDocList() {
    const ul = $("#docList"); if (!ul) return;
    const docs = [...S.docs.values()];
    const folders = [...new Set([...S.meta.folders, ...docs.map((d) => d.folder || "").filter(Boolean)])];
    ul.innerHTML = docs.length || folders.length ? "" : `<li class="hint">Aucun document pour l'instant.</li>`;
    const groups = [["", "Sans dossier"], ...folders.map((f) => [f, f])];
    for (const [key, label] of groups) {
      const inG = docs.filter((d) => (d.folder || "") === key);
      if (key === "" && !inG.length) continue;
      if (folders.length) {
        const h = document.createElement("li");
        h.className = "folder";
        h.innerHTML = `<span class="grow"><b>${esc(label)}</b> <span class="sub">${inG.length} document${inG.length > 1 ? "s" : ""}</span></span>${key ? `<button class="ico sm" title="Supprimer le dossier (les documents sont conservés)"><svg viewBox="0 0 20 20"><path d="M5 5l10 10M15 5L5 15"/></svg></button>` : ""}`;
        const b = h.querySelector("button");
        if (b) b.onclick = () => { S.meta.folders = S.meta.folders.filter((f) => f !== key); for (const d of inG) { d.folder = ""; setDocMeta(d); } persistSoon(); renderDocList(); };
        ul.appendChild(h);
      }
      for (const d of inG) ul.appendChild(docItem(d, folders));
    }
    const fs = $("#importFolderSel");
    if (fs) { fs.innerHTML = `<option value="">Sans dossier</option>` + folders.map((f) => `<option value="${esc(f)}">${esc(f)}</option>`).join(""); fs.value = folders.includes(S.importFolder) ? S.importFolder : ""; }
  }
  function docItem(d, folders) {
    const li = document.createElement("li");
    li.className = "doc" + (d.id === S.cur ? " cur" : "");
    const nSn = [...S.snips.values()].filter((s) => s.rects.some((r) => (r.docId || s.docId) === d.id)).length;
    const p1 = d.pages && d.pages[1];
    const txt = p1 ? (p1.needsOcr ? (p1.ocrDone ? "scan lu (OCR)" : "scan sans texte") : "") : "";
    const ext = (/\.([a-z0-9]{2,4})$/i.exec(d.name) || [, d.kind === "pdf" ? "pdf" : "img"])[1].toUpperCase();
    const meta = [d.numPages ? d.numPages + (d.numPages > 1 ? " pages" : " page") : "", nSn ? nSn + (nSn > 1 ? " snips" : " snip") : "aucun snip", d.stored ? "" : "non enregistré dans le classeur", txt].filter(Boolean).join(", ");
    li.innerHTML = `<span class="fico" data-ext="${esc(ext)}">${esc(ext.slice(0, 3))}</span>
      <div class="grow click"><div class="ttl">${esc(d.name)}</div><div class="sub">${esc(meta)}</div></div>
      <select data-f title="Ranger dans un dossier" aria-label="Dossier">${["", ...folders].map((f) => `<option value="${esc(f)}" ${f === (d.folder || "") ? "selected" : ""}>${f ? esc(f) : "Sans dossier"}</option>`).join("")}</select>
      <button class="ico sm" data-ren title="Renommer"><svg viewBox="0 0 20 20"><path d="M4 16l1-4 8.5-8.5 3 3L8 15z"/><path d="M11.5 5.5l3 3"/></svg></button>
      <button class="ico sm del" data-del title="Retirer ce document"><svg viewBox="0 0 20 20"><path d="M4.5 6h11M8 6V4h4v2M6 6l.8 10.5h6.4L14 6"/></svg></button>`;
    if (!folders.length) li.querySelector("[data-f]").hidden = true;
    const g = li.querySelector(".grow");
    g.onclick = () => { openDoc(d.id); switchTab("viewer"); };
    li.querySelector("[data-f]").onchange = (e) => { d.folder = e.target.value; setDocMeta(d); persistSoon(); renderDocList(); };
    li.querySelector("[data-ren]").onclick = () => {
      g.onclick = null;
      g.innerHTML = `<input type="text" value="${esc(d.name)}" style="width:100%">`;
      const inp = g.querySelector("input"); inp.focus(); inp.select();
      let done = false;
      const ok = () => {
        if (done) return; done = true;
        const v = inp.value.trim();
        if (v && v !== d.name) { d.name = v; setDocMeta(d); for (const sn of S.snips.values()) if (sn.docId === d.id) sn.docName = v; persistSoon(); }
        refreshDocUI();
      };
      inp.onkeydown = (e) => { if (e.key === "Enter") ok(); if (e.key === "Escape") { done = true; refreshDocUI(); } };
      inp.onblur = ok;
    };
    const del = li.querySelector("[data-del]");
    del.onclick = () => {
      if (del.dataset.confirm) { deleteDoc(d.id); return; }
      del.dataset.confirm = "1"; del.classList.add("confirm"); del.title = "Cliquez encore pour retirer";
      status(`Cliquez encore sur la corbeille pour retirer « ${d.name} ».`, "warn");
      setTimeout(() => { del.dataset.confirm = ""; del.classList.remove("confirm"); del.title = "Retirer ce document"; }, 3000);
    };
    return li;
  }

  async function replaceDocBytes(doc, bytes, mime) {
    doc.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (mime) doc.mime = mime;
    doc.pdf = null; doc.img = null; doc.ready = null; doc.pages = {}; doc.dims = null; doc.rev = (doc.rev || 0) + 1;
    if (S.inOffice && doc.stored) {
      await xl.deleteDoc(doc.id);
      await xl.saveDoc(doc, (pct) => status(`Mise à jour de « ${doc.name} » dans le classeur… ${pct} %`, "saving"));
    }
    await openDocData(doc);
    if (S.cur === doc.id) { S.page = clamp(S.page, 1, doc.numPages); await render(); }
  }

  /* Retirer les pages sans snips / sans exceptions */
  async function removePages(doc, onlyExceptions) {
    await openDocData(doc);
    if (doc.kind !== "pdf" || doc.numPages < 2) return status(`« ${doc.name} » n'a qu'une page : rien à retirer.`, "warn");
    const keep = new Set();
    for (const sn of S.snips.values()) {
      if (onlyExceptions && sn.type !== "except") continue;
      for (const r of sn.rects) if ((r.docId || sn.docId) === doc.id) keep.add(r.page);
    }
    const kept = [...keep].filter((p) => p <= doc.numPages).sort((a, b) => a - b);
    if (!kept.length) return status(`Aucune page de « ${doc.name} » n'a de ${onlyExceptions ? "snip Exception" : "snip"} : rien n'a été retiré.`, "warn");
    if (kept.length === doc.numPages) return status(`Toutes les pages de « ${doc.name} » ont des snips : rien à retirer.`, "ok");
    const { PDFDocument } = libPdf();
    const src = await PDFDocument.load(doc.bytes);
    const out = await PDFDocument.create();
    (await out.copyPages(src, kept.map((p) => p - 1))).forEach((pg) => out.addPage(pg));
    const map = new Map(kept.map((p, i) => [p, i + 1]));
    const orphan = [];
    for (const sn of S.snips.values()) {
      sn.rects = sn.rects.filter((r) => (r.docId || sn.docId) !== doc.id || map.has(r.page))
        .map((r) => ((r.docId || sn.docId) === doc.id ? { ...r, page: map.get(r.page) } : r));
      if (!sn.rects.length) orphan.push(sn.id);
    }
    const remap = (arr) => arr.filter((m) => m.docId !== doc.id || map.has(m.page)).map((m) => (m.docId === doc.id ? { ...m, page: map.get(m.page) } : m));
    S.meta.markups = remap(S.meta.markups); S.meta.notes = remap(S.meta.notes);
    const removed = doc.numPages - kept.length;
    await replaceDocBytes(doc, await out.save());
    if (orphan.length) await deleteSnips(orphan); else await persistSnips();
    drawOverlay(); refreshDocUI();
    status(`${removed} page(s) retirée(s) de « ${doc.name} »${orphan.length ? ` · ${orphan.length} snip(s) sans preuve supprimé(s)` : ""}.`, "ok");
  }

  async function compressDoc(doc) {
    await openDocData(doc);
    const before = doc.bytes.length;
    let bytes, mime = null;
    if (doc.kind === "image") {
      const c = document.createElement("canvas"); c.width = doc.img.naturalWidth; c.height = doc.img.naturalHeight;
      const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height); x.drawImage(doc.img, 0, 0);
      bytes = new Uint8Array(await (await new Promise((r) => c.toBlob(r, "image/jpeg", 0.72))).arrayBuffer()); mime = "image/jpeg";
    } else {
      const { PDFDocument } = libPdf();
      bytes = await (await PDFDocument.load(doc.bytes)).save({ useObjectStreams: true });
    }
    if (bytes.length < before * 0.95) { await replaceDocBytes(doc, bytes, mime); return [before, bytes.length]; }
    return [before, before];
  }
  async function compressAll() {
    let b = 0, a = 0;
    for (const d of S.docs.values()) {
      try { const [x, y] = await compressDoc(d); b += x; a += y; } catch (e) { status(`« ${d.name} » : ${e.message}`, "warn"); }
    }
    const mb = (n) => (n / 1048576).toFixed(1).replace(".", ",");
    status(b ? `Compression : ${mb(b)} Mo → ${mb(a)} Mo${a >= b ? " (déjà optimisés)" : ""}. Enregistrez le classeur.` : "Aucun document à compresser.", "ok");
  }

  /* Transfert du travail vers une nouvelle version d'un document */
  function offerVersionTransfer(oldDoc, newDoc) {
    const n = [...S.snips.values()].filter((sn) => sn.rects.some((r) => (r.docId || sn.docId) === oldDoc.id)).length
      + S.meta.markups.filter((m) => m.docId === oldDoc.id).length + S.meta.notes.filter((m) => m.docId === oldDoc.id).length;
    if (!n) return;
    switchTab("viewer");
    openPanel("version", `<div><strong>Nouvelle version détectée</strong> : « ${esc(newDoc.name)} » existe déjà avec un contenu différent.</div>
      <div class="hint">${n} snip(s) ou annotation(s) pointent vers l'ancienne version. SnipCompta peut les transférer vers la nouvelle (en les recalant sur le même texte quand c'est possible) puis retirer l'ancienne.</div>
      <div class="row"><button class="primary" id="vTr">Transférer vers la nouvelle version</button><button id="vKeep">Garder les deux</button></div>`);
    $("#vKeep").onclick = () => closePanel("version");
    $("#vTr").onclick = () => runAction(() => transferVersion(oldDoc, newDoc), { label: "Transfert impossible" }).finally(() => closePanel("version"));
  }
  async function transferVersion(oldDoc, newDoc) {
    await openDocData(oldDoc); await openDocData(newDoc);
    let moved = 0, kept = 0;
    const relocate = async (r) => {
      const page = Math.min(r.page, newDoc.numPages);
      const opg = await pageText(oldDoc, r.page, { allowOcr: false });
      const txt = norm(joinWords(wordsInRect(opg.words, r))).trim();
      const probe = txt.split(/\s+/).slice(0, 6).join(" ");
      if (probe.length >= 2) {
        const order = [page, ...Array.from({ length: newDoc.numPages }, (_, i) => i + 1).filter((x) => x !== page)];
        for (const p of order) {
          const npg = await pageText(newDoc, p, { allowOcr: false });
          let best = null;
          for (const L of npg.lines) {
            const k = norm(L.text).indexOf(probe);
            if (k < 0) continue;
            const bb = spanBBox(L, k, k + probe.length);
            const d = Math.abs(bb.y0 - r.y0) + Math.abs(bb.x0 - r.x0);
            if (!best || d < best.d) best = { d, bb };
          }
          if (best && (probe.length >= 8 || (p === r.page && best.d < 60))) { moved++; return { ...r, docId: newDoc.id, page: p, x0: best.bb.x0, y0: best.bb.y0, x1: best.bb.x0 + (r.x1 - r.x0), y1: best.bb.y0 + (r.y1 - r.y0) }; }
        }
      }
      kept++;
      return { ...r, docId: newDoc.id, page };
    };
    for (const sn of S.snips.values()) {
      if (!sn.rects.some((r) => (r.docId || sn.docId) === oldDoc.id)) continue;
      const nr = [];
      for (const r of sn.rects) nr.push((r.docId || sn.docId) === oldDoc.id ? await relocate({ ...r }) : r);
      sn.rects = nr;
      if (sn.docId === oldDoc.id) { sn.docId = newDoc.id; sn.docName = newDoc.name; }
    }
    for (const m of S.meta.markups) if (m.docId === oldDoc.id) { const nr = await relocate({ ...normRect(m), page: m.page }); const dx = nr.x0 - Math.min(m.x0, m.x1), dy = nr.y0 - Math.min(m.y0, m.y1); Object.assign(m, { x0: m.x0 + dx, x1: m.x1 + dx, y0: m.y0 + dy, y1: m.y1 + dy, page: nr.page, docId: newDoc.id }); }
    for (const n of S.meta.notes) if (n.docId === oldDoc.id) { const nr = await relocate({ page: n.page, x0: n.x - 4, y0: n.y - 4, x1: n.x + 4, y1: n.y + 4 }); Object.assign(n, { page: nr.page, x: nr.x0 + 4, y: nr.y0 + 4, docId: newDoc.id }); }
    newDoc.folder = oldDoc.folder || ""; setDocMeta(newDoc);
    await deleteDoc(oldDoc.id);
    await persistSnips();
    await openDoc(newDoc.id, S.page);
    refreshSnipList();
    status(`Travail transféré vers la nouvelle version : ${moved} zone(s) recalée(s) sur le texte, ${kept} gardée(s) à la même position.`, "ok");
  }

  /* ---------------- Rotation des pages ------------------------------------ */
  /** Point (x, y) d'une page W×H après q quarts de tour dans le sens des aiguilles d'une montre. */
  function rotPt(x, y, W, H, q) {
    if (q === 1) return [H - y, x];
    if (q === 2) return [W - x, H - y];
    return [y, W - x];
  }
  function rotRect(r, W, H, q) {
    const [ax, ay] = rotPt(r.x0, r.y0, W, H, q), [bx, by] = rotPt(r.x1, r.y1, W, H, q);
    return { x0: Math.min(ax, bx), y0: Math.min(ay, by), x1: Math.max(ax, bx), y1: Math.max(ay, by) };
  }
  /** Pivote une page (ou "all") de q quarts de tour : la pièce est vraiment tournée, les snips suivent. */
  async function rotateDoc(doc, which, q) {
    q = ((q % 4) + 4) % 4; if (!q) return;
    await openDocData(doc);
    const dims = (await pageDims(doc)).map((d) => ({ ...d }));
    const pages = which === "all" ? Array.from({ length: doc.numPages }, (_, i) => i + 1)
      : Array.isArray(which) ? which.filter((p) => p >= 1 && p <= doc.numPages) : [clamp(+which || 1, 1, doc.numPages)];
    if (!pages.length) return;
    status(pages.length > 1 ? `Rotation des ${pages.length} pages…` : `Rotation de la page ${pages[0]}…`, "busy");
    let bytes, mime = null;
    if (doc.kind === "pdf") {
      const { PDFDocument, degrees } = libPdf();
      let pdf;
      try { pdf = await PDFDocument.load(doc.bytes); }
      catch (e) { throw new Error(/encrypt/i.test(e.message) ? "ce PDF est protégé : il ne peut pas être modifié" : e.message); }
      for (const p of pages) { const pg = pdf.getPage(p - 1); pg.setRotation(degrees((((pg.getRotation().angle || 0) + 90 * q) % 360 + 360) % 360)); }
      bytes = await pdf.save();
    } else {
      const img = doc.img, W = img.naturalWidth, H = img.naturalHeight;
      const c = document.createElement("canvas"); c.width = q % 2 ? H : W; c.height = q % 2 ? W : H;
      const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
      if (q === 1) { ctx.translate(H, 0); ctx.rotate(Math.PI / 2); }
      else if (q === 2) { ctx.translate(W, H); ctx.rotate(Math.PI); }
      else { ctx.translate(0, W); ctx.rotate(-Math.PI / 2); }
      ctx.drawImage(img, 0, 0);
      mime = /png/i.test(doc.mime || "") ? "image/png" : "image/jpeg";
      bytes = new Uint8Array(await (await new Promise((r) => c.toBlob(r, mime, 0.93))).arrayBuffer());
    }
    // snips, annotations et bulles suivent la page
    const on = (docId, page) => docId === doc.id && pages.includes(page);
    for (const sn of S.snips.values()) sn.rects = sn.rects.map((r) => (on(r.docId || sn.docId, r.page) ? { ...r, ...rotRect(r, dims[r.page - 1].w, dims[r.page - 1].h, q) } : r));
    for (const m of S.meta.markups) if (on(m.docId, m.page)) {
      const d = dims[m.page - 1], [a, b] = rotPt(m.x0, m.y0, d.w, d.h, q), [c, e] = rotPt(m.x1, m.y1, d.w, d.h, q);
      Object.assign(m, { x0: a, y0: b, x1: c, y1: e });
    }
    for (const n of S.meta.notes) if (on(n.docId, n.page)) { const d = dims[n.page - 1]; [n.x, n.y] = rotPt(n.x, n.y, d.w, d.h, q); }
    if (S.table && on(S.table.docId, S.table.page)) S.table = null;
    S.sumParts = S.sumParts.filter((p) => !on(p.docId, p.page));
    if (S.hits.some((h) => on(h.docId, h.page))) clearHits();
    const keep = S.page;
    await replaceDocBytes(doc, bytes, mime);
    await persistSnips();
    if (S.cur === doc.id) scrollToPage(clamp(keep, 1, doc.numPages));
    drawOverlay(); refreshSnipList();
    docChangedElsewhere(doc);
    const what = pages.length > 1 ? "Document pivoté" : `Page ${pages[0]} pivotée`;
    status(`${what}${doc.stored ? " et enregistré dans le classeur" : ""}.`, "ok",
      pages.length === 1 && doc.numPages > 1 ? [{ label: "Pivoter aussi les autres pages", fn: () => runAction(() => rotateDoc(doc, Array.from({ length: doc.numPages }, (_, i) => i + 1).filter((p) => p !== pages[0]), q), { label: "Rotation impossible" }) }] : null);
  }
  /** La pièce a changé ici : l'autre vue (panneau ou fenêtre séparée) la recharge. */
  function docChangedElsewhere(doc) {
    const msg = { event: "docChanged", doc: doc.id };
    if (S.isWindow) parentSend(msg); else if (S.win) winSend(msg);
  }
  async function onDocChanged(id) {
    const d = S.docs.get(id); if (!d) return;
    d.bytes = null; d.pdf = null; d.img = null; d.ready = null; d.pages = {}; d.dims = null; d.rev = (d.rev || 0) + 1;
    try { await reloadSnips(); } catch (_) { /* ignore */ }
    if (S.cur === id) await render(); else drawOverlay();
    refreshSnipList();
  }

  /* ---------------- Export des pièces annotées --------------------------- */
  const ASCII_LABEL = { text: "T", sum: "S", valid: "V", except: "X", table: "TB", comment: "C" };
  function hexRgb(hex) { const { rgb } = libPdf(); return rgb(parseInt(hex.slice(1, 3), 16) / 255, parseInt(hex.slice(3, 5), 16) / 255, parseInt(hex.slice(5, 7), 16) / 255); }
  async function embedDocImage(pdf, doc) {
    if (/png/.test(doc.mime)) return pdf.embedPng(doc.bytes);
    if (/jpe?g/.test(doc.mime)) return pdf.embedJpg(doc.bytes);
    const c = document.createElement("canvas"); c.width = doc.img.naturalWidth; c.height = doc.img.naturalHeight;
    c.getContext("2d").drawImage(doc.img, 0, 0);
    return pdf.embedPng(new Uint8Array(await (await new Promise((r) => c.toBlob(r, "image/png"))).arrayBuffer()));
  }
  async function annotatedPdf(doc) {
    const { PDFDocument, StandardFonts } = libPdf();
    await openDocData(doc);
    let pdf;
    if (doc.kind === "pdf") pdf = await PDFDocument.load(doc.bytes);
    else {
      pdf = await PDFDocument.create();
      const img = await embedDocImage(pdf, doc);
      const w = doc.img.naturalWidth, h = doc.img.naturalHeight;
      pdf.addPage([w, h]).drawImage(img, { x: 0, y: 0, width: w, height: h });
    }
    // scans de travers : on exporte la page redressée, comme à l'écran (les cadres tombent juste)
    for (let p = 1; p <= doc.numPages; p++) {
      const pg = await pageText(doc, p, { allowOcr: false });
      const skew = await pageSkew(doc, p, pg);
      if (!skew) continue;
      const f = doc.kind === "pdf" ? clamp(2200 / pg.w, 1.5, 3) : 1;
      const c = document.createElement("canvas"); c.width = Math.ceil(pg.w * f); c.height = Math.ceil(pg.h * f);
      await paintPage(c.getContext("2d"), doc, p, f, skew);
      const img = await pdf.embedJpg(new Uint8Array(await (await new Promise((r) => c.toBlob(r, "image/jpeg", 0.88))).arrayBuffer()));
      const { width, height } = pdf.getPage(p - 1).getSize();
      pdf.insertPage(p - 1, [width, height]).drawImage(img, { x: 0, y: 0, width, height });
      pdf.removePage(p);
    }
    const font = await pdf.embedFont(StandardFonts.Helvetica);
    const pages = pdf.getPages();
    const P = (n) => pages[n - 1];
    const fy = (page, y) => page.getHeight() - y;
    for (const sn of S.snips.values()) for (const r of sn.rects) {
      const page = (r.docId || sn.docId) === doc.id && P(r.page); if (!page) continue;
      const st = STYLE[sn.type] || STYLE.text, c = hexRgb(st.color);
      page.drawRectangle({ x: r.x0, y: fy(page, r.y1), width: r.x1 - r.x0, height: r.y1 - r.y0, borderColor: c, borderWidth: 1.1, color: c, opacity: 0.12, borderOpacity: 0.9 });
      page.drawText(winAnsi(`${ASCII_LABEL[sn.type] || ""} ${sn.sheet}!${sn.address}`), { x: r.x0, y: fy(page, r.y0) + 1.5, size: 5.5, font, color: c });
    }
    for (const m of S.meta.markups) {
      const page = m.docId === doc.id && P(m.page); if (!page) continue;
      const r = normRect(m), c = hexRgb(m.color || "#000000");
      if (m.kind === "highlight") page.drawRectangle({ x: r.x0, y: fy(page, r.y1), width: r.x1 - r.x0, height: r.y1 - r.y0, color: c, opacity: 0.35 });
      else if (m.kind === "redact") page.drawRectangle({ x: r.x0, y: fy(page, r.y1), width: r.x1 - r.x0, height: r.y1 - r.y0, color: hexRgb("#000000") });
      else if (m.kind === "text") page.drawText(winAnsi(m.text || ""), { x: r.x0 + 1, y: fy(page, r.y1) + 3, size: 9, font, color: c });
      else if (m.kind === "check") { page.drawLine({ start: { x: r.x0, y: fy(page, (r.y0 + r.y1) / 2) }, end: { x: r.x0 + (r.x1 - r.x0) * 0.4, y: fy(page, r.y1) }, thickness: 1.6, color: c }); page.drawLine({ start: { x: r.x0 + (r.x1 - r.x0) * 0.4, y: fy(page, r.y1) }, end: { x: r.x1, y: fy(page, r.y0) }, thickness: 1.6, color: c }); }
      else if (m.kind === "cross") { page.drawLine({ start: { x: r.x0, y: fy(page, r.y0) }, end: { x: r.x1, y: fy(page, r.y1) }, thickness: 1.6, color: c }); page.drawLine({ start: { x: r.x0, y: fy(page, r.y1) }, end: { x: r.x1, y: fy(page, r.y0) }, thickness: 1.6, color: c }); }
      else if (m.kind === "arrow") {
        const a = { x: m.x0, y: fy(page, m.y0) }, b = { x: m.x1, y: fy(page, m.y1) };
        page.drawLine({ start: a, end: b, thickness: 1.4, color: c });
        const ang = Math.atan2(b.y - a.y, b.x - a.x);
        for (const s of [-0.45, 0.45]) page.drawLine({ start: b, end: { x: b.x - 7 * Math.cos(ang + s), y: b.y - 7 * Math.sin(ang + s) }, thickness: 1.4, color: c });
      }
    }
    for (const n of S.meta.notes) {
      const page = n.docId === doc.id && P(n.page); if (!page) continue;
      const lines = winAnsi(n.text).match(/.{1,48}(\s|$)|\S+/g) || [""];
      const w = 150, h = 9 * lines.length + 6, x = Math.min(n.x + 6, page.getWidth() - w - 4), y = fy(page, n.y) - h;
      page.drawRectangle({ x, y, width: w, height: h, color: hexRgb("#FFF4C2"), borderColor: hexRgb("#E39B00"), borderWidth: 0.8 });
      lines.forEach((l, i) => page.drawText(l.trim(), { x: x + 3, y: y + h - 10 - i * 9, size: 7, font }));
    }
    return pdf.save();
  }
  function download(bytes, name, mime) {
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const a = document.createElement("a"); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 120000);
    status(`« ${name} » est prêt : regardez dans vos Téléchargements.`, "ok");
  }
  async function exportAnnotated(all) {
    const docs = all ? [...S.docs.values()] : [curDoc()].filter(Boolean);
    if (!docs.length) return status("Aucun document à exporter.", "warn");
    const outs = [];
    for (const d of docs) {
      status(`Préparation de « ${d.name} »…`, "busy");
      try { outs.push({ name: d.name.replace(/\.[^.]+$/, "") + " - annoté.pdf", bytes: await annotatedPdf(d) }); }
      catch (e) { status(`« ${d.name} » : ${e.message}`, "warn"); }
    }
    if (!outs.length) return;
    if (outs.length === 1) return download(outs[0].bytes, outs[0].name, "application/pdf");
    const zip = new JSZip(); outs.forEach((o) => zip.file(o.name, o.bytes));
    download(await zip.generateAsync({ type: "uint8array" }), "Pièces annotées.zip", "application/zip");
  }

  /* ---------------- Annotations : surlignage, texte, flèches, caviardage -- */
  function drawExtras(ov, doc, p) {
    const sc = S.scale;
    const place = (el, r) => { el.style.left = r.x0 * sc + "px"; el.style.top = r.y0 * sc + "px"; el.style.width = Math.max(2, (r.x1 - r.x0) * sc) + "px"; el.style.height = Math.max(2, (r.y1 - r.y0) * sc) + "px"; return el; };
    for (const m of S.meta.markups) {
      if (m.docId !== doc.id || m.page !== p) continue;
      const sel = S.selMarkup === m.id;
      if (m.kind === "arrow") {
        const len = Math.hypot(m.x1 - m.x0, m.y1 - m.y0) * sc, ang = Math.atan2(m.y1 - m.y0, m.x1 - m.x0);
        const el = document.createElement("div");
        el.className = "mk mk-arrow" + (sel ? " sel" : "");
        el.style.left = m.x0 * sc + "px"; el.style.top = m.y0 * sc + "px"; el.style.width = len + "px"; el.style.color = m.color;
        el.style.transform = `rotate(${ang}rad)`;
        ov.appendChild(el); continue;
      }
      const el = place(document.createElement("div"), normRect(m));
      el.className = "mk mk-" + m.kind + (sel ? " sel" : "");
      if (m.kind === "highlight") el.style.background = m.color + "66";
      else if (m.kind === "text") { el.textContent = m.text; el.style.color = m.color; el.style.fontSize = Math.max(8, 9 * sc) + "px"; }
      else if (m.kind === "check" || m.kind === "cross") { el.textContent = m.kind === "check" ? "✓" : "✗"; el.style.color = m.color; el.style.fontSize = Math.max(10, (Math.abs(m.y1 - m.y0)) * sc) + "px"; }
      ov.appendChild(el);
    }
    for (const n of S.meta.notes) {
      if (n.docId !== doc.id || n.page !== p) continue;
      const el = document.createElement("div");
      el.className = "note" + (n.resolved ? " done" : "") + (S.selNote === n.id ? " sel" : "");
      el.style.left = n.x * sc - 9 + "px"; el.style.top = n.y * sc - 9 + "px"; el.textContent = "💬"; el.title = n.text;
      ov.appendChild(el);
    }
    const marks = [...S.diffMarks.map((m) => ({ ...m, cls: m.kind === "add" ? "dif-add" : "dif-del" })), ...S.sumMarks.map((m) => ({ ...m, cls: m.kind === "total" ? "sum-total" : "sum-part" }))];
    for (const m of marks) if (m.docId === doc.id && m.page === p) { const el = place(document.createElement("div"), m.rect); el.className = "ovbox " + m.cls; ov.appendChild(el); }
    const F = S.form;
    if (F.refDoc === doc.id) for (const f of F.fields) if (f.page === p) {
      const el = place(document.createElement("div"), f.rect); el.className = "ovbox form-val";
      el.innerHTML = `<span class="badge" style="background:#0F8B8D">${esc(f.name.slice(0, 16))}</span>`; ov.appendChild(el);
      if (f.key) { const k = place(document.createElement("div"), f.key.rect); k.className = "ovbox form-key"; ov.appendChild(k); }
    }
  }
  function markupAt(pt) {
    const doc = curDoc(); if (!doc) return null;
    for (let i = S.meta.markups.length - 1; i >= 0; i--) {
      const m = S.meta.markups[i];
      if (m.docId === doc.id && m.page === S.page && inR(pt, normRect(m), 3)) return m;
    }
    return null;
  }
  function addMarkup(o) {
    const doc = curDoc(); if (!doc) return;
    const m = { id: uid(), docId: doc.id, page: S.page, color: o.kind === "redact" ? "#000000" : S.markColor, ...o };
    S.meta.markups.push(m); S.selMarkup = m.id;
    persistSoon(); drawOverlay(); renderAnnotList();
    if (S.mode === "markup") showMarkupPanel();
  }
  function addMarkupAt(pt, rect) {
    if (S.markTool === "text") {
      openPanel("mktext", `<div><strong>Texte sur la pièce</strong></div><div class="row"><input type="text" id="mkTxt" style="flex:1" placeholder="Ex. : vu, conforme au contrat"><button class="primary" id="mkTxtOk">Ajouter</button><button id="mkTxtNo">Annuler</button></div>`);
      const go = () => {
        const t = $("#mkTxt").value.trim(); if (!t) return;
        const w = Math.max(30, t.length * 4.8 + 6);
        const r = rect || { x0: pt.x, y0: pt.y - 6, x1: pt.x + w, y1: pt.y + 8 };
        closePanel("mktext"); addMarkup({ kind: "text", text: t, ...r });
      };
      $("#mkTxtOk").onclick = go; $("#mkTxt").onkeydown = (e) => { if (e.key === "Enter") go(); };
      $("#mkTxtNo").onclick = () => { closePanel("mktext"); showMarkupPanel(); };
      setTimeout(() => $("#mkTxt").focus(), 0);
      return;
    }
    const s = 8;
    addMarkup({ kind: S.markTool, x0: pt.x - s, y0: pt.y - s, x1: pt.x + s, y1: pt.y + s });
  }
  function deleteMarkup(id) {
    S.meta.markups = S.meta.markups.filter((m) => m.id !== id);
    if (S.selMarkup === id) S.selMarkup = null;
    persistSoon(); drawOverlay(); renderAnnotList();
    if (S.mode === "markup") showMarkupPanel();
    status("Annotation supprimée.", "ok");
  }
  function showMarkupPanel() {
    const tools = [["highlight", "▮ Surligner"], ["text", "A Texte"], ["arrow", "➜ Flèche"], ["check", "✓ Coche"], ["cross", "✗ Croix"], ["redact", "■ Caviarder"]];
    const colors = ["#FFD400", "#22A04B", "#D93025", "#2F6FD6", "#111111"];
    const pending = S.meta.markups.filter((m) => m.kind === "redact" && m.docId === S.cur).length;
    const panel = openPanel("markup", `<div class="row">${tools.map(([k, l]) => `<button data-tool="${k}" class="${S.markTool === k ? "primary" : ""}">${l}</button>`).join("")}</div>
      <div class="row">Couleur ${colors.map((c) => `<button data-col="${c}" class="swatch${S.markColor === c ? " on" : ""}" style="background:${c}" title="${c}"></button>`).join("")}
        <button id="mkDel" ${S.selMarkup ? "" : "disabled"}>Supprimer la sélection</button></div>
      ${pending ? `<div class="row"><span class="ko">${pending} zone(s) à caviarder sur ce document.</span><button class="primary" id="mkRedact">Appliquer le caviardage (définitif)</button></div>
        <div class="hint">Le caviardage remplace la page par une image noircie : le texte masqué disparaît vraiment du document enregistré.</div>` : ""}
      <div class="hint">Cliquez une annotation pour la sélectionner (touche Suppr pour l'effacer). Les annotations sont enregistrées dans le classeur et incluses à l'export des pièces.</div>`);
    $$("[data-tool]", panel).forEach((b) => (b.onclick = () => { S.markTool = b.dataset.tool; showMarkupPanel(); drawOverlay(); }));
    $$("[data-col]", panel).forEach((b) => (b.onclick = () => { S.markColor = b.dataset.col; const m = S.meta.markups.find((x) => x.id === S.selMarkup); if (m && m.kind !== "redact") { m.color = S.markColor; persistSoon(); drawOverlay(); } showMarkupPanel(); }));
    $("#mkDel").onclick = () => S.selMarkup && deleteMarkup(S.selMarkup);
    const rb = $("#mkRedact");
    if (rb) rb.onclick = () => {
      if (!rb.dataset.confirm) { rb.dataset.confirm = "1"; rb.textContent = "Confirmer : c'est définitif"; return; }
      runAction(() => applyRedactions(curDoc()), { label: "Caviardage impossible" });
    };
  }
  async function applyRedactions(doc) {
    const zones = S.meta.markups.filter((m) => m.kind === "redact" && m.docId === doc.id);
    if (!zones.length) return;
    await openDocData(doc);
    status("Caviardage en cours…", "busy");
    const byPage = new Map();
    for (const z of zones) { if (!byPage.has(z.page)) byPage.set(z.page, []); byPage.get(z.page).push(normRect(z)); }
    if (doc.kind === "image") {
      const c = document.createElement("canvas"); c.width = doc.img.naturalWidth; c.height = doc.img.naturalHeight;
      const x = c.getContext("2d");
      await paintPage(x, doc, 1, 1, await pageSkew(doc, 1, await pageText(doc, 1, { allowOcr: false })));
      x.fillStyle = "#000";
      for (const r of byPage.get(1) || []) x.fillRect(r.x0, r.y0, r.x1 - r.x0, r.y1 - r.y0);
      await replaceDocBytes(doc, new Uint8Array(await (await new Promise((r) => c.toBlob(r, "image/png"))).arrayBuffer()), "image/png");
    } else {
      const { PDFDocument } = libPdf();
      const pdf = await PDFDocument.load(doc.bytes);
      for (const [p, rects] of [...byPage.entries()].sort((a, b) => a[0] - b[0])) {
        const page = await doc.pdf.getPage(p);
        const f = clamp(2200 / page.getViewport({ scale: 1 }).width, 1.5, 3);
        const vp = page.getViewport({ scale: f });
        const c = document.createElement("canvas"); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
        const ctx = c.getContext("2d");
        await paintPage(ctx, doc, p, f, await pageSkew(doc, p, await pageText(doc, p, { allowOcr: false })));
        ctx.fillStyle = "#000";
        for (const r of rects) ctx.fillRect(r.x0 * f, r.y0 * f, (r.x1 - r.x0) * f, (r.y1 - r.y0) * f);
        const png = new Uint8Array(await (await new Promise((res) => c.toBlob(res, "image/png"))).arrayBuffer());
        const img = await pdf.embedPng(png);
        const { width, height } = pdf.getPage(p - 1).getSize();
        pdf.insertPage(p - 1, [width, height]).drawImage(img, { x: 0, y: 0, width, height });
        pdf.removePage(p);
      }
      await replaceDocBytes(doc, await pdf.save());
    }
    S.meta.markups = S.meta.markups.filter((m) => !(m.kind === "redact" && m.docId === doc.id));
    await persistSnips();
    drawOverlay(); renderAnnotList(); if (S.mode === "markup") showMarkupPanel();
    status(`Caviardage appliqué (${zones.length} zone(s)) sur « ${doc.name} » : le texte masqué a été supprimé.`, "ok");
  }

  /* Commentaires sur la pièce (bulles) */
  function startNote(pt) {
    const doc = curDoc();
    S.noteDraft = { docId: doc.id, page: S.page, x: pt.x, y: pt.y };
    openPanel("note", `<div><strong>Commentaire sur la pièce</strong> — page ${S.page}</div>
      <textarea id="nText" rows="3" class="ta" placeholder="Ex. : pièce manquante pour ce montant, à demander au client"></textarea>
      <div class="row"><button class="primary" id="nAdd">Ajouter</button><button id="nNo">Annuler</button></div>`);
    setTimeout(() => $("#nText").focus(), 0);
    $("#nNo").onclick = () => { S.noteDraft = null; closePanel("note"); };
    $("#nAdd").onclick = () => {
      const t = $("#nText").value.trim(); if (!t || !S.noteDraft) return;
      S.meta.notes.push({ id: uid(), ...S.noteDraft, text: t, created: Date.now(), resolved: false });
      S.noteDraft = null; closePanel("note"); persistSoon(); drawOverlay(); renderAnnotList();
      status("Commentaire ajouté sur la pièce.", "ok");
    };
  }
  function showNote(n) {
    S.selNote = n.id; drawOverlay();
    openPanel("note", `<div><strong>Commentaire</strong> <span class="hint">· ${new Date(n.created).toLocaleString("fr-FR")}${n.resolved ? " · résolu" : ""}</span></div>
      <div class="notetext">${esc(n.text)}</div>
      <div class="row"><button id="nRes">${n.resolved ? "Rouvrir" : "Marquer comme résolu"}</button><button class="danger" id="nDel">Supprimer</button><button id="nClose">Fermer</button></div>`);
    $("#nRes").onclick = () => { n.resolved = !n.resolved; persistSoon(); showNote(n); renderAnnotList(); };
    $("#nDel").onclick = () => { S.meta.notes = S.meta.notes.filter((x) => x.id !== n.id); S.selNote = null; persistSoon(); closePanel("note"); drawOverlay(); renderAnnotList(); };
    $("#nClose").onclick = () => { S.selNote = null; closePanel("note"); drawOverlay(); };
  }
  function renderAnnotList() {
    const ul = $("#annotList"); if (!ul) return;
    const items = [
      ...S.meta.notes.map((n) => ({ kind: "note", o: n, label: (n.resolved ? "✔ " : "💬 ") + n.text })),
      ...S.meta.markups.map((m) => ({ kind: "markup", o: m, label: ({ highlight: "▮ Surlignage", text: "A " + (m.text || ""), arrow: "➜ Flèche", check: "✓ Coche", cross: "✗ Croix", redact: "■ Zone à caviarder" })[m.kind] || m.kind })),
    ];
    $("#annotCount").textContent = items.length ? `${items.length} annotation(s) sur les pièces.` : "Aucune annotation.";
    ul.innerHTML = "";
    for (const it of items.slice(0, 300)) {
      const li = document.createElement("li");
      li.innerHTML = `<div class="grow click"><div class="ttl">${esc(it.label)}</div><div class="sub">${esc((S.docs.get(it.o.docId) || {}).name || "document absent")} · p. ${it.o.page}</div></div><button class="danger" title="Supprimer">×</button>`;
      li.querySelector(".grow").onclick = () => {
        const r = it.kind === "note" ? { x0: it.o.x - 6, y0: it.o.y - 6, x1: it.o.x + 6, y1: it.o.y + 6 } : normRect(it.o);
        goToLocation(it.o.docId, it.o.page, r).then(() => { if (it.kind === "note") showNote(it.o); });
      };
      li.querySelector("button").onclick = () => {
        if (it.kind === "note") { S.meta.notes = S.meta.notes.filter((x) => x.id !== it.o.id); persistSoon(); drawOverlay(); renderAnnotList(); }
        else deleteMarkup(it.o.id);
      };
      ul.appendChild(li);
    }
  }

  /* ---------------- Clic, glisser, clic droit ----------------------------- */
  function extraClick(pt) {
    const doc = curDoc();
    const n = S.meta.notes.find((x) => x.docId === doc.id && x.page === S.page && Math.hypot(x.x - pt.x, x.y - pt.y) * S.scale < 13);
    if (n) { showNote(n); return true; }
    if (S.mode === "note") { startNote(pt); return true; }
    if (S.mode === "markup" && ["text", "check", "cross"].includes(S.markTool)) { addMarkupAt(pt); return true; }
    const m = markupAt(pt);
    if (m && (S.mode === "nav" || S.mode === "markup")) {
      S.selMarkup = m.id; drawOverlay(); if (S.mode === "markup") showMarkupPanel();
      status("Annotation sélectionnée — touche Suppr pour l'effacer.");
      return true;
    }
    if (S.selMarkup) { S.selMarkup = null; drawOverlay(); if (S.mode === "markup") showMarkupPanel(); }
    return false;
  }
  async function extraRect(r, opts) {
    if (!curDoc()) return false;
    if (S.mode === "markup") {
      if (S.markTool === "arrow") addMarkup({ kind: "arrow", x0: opts.from.x, y0: opts.from.y, x1: opts.to.x, y1: opts.to.y });
      else if (S.markTool === "text") addMarkupAt({ x: r.x0, y: r.y0 }, r);
      else addMarkup({ kind: S.markTool, ...r });
      return true;
    }
    if (S.mode === "note") { startNote({ x: r.x0, y: r.y0 }); return true; }
    if (S.mode === "formfield") { await formAddField(r); return true; }
    if (S.mode === "formkey") { await formSetKey(r); return true; }
    return false;
  }
  function modeExtras(m) {
    if (m === "markup") showMarkupPanel();
    else if (S.panelKind === "markup") closePanel("markup");
  }

  /* Plusieurs preuves pour une cellule (Ctrl + glisser) */
  async function attachEvidence(r) {
    const doc = curDoc(), p = S.page;
    const info = await xl.selectionInfo();
    let hit = null;
    for (const [id, ref] of Object.entries(info.links)) {
      if (!ref || ref.sheet !== info.sheet) continue;
      if (info.row >= ref.r1 && info.row <= ref.r2 && info.col >= ref.c1 && info.col <= ref.c2) {
        const sn = S.snips.get(id); if (sn && (!hit || sn.created > hit.created)) hit = sn;
      }
    }
    if (!hit) return status("Ctrl + glisser ajoute une preuve au snip de la cellule active, mais cette cellule n'a pas encore de snip.", "warn");
    let rect = r;
    try { const pg = await pageText(doc, p, { allowOcr: true }); const ws = wordsInRect(pg.words, r); if (ws.length) rect = bboxOf(ws, 2); } catch (_) { /* zone dessinée */ }
    hit.rects.push({ docId: doc.id, page: p, ...rect });
    await persistSnips(); S.selSnip = hit.id; drawOverlay();
    status(`Preuve ajoutée au snip ${hit.sheet}!${hit.address} (${hit.rects.length} zones liées).`, "ok");
  }
  /* Validation / Exception sur une cellule déjà remplie : on garde la valeur */
  async function linkExistingCell(type, ac, rects, cur) {
    const address = a1(ac.row, ac.col);
    await xl.format([{ sheet: ac.sheet, address, type }]);
    const id = uid();
    const sn = { id, group: id, auto: true, type, sheet: ac.sheet, address, docId: rects[0].docId, docName: (S.docs.get(rects[0].docId) || {}).name || "",
      rects, value: cur, cellVals: [cur], created: Date.now() };
    S.snips.set(id, sn);
    await xl.addLinks([sn]); await persistSnips();
    if (S.opts.advance) await xl.select(ac.sheet, ac.row + 1, ac.col);
    S.selSnip = id; drawOverlay(); refreshSnipList();
    confirmSnip(fmtVal(cur), address, type);
  }
  function snipAt(pt) {
    const doc = curDoc(); let hit = null;
    for (const sn of S.snips.values()) for (const r of sn.rects)
      if ((r.docId || sn.docId) === doc.id && r.page === S.page && inR(pt, r) && (!hit || sn.created > hit.created)) hit = sn;
    return hit;
  }
  async function convertSnip(sn, type) {
    sn.type = type; sn.auto = true;
    await xl.format([{ sheet: sn.sheet, address: sn.address, type }]);
    await persistSnips(); drawOverlay(); refreshSnipList();
    status(`Snip converti en ${STYLE[type].name} (la valeur de la cellule est conservée).`, "ok");
  }
  function showMenu(x, y, items) {
    const m = $("#ctxMenu"); m.innerHTML = "";
    for (const [label, fn] of items) {
      const b = document.createElement("button"); b.textContent = label;
      b.onclick = () => { hideMenu(); runAction(fn); };
      m.appendChild(b);
    }
    m.hidden = false;
    m.style.left = Math.max(4, Math.min(x, innerWidth - 230)) + "px";
    m.style.top = Math.max(4, Math.min(y, innerHeight - items.length * 30 - 12)) + "px";
  }
  function hideMenu() { const m = $("#ctxMenu"); if (m) m.hidden = true; }
  function bindContextMenu() {
    $("#pages").addEventListener("contextmenu", (e) => {
      const ov = e.target.closest(".ov");
      if (!curDoc() || !ov) return;
      S.page = +ov.dataset.p;
      const pt = evPt(e), sn = snipAt(pt), mk = markupAt(pt);
      if (!sn && !mk) return;
      e.preventDefault();
      const items = [];
      if (sn) {
        items.push([`Aller à la cellule ${sn.sheet}!${sn.address}`, () => xl.selectAddr(sn.sheet, sn.address)]);
        if (sn.type !== "table") for (const t of ["text", "valid", "except"]) if (t !== sn.type) items.push([`Convertir en ${STYLE[t].name}`, () => convertSnip(sn, t)]);
        items.push(["Supprimer le snip", () => deleteSnips([sn.id]).then(() => status("Snip supprimé (la valeur reste dans la cellule).", "ok"))]);
      }
      if (mk) items.push(["Supprimer l'annotation", () => deleteMarkup(mk.id)]);
      showMenu(e.clientX, e.clientY, items);
    });
    document.addEventListener("click", (e) => { if (!e.target.closest("#ctxMenu")) hideMenu(); });
  }

  /* ---------------- Recherche intelligente --------------------------------- */
  const isDateLike = (q) => /^\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*$/.test(q) || /^\s*\d{4}-\d{2}-\d{2}\s*$/.test(q);
  function dateVariants(q) {
    let d, m, y, mm;
    if ((mm = String(q).trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/))) { d = +mm[1]; m = +mm[2]; y = +mm[3]; }
    else if ((mm = String(q).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/))) { y = +mm[1]; m = +mm[2]; d = +mm[3]; }
    else return [];
    if (y < 100) y += 2000;
    const p2 = (n) => String(n).padStart(2, "0"), out = new Set();
    const add = (D, M) => {
      if (M < 1 || M > 12 || D < 1 || D > 31) return;
      for (const sep of ["/", ".", "-"]) { out.add(`${p2(D)}${sep}${p2(M)}${sep}${y}`); out.add(`${D}${sep}${M}${sep}${y}`); out.add(`${p2(D)}${sep}${p2(M)}${sep}${String(y).slice(2)}`); }
      out.add(`${y}-${p2(M)}-${p2(D)}`);
    };
    add(d, m); add(m, d);
    return [...out];
  }
  function searchTerms(q) { return [...new Set([norm(String(q).trim()), ...dateVariants(q).map(norm)])].filter(Boolean); }
  function findInLine(L, terms, qAmt) {
    const ln = norm(L.text), found = [];
    for (const t of terms) {
      let k = ln.indexOf(t);
      while (k >= 0) {
        const bad = (/^\d/.test(t) && (/\d/.test(ln[k - 1] || "") || /\d[.,]$/.test(ln.slice(Math.max(0, k - 2), k)))) ||
          (/\d$/.test(t) && (/\d/.test(ln[k + t.length] || "") || /^[.,]\d/.test(ln.slice(k + t.length, k + t.length + 2))));
        if (!bad && !found.some(([s, e]) => s < k + t.length && e > k)) found.push([k, k + t.length]);
        k = ln.indexOf(t, k + 1);
      }
    }
    if (qAmt != null) for (const a of findAmounts(L.text)) if (Math.abs(Math.abs(a.value) - Math.abs(qAmt)) < 0.005 && !found.some(([s, e]) => s < a.end && e > a.start)) found.push([a.start, a.end]);
    return found;
  }
  async function snipSearchResult(r) {
    await goToLocation(r.doc.id, r.p, r.rect);
    const tv = typed(r.L.text.slice(r.s, r.e));
    await placeSnip("text", tv.v, [{ docId: r.doc.id, page: r.p, ...r.rect }], tv.f);
  }
  /* --- Clic sur une cellule Excel : surligner sa valeur dans la pièce ouverte --- */
  let autoTimer = null, autoSeq = 0;
  function cellQuery(text) {
    const raw = String(text ?? "").trim();
    const clean = raw.replace(/\s?(€|EUR|\$|USD|£|CHF)$/i, "").replace(/^(€|\$|£)\s?/, "").trim();
    const date = isDateLike(clean);
    const amt = date ? null : parseAmount(clean);
    return { raw, clean, amt, terms: amt != null ? [] : searchTerms(clean) };
  }
  function hitsInDoc(doc, q, pages) {
    const out = [];
    for (const { p, pg } of pages) for (const L of pg.lines) {
      for (const [s, e] of findInLine(L, q.terms, null)) out.push({ docId: doc.id, page: p, rect: spanBBox(L, s, e) });
      if (q.amt != null) for (const a of findAmounts(L.text)) {
        const ws = L.words.filter((w) => w.end > a.start && w.start < a.end);
        if (ws.length === 1 && /[A-Za-zÀ-ÿ]/.test(ws[0].str)) continue;
        if (Math.abs(Math.abs(a.value) - Math.abs(q.amt)) < 0.005) out.push({ docId: doc.id, page: p, rect: spanBBox(L, a.start, a.end) });
      }
    }
    return out.sort((a, b) => a.page - b.page || a.rect.y0 - b.rect.y0 || a.rect.x0 - b.rect.x0);
  }
  async function docPages(doc) {
    await openDocData(doc);
    const out = [];
    for (let p = 1; p <= doc.numPages; p++) out.push({ p, pg: await pageText(doc, p, { allowOcr: false }) });
    return out;
  }
  function autoSearchCell(text) {
    clearTimeout(autoTimer);
    const seq = ++autoSeq;
    autoTimer = setTimeout(() => cellHighlight(text, seq).catch((e) => console.warn(e)), 120);
  }
  async function cellHighlight(text, seq = ++autoSeq) {
    const q = cellQuery(text);
    const doc = curDoc();
    if (!doc || q.clean.length < (q.amt != null ? 1 : 2) || (q.amt != null && q.amt === 0)) { clearHits(); return; }
    const hits = hitsInDoc(doc, q, await docPages(doc));
    if (seq !== autoSeq) return;                       // une autre cellule a été cliquée entre-temps
    S.hits = hits; S.hitIdx = 0; S.hitQuery = q.raw;
    let others = [];
    if (!hits.length) {
      for (const d of S.docs.values()) {
        if (d === doc || (!d.bytes && !d.pdf && !d.img && !d.stored)) continue;
        try { const h = hitsInDoc(d, q, await docPages(d)); if (h.length) others.push({ doc: d, n: h.length, first: h[0] }); } catch (_) { /* pièce illisible */ }
        if (seq !== autoSeq) return;
      }
    }
    showHitBar(others);
    drawOverlay();
    if (hits.length && !hitVisible(hits[0])) gotoHit(0);
  }
  function hitVisible(h) {
    const v = $("#viewer"), el = pageEl(h.page); if (!el || h.docId !== S.cur) return false;
    const top = el.offsetTop + h.rect.y0 * S.scale, bot = el.offsetTop + h.rect.y1 * S.scale;
    return top >= v.scrollTop && bot <= v.scrollTop + v.clientHeight - 30;
  }
  function gotoHit(i) {
    if (!S.hits.length) return;
    S.hitIdx = (i + S.hits.length) % S.hits.length;
    const h = S.hits[S.hitIdx];
    scrollToPage(h.page, h.rect);
    drawOverlay(); showHitBar();
  }
  function clearHits() {
    autoSeq++;
    if (S.hits.length || S.hitQuery) { S.hits = []; S.hitQuery = ""; drawOverlay(); }
    const b = $("#hitBar"); if (b) b.hidden = true;
  }
  function showHitBar(others = null) {
    const b = $("#hitBar"); if (!b) return;
    if (!S.hitQuery) { b.hidden = true; return; }
    const n = S.hits.length;
    b.hidden = false;
    b.classList.toggle("none", !n);
    let html = `<span class="hq">« ${esc(S.hitQuery.length > 24 ? S.hitQuery.slice(0, 22) + "…" : S.hitQuery)} »</span>`;
    if (n) html += `<span class="hn">${n > 1 ? `${S.hitIdx + 1} / ${n}` : "1 fois"}</span>` +
      (n > 1 ? `<button class="ico sm" data-h="-1" title="Précédent (Maj+F3)"><svg viewBox="0 0 20 20"><path d="M5 12l5-5 5 5"/></svg></button><button class="ico sm" data-h="1" title="Suivant (F3)"><svg viewBox="0 0 20 20"><path d="M5 8l5 5 5-5"/></svg></button>` : "");
    else {
      html += `<span class="hn">absent de cette pièce</span>`;
      if (others && others.length) html += others.slice(0, 2).map((o, i) => `<button class="link" data-o="${i}" title="Ouvrir cette pièce">${esc(o.doc.name)} (${o.n})</button>`).join("");
    }
    html += `<button class="ico sm" data-x title="Effacer le surlignage (Échap)"><svg viewBox="0 0 20 20"><path d="M5 5l10 10M15 5L5 15"/></svg></button>`;
    b.innerHTML = html;
    $$("[data-h]", b).forEach((x) => (x.onclick = () => gotoHit(S.hitIdx + +x.dataset.h)));
    $$("[data-o]", b).forEach((x) => (x.onclick = async () => { const o = others[+x.dataset.o]; await openDoc(o.doc.id, o.first.page); cellHighlight(S.hitQuery); }));
    $("[data-x]", b).onclick = clearHits;
  }

  /* ---------------- Liste des snips : filtres, liens multiples, tri -------- */
  function filteredSnips() {
    const all = [...S.snips.values()];
    if (S.snipFilter === "doc") return all.filter((s) => s.rects.some((r) => (r.docId || s.docId) === S.cur));
    if (S.snipFilter === "sel" && S.selBox) {
      const b = S.selBox;
      return all.filter((s) => { const r = refOf(s); return r && r.sheet === b.sheet && r.r1 <= b.r2 && r.r2 >= b.r1 && r.c1 <= b.c2 && r.c2 >= b.c1; });
    }
    return all;
  }
  async function linkSnipToSelection() {
    const sn = S.snips.get(S.selSnip);
    if (!sn) return status("Sélectionnez d'abord un snip : cliquez sa zone sur la pièce ou dans la liste.", "warn");
    const sel = await xl.selected();
    const made = [], formats = [], group = uid();
    sel.values.forEach((row, i) => row.forEach((v, j) => {
      const address = a1(sel.row + i, sel.col + j);
      if (sel.sheet === sn.sheet && address === sn.address) return;
      made.push({ ...JSON.parse(JSON.stringify(sn)), id: group + "_" + made.length, group, auto: true, sheet: sel.sheet, address, value: v, cellVals: [v], created: Date.now() + made.length / 1000 });
      formats.push({ sheet: sel.sheet, address, type: sn.type });
    }));
    if (!made.length) return status("Sélectionnez dans Excel les cellules à relier à ce snip.", "warn");
    await xl.format(formats);
    for (const m of made) S.snips.set(m.id, m);
    await xl.addLinks(made); await persistSnips(); refreshSnipList();
    status(`Snip relié à ${made.length} autre(s) cellule(s).`, "ok");
  }
  function looseEq(a, b) {
    if (typeof a === "number" || typeof b === "number") {
      const x = typeof a === "number" ? a : parseAmount(a), y = typeof b === "number" ? b : parseAmount(b);
      return x != null && y != null && Math.abs(x - y) < 1e-6;
    }
    return norm(String(a ?? "")).trim() === norm(String(b ?? "")).trim();
  }
  async function realign(sheet, address) {
    const box = parseRef(`'${sheet.replace(/'/g, "''")}'!${String(address).split("!").pop()}`);
    if (!box) return 0;
    const vals = await xl.readRange(sheet, box.address);
    const cand = [...S.snips.values()].filter((sn) => { const r = refOf(sn); return r && r.sheet === sheet && r.r1 >= box.r1 && r.r2 <= box.r2 && r.c1 >= box.c1 && r.c2 <= box.c2; });
    const used = new Set(), changes = [];
    for (const sn of cand) {
      const r = refOf(sn), want = sn.cellVals || [sn.value];
      const rowOk = (row) => want.every((w, k) => looseEq((vals[row - box.r1] || [])[r.c1 - box.c1 + k], w));
      if (rowOk(r.r1) && !used.has(r.r1)) { used.add(r.r1); continue; }
      let target = -1;
      for (let row = box.r1; row <= box.r2; row++) if (!used.has(row) && rowOk(row)) { target = row; break; }
      if (target < 0) continue;
      used.add(target);
      changes.push({ sn, address: rangeA1(target, r.c1, r.r2 - r.r1 + 1, r.c2 - r.c1 + 1) });
    }
    if (changes.length) {
      await xl.relink(changes);
      for (const c of changes) c.sn.address = c.address;
      await persistSnips(); drawOverlay(); refreshSnipList();
    }
    return changes.length;
  }
  async function onRowSorted(ev) {
    if (!S.opts.sortFollow) return;
    try {
      const sheet = await xl.sheetNameById(ev.worksheetId);
      const n = await realign(sheet, ev.address);
      if (n) status(`Tri détecté : ${n} lien(s) déplacé(s) avec leurs valeurs.`, "ok");
    } catch (e) { console.warn("Tri :", e); }
  }
  async function relinkActiveSheet() {
    const ac = await xl.activeCell();
    const addr = await xl.usedAddress(ac.sheet);
    if (!addr) return status("Feuille vide.", "warn");
    const n = await realign(ac.sheet, addr);
    status(n ? `${n} lien(s) recalé(s) sur leurs valeurs après le tri.` : "Tous les liens de cette feuille sont déjà à leur place.", "ok");
  }

  /* ---------------- Écriture d'une nouvelle feuille de résultats ---------- */
  async function writeSheet(base, cells) {
    const { name, row0 } = await xl.newSheet(base);
    const shifted = cells.map((c) => ({ ...c, row: c.row + row0 }));
    await xl.writeCells(name, shifted);
    return { name, row0 };
  }
  async function linkCells(sheetName, row0, entries, group) {
    // entries: [{row, col, type, value, rects:[{docId,page,...}], docName}]
    const made = entries.filter((e) => e.rects && e.rects.length).map((e, i) => ({
      id: group + "_" + i, group, auto: true, type: e.type, sheet: sheetName, address: a1(e.row + row0, e.col),
      docId: e.rects[0].docId, docName: e.docName || (S.docs.get(e.rects[0].docId) || {}).name || "", rects: e.rects,
      value: e.value, cellVals: [e.value], created: Date.now() + i / 1000,
    }));
    for (const m of made) S.snips.set(m.id, m);
    if (made.length) { await xl.addLinks(made); await persistSnips(); }
    return made;
  }

  /* ---------------- Document Matching ligne par ligne -------------------- */
  function similarity(a, b) {
    if (a === b) return 1;
    const m = a.length, n = b.length; if (!m || !n) return 0;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return 1 - prev[n] / Math.max(m, n);
  }
  function wordStarts(s) { const out = []; const re = /\S+/g; let mm; while ((mm = re.exec(s))) out.push([mm.index, mm.index + mm[0].length]); return out; }
  function buildIndex(doc, pages) {
    const lines = [];
    for (const { doc: d, p, pg } of pages) if (d.id === doc.id) for (const L of pg.lines) lines.push({ p, L, n: norm(L.text), amts: findAmounts(L.text) });
    return { doc, lines, nameN: norm(doc.name) };
  }
  function matchField(text, raw, c, idx) {
    const s = String(text ?? "").trim(); if (!s) return null;
    const date = isDateLike(s);
    const num = !date ? (typeof raw === "number" ? raw : parseAmount(s)) : null;
    let best = null;
    const take = (cand) => { if (!best || cand.score > best.score) best = cand; };
    if (num != null) {
      const tolS = String(c.tol || "").trim();
      const tol = /%$/.test(tolS) ? Math.abs(num) * (parseAmount(tolS.replace("%", "")) || 0) / 100 : (parseAmount(tolS) ?? 0.005);
      for (const it of idx.lines) for (const a of it.amts) {
        const d = Math.abs(Math.abs(a.value) - Math.abs(num));
        if (d <= tol + 1e-9) take({ score: 1 - (tol ? 0.1 * d / tol : 0), page: it.p, rect: spanBBox(it.L, a.start, a.end), text: it.L.text.slice(a.start, a.end) });
      }
      if (best) return best;
    }
    for (const it of idx.lines) for (const t of searchTerms(s)) {
      const k = it.n.indexOf(t);
      if (k >= 0) take({ score: 1, page: it.p, rect: spanBBox(it.L, k, k + t.length), text: it.L.text.slice(k, k + t.length) });
    }
    if (best) return best;
    if (s.length >= 3 && idx.nameN.includes(norm(s))) return { score: 0.9, page: 1, rect: { x0: 4, y0: 4, x1: 70, y1: 16 }, text: s, byName: true };
    if (c.fuzzy && !num) {
      const tn = norm(s), n = tn.split(/\s+/).length;
      for (const it of idx.lines) {
        const ws = wordStarts(it.n);
        for (let k = 0; k + n <= ws.length; k++) {
          const st = ws[k][0], en = ws[k + n - 1][1], win = it.n.slice(st, en);
          const sim = similarity(tn, win);
          if (sim >= 0.8) take({ score: 0.8 * sim, page: it.p, rect: spanBBox(it.L, st, en), text: it.L.text.slice(st, en) });
        }
      }
      if (best) return best;
    }
    if (c.partial && !num) {
      const toks = norm(s).split(/\s+/).filter((t) => t.length >= 3);
      for (const it of idx.lines) {
        let hit = 0, first = -1, last = -1;
        for (const t of toks) { const k = it.n.indexOf(t); if (k >= 0) { hit++; if (first < 0 || k < first) first = k; last = Math.max(last, k + t.length); } }
        if (hit) take({ score: 0.6 * hit / toks.length, page: it.p, rect: spanBBox(it.L, first, last), text: it.L.text.slice(first, last) });
      }
    }
    return best;
  }
  function loadTpl(key) { try { return JSON.parse(localStorage.getItem(key) || "{}"); } catch (_) { return {}; } }
  function saveTpl(key, obj) { try { localStorage.setItem(key, JSON.stringify(obj)); } catch (_) { status("Impossible d'enregistrer le modèle sur ce PC.", "warn"); } }
  function renderTplSelect(selId, key) { const el = $(selId); if (!el) return; const t = loadTpl(key); el.innerHTML = `<option value="">— modèles enregistrés —</option>` + Object.keys(t).map((k) => `<option>${esc(k)}</option>`).join(""); }

  async function prepareRowMatch() {
    const sel = await xl.selected();
    const header = $("#rmHeader").checked;
    const n = (sel.values[0] || []).length;
    const heads = header ? sel.texts[0] : Array.from({ length: n }, (_, j) => "Colonne " + colLetters(sel.col + j));
    const first = sel.values[header ? 1 : 0] || [];
    const old = S.rm.cols;
    S.rm.cols = heads.map((h, j) => old[j] && old[j].name === String(h) ? old[j] : ({ name: String(h || "Colonne " + colLetters(sel.col + j)), required: false, fuzzy: false, partial: false, tol: typeof first[j] === "number" ? "0,01" : "" }));
    renderRowMatchCols();
    $("#rmInfo").textContent = `${sel.values.length - (header ? 1 : 0)} ligne(s) × ${n} colonne(s) sélectionnée(s) en ${sel.sheet}!${rangeA1(sel.row, sel.col, sel.values.length, n)}.`;
  }
  function renderRowMatchCols() {
    const box = $("#rmCols");
    if (!S.rm.cols.length) { box.innerHTML = `<div class="hint">Sélectionnez la plage dans Excel puis « Préparer les colonnes ».</div>`; return; }
    box.innerHTML = `<table class="opt"><tr><th>Colonne</th><th title="La ligne n'est validée que si ce champ est trouvé">Oblig.</th><th title="Accepte les petites fautes / abréviations">Approx.</th><th title="Accepte une partie du texte">Partiel</th><th title="Écart accepté sur les montants (ex. 0,05 ou 1%)">Tolérance</th></tr>
      ${S.rm.cols.map((c, j) => `<tr><td>${esc(c.name)}</td><td><input type="checkbox" data-j="${j}" data-k="required" ${c.required ? "checked" : ""}></td><td><input type="checkbox" data-j="${j}" data-k="fuzzy" ${c.fuzzy ? "checked" : ""}></td><td><input type="checkbox" data-j="${j}" data-k="partial" ${c.partial ? "checked" : ""}></td><td><input type="text" size="5" data-j="${j}" data-k="tol" value="${esc(c.tol || "")}"></td></tr>`).join("")}</table>`;
    $$("input", box).forEach((inp) => (inp.onchange = () => { const c = S.rm.cols[+inp.dataset.j]; c[inp.dataset.k] = inp.type === "checkbox" ? inp.checked : inp.value; }));
  }
  async function runRowMatch() {
    const ul = $("#matchResults"), info = $("#matchInfo");
    ul.innerHTML = "";
    if (!S.docs.size) { info.textContent = "Importez d'abord les pièces justificatives."; return; }
    const sel = await xl.selected();
    const header = $("#rmHeader").checked, once = $("#rmOnce").checked;
    const n = (sel.values[0] || []).length;
    if (S.rm.cols.length !== n) await prepareRowMatch();
    const cols = S.rm.cols;
    const off = header ? 1 : 0;
    const scope = $("#matchDoc").value;
    const docs = scope === "*" ? [...S.docs.values()] : [S.docs.get(scope)].filter(Boolean);
    const pages = await allPages(docs, true, (d, p) => status(`Lecture de « ${d.name} », p. ${p}…`, "busy"));
    const idxs = docs.map((d) => buildIndex(d, pages));
    const usedDocs = new Set(), results = [];
    for (let i = off; i < sel.values.length; i++) {
      const vals = sel.values[i], texts = sel.texts[i];
      if (vals.every((v) => v === "" || v == null)) { results.push(null); continue; }
      status(`Rapprochement ligne ${i - off + 1}/${sel.values.length - off}…`, "busy");
      let best = null;
      for (const idx of idxs) {
        if (once && usedDocs.has(idx.doc.id)) continue;
        const fields = cols.map((c, j) => matchField(texts[j], vals[j], c, idx));
        if (cols.some((c, j) => c.required && String(texts[j] ?? "").trim() && !fields[j])) continue;
        const found = fields.filter(Boolean); if (!found.length) continue;
        const pc = {}; found.forEach((f) => (pc[f.page] = (pc[f.page] || 0) + 1));
        const major = +Object.entries(pc).sort((a, b) => b[1] - a[1])[0][0];
        const score = found.reduce((s, f) => s + f.score, 0) + 0.15 * found.filter((f) => f.page === major).length;
        if (!best || score > best.score) best = { idx, fields, score };
      }
      if (best && once) usedDocs.add(best.idx.doc.id);
      results.push(best);
    }
    // écriture des colonnes de sortie à droite de la sélection
    const outCol = sel.col + n;
    const cells = [], entries = [];
    if (header) { cells.push({ row: sel.row, col: outCol, value: "Document", type: "head" }); cols.forEach((c, j) => cells.push({ row: sel.row, col: outCol + 1 + j, value: c.name + " (pièce)", type: "head" })); }
    let ok = 0, partial = 0;
    results.forEach((res, k) => {
      const row = sel.row + off + k;
      if (!res) { if (sel.values[off + k] && sel.values[off + k].some((v) => v !== "" && v != null)) { cells.push({ row, col: outCol, value: "✗ aucune pièce", type: "except" }); } return; }
      const all = res.fields.every((f, j) => f || !String(sel.texts[off + k][j] ?? "").trim());
      all ? ok++ : partial++;
      const first = res.fields.find(Boolean);
      cells.push({ row, col: outCol, value: res.idx.doc.name, type: all ? "valid" : "except" });
      entries.push({ row, col: outCol, type: all ? "valid" : "except", value: res.idx.doc.name, rects: [{ docId: res.idx.doc.id, page: first.page, ...first.rect }] });
      res.fields.forEach((f, j) => {
        if (f) { const tv = typed(f.text); cells.push({ row, col: outCol + 1 + j, value: tv.v, fmt: tv.f, type: "text" }); entries.push({ row, col: outCol + 1 + j, type: "text", value: tv.v, rects: [{ docId: res.idx.doc.id, page: f.page, ...f.rect }] }); }
        else if (String(sel.texts[off + k][j] ?? "").trim()) cells.push({ row, col: outCol + 1 + j, value: "✗", type: "except" });
      });
    });
    await xl.writeCells(sel.sheet, cells);
    await linkCells(sel.sheet, 0, entries, uid());
    const total = results.filter((r, k) => r || (sel.values[off + k] || []).some((v) => v !== "" && v != null)).length;
    info.innerHTML = `<span class="ok">${ok} ligne(s) entièrement justifiée(s)</span> · ${partial} partielle(s) · <span class="ko">${total - ok - partial} sans pièce</span>. Résultats écrits à droite de la sélection (colonne ${colLetters(outCol)}).`;
    refreshSnipList(); drawOverlay();
    status(`Document Matching terminé : ${ok}/${total} ligne(s) justifiée(s).`, "ok");
  }

  /* ---------------- Sommes du document (Find all sums) ------------------ */
  async function findDocSums() {
    const doc = curDoc();
    if (!doc) return status("Ouvrez d'abord le document à analyser (état financier, relevé…).", "warn");
    await openDocData(doc);
    const results = [];
    for (let p = 1; p <= doc.numPages; p++) {
      status(`Analyse des sommes — page ${p}/${doc.numPages}…`, "busy");
      const pg = await pageText(doc, p, { allowOcr: true });
      const items = [];
      for (const L of pg.lines) for (const a of findAmounts(L.text)) if (a.hasDec) items.push({ v: a.value, rect: spanBBox(L, a.start, a.end) });
      const cols = [];
      for (const it of items) { let c = cols.find((x) => Math.abs(x.x1 - it.rect.x1) < 6); if (!c) { c = { x1: it.rect.x1, items: [] }; cols.push(c); } c.items.push(it); }
      for (const c of cols) {
        const blocks = [];
        for (const it of c.items.sort((a, b) => a.rect.y0 - b.rect.y0)) {
          let sum = 0, found = false;
          for (let j = blocks.length - 1; j >= 0 && blocks.length - j <= 40; j--) {
            sum = round2(sum + blocks[j].v);
            if (blocks.length - j >= 2 && it.v !== 0 && Math.abs(sum - it.v) < 0.005) {
              const comps = blocks.slice(j);
              results.push({ page: p, total: it, parts: comps.map((b) => b.item) });
              blocks.splice(j, comps.length, { v: it.v, item: it }); found = true; break;
            }
          }
          if (!found) blocks.push({ v: it.v, item: it });
        }
      }
    }
    S.sumMarks = results.flatMap((r) => [{ docId: doc.id, page: r.page, rect: r.total.rect, kind: "total" }, ...r.parts.map((x) => ({ docId: doc.id, page: r.page, rect: x.rect, kind: "part" }))]);
    drawOverlay();
    const ul = $("#sumsOut"); if (ul) ul.innerHTML = "";
    if (!results.length) return status(`Aucune somme détectée dans « ${doc.name} ».`, "warn");
    const cells = [{ row: 0, col: 0, value: "Total", type: "head" }, { row: 0, col: 1, value: "Composant", type: "head" }, { row: 0, col: 2, value: "Page", type: "head" }];
    const entries = []; let row = 1;
    for (const r of results) {
      cells.push({ row, col: 0, value: r.total.v, type: "text" }, { row, col: 2, value: r.page });
      entries.push({ row, col: 0, type: "text", value: r.total.v, rects: [{ docId: doc.id, page: r.page, ...r.total.rect }] });
      row++;
      for (const x of r.parts) { cells.push({ row, col: 1, value: x.v, type: "sum" }); entries.push({ row, col: 1, type: "sum", value: x.v, rects: [{ docId: doc.id, page: r.page, ...x.rect }] }); row++; }
    }
    const { name, row0 } = await writeSheet("Sommes trouvées", cells);
    await linkCells(name, row0, entries, uid());
    if (ul) for (const r of results.slice(0, 200)) {
      const li = document.createElement("li"); li.className = "click";
      li.innerHTML = `<span class="chip" style="background:#2F6FD6">Σ</span><div class="grow"><div class="ttl">${fmtNum(r.total.v)} = ${r.parts.map((x) => fmtNum(x.v)).join(" + ")}</div><div class="sub">page ${r.page}</div></div>`;
      li.onclick = () => goToLocation(doc.id, r.page, r.total.rect);
      ul.appendChild(li);
    }
    refreshSnipList();
    status(`${results.length} somme(s) trouvée(s) dans « ${doc.name} » : feuille « ${name} » créée (totaux en bleu, composants en orange).`, "ok");
  }

  /* ---------------- Extraction de formulaire (modèle) -------------------- */
  const cleanVal = (t) => String(t || "").replace(/^[\s:;,.\-–—]+/, "").trim();
  function autoKey(pg, rect) {
    const h = rect.y1 - rect.y0, yc = (rect.y0 + rect.y1) / 2;
    const left = pg.words.filter((w) => w.x1 <= rect.x0 + 1 && rect.x0 - w.x1 < 220 && Math.abs((w.y0 + w.y1) / 2 - yc) < Math.max(4, h * 0.6)).sort((a, b) => b.x1 - a.x1);
    if (left.length) {
      const chain = [left[0]];
      for (let i = 1; i < left.length && chain.length < 4; i++) { if (chain[chain.length - 1].x0 - left[i].x1 < 12) chain.push(left[i]); else break; }
      const ws = chain.reverse();
      return { text: joinWords(ws), rect: bboxOf(ws, 1) };
    }
    const above = pg.words.filter((w) => w.y1 <= rect.y0 + 1 && rect.y0 - w.y1 < 30 && w.x1 > rect.x0 - 10 && w.x0 < rect.x1 + 10);
    if (above.length) {
      const maxY = Math.max(...above.map((w) => w.y1));
      const ws = above.filter((w) => maxY - w.y1 < 3).sort((a, b) => a.x0 - b.x0).slice(0, 5);
      return { text: joinWords(ws), rect: bboxOf(ws, 1) };
    }
    return null;
  }
  async function formAddField(r) {
    const doc = curDoc(), p = S.page, F = S.form;
    if (F.fields.length && F.refDoc && F.refDoc !== doc.id) {
      setMode("nav");
      return status(`Les champs sont définis sur « ${(S.docs.get(F.refDoc) || {}).name} ». Rouvrez ce document ou effacez les champs.`, "warn");
    }
    const pg = await pageText(doc, p, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    const rect = ws.length ? bboxOf(ws, 2) : r;
    F.refDoc = doc.id;
    const key = autoKey(pg, rect);
    F.fields.push({ id: uid(), name: key ? cleanVal(key.text).replace(/\s*[:：]\s*$/, "").slice(0, 30) || "Champ " + (F.fields.length + 1) : "Champ " + (F.fields.length + 1), page: p, rect, key, sample: cleanVal(joinWords(ws)) });
    setMode("nav"); renderForm(); drawOverlay();
    status("Champ ajouté. Renommez-le si besoin, puis ajoutez les suivants ou lancez l'aperçu.", "ok");
  }
  async function formSetKey(r) {
    const F = S.form, f = F.fields.find((x) => x.id === F.pendingKeyFor);
    setMode("nav");
    if (!f) return;
    const pg = await pageText(curDoc(), S.page, { allowOcr: true });
    const ws = wordsInRect(pg.words, r);
    if (!ws.length) return status("Aucun texte dans la zone choisie comme repère.", "warn");
    f.key = { text: joinWords(ws), rect: bboxOf(ws, 1) };
    F.pendingKeyFor = null; renderForm(); drawOverlay();
    status(`Repère du champ « ${f.name} » : « ${f.key.text} ».`, "ok");
  }
  async function applyField(f, doc, onlyPage) {
    const scan = onlyPage ? [onlyPage] : Array.from({ length: doc.numPages }, (_, i) => i + 1);
    if (f.key && f.key.text) {
      const kn = norm(f.key.text);
      for (const p of scan) {
        const pg = await pageText(doc, p, { allowOcr: true });
        for (const L of pg.lines) {
          const k = norm(L.text).indexOf(kn); if (k < 0) continue;
          const kb = bboxOf(L.words.filter((w) => w.end > k && w.start < k + kn.length), 1);
          const dx = kb.x0 - f.key.rect.x0, dy = kb.y0 - f.key.rect.y0;
          const rect = { x0: f.rect.x0 + dx, y0: f.rect.y0 + dy, x1: f.rect.x1 + dx, y1: f.rect.y1 + dy };
          return { text: cleanVal(joinWords(wordsInRect(pg.words, rect))), page: p, rect };
        }
      }
    }
    const p = onlyPage || Math.min(f.page, doc.numPages);
    const pg = await pageText(doc, p, { allowOcr: true });
    return { text: cleanVal(joinWords(wordsInRect(pg.words, f.rect))), page: p, rect: f.rect, fixed: true };
  }
  function checkedDocs(containerId) {
    const ids = $$(`#${containerId} input[type=checkbox]:checked`).map((i) => i.value);
    return ids.map((id) => S.docs.get(id)).filter(Boolean);
  }
  function renderDocChecks(containerId) {
    const box = $("#" + containerId); if (!box) return;
    const prev = new Set($$(`#${containerId} input:checked`).map((i) => i.value));
    const fresh = !box.dataset.init;
    box.dataset.init = "1";
    box.innerHTML = [...S.docs.values()].map((d) => `<label class="chk"><input type="checkbox" value="${d.id}" ${fresh || prev.has(d.id) ? "checked" : ""}> ${esc(d.name)}</label>`).join("") || `<div class="hint">Aucun document importé.</div>`;
  }
  function renderAutoDocs() { renderDocChecks("formDocs"); renderDocChecks("invDocs"); }
  function renderForm() {
    const F = S.form, ul = $("#formFields"); if (!ul) return;
    ul.innerHTML = F.fields.length ? "" : `<li class="hint">Aucun champ. Ouvrez un document type puis « + Ajouter un champ ».</li>`;
    F.fields.forEach((f, i) => {
      const li = document.createElement("li");
      li.innerHTML = `<div class="grow"><input type="text" value="${esc(f.name)}" style="width:100%"><div class="sub">p. ${f.page} · repère : ${f.key ? "« " + esc(f.key.text) + " »" : "position fixe"}${f.sample ? " · ex. « " + esc(f.sample.slice(0, 40)) + " »" : ""}</div></div>
        <button data-key title="Choisir un autre libellé comme repère">Repère</button><button class="danger" data-rm>×</button>`;
      li.querySelector("input").onchange = (e) => { f.name = e.target.value.trim() || f.name; drawOverlay(); };
      li.querySelector("[data-key]").onclick = () => { F.pendingKeyFor = f.id; if (S.cur !== F.refDoc) openDoc(F.refDoc, f.page); switchTab("viewer"); setMode("formkey"); };
      li.querySelector("[data-rm]").onclick = () => { F.fields.splice(i, 1); if (!F.fields.length) F.refDoc = null; renderForm(); drawOverlay(); };
      ul.appendChild(li);
    });
  }
  async function formCompute() {
    const F = S.form;
    if (!F.fields.length) throw new Error("ajoutez au moins un champ");
    const docs = checkedDocs("formDocs");
    if (!docs.length) throw new Error("cochez au moins un document");
    const rows = [];
    for (const d of docs) {
      await openDocData(d);
      const pagesList = F.mode === "page" ? Array.from({ length: d.numPages }, (_, i) => i + 1) : [null];
      for (const p of pagesList) {
        status(`Extraction — ${d.name}${p ? " p. " + p : ""}…`, "busy");
        const cells = [];
        for (const f of F.fields) cells.push(await applyField(f, d, p));
        rows.push({ doc: d, page: p, cells });
      }
    }
    F.rows = rows;
    return rows;
  }
  async function formPreview() {
    const rows = await formCompute(), F = S.form;
    $("#formOut").innerHTML = `<table class="opt"><tr><th>Document</th>${F.mode === "page" ? "<th>Page</th>" : ""}${F.fields.map((f) => `<th>${esc(f.name)}</th>`).join("")}</tr>
      ${rows.slice(0, 50).map((r) => `<tr><td>${esc(r.doc.name)}</td>${F.mode === "page" ? `<td>${r.page}</td>` : ""}${r.cells.map((c) => `<td class="${c.text ? "" : "ko"}">${esc(c.text || "—")}</td>`).join("")}</tr>`).join("")}</table>
      ${rows.length > 50 ? `<div class="hint">… ${rows.length - 50} ligne(s) de plus</div>` : ""}`;
    status(`Aperçu : ${rows.length} ligne(s).`, "ok");
  }
  async function formExport() {
    const F = S.form, rows = await formCompute();
    const pc = F.mode === "page" ? 1 : 0;
    const cells = [{ row: 0, col: 0, value: "Document", type: "head" }];
    if (pc) cells.push({ row: 0, col: 1, value: "Page", type: "head" });
    F.fields.forEach((f, j) => cells.push({ row: 0, col: 1 + pc + j, value: f.name, type: "head" }));
    const entries = [];
    rows.forEach((r, i) => {
      cells.push({ row: i + 1, col: 0, value: r.doc.name });
      if (pc) cells.push({ row: i + 1, col: 1, value: r.page });
      r.cells.forEach((c, j) => {
        if (!c.text) return;
        const tv = typed(c.text);
        cells.push({ row: i + 1, col: 1 + pc + j, value: tv.v, fmt: tv.f, type: "text" });
        entries.push({ row: i + 1, col: 1 + pc + j, type: "text", value: tv.v, rects: [{ docId: r.doc.id, page: c.page, ...c.rect }] });
      });
    });
    const { name, row0 } = await writeSheet("Extraction", cells);
    await linkCells(name, row0, entries, uid());
    refreshSnipList();
    status(`Extraction exportée : ${rows.length} ligne(s) dans la feuille « ${name} », chaque valeur liée à sa pièce.`, "ok");
  }

  /* ---------------- Extraction de factures et reçus (règles, sans IA) ---- */
  const FR_MONTHS = "janv(?:ier)?|f[ée]vr(?:ier)?|mars|avr(?:il)?|mai|juin|juil(?:let)?|ao[uû]t|sept(?:embre)?|oct(?:obre)?|nov(?:embre)?|d[ée]c(?:embre)?";
  const DATE_RE = new RegExp(`(\\d{1,2}[\\/.\\-]\\d{1,2}[\\/.\\-]\\d{2,4}|\\d{1,2}\\s+(?:${FR_MONTHS})\\.?\\s+\\d{4}|\\d{4}-\\d{2}-\\d{2})`, "i");
  async function docLines(doc, maxPages) {
    const out = [];
    for (let p = 1; p <= Math.min(doc.numPages, maxPages); p++) { const pg = await pageText(doc, p, { allowOcr: true }); pg.lines.forEach((L) => out.push({ p, L, pg })); }
    return out;
  }
  const hitAt = (it, s, e) => ({ text: it.L.text.slice(s, e).trim(), page: it.p, rect: spanBBox(it.L, s, e) });
  function findAfterLabel(lines, labelRe, valueRe) {
    for (let i = 0; i < lines.length; i++) {
      const it = lines[i], m = labelRe.exec(it.L.text); if (!m) continue;
      const after = m.index + m[0].length, vm = valueRe.exec(it.L.text.slice(after));
      if (vm) return hitAt(it, after + vm.index, after + vm.index + vm[0].length);
      const nx = lines[i + 1];
      if (nx && nx.p === it.p) { const v2 = valueRe.exec(nx.L.text); if (v2) return hitAt(nx, v2.index, v2.index + v2[0].length); }
    }
    return null;
  }
  function amountAfterLabel(lines, labelRe, exclRe) {
    let res = null;
    for (const it of lines) {
      const t = it.L.text, m = labelRe.exec(t);
      if (!m || (exclRe && exclRe.test(t))) continue;
      const a = findAmounts(t).find((x) => x.start >= m.index + m[0].length - 1 && x.hasDec && !/^\s?%/.test(t.slice(x.end, x.end + 2)));
      if (a) res = { ...hitAt(it, a.start, a.end), value: a.value };
    }
    return res;
  }
  function regexAnywhere(lines, re, filter) {
    for (const it of lines) { const m = re.exec(it.L.text); if (m && (!filter || filter(m[0]))) return hitAt(it, m.index, m.index + m[0].length); }
    return null;
  }
  function segmentsOf(L) {
    const segs = []; let cur = [];
    for (const w of L.words) {
      if (cur.length && w.x0 - cur[cur.length - 1].x1 > Math.max(18, 2.5 * L.h)) { segs.push(cur); cur = []; }
      cur.push(w);
    }
    if (cur.length) segs.push(cur);
    return segs.map((ws) => ({ ws, text: joinWords(ws), h: Math.max(...ws.map((w) => w.y1 - w.y0)) }));
  }
  function topName(lines) {
    const cands = [];
    for (const it of lines) {
      if (it.p !== 1 || it.L.y1 > it.pg.h * 0.25) continue;
      for (const sg of segmentsOf(it.L)) if (/[a-zà-ÿ]{3}/i.test(sg.text) && !/facture|invoice|devis|avoir|date|page|ticket|re[çc]u|client|siret|tva|t[ée]l|e-?mail|rue|avenue|bd\b/i.test(sg.text)) cands.push(sg);
    }
    if (!cands.length) return null;
    const b = cands.reduce((a, c) => (c.h > a.h + 0.5 ? c : a));
    return { text: b.text.trim(), page: 1, rect: bboxOf(b.ws, 1.5) };
  }
  async function extractInvoice(doc) {
    await openDocData(doc);
    const lines = await docLines(doc, 3), f = {};
    f.supplier = topName(lines);
    const NUM = /[A-Z0-9][A-Z0-9\-\/_.]{1,}/i;
    f.number = findAfterLabel(lines, /(facture|invoice|avoir)\s*(n[°o]|num[ée]ro|no\.?|#|number)\s*[:.\-]?\s*/i, NUM)
      || findAfterLabel(lines, /\bn[°o]\s*(de\s*)?(la\s*)?facture\s*[:.\-]?\s*/i, NUM)
      || findAfterLabel(lines, /\br[ée]f[ée]rence\s*(facture)?\s*[:.\-]\s*/i, NUM);
    f.date = findAfterLabel(lines, /date(?!\s*(d['’]\s*)?[ée]ch[ée]ance|\s*limite)\s*(de\s*(la\s*)?facture|d['’]?\s*[ée]mission|de\s*facturation)?\s*[:.\-]?\s*/i, DATE_RE)
      || regexAnywhere(lines.filter((x) => x.p === 1), DATE_RE);
    f.due = findAfterLabel(lines, /([ée]ch[ée]ance|date\s*limite(\s*de\s*(paiement|r[èe]glement))?|payable\s*(avant\s*)?le|due\s*date)\s*[:.\-]?\s*/i, DATE_RE);
    f.ht = amountAfterLabel(lines, /(total|montant|net)\s*h\.?\s*t\.?(?![a-z])|total\s*hors\s*taxes?|sous[- ]total/i);
    f.tva = amountAfterLabel(lines, /(?:total\s*|montant\s*)?\b(t\.?\s*v\.?\s*a\.?|vat)\b/i, /intra|n[°o]\s*(de\s*)?tva|identifiant|num[ée]ro\s*(de\s*)?tva|tva\s*:\s*fr/i);
    f.ttc = amountAfterLabel(lines, /(total|montant)\s*t\.?\s*t\.?\s*c\.?|net\s*[àa]\s*payer|total\s*[àa]\s*payer|reste\s*[àa]\s*payer|amount\s*due|total\s*due|montant\s*d[uû]/i);
    f.siret = regexAnywhere(lines, /\b\d{3}\s?\d{3}\s?\d{3}\s?\d{5}\b/);
    f.vat = regexAnywhere(lines, /\bFR\s?[0-9A-Z]{2}\s?\d{3}\s?\d{3}\s?\d{3}\b/);
    f.iban = regexAnywhere(lines, /\b[A-Z]{2}\d{2}(?:\s?[A-Z0-9]{4}){3,7}(?:\s?[A-Z0-9]{1,3})?\b/, (s) => /^(FR|BE|LU|DE|ES|IT|CH|MC|NL|PT|GB)/.test(s));
    if (f.ht && f.tva && f.ttc) {
      const d = round2(f.ht.value + f.tva.value - f.ttc.value);
      f.check = Math.abs(d) < 0.011 ? { ok: true, text: "✓ HT + TVA = TTC" } : { ok: false, text: `✗ écart de ${fmtNum(d)}` };
    } else f.check = { ok: false, text: "à vérifier (montants incomplets)" };
    return f;
  }
  async function extractReceipt(doc) {
    await openDocData(doc);
    const lines = await docLines(doc, 2), f = {};
    f.merchant = topName(lines);
    f.date = regexAnywhere(lines, DATE_RE);
    f.time = regexAnywhere(lines, /\b([01]?\d|2[0-3])[:h][0-5]\d\b/);
    f.total = amountAfterLabel(lines, /total\s*(t\.?t\.?c\.?|[àa]\s*payer|eur|€)?|montant\s*(d[uû]|pay[ée])|[àa]\s*payer|net\s*[àa]\s*payer/i, /sous[- ]total|total\s*h\.?t|tva/i);
    f.tva = amountAfterLabel(lines, /\b(t\.?\s*v\.?\s*a\.?|vat)\b/i, /intra|n[°o]/i);
    return f;
  }
  const INV_COLS = [["supplier", "Fournisseur"], ["number", "N° facture"], ["date", "Date"], ["due", "Échéance"], ["ht", "Total HT"], ["tva", "TVA"], ["ttc", "Total TTC"], ["siret", "SIRET"], ["vat", "N° TVA intracom."], ["iban", "IBAN"], ["check", "Contrôle"]];
  const REC_COLS = [["merchant", "Commerçant"], ["date", "Date"], ["time", "Heure"], ["total", "Total"], ["tva", "TVA"]];
  async function runDocExtraction(kind) {
    const docs = checkedDocs("invDocs");
    if (!docs.length) return status("Cochez au moins un document dans la section Factures & reçus.", "warn");
    const COLS = kind === "invoice" ? INV_COLS : REC_COLS;
    const cells = [{ row: 0, col: 0, value: "Document", type: "head" }, ...COLS.map(([, l], j) => ({ row: 0, col: 1 + j, value: l, type: "head" }))];
    const entries = []; let found = 0;
    for (let i = 0; i < docs.length; i++) {
      const d = docs[i];
      status(`Lecture de « ${d.name} » (${i + 1}/${docs.length})…`, "busy");
      let f;
      try { f = kind === "invoice" ? await extractInvoice(d) : await extractReceipt(d); } catch (e) { status(`« ${d.name} » : ${e.message}`, "warn"); continue; }
      cells.push({ row: i + 1, col: 0, value: d.name });
      COLS.forEach(([k], j) => {
        const v = f[k]; if (!v) return;
        if (k === "check") { cells.push({ row: i + 1, col: 1 + j, value: v.text, type: v.ok ? "valid" : "except" }); return; }
        const tv = ["ht", "tva", "ttc", "total"].includes(k) ? { v: v.value, f: NUMFMT } : (["siret", "vat", "iban", "number"].includes(k) ? { v: v.text, f: "@" } : typed(v.text));
        found++;
        cells.push({ row: i + 1, col: 1 + j, value: tv.v, fmt: tv.f, type: "text" });
        entries.push({ row: i + 1, col: 1 + j, type: "text", value: tv.v, rects: [{ docId: d.id, page: v.page, ...v.rect }] });
      });
    }
    const { name, row0 } = await writeSheet(kind === "invoice" ? "Factures" : "Reçus", cells);
    await linkCells(name, row0, entries, uid());
    refreshSnipList();
    status(`${docs.length} document(s) lus, ${found} valeur(s) extraite(s) dans la feuille « ${name} ». Vérifiez-les : chaque valeur est liée à sa pièce.`, "ok");
  }

  /* ---------------- Comparaison de versions ------------------------------ */
  async function allLinesOf(doc) {
    await openDocData(doc);
    const out = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const pg = await pageText(doc, p, { allowOcr: true });
      for (const L of pg.lines) { const k = norm(L.text).replace(/\s+/g, " ").trim(); if (k) out.push({ p, L, k }); }
    }
    return out;
  }
  async function compareVersions() {
    const A = S.docs.get($("#cmpA").value), B = S.docs.get($("#cmpB").value);
    if (!A || !B || A === B) return status("Choisissez deux documents différents (ancienne et nouvelle version).", "warn");
    status("Comparaison en cours…", "busy");
    const la = await allLinesOf(A), lb = await allLinesOf(B);
    const n = la.length, m = lb.length;
    if (n * m > 12e6) return status("Documents trop longs pour une comparaison ligne à ligne.", "warn");
    const W = m + 1, dp = new Uint16Array((n + 1) * W);
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i * W + j] = la[i].k === lb[j].k ? dp[(i + 1) * W + j + 1] + 1 : Math.max(dp[(i + 1) * W + j], dp[i * W + j + 1]);
    const ops = []; let i = 0, j = 0;
    while (i < n && j < m) {
      if (la[i].k === lb[j].k) { i++; j++; }
      else if (dp[(i + 1) * W + j] >= dp[i * W + j + 1]) ops.push({ kind: "del", it: la[i++], doc: A });
      else ops.push({ kind: "add", it: lb[j++], doc: B });
    }
    while (i < n) ops.push({ kind: "del", it: la[i++], doc: A });
    while (j < m) ops.push({ kind: "add", it: lb[j++], doc: B });
    S.diffMarks = ops.map((o) => ({ docId: o.doc.id, page: o.it.p, rect: bboxOf(o.it.L.words, 1.5), kind: o.kind }));
    S.cmp = { A, B, ops };
    const ul = $("#cmpOut"); ul.innerHTML = "";
    $("#cmpInfo").innerHTML = ops.length ? `<span class="ko">${ops.filter((o) => o.kind === "del").length} ligne(s) supprimée(s)</span> · <span class="ok">${ops.filter((o) => o.kind === "add").length} ligne(s) ajoutée(s)</span>. Surlignées en rouge et vert sur les pièces.` : "Aucune différence de texte entre les deux versions.";
    for (const o of ops.slice(0, 300)) {
      const li = document.createElement("li"); li.className = "click";
      li.innerHTML = `<span class="chip" style="background:${o.kind === "add" ? "#22A04B" : "#D93025"}">${o.kind === "add" ? "+" : "−"}</span><div class="grow"><div class="ttl">${esc(o.it.L.text)}</div><div class="sub">${esc(o.doc.name)} · p. ${o.it.p}</div></div>`;
      li.onclick = () => goToLocation(o.doc.id, o.it.p, bboxOf(o.it.L.words, 1.5));
      ul.appendChild(li);
    }
    drawOverlay();
    status(`Comparaison terminée : ${ops.length} différence(s).`, "ok");
  }
  async function exportComparison() {
    if (!S.cmp) return status("Lancez d'abord la comparaison.", "warn");
    const cells = [["Type", 0], ["Texte", 1], ["Document", 2], ["Page", 3]].map(([v, c]) => ({ row: 0, col: c, value: v, type: "head" }));
    const entries = [];
    S.cmp.ops.forEach((o, i) => {
      const type = o.kind === "add" ? "valid" : "except";
      cells.push({ row: i + 1, col: 0, value: o.kind === "add" ? "Ajouté" : "Supprimé", type }, { row: i + 1, col: 1, value: o.it.L.text, fmt: "@" }, { row: i + 1, col: 2, value: o.doc.name }, { row: i + 1, col: 3, value: o.it.p });
      entries.push({ row: i + 1, col: 1, type: "text", value: o.it.L.text, rects: [{ docId: o.doc.id, page: o.it.p, ...bboxOf(o.it.L.words, 1.5) }] });
    });
    const { name, row0 } = await writeSheet("Comparaison", cells);
    await linkCells(name, row0, entries, uid());
    refreshSnipList();
    status(`Différences exportées dans la feuille « ${name} ».`, "ok");
  }

  /* ---------------- Adaptateurs Excel : fonctions supplémentaires -------- */
  Object.assign(OfficeXL, {
    async newSheet(base) {
      return Excel.run(async (ctx) => {
        const wss = ctx.workbook.worksheets; wss.load("items/name"); await ctx.sync();
        const names = new Set(wss.items.map((w) => w.name));
        let n = base.slice(0, 31), k = 2;
        while (names.has(n)) n = `${base.slice(0, 25)} (${k++})`;
        const ws = wss.add(n); ws.activate(); await ctx.sync();
        return { name: n, row0: 0 };
      });
    },
    async writeCells(sheet, cells) {
      for (let i = 0; i < cells.length; i += 400) {
        const part = cells.slice(i, i + 400);
        await Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItem(sheet);
          for (const c of part) {
            const rg = ws.getRangeByIndexes(c.row, c.col, 1, 1), v = cellOut(c.value);
            rg.numberFormat = [[c.fmt || (typeof v === "number" ? NUMFMT : "General")]];
            rg.values = [[v]];
            if (c.type) applyStyle(rg, c.type);
          }
          await ctx.sync();
        });
      }
      try { await Excel.run(async (ctx) => { ctx.workbook.worksheets.getItem(sheet).getUsedRange().format.autofitColumns(); await ctx.sync(); }); } catch (_) { /* facultatif */ }
    },
    async readRange(sheet, address) {
      return Excel.run(async (ctx) => { const rg = ctx.workbook.worksheets.getItem(sheet).getRange(address); rg.load("values"); await ctx.sync(); return rg.values; });
    },
    async usedAddress(sheet) {
      return Excel.run(async (ctx) => { const ur = ctx.workbook.worksheets.getItem(sheet).getUsedRangeOrNullObject(true); ur.load("address"); await ctx.sync(); return ur.isNullObject ? null : ur.address.split("!").pop(); });
    },
    async sheetNameById(id) {
      return Excel.run(async (ctx) => { const ws = ctx.workbook.worksheets.getItem(id); ws.load("name"); await ctx.sync(); return ws.name; });
    },
    async relink(changes) {
      await Excel.run(async (ctx) => {
        const nms = changes.map((c) => ctx.workbook.names.getItemOrNullObject(NAME_PREFIX + c.sn.id));
        nms.forEach((n) => n.load("name"));
        await ctx.sync();
        nms.forEach((n) => { if (!n.isNullObject) n.delete(); });
        await ctx.sync();
      });
      await OfficeXL.addLinks(changes.map((c) => ({ ...c.sn, address: c.address })));
    },
    async onSort(cb) {
      if (!Office.context.requirements.isSetSupported("ExcelApi", "1.10")) return;
      await Excel.run(async (ctx) => { ctx.workbook.worksheets.onRowSorted.add(cb); await ctx.sync(); });
    },
  });
  Object.assign(DemoXL, {
    async newSheet() {
      let max = -1;
      for (const k of DemoXL._cells.keys()) max = Math.max(max, +k.split(",")[0]);
      return { name: "Feuil1", row0: max + 2 };
    },
    async writeCells(sheet, cells) { for (const c of cells) await DemoXL.write(sheet, c.row, c.col, [[c.value]], c.type || null); },
    async readRange(sheet, address) {
      const r = parseRef("Feuil1!" + address), out = [];
      for (let i = r.r1; i <= r.r2; i++) { const row = []; for (let j = r.c1; j <= r.c2; j++) { const v = DemoXL._cells.get(i + "," + j); row.push(v ? v.v : ""); } out.push(row); }
      return out;
    },
    async usedAddress() {
      let r2 = 0, c2 = 0;
      for (const k of DemoXL._cells.keys()) { const [r, c] = k.split(",").map(Number); r2 = Math.max(r2, r); c2 = Math.max(c2, c); }
      return rangeA1(0, 0, r2 + 1, c2 + 1);
    },
    async sheetNameById() { return "Feuil1"; },
    async relink() {},
  });

  /* ---------------- Liaison de l'interface V2 ---------------------------- */
  function showImport() { const sh = $("#importSheet"); if (sh && S.docs.size) { sh.hidden = false; setTimeout(() => $("#importBig").focus(), 0); } }
  function hideImport() { const sh = $("#importSheet"); if (sh) sh.hidden = true; }
  function bindPanel() {
    const menu = $("#mainMenu");
    $("#menuBtn").onclick = (e) => { e.stopPropagation(); menu.hidden = !menu.hidden; };
    document.addEventListener("click", (e) => { if (!menu.hidden && !e.target.closest("#mainMenu") && !e.target.closest("#menuBtn")) menu.hidden = true; });
    $$("#mainMenu [data-tab]").forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
    $$("#mainMenu [data-action]").forEach((b) => b.addEventListener("click", () => { menu.hidden = true; }));
    $$("[data-back]").forEach((b) => (b.onclick = () => switchTab("viewer")));
    $("#topSearch").onsubmit = (e) => {
      e.preventDefault();
      const q = $("#topSearchInput").value.trim(); if (!q) return;
      $("#searchInput").value = q; switchTab("search");
      runAction(() => doSearch(q), { label: "Recherche impossible" });
    };
    const tb = $(".topbar");
    const openSearch = () => { tb.classList.add("searching"); setTimeout(() => $("#topSearchInput").focus(), 0); };
    const closeSearch = () => { tb.classList.remove("searching"); $("#topSearchInput").value = ""; };
    $("#searchToggle").onclick = openSearch;
    $("#searchClose").onclick = closeSearch;
    $("#topSearchInput").onkeydown = (e) => { if (e.key === "Escape") { e.stopPropagation(); closeSearch(); } };
    document.addEventListener("keydown", (e) => { if (e.ctrlKey && (e.key === "f" || e.key === "F")) { e.preventDefault(); openSearch(); } });
    $("#modeChip").onclick = () => setMode("nav");
    $("#importClose").onclick = hideImport;
    if (S.isWindow) document.body.classList.add("window");
    else if (S.inOffice) document.body.classList.add("in-office"); else document.body.classList.add("demo");
    $("#winBtn").onclick = () => (S.isWindow ? parentSend({ event: "dock" }) : openWindow());
    $("#winBack").onclick = () => closeWindow();
  }
  function bindExtras() {
    bindPanel();
    bindContextMenu();
    renderStyleBox();
    // import de dossiers, dossiers de rangement
    $$("[data-action=importFolder]").forEach((b) => (b.onclick = () => $("#folderInput").click()));
    $("#folderInput").onchange = (e) => {
      const fs = [...e.target.files].filter((f) => SUPPORTED_RE.test(f.name) || TEXT_RE.test(f.name) || /\.zip$/i.test(f.name));
      e.target.value = ""; hideImport();
      if (!fs.length) return status("Aucun fichier pris en charge dans ce dossier.", "warn");
      runAction(() => importFiles(fs), { label: "Import impossible" });
    };
    $("#importFolderSel").onchange = (e) => { S.importFolder = e.target.value; };
    $("#addFolder").onclick = () => {
      const v = $("#newFolder").value.trim(); if (!v) return;
      if (!S.meta.folders.includes(v)) S.meta.folders.push(v);
      $("#newFolder").value = ""; S.importFolder = v; persistSoon(); renderDocList();
    };
    // actions sur les documents
    const run = (fn, label) => () => runAction(fn, { label: label + " impossible" });
    $("#expAnnCur").onclick = run(() => exportAnnotated(false), "Export");
    $("#expAnnAll").onclick = run(() => exportAnnotated(true), "Export");
    $("#rmNoSnips").onclick = run(() => { const d = curDoc(); if (!d) throw new Error("ouvrez un document"); return removePages(d, false); }, "Suppression de pages");
    $("#rmNoExc").onclick = run(() => { const d = curDoc(); if (!d) throw new Error("ouvrez un document"); return removePages(d, true); }, "Suppression de pages");
    $("#compressBtn").onclick = run(compressAll, "Compression");
    // collage de fichiers (Ctrl+V)
    document.addEventListener("paste", (e) => {
      if (!S.opts.pasteImport || e.target.matches("input, textarea")) return;
      const fs = [...((e.clipboardData && e.clipboardData.files) || [])];
      if (fs.length) { e.preventDefault(); runAction(() => importFiles(fs), { label: "Import impossible" }); }
    });
    // liste des snips
    $("#snipFilter").onchange = async (e) => {
      S.snipFilter = e.target.value;
      if (S.snipFilter === "sel") {
        try { const s = await xl.selected(); S.selBox = { sheet: s.sheet, r1: s.row, c1: s.col, r2: s.row + s.values.length - 1, c2: s.col + (s.values[0] || []).length - 1 }; }
        catch (err) { status(err.message, "err"); }
      }
      refreshSnipList();
    };
    $("#linkSelBtn").onclick = run(linkSnipToSelection, "Liaison");
    $("#relinkBtn").onclick = run(relinkActiveSheet, "Recalage");
    // rapprochement : mode ligne par ligne
    $$("input[name=mmode]").forEach((r) => (r.onchange = () => { const row = $("input[name=mmode]:checked").value === "row"; $("#rowMatchBox").hidden = !row; $("#cellMatchBox").hidden = row; }));
    $("#rmPrep").onclick = run(prepareRowMatch, "Préparation");
    $("#rmRun").onclick = run(runRowMatch, "Document Matching");
    renderTplSelect("#rmTpl", "snipcompta.matchTpl");
    $("#rmSave").onclick = () => {
      const name = $("#rmTplName").value.trim(); if (!name) return status("Donnez un nom au modèle.", "warn");
      const t = loadTpl("snipcompta.matchTpl"); t[name] = { header: $("#rmHeader").checked, once: $("#rmOnce").checked, cols: S.rm.cols };
      saveTpl("snipcompta.matchTpl", t); renderTplSelect("#rmTpl", "snipcompta.matchTpl"); status(`Modèle « ${name} » enregistré sur ce PC.`, "ok");
    };
    $("#rmLoad").onclick = () => {
      const t = loadTpl("snipcompta.matchTpl")[$("#rmTpl").value]; if (!t) return;
      $("#rmHeader").checked = !!t.header; $("#rmOnce").checked = !!t.once; S.rm.cols = t.cols || []; renderRowMatchCols();
      status("Modèle chargé. Sélectionnez une plage avec les mêmes colonnes puis lancez.", "ok");
    };
    renderRowMatchCols();
    // automatisation
    $("#sumsDocBtn").onclick = run(findDocSums, "Recherche des sommes");
    $("#formMode").onchange = (e) => { S.form.mode = e.target.value; };
    $("#formAdd").onclick = () => { if (!curDoc()) return status("Ouvrez d'abord un document type dans la Visionneuse.", "warn"); switchTab("viewer"); setMode("formfield"); };
    $("#formClear").onclick = () => { S.form.fields = []; S.form.refDoc = null; S.form.rows = null; $("#formOut").innerHTML = ""; renderForm(); drawOverlay(); };
    $("#formPreview").onclick = run(formPreview, "Aperçu");
    $("#formExport").onclick = run(formExport, "Export");
    renderTplSelect("#formTpl", "snipcompta.formTpl");
    $("#formSave").onclick = () => {
      const name = $("#formTplName").value.trim(); if (!name || !S.form.fields.length) return status("Ajoutez des champs et donnez un nom au modèle.", "warn");
      const t = loadTpl("snipcompta.formTpl"); t[name] = { mode: S.form.mode, fields: S.form.fields.map((f) => ({ ...f, id: undefined })) };
      saveTpl("snipcompta.formTpl", t); renderTplSelect("#formTpl", "snipcompta.formTpl"); status(`Modèle « ${name} » enregistré sur ce PC.`, "ok");
    };
    $("#formLoad").onclick = () => {
      const t = loadTpl("snipcompta.formTpl")[$("#formTpl").value]; if (!t) return;
      S.form.mode = t.mode || "doc"; $("#formMode").value = S.form.mode;
      S.form.fields = (t.fields || []).map((f) => ({ ...f, id: uid() })); S.form.refDoc = S.cur;
      renderForm(); drawOverlay(); status("Modèle chargé : il sera appliqué à tous les documents cochés.", "ok");
    };
    $("#formAll").onclick = () => $$("#formDocs input").forEach((i) => (i.checked = true));
    $("#invAll").onclick = () => $$("#invDocs input").forEach((i) => (i.checked = true));
    $("#invRun").onclick = run(() => runDocExtraction("invoice"), "Extraction");
    $("#recRun").onclick = run(() => runDocExtraction("receipt"), "Extraction");
    $("#cmpRun").onclick = run(compareVersions, "Comparaison");
    $("#cmpExport").onclick = run(exportComparison, "Export");
    $("#cmpClear").onclick = () => { S.diffMarks = []; S.cmp = null; $("#cmpOut").innerHTML = ""; $("#cmpInfo").textContent = ""; drawOverlay(); };
    $("#sumsClear").onclick = () => { S.sumMarks = []; $("#sumsOut").innerHTML = ""; drawOverlay(); };
    renderForm(); renderAnnotList();
  }

  /* ====================================================================== */
  /* Fenêtre de document indépendante (boîte de dialogue Office)            */
  /*  - la fenêtre affiche la visionneuse complète ;                        */
  /*  - elle ne peut pas écrire dans Excel : chaque opération Excel est     */
  /*    relayée au panneau (« hub »), qui l'exécute et renvoie le résultat. */
  /* ====================================================================== */
  const IS_WINDOW = /[?&]window=1(&|$)/.test(location.search);
  S.win = null; S.isWindow = IS_WINDOW;
  let RPC_CH = 48000;
  function rpcSend(str, fn) {
    if (str.length <= RPC_CH) { fn(str); return; }
    const cid = uid(), n = Math.ceil(str.length / RPC_CH);
    for (let i = 0; i < n; i++) fn(JSON.stringify({ __c: cid, i, n, d: str.slice(i * RPC_CH, (i + 1) * RPC_CH) }));
  }
  const rpcBuf = new Map();
  function rpcRecv(raw, onMsg) {
    let m;
    try { m = JSON.parse(raw); } catch (_) { return; }
    if (m && m.__c) {
      let b = rpcBuf.get(m.__c);
      if (!b) { b = { parts: new Array(m.n), got: 0 }; rpcBuf.set(m.__c, b); }
      if (b.parts[m.i] == null) { b.parts[m.i] = m.d; b.got++; }
      if (b.got === m.n) { rpcBuf.delete(m.__c); rpcRecv(b.parts.join(""), onMsg); }
      return;
    }
    onMsg(m);
  }

  /* ---- côté panneau ---- */
  let HUBXL = null;
  let winSendRaw = (s) => { if (S.win) S.win.messageChild(s); };
  function winSend(obj) { if (S.win) rpcSend(JSON.stringify(obj), winSendRaw); }
  async function hubHandle(m) {
    if (!m) return;
    if (m.event === "ready") { applyTheme(); return; }
    if (m.event === "winSize") { storeWinSize({ w: m.w, h: m.h }); return; }
    if (m.event === "dock") { closeWindow(); return; }
    if (m.event === "docChanged") { await onDocChanged(m.doc); return; }
    const { id, op, args = [] } = m;
    if (!id) return;
    const reply = (o) => winSend({ id, ...o });
    try {
      let res;
      if (op === "hubDocs") res = [...S.docs.values()].map((d) => ({ id: d.id, name: d.name, mime: d.mime, stored: !!d.stored, folder: d.folder || "" }));
      else if (op === "hubDocBytes") {
        const d = S.docs.get(args[0]);
        let b = d && d.bytes;
        if (!b && d) b = await HUBXL.loadDocBytes(d.id);
        res = b ? toB64(b) : null;
      } else if (op === "saveDoc") {
        const d = args[0];
        await HUBXL.saveDoc({ id: d.id, name: d.name, mime: d.mime, bytes: fromB64(d.b64) });
        res = true;
      } else if (typeof HUBXL[op] === "function") res = await HUBXL[op](...args);
      else throw new Error("opération inconnue : " + op);
      reply({ ok: true, res: res === undefined ? null : res });
    } catch (e) { reply({ ok: false, err: (e && e.message) || String(e), code: (e && e.code) || "" }); }
  }
  function updateWinBtn() {
    const b = $("#winBtn"); if (!b) return;
    b.title = S.isWindow ? "Ramener le document dans le panneau" : (S.win ? "Le document est ouvert dans la fenêtre séparée" : "Ouvrir le document dans une fenêtre séparée (agrandissable)");
    b.classList.toggle("on", !!S.win);
  }
  function openWindow() {
    if (S.isWindow) return;
    if (!S.inOffice || typeof Office === "undefined" || !Office.context || !Office.context.ui || !Office.context.ui.displayDialogAsync) return status("La fenêtre séparée n'est disponible que dans Excel.", "warn");
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return status("Votre version d'Excel ne permet pas la fenêtre séparée (mise à jour Microsoft 365 nécessaire).", "warn");
    if (S.win) return status("Le document est déjà ouvert dans la fenêtre séparée.", "ok");
    HUBXL = OfficeXL;
    const url = new URL("taskpane.html", location.href);
    url.search = "";
    url.searchParams.set("window", "1");
    if (S.cur) { url.searchParams.set("doc", S.cur); url.searchParams.set("page", String(S.page)); }
    url.searchParams.set("v", BUILD);
    const size = savedWinSize();
    Office.context.ui.displayDialogAsync(url.href, { width: size.w, height: size.h, displayInIframe: false }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return status("Ouverture de la fenêtre impossible : " + ((r.error && r.error.message) || "erreur inconnue"), "err");
      S.win = r.value;
      S.win.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => rpcRecv(arg.message, hubHandle));
      S.win.addEventHandler(Office.EventType.DialogEventReceived, () => onWindowClosed());
      $("#winBanner").hidden = false; updateWinBtn();
      status("Document ouvert dans une fenêtre séparée : déplacez-la, agrandissez-la ou posez-la sur un autre écran.", "ok");
    });
  }
  function closeWindow() {
    const w = S.win;
    if (w) { try { w.close(); } catch (_) { /* déjà fermée */ } }
    onWindowClosed();
  }
  async function onWindowClosed() {
    if (!S.win) return;
    S.win = null;
    $("#winBanner").hidden = true; updateWinBtn();
    try { await hubRefresh(); } catch (e) { console.warn(e); }
  }
  async function hubRefresh() {
    const docs = await xl.listDocs();
    for (const d of docs) if (!S.docs.has(d.id)) S.docs.set(d.id, { ...d, kind: kindOf(d.name, d.mime), bytes: null, pages: {}, stored: true });
    for (const [id, d] of [...S.docs]) if (d.stored && !docs.some((x) => x.id === id)) S.docs.delete(id);
    applyDocMeta();
    await reloadSnips(); refreshDocUI(); render();
  }

  /* ---- côté fenêtre ---- */
  const pending = new Map();
  let winSelCb = null;
  let parentSendRaw = null; // remplaçable (tests)
  function parentSend(obj) {
    const f = parentSendRaw || window.__rpcToHub || ((s) => Office.context.ui.messageParent(s));
    rpcSend(JSON.stringify(obj), f);
  }
  function rcall(op, args) {
    return new Promise((resolve, reject) => {
      const id = uid();
      pending.set(id, { resolve, reject });
      parentSend({ id, op, args });
    });
  }
  async function windowRecv(m) {
    if (!m) return;
    if (m.id && pending.has(m.id)) {
      const p = pending.get(m.id); pending.delete(m.id);
      if (m.ok) p.resolve(m.res); else { const er = new Error(m.err); if (m.code) er.code = m.code; p.reject(er); }
      return;
    }
    if (m.event === "selection") { if (winSelCb) winSelCb(); }
    else if (m.event === "action") { const fn = getActions()[m.name]; if (fn) runAction(fn); }
    else if (m.event === "open") { if (S.docs.has(m.doc)) openDoc(m.doc, m.page || 1); }
    else if (m.event === "reloadDocs") { await windowLoadDocs(); }
    else if (m.event === "docChanged") { await onDocChanged(m.doc); }
    else if (m.event === "status") { status(m.msg, m.kind); }
    else if (m.event === "theme") { S.hubDark = !!m.dark; applyTheme(); }
  }
  const REMOTE_OPS = ["activeCell", "write", "select", "selectAddr", "selected", "format", "addLinks", "removeLinks", "selectionInfo", "saveSnips", "loadSnips",
    "listDocs", "deleteDoc", "readCell", "addComment", "exportAudit", "newSheet", "writeCells", "readRange", "usedAddress", "sheetNameById", "relink"];
  const RemoteXL = {};
  for (const op of REMOTE_OPS) RemoteXL[op] = (...args) => rcall(op, args);
  RemoteXL.saveDoc = (doc) => rcall("saveDoc", [{ id: doc.id, name: doc.name, mime: doc.mime, b64: toB64(doc.bytes) }]);
  RemoteXL.loadDocBytes = async (id) => { const b = await rcall("hubDocBytes", [id]); return b ? fromB64(b) : null; };
  RemoteXL.onSelection = (cb) => { winSelCb = cb; };
  async function windowLoadDocs() {
    const docs = await rcall("hubDocs", []);
    let added = 0;
    for (const d of docs) if (!S.docs.has(d.id)) { S.docs.set(d.id, { ...d, kind: kindOf(d.name, d.mime), bytes: null, pages: {} }); added++; }
    refreshDocUI();
    if (added && !S.cur) await openDoc(docs[0].id);
    return docs;
  }
  async function initWindow() {
    document.body.classList.add("window");
    if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.addHandlerAsync)
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => rpcRecv(arg.message, windowRecv));
    parentSend({ event: "ready" });
    status("Chargement des documents…", "busy");
    const docs = await windowLoadDocs();
    await reloadSnips();
    refreshDocUI();
    const q = new URLSearchParams(location.search), id = q.get("doc");
    if (id && S.docs.has(id)) await openDoc(id, +q.get("page") || 1);
    else if (docs.length) await openDoc(docs[0].id);
    status(docs.length ? "Fenêtre prête : agrandissez-la comme vous voulez." : "Importez vos pièces pour commencer.", "ok");
  }

  /* ====================================================================== */
  /* Ergonomie : snip en un clic, survol, bulles de confirmation            */
  /* ====================================================================== */
  const CLICK_MODES = new Set(["text", "sum", "valid", "except", "comment", "table"]);
  function lineAt(pg, pt, tol = 1.5) {
    let best = null, bd = Infinity;
    for (const L of pg.lines || []) {
      if (pt.y < L.y0 - tol || pt.y > L.y1 + tol) continue;
      const dx = pt.x < L.x0 ? L.x0 - pt.x : pt.x > L.x1 ? pt.x - L.x1 : 0;
      if (dx < bd) { bd = dx; best = L; }
    }
    return best && bd <= 6 ? best : null;
  }
  const cleanWord = (s) => s.replace(/^[(\[«"'“]+|[)\]»"'”,;:!?]+$/g, "").replace(/\.$/, "") || s;
  /** Ce qui se trouve sous le pointeur : un montant, sinon un mot (ou la ligne entière). */
  function targetAt(pg, pt, level = "auto") {
    const L = lineAt(pg, pt); if (!L) return null;
    if (level === "line") return { kind: "line", text: L.text, rect: bboxOf(L.words, 2), line: L };
    for (const a of findAmounts(L.text)) {
      const ws = L.words.filter((w) => w.end > a.start && w.start < a.end);
      if (ws.length === 1 && /[A-Za-zÀ-ÿ]/.test(ws[0].str)) continue;   // « FAC-2024-001 » n'est pas un montant
      const bb = spanBBox(L, a.start, a.end);
      if (pt.x >= bb.x0 - 2 && pt.x <= bb.x1 + 2) return { kind: "amount", text: L.text.slice(a.start, a.end), value: a.value, rect: bb, line: L };
    }
    let w = null, wd = Infinity;
    for (const x of L.words) { const d = pt.x < x.x0 ? x.x0 - pt.x : pt.x > x.x1 ? pt.x - x.x1 : 0; if (d < wd) { wd = d; w = x; } }
    if (!w || wd > 3) return null;
    return { kind: "word", text: cleanWord(w.str), rect: bboxOf([w], 1.5), line: L };
  }
  function amountsIn(ws) {
    let amts = [];
    for (const L of groupLines(ws)) for (const a of findAmounts(L.text)) amts.push({ ...a, bbox: spanBBox(L, a.start, a.end) });
    if (amts.some((a) => a.hasDec)) amts = amts.filter((a) => a.hasDec);
    return amts;
  }

  /* --- survol : on surligne ce qui sera capturé --- */
  let hovEl = null;
  function hideHover() { if (hovEl) { hovEl.remove(); hovEl = null; } }
  function hoverAt(e) {
    const ov = e.target && e.target.closest && e.target.closest(".ov");
    const doc = curDoc();
    if (!ov || !doc || !CLICK_MODES.has(S.mode) || e.target.classList.contains("sep")) { hideHover(); return; }
    const p = +ov.dataset.p, pg = doc.pages[p];
    if (!pg || !pg.lines) { if (!pg) pageText(doc, p, { allowOcr: false }).catch(() => {}); return; }
    const pt = evPt(e, ov);
    let r = null, kind = "";
    if (S.mode === "table") {
      const t = S.table;
      if (t && t.docId === doc.id && t.page === p && pt.x > t.rect.x0 && pt.x < t.rect.x1 && pt.y > t.rect.y0 && pt.y < t.rect.y1) { hideHover(); return; }
      r = tableRectAt(pg, pt); kind = "table";
    } else { const t = targetAt(pg, pt); if (t) { r = t.rect; kind = t.kind; } }
    if (!r) { hideHover(); return; }
    if (!hovEl) { hovEl = document.createElement("div"); hovEl.className = "hov"; }
    if (hovEl.parentNode !== ov) ov.appendChild(hovEl);
    hovEl.dataset.kind = kind;
    hovEl.style.left = r.x0 * S.scale + "px"; hovEl.style.top = r.y0 * S.scale + "px";
    hovEl.style.width = (r.x1 - r.x0) * S.scale + "px"; hovEl.style.height = (r.y1 - r.y0) * S.scale + "px";
  }

  /* --- valeur en direct pendant le glisser --- */
  function liveValue(r, e) {
    const el = $("#live"); if (!el) return;
    if (!["text", "sum"].includes(S.mode)) return;
    const doc = curDoc(), pg = doc && doc.pages[S.page]; if (!pg || !pg.words) return;
    const ws = wordsInRect(pg.words, r);
    let txt = "";
    if (S.mode === "sum") { const a = amountsIn(ws); if (a.length) txt = `Σ ${fmtNum(round2(a.reduce((s, x) => s + x.value, 0)))}  (${a.length} montant${a.length > 1 ? "s" : ""})`; }
    else { const s = joinWords(clipWords(pg.words, r)); if (s) { const tv = typed(s); txt = tv.f === NUMFMT ? fmtNum(tv.v) : String(tv.v); } }
    if (!txt) { el.hidden = true; return; }
    el.textContent = txt.length > 70 ? txt.slice(0, 68) + "…" : txt;
    el.hidden = false;
    placeFloat(el, e.clientX + 16, e.clientY + 18);
  }
  function hideLive() { const el = $("#live"); if (el) el.hidden = true; }
  function placeFloat(el, x, y) {
    const w = el.offsetWidth || 160, h = el.offsetHeight || 30;
    el.style.left = clamp(x, 6, Math.max(6, innerWidth - w - 6)) + "px";
    el.style.top = clamp(y, 6, Math.max(6, innerHeight - h - 6)) + "px";
  }

  /* --- bulle « 1 484,86 → B3 ✓  Annuler » --- */
  let confirmTimer = null;
  function confirmSnip(valueText, cell, kind) {
    const el = $("#confirm"); if (!el) return;
    status("");
    el.innerHTML = `<span class="cv"></span><span class="arr">→</span><b></b><span class="tick">✓</span><button class="link">Annuler</button>`;
    $(".cv", el).textContent = String(valueText).length > 40 ? String(valueText).slice(0, 38) + "…" : valueText;
    $("b", el).textContent = cell;
    el.dataset.kind = kind || "";
    $("button", el).onclick = () => { el.hidden = true; runAction(undoLast, { label: "Annulation impossible" }); };
    el.hidden = false; el.classList.remove("show"); void el.offsetWidth; el.classList.add("show");
    const anchor = S.selSnip && $(`#pages .ovbox.sel`);
    if (anchor) { const r = anchor.getBoundingClientRect(); placeFloat(el, r.left, r.top - el.offsetHeight - 10); }
    else if (S.lastPointer) placeFloat(el, S.lastPointer.x + 14, S.lastPointer.y - 46);
    else placeFloat(el, innerWidth / 2 - el.offsetWidth / 2, innerHeight - el.offsetHeight - 70);
    clearTimeout(confirmTimer);
    const arm = () => { confirmTimer = setTimeout(() => { el.hidden = true; }, 4000); };
    el.onmouseenter = () => clearTimeout(confirmTimer); el.onmouseleave = arm;
    arm();
  }

  /* --- clic simple = montant ou mot ; double-clic = ligne entière --- */
  let clickTimer = null, lastClick = null;
  function queueClickSnip(pt, e) {
    const p = S.page, now = Date.now(), ctrl = !!(e && e.ctrlKey);
    if (S.mode === "table") { runAction(() => clickSnip(pt, p, "auto", ctrl)); return; }
    if (lastClick && now - lastClick.t < 330 && lastClick.p === p && Math.hypot(pt.x - lastClick.pt.x, pt.y - lastClick.pt.y) * S.scale < 10) {
      clearTimeout(clickTimer); lastClick = null;
      runAction(() => clickSnip(pt, p, "line", ctrl));
      return;
    }
    lastClick = { t: now, pt, p };
    clearTimeout(clickTimer);
    clickTimer = setTimeout(() => { lastClick = null; runAction(() => clickSnip(pt, p, "auto", ctrl)); }, 280);
  }
  async function clickSnip(pt, p, level, ctrl) {
    const doc = curDoc(); if (!doc) return;
    S.page = p;
    const pg = await pageText(doc, p, { allowOcr: true });
    if (S.mode === "table") return detectTableAt(pg, pt, p);
    const t = targetAt(pg, pt, level);
    if (!t) {
      if (S.mode === "valid" || S.mode === "except") return handleRect({ x0: pt.x - 7, y0: pt.y - 7, x1: pt.x + 7, y1: pt.y + 7 }, { click: true, ctrl });
      return status(pg.needsOcr && !pg.ocrDone ? "Page scannée : le texte n'a pas encore été reconnu (OCR désactivé dans les réglages ?)." : "Cliquez sur un montant ou un mot (double-clic = toute la ligne), ou dessinez un cadre.", "warn");
    }
    if (S.mode === "text" && !ctrl) return doTextSnip(t.rect, t);
    return handleRect(t.rect, { click: true, target: t, ctrl });
  }

  /* ====================================================================== */
  /* Tableau détecté en un clic                                             */
  /* ====================================================================== */
  function lineSegments(L) {
    let n = 1;
    for (let i = 1; i < L.words.length; i++) if (L.words[i].x0 - L.words[i - 1].x1 > Math.max(6, 1.1 * L.h)) n++;
    return n;
  }
  const isTabular = (L) => lineSegments(L) >= 2 || findAmounts(L.text).some((a) => a.hasDec);
  function tableRectAt(pg, pt) {
    pg._tbl = pg._tbl || new Map();
    const lines = pg.lines || [];
    let L0 = lineAt(pg, pt, 4);
    if (!L0) {
      let bd = 14;
      for (const L of lines) { const d = Math.min(Math.abs(pt.y - L.y0), Math.abs(pt.y - L.y1)); if (d < bd && pt.x >= L.x0 - 20 && pt.x <= L.x1 + 20) { bd = d; L0 = L; } }
    }
    if (!L0) return null;
    if (pg._tbl.has(L0)) return pg._tbl.get(L0);
    const i0 = lines.indexOf(L0);
    const hs = lines.map((L) => L.y1 - L.y0).sort((a, b) => a - b);
    const med = hs[Math.floor(hs.length / 2)] || 10;
    const near = (a, b) => b.y0 - a.y1 <= Math.max(2.4 * med, 14) && Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) > -10;
    let lo = i0, hi = i0;
    for (let j = i0 + 1, miss = 0; j < lines.length; j++) {
      if (!near(lines[j - 1], lines[j])) break;
      if (isTabular(lines[j])) { hi = j; miss = 0; } else if (++miss > 1) break;
    }
    for (let j = i0 - 1, miss = 0; j >= 0; j--) {
      if (!near(lines[j], lines[j + 1])) break;
      if (isTabular(lines[j])) { lo = j; miss = 0; } else if (++miss > 1) break;
    }
    let r = null;
    if (hi > lo || isTabular(L0)) r = bboxOf(lines.slice(lo, hi + 1).flatMap((L) => L.words), 3);
    pg._tbl.set(L0, r);
    return r;
  }
  async function detectTableAt(pg, pt, p) {
    const r = tableRectAt(pg, pt);
    if (!r) return status("Aucun tableau ici : cliquez sur une ligne du tableau, ou dessinez un cadre autour.", "warn");
    hideHover();
    await startTable(r, p);
  }

  /* ====================================================================== */
  /* Progression, annulation, messages clairs, enregistrement garanti       */
  /* ====================================================================== */
  function cancelErr() { const e = new Error("Opération annulée"); e.name = "CancelError"; return e; }
  const isCancel = (e) => e && e.name === "CancelError";
  function showProgress(msg, cancellable) {
    const el = $("#progress"); if (!el) return;
    if (el.hidden) { el.hidden = false; el.classList.remove("show"); void el.offsetWidth; el.classList.add("show"); }
    $(".pmsg", el).textContent = msg;
    let frac = null;
    const m1 = msg.match(/(\d+)\s*%/), m2 = msg.match(/(\d+)\s*\/\s*(\d+)/);
    if (m1) frac = +m1[1] / 100; else if (m2 && +m2[2] > 0) frac = +m2[1] / +m2[2];
    else if (S.prog && S.prog.total > 1) frac = S.prog.i / S.prog.total;
    const bar = $(".pbar", el);
    bar.classList.toggle("indet", frac == null);
    $("i", bar).style.width = frac == null ? "" : Math.round(clamp(frac, 0.03, 1) * 100) + "%";
    const c = $(".pcancel", el);
    c.hidden = !cancellable;
    if (cancellable && !(S.op && S.op.cancelled)) { c.disabled = false; c.textContent = "Annuler"; }
  }
  function hideProgress() { const el = $("#progress"); if (el) el.hidden = true; }
  function requestCancel() {
    const op = S.op; if (!op) return;
    op.cancelled = true;
    const c = $("#progress .pcancel"); if (c) { c.disabled = true; c.textContent = "Annulation…"; }
    for (const h of op.hooks.splice(0)) { try { h(); } catch (_) { /* ignore */ } }
  }
  /** Promesse rejetée dès que l'utilisateur clique sur « Annuler ». */
  function onCancel() {
    const op = S.op; if (!op) return new Promise(() => {});
    if (op.cancelled) return Promise.reject(cancelErr());
    return new Promise((_, rej) => op.hooks.push(() => rej(cancelErr())));
  }
  async function runAction(fn, { label = "" } = {}) {
    const op = { cancelled: false, hooks: [] }, prev = S.op;
    S.op = op;
    try { return await fn(); }
    catch (e) { reportError(e, () => runAction(fn, { label }), label); }
    finally { if (S.op === op) { S.op = prev || null; if (!S.op) { const el = $("#progress"); if (el && !el.hidden) hideProgress(); } } }
  }

  /** Traduit une erreur technique en message clair, avec la solution. */
  function friendly(e) {
    const code = (e && e.code) || "", raw = (e && e.message) || String(e || "");
    const m = code + " " + raw;
    if (/CellEditMode|edit mode|mode (de )?modification|cell.?edit/i.test(m))
      return { msg: "Excel est en train de modifier une cellule. Appuyez sur Entrée ou Échap dans Excel, puis cliquez sur Réessayer." };
    if (/AccessDenied|protect|protégé/i.test(m))
      return { msg: "La feuille est protégée. Ôtez la protection (onglet Révision → Ôter la protection de la feuille), puis réessayez." };
    if (/ItemNotFound|introuvable|does not exist|n'existe pas/i.test(m))
      return { msg: "La feuille ou la cellule liée n'existe plus (renommée ou supprimée ?). Vérifiez le classeur, puis réessayez." };
    if (/PayloadSize|too large|trop volumineux|quota/i.test(m))
      return { msg: "Le document est trop lourd pour être enregistré dans le classeur.", fix: { label: "Compresser les PDF", fn: () => runAction(compressAll, { label: "Compression impossible" }) } };
    if (/ApiNotFound|not supported|non pris en charge|requirement/i.test(m))
      return { msg: "Votre version d'Excel ne connaît pas encore cette fonction. Mettez Office à jour (Fichier → Compte → Options de mise à jour → Mettre à jour).", retry: false };
    if (/Invalid PDF|InvalidPDF|PDF structure|corrupt|endommag/i.test(m))
      return { msg: "Ce PDF est abîmé. Ouvrez-le dans votre lecteur PDF, imprimez-le vers « Microsoft Print to PDF », puis réimportez le nouveau fichier.", retry: false };
    if (/Failed to fetch|NetworkError|network|réseau|load failed|chargé/i.test(m))
      return { msg: "Un module n'a pas pu être chargé (connexion internet coupée ?). Vérifiez la connexion, puis réessayez." };
    if (/mot de passe|password/i.test(m)) return { msg: raw, retry: false };
    return { msg: raw || "Erreur inconnue." };
  }
  function reportError(e, retry, label) {
    if (isCancel(e)) { hideProgress(); return status("Opération annulée.", "warn"); }
    console.error(e);
    const f = friendly(e), acts = [];
    if (f.fix) acts.push(f.fix);
    if (retry && f.retry !== false) acts.push({ label: "Réessayer", fn: retry });
    status((label ? label + " : " : "") + f.msg, "err", acts);
  }

  /* Excel occupé (cellule en cours de saisie…) : on réessaie tout seul. */
  const BUSY_RE = /CellEditMode|RequestAborted|ServiceNotAvailable|Busy|occupé|mode (de )?modification/i;
  function patchExcelRun() {
    if (typeof Excel === "undefined" || !Excel.run || Excel.__snipRetry) return;
    const orig = Excel.run;
    Excel.run = function (...a) { return withRetry(() => orig.apply(Excel, a)); };
    Excel.__snipRetry = true;
  }
  async function withRetry(fn) {
    for (let i = 0; ; i++) {
      try {
        const r = await fn();
        if (i > 0) notify("Excel a répondu : c'est enregistré.", "ok");
        return r;
      } catch (e) {
        if (!BUSY_RE.test(((e && e.code) || "") + " " + ((e && e.message) || "")) || i >= 90) throw e;
        if (i === 0) notify("Excel est occupé (une cellule est peut-être en cours de saisie). Terminez la saisie avec Entrée ou Échap — SnipCompta réessaie tout seul…", "saving");
        await sleep(1500);
      }
    }
  }
  function notify(msg, kind) { status(msg, kind); if (S.win) winSend({ event: "status", msg, kind }); }

  /* Enregistrement des snips : une file d'attente, réessais automatiques. */
  let saveRun = null, saveDirty = false, saveRetryTimer = null;
  function persistSnips() {
    saveDirty = true;
    if (!saveRun) saveRun = (async () => {
      try {
        while (saveDirty) {
          saveDirty = false;
          for (let attempt = 1; ; attempt++) {
            try { await xl.saveSnips([...S.snips.values()], S.meta); S.unsaved = false; break; }
            catch (e) {
              if (attempt < 4) { await sleep(1000 * attempt * attempt); continue; }
              S.unsaved = true;
              console.error(e);
              status("Les snips n'ont pas pu être enregistrés dans le classeur (" + friendly(e).msg + "). Ils sont gardés en mémoire et SnipCompta réessaiera dans 30 s.", "warn",
                [{ label: "Réessayer maintenant", fn: () => persistSnips() }]);
              clearTimeout(saveRetryTimer);
              saveRetryTimer = setTimeout(() => { if (S.unsaved) persistSnips(); }, 30000);
              return;
            }
          }
        }
      } finally { saveRun = null; }
    })();
    return saveRun;
  }

  /* ====================================================================== */
  /* Thème sombre automatique, taille de fenêtre mémorisée, mises à jour    */
  /* ====================================================================== */
  function isDarkColor(hex) {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || "").trim()); if (!m) return null;
    const n = parseInt(m[1], 16), r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.45;
  }
  function officeDark() {
    try { const t = Office.context.officeTheme; if (t) { const d = isDarkColor(t.bodyBackgroundColor); if (d != null) return d; } } catch (_) { /* hors Office */ }
    return null;
  }
  function applyTheme(forced) {
    let dark;
    const pref = S.opts.theme || "auto";
    if (forced != null && pref === "auto") dark = forced;
    else if (pref === "dark") dark = true;
    else if (pref === "light") dark = false;
    else { const o = officeDark(); dark = o != null ? o : !!(window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches); }
    if (S.isWindow && pref === "auto" && forced == null && S.hubDark != null) dark = S.hubDark;
    document.body.classList.toggle("dark", dark);
    if (S.win) winSend({ event: "theme", dark });
    return dark;
  }
  function watchTheme() {
    applyTheme();
    try { matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => applyTheme()); } catch (_) { /* ancien navigateur */ }
    try { Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, () => applyTheme(), () => {}); } catch (_) { /* pas pris en charge */ }
    let last = document.body.classList.contains("dark");
    setInterval(() => { if (document.hidden) return; const d = applyTheme(); if (d !== last) last = d; }, 6000);
  }

  function savedWinSize() {
    try { const s = JSON.parse(localStorage.getItem("snipcompta.winSize") || "null"); if (s && s.w >= 30 && s.h >= 30) return { w: Math.min(100, s.w), h: Math.min(100, s.h) }; } catch (_) { /* ignore */ }
    return { w: 70, h: 85 };
  }
  function storeWinSize(s) { try { localStorage.setItem("snipcompta.winSize", JSON.stringify(s)); } catch (_) { /* ignore */ } }
  function trackWinSize() {
    let t = null;
    addEventListener("resize", () => {
      clearTimeout(t);
      t = setTimeout(() => {
        const sw = screen.availWidth || screen.width, sh = screen.availHeight || screen.height;
        if (!sw || !sh) return;
        const s = { w: Math.round((outerWidth / sw) * 100), h: Math.round((outerHeight / sh) * 100) };
        if (s.w < 30 || s.h < 30) return;
        storeWinSize(s); parentSend({ event: "winSize", ...s });
        render();
      }, 400);
    });
  }

  /* Mise à jour automatique : on compare avec version.json sur le site. */
  async function checkUpdate(force) {
    if (location.protocol !== "https:" && !force) return;
    try {
      const r = await fetch(new URL("version.json", location.href).href + "?t=" + Date.now(), { cache: "no-store" });
      if (!r.ok) return;
      const v = await r.json();
      if (!v || !v.build || v.build === BUILD) return;
      const busy = S.table || S.sumParts.length || S.commentDraft || (S.op != null) || S.unsaved;
      if (new URLSearchParams(location.search).get("v") === v.build) { console.warn("SnipCompta : version.json annonce", v.build, "mais taskpane.js est", BUILD, "(fichier oublié sur GitHub ?)"); return; }
      const go = () => { const u = new URL(location.href); u.searchParams.set("v", v.build); location.replace(u.href); };
      if (!S.userActed && !busy && !S.win) { go(); return; }
      status(`Une nouvelle version de SnipCompta est disponible (${v.label || v.build}).`, "ok", [{ label: "Mettre à jour", fn: go }]);
    } catch (_) { /* hors ligne : on réessaiera */ }
  }


  /* ====================================================================== */
  /* Liste des snips de la pièce, à droite du document (fenêtre large)      */
  /* ====================================================================== */
  let slSig = "", slFilter = "all";
  const sideWanted = () => { try { const v = localStorage.getItem("snipcompta.side"); if (v != null) return v === "1"; } catch (_) { /* ignore */ } return true; };
  function sideFits() { const w = $("#viewWrap"); return !!w && w.clientWidth >= 760; }
  function updateSide() {
    const el = $("#sideList"), btn = $("#sideBtn"); if (!el) return;
    const fits = sideFits(), show = fits && sideWanted() && !!curDoc();
    if (btn) { btn.hidden = !fits; btn.classList.toggle("on", show); }
    if (el.hidden === show) { el.hidden = !show; slSig = ""; if (show) setTimeout(render, 0); }
    if (show) syncSideList();
  }
  function setSideWanted(v) { try { localStorage.setItem("snipcompta.side", v ? "1" : "0"); } catch (_) { /* ignore */ } updateSide(); }
  function sideItems(doc) {
    const groups = new Map();
    for (const sn of S.snips.values()) {
      const r = sn.rects.find((x) => (x.docId || sn.docId) === doc.id); if (!r) continue;
      const g = groups.get(sn.group || sn.id);
      if (g) { g.list.push(sn); if (r.page < g.page || (r.page === g.page && r.y0 < g.rect.y0)) { g.page = r.page; g.rect = r; } }
      else groups.set(sn.group || sn.id, { list: [sn], page: r.page, rect: r });
    }
    return [...groups.values()].sort((a, b) => a.page - b.page || a.rect.y0 - b.rect.y0 || a.rect.x0 - b.rect.x0);
  }
  function syncSideList(force) {
    const el = $("#sideList"); if (!el || el.hidden) return;
    const doc = curDoc(); if (!doc) return;
    const sig = [doc.id, S.snips.size, S.selSnip, S.snipRev || 0, slFilter].join("|");
    if (!force && sig === slSig) return;
    slSig = sig;
    const all = sideItems(doc), items = slFilter === "all" ? all : all.filter((g) => g.list[0].type === slFilter);
    $("#slCount").textContent = all.length ? String(all.length) : "";
    const ul = $("#slItems");
    if (!items.length) {
      ul.innerHTML = `<li class="sl-empty">${all.length ? "Aucun snip de ce type sur cette pièce." : "Aucun snip sur cette pièce.<br>Choisissez un outil, puis cliquez un montant."}</li>`;
      return;
    }
    ul.innerHTML = "";
    let prevPage = 0;
    for (const g of items) {
      const sn = g.list[0], st = STYLE[sn.type] || STYLE.text;
      if (g.page !== prevPage && new Set(items.map((x) => x.page)).size > 1) {
        const h = document.createElement("li"); h.className = "sl-page"; h.textContent = "Page " + g.page; ul.appendChild(h); prevPage = g.page;
      }
      const isTable = g.list.length > 1;
      const last = g.list[g.list.length - 1];
      const addr = isTable ? `${String(sn.address).split(":")[0]}:${String(last.address).split(":").pop()}` : sn.address;
      let val = isTable ? `${g.list.length} lignes` : fmtVal(sn.value);
      if (!isTable && (sn.type === "valid" || sn.type === "except")) {
        const pg = doc.pages[g.page];
        const t = pg && pg.words ? joinWords(wordsInRect(pg.words, g.rect)) : "";
        if (t) val = t;
      }
      const li = document.createElement("li");
      li.className = "sl-item" + (g.list.some((x) => x.id === S.selSnip) ? " on" : "");
      li.tabIndex = 0;
      li.innerHTML = `<span class="sl-chip" style="background:${st.color}">${esc(st.label)}</span><span class="sl-val">${esc(String(val).slice(0, 60))}</span><span class="sl-cell">${esc(addr)}</span>`;
      li.title = `${st.name} — ${sn.sheet}!${sn.address}${sn.rects.length > 1 ? ` (${sn.rects.length} zones)` : ""}`;
      const go = () => { S.selSnip = sn.id; scrollToPage(g.page, g.rect); S.flash = { docId: doc.id, page: g.page, rect: g.rect }; drawOverlay(); setTimeout(() => { if (S.flash && S.flash.rect === g.rect) { S.flash = null; drawOverlay(); } }, 1600); xl.selectAddr(sn.sheet, sn.address).catch(() => {}); };
      li.onclick = go;
      li.onkeydown = (e) => { if (e.key === "Enter") go(); };
      ul.appendChild(li);
    }
    const on = $(".sl-item.on", ul); if (on) on.scrollIntoView({ block: "nearest" });
  }
  function bindSideList() {
    $("#sideBtn").onclick = () => setSideWanted(!( !$("#sideList").hidden));
    $("#slClose").onclick = () => setSideWanted(false);
    $$("[data-slf]").forEach((b) => (b.onclick = () => { slFilter = b.dataset.slf; $$("[data-slf]").forEach((x) => x.classList.toggle("on", x === b)); syncSideList(true); }));
    if (typeof ResizeObserver !== "undefined") new ResizeObserver(() => updateSide()).observe($("#viewWrap"));
    updateSide();
  }

  /* ====================================================================== */
  /* Boutons du ruban (onglet SNIPCOMPTA)                                   */
  /* ====================================================================== */
  let ACTIONS = null;
  const getActions = () => ACTIONS || (ACTIONS = buildActions());
  const LOCAL_ACTIONS = new Set(["openViewer", "openWindow", "toggleStore", "openSettings", "openHelp"]);
  function registerRibbon() {
    if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) return;
    const show = () => { try { if (Office.addin && Office.addin.showAsTaskpane) Office.addin.showAsTaskpane().catch(() => {}); } catch (_) { /* ancien Excel */ } };
    for (const [name, fn] of Object.entries(getActions())) {
      Office.actions.associate(name, (event) => {
        if (S.win && !LOCAL_ACTIONS.has(name)) winSend({ event: "action", name });
        else { show(); runAction(fn); }
        if (event && event.completed) event.completed();
      });
    }
  }
  function buildActions() {
    const A = {};
    const act = (name, fn) => { A[name] = fn; };
    act("openWindow", () => openWindow());
    const mode = (m) => () => { switchTab("viewer"); setMode(m); };
    act("openViewer", () => switchTab("viewer"));
    act("importDocs", () => { switchTab("viewer"); showImport(); try { $("#fileInput").click(); } catch (_) { /* bloqué : le grand bouton prend le relais */ } });
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
    act("openSettings", () => switchTab("settings"));
    act("openHelp", () => { switchTab("settings"); $("#helpBox").scrollIntoView(); });
    act("importFolder", () => { switchTab("viewer"); showImport(); try { $("#folderInput").click(); } catch (_) { /* bloqué */ } });
    act("modeMarkup", mode("markup"));
    act("modeNote", mode("note"));
    act("modeRedact", () => { S.markTool = "redact"; switchTab("viewer"); setMode("markup"); });
    act("findDocSums", () => { switchTab("auto"); return findDocSums(); });
    act("formExtraction", () => { switchTab("auto"); $("#formBox").scrollIntoView(); });
    act("invoiceExtraction", () => { switchTab("auto"); $("#invBox").scrollIntoView(); status("Cochez les factures puis « Extraire les factures »."); });
    act("receiptExtraction", () => { switchTab("auto"); $("#invBox").scrollIntoView(); status("Cochez les reçus puis « Extraire les reçus »."); });
    act("versionCompare", () => { switchTab("auto"); $("#cmpBox").scrollIntoView(); });
    act("rowMatching", () => { switchTab("match"); const r = $("input[name=mmode][value=row]"); r.checked = true; r.onchange(); });
    act("exportAnnotated", () => exportAnnotated(true));
    act("removePagesNoSnips", () => { const d = curDoc(); if (!d) throw new Error("ouvrez un document"); return removePages(d, false); });
    act("removePagesNoExceptions", () => { const d = curDoc(); if (!d) throw new Error("ouvrez un document"); return removePages(d, true); });
    act("compressDocs", () => compressAll());
    act("relinkSort", () => relinkActiveSheet());
    act("linkSnipCells", () => linkSnipToSelection());
    return A;
  }

  let started = false;
  async function init(inOffice) {
    if (started) return; started = true;
    S.inOffice = inOffice || IS_WINDOW;
    if (inOffice || IS_WINDOW) patchExcelRun();
    xl = IS_WINDOW ? RemoteXL : (inOffice ? OfficeXL : DemoXL);
    if (typeof pdfjsLib !== "undefined") pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
    loadOpts();
    loadStyles();
    bindUI();
    bindExtras();
    bindSideList();
    watchTheme();
    if (IS_WINDOW) trackWinSize();
    for (const ev of ["pointerdown", "keydown"]) addEventListener(ev, () => { S.userActed = true; }, { once: true, capture: true });
    if (!S.inOffice) { $("#demoBar").hidden = false; DemoXL.bind(); }
    xl.onSelection(onSelectionChanged);
    if (inOffice && !IS_WINDOW) { try { registerRibbon(); } catch (e) { console.warn("Ruban :", e); } }
    if (inOffice && !IS_WINDOW) { try { xl.onSort && xl.onSort(onRowSorted); } catch (e) { console.warn("Tri :", e); } }
    updateWinBtn();
    refreshDocUI();
    try {
      if (IS_WINDOW) {
        await initWindow();
      } else if (inOffice) {
        status("Lecture des snips et documents du classeur…", "saving");
        const docs = await xl.listDocs();
        for (const d of docs) S.docs.set(d.id, { ...d, kind: kindOf(d.name, d.mime), bytes: null, pages: {}, stored: true });
        await reloadSnips();
        refreshDocUI();
        if (docs.length) await openDoc(docs[0].id);
        status(docs.length || S.snips.size ? `Prêt — ${docs.length} document(s), ${S.snips.size} snip(s) dans ce classeur.` : "Prêt. Importez un document pour commencer.", "ok");
      } else {
        status("Mode démo : ouvert hors d'Excel. Les valeurs vont dans la mini-feuille en bas.", "warn");
      }
    } catch (e) { reportError(e, () => location.reload(), "Démarrage incomplet"); }
    render();
    if (!IS_WINDOW) { checkUpdate(); setInterval(checkUpdate, 30 * 60 * 1000); }
  }

  // exposé pour les tests automatisés
  window.SnipCompta = { S, switchTab, checkUpdate, cellHighlight, gotoHit, clearHits, onSelectionChanged, runAction, status, withRetry, persistSnips, applyTheme, targetAt, tableRectAt, friendly, setZoom, rpcSend, rpcRecv, hubHandle, windowRecv, RemoteXL, setRpcChunk: (n) => { RPC_CH = n; }, setTransport: (toHub, toWin) => { parentSendRaw = toHub; winSendRaw = toWin; }, setHubXL: (x) => { HUBXL = x; }, initWindow, getActions, registerRibbon, findDocSums, formAddField, formPreview, formExport, runDocExtraction, extractInvoice, extractReceipt, compareVersions, exportComparison, runRowMatch, prepareRowMatch, annotatedPdf, removePages, compressAll, applyRedactions, addMarkup, transferVersion, realign, linkSnipToSelection, convertSnip, attachEvidence, expandFiles, textToPdf, searchTerms, autoSearchCell, findSums, exportSnips, deleteSnipsInSelection, deleteAllSnips, insertComment, sumRuns, typed, parseAmount, findAmounts, groupLines, autoColumns, extractTable, wordsInRect, DemoXL, setMode, handleRect, insertSum, insertTable, importFiles, doSearch, runMatch, pageText, rotateDoc, resizeSnip, clipWords, paintPage, cleanForOcr, pageSkew, detectSkew };

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady((info) => init(!!info && info.host === Office.HostType.Excel));
  } else {
    document.addEventListener("DOMContentLoaded", () => init(false));
    if (document.readyState !== "loading") init(false);
  }
})();
