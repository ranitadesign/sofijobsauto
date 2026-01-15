// server.js
// ✅ Genera PDF desde PPTX con Docxtemplater + ImageModule
// ✅ FOTO: en el PPTX usar placeholder EXACTO: {{%photo}}
// ✅ Texto: usar {{name}}, {{title}}, etc.
// ✅ Acepta foto por: photo_base64 (recomendado) o photo_url
// ✅ Nunca "undefined": nullGetter + normalización
// ✅ Exp bullets hasta 5 por experiencia (ahora hasta 7 experiencias)
// ✅ Variantes por plantilla según cantidad de experiencias:
//    - AHORA soporta v1, v2, v3, v4, v5 (elige según expCount)
// ✅ Color: reemplaza SENTINEL_HEX (c0504d) por el color elegido en TODOS los XML
// ✅ Texto sidebar auto (blanco/negro): reemplaza TEXT_SENTINEL_HEX (00FFFF por default)
//
// ✅ FIX FOTO DEFINITIVO (tu caso):
// - Si Gemini devuelve un PNG con fondo blanco pegado, hacemos keying de fondo y alpha
// - Luego recortamos a círculo con alpha afuera
// - Nunca flatten en la foto final

const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { execFile } = require("child_process");
const http = require("http");
const https = require("https");

const express = require("express");
const cors = require("cors");
const sharp = require("sharp");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");

const app = express();
app.use(cors());
app.use(express.json({ limit: "60mb" }));

app.get("/health", (_, res) => res.json({ ok: true }));

/**
 * LibreOffice / soffice:
 * - Windows: ruta típica
 * - Linux (Railway/Render): normalmente "soffice"
 */
const DEFAULT_SOFFICE =
  process.platform === "win32"
    ? "C:\\ Program Files\\LibreOffice\\program\\soffice.exe".replace("C:\\ ", "C:\\")
    : "soffice";

const SOFFICE_PATH = process.env.SOFFICE_PATH || DEFAULT_SOFFICE;

// Carpeta de plantillas PPTX
const TEMPLATES_DIR = path.join(__dirname, "templates");

// Template default (si no mandás template_id)
const DEFAULT_TEMPLATE_ID = process.env.DEFAULT_TEMPLATE_ID || "1";

/**
 * ✅ SENTINELA de color de FONDO (sidebar / acentos)
 * IMPORTANTE: en LibreOffice poné EXACTAMENTE este color en sidebar/títulos/lineas/etc.
 */
const SENTINEL_HEX = (process.env.SENTINEL_HEX || "c0504d")
  .replace("#", "")
  .toUpperCase();

/**
 * ✅ SENTINELA de color para TEXTO del sidebar (auto blanco/negro)
 * En tus PPTX: poné este color a TODO el texto del sidebar que quieras auto-contraste
 */
const TEXT_SENTINEL_HEX = (process.env.TEXT_SENTINEL_HEX || "543F3F")
  .replace("#", "")
  .toUpperCase();

const DEBUG_COLOR = (process.env.DEBUG_COLOR || "").trim() === "1";

/* =========================================================================
   1) LIMITES GLOBALES (modo “no cortar por código”)
   ========================================================================= */

const LIMITS = {
  ENABLE_CLAMP: false, // ✅ en false: NO recorta

  NAME_MAX_CHARS: 200,
  TITLE_MAX_CHARS: 240,
  ABOUT_MAX_CHARS_DEFAULT: 6000,

  CONTACT_EMAIL_MAX: 180,
  CONTACT_PHONE_MAX: 80,
  CONTACT_LOCATION_MAX: 240,
  CONTACT_WEBSITE_MAX: 220,

  EXP_ROLE_MAX: 400,
  EXP_COMPANY_MAX: 400,
  EXP_DATES_MAX: 120,
  EXP_BULLET_MAX: 220,

  EDU_SCHOOL_MAX: 380,
  EDU_DEGREE_MAX: 380,
  EDU_YEARS_MAX: 120,

  SKILL_MAX: 200,
  ITEM_MAX: 260, // idiomas/it/cursos
};

function safeStr(v) {
  if (v === null || v === undefined) return "";
  const s = String(v);
  if (s.toLowerCase() === "undefined" || s.toLowerCase() === "null") return "";
  return s.normalize("NFC");
}

function maybeClampPlain(s, maxChars) {
  const text = safeStr(s).replace(/\s+/g, " ").trim();
  if (!text) return "";
  if (!LIMITS.ENABLE_CLAMP) return text;
  if (!maxChars || maxChars <= 0) return text;
  return text.length > maxChars
    ? text.slice(0, Math.max(0, maxChars - 1)).trimEnd() + "…"
    : text;
}

function getAny(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

function splitByCommonDelimiters(s) {
  const raw = safeStr(s).trim();
  if (!raw) return [];
  return raw
    .split(/[\n,;•]+/g)
    .map((x) => x.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function buildBulletsBlock(data, expIndex) {
  const arr = [];
  for (let b = 1; b <= 5; b++) {
    const v = safeStr(data[`exp_${expIndex}_b${b}`]).trim();
    if (v) arr.push("• " + v);
  }
  return arr.join("\n");
}


/* =========================================================================
   1.1) SIDEBAR (TÍTULOS + CUERPO) — sin huecos, loop por sección
   ========================================================================= */

function nonEmptyLines(arr) {
  return (arr || []).map((x) => safeStr(x).trim()).filter(Boolean);
}

// 🔧 Mantiene "gaps" (líneas vacías) cuando vos las insertás a propósito
function linesPreserveGaps(arr) {
  return (arr || []).map((x) => safeStr(x).replace(/\r/g, "").trimEnd());
}

// ✅ Solo la primera letra en mayúscula, NO rompe acentos, NO fuerza minúsculas
function capitalizeFirst(s) {
  const t = safeStr(s).trim();
  if (!t) return "";
  return t.charAt(0).toUpperCase() + t.slice(1);
}

function asBulletLines(items) {
  return nonEmptyLines(items).map((x) => "• " + x);
}

/**
 * Educación: devuelve un array de líneas donde cada educación queda en un "bloque"
 * y entre bloques mete N líneas vacías (gap).
 */
function buildEducationLines(data, { gapLines = 1 } = {}) {
  const out = [];
  const gap = Math.max(0, Number(gapLines || 0));

  for (let i = 1; i <= 3; i++) {
    const degree = safeStr(data[`edu_${i}_degree`]).trim();
    const school = safeStr(data[`edu_${i}_school`]).trim();
    const years = safeStr(data[`edu_${i}_years`]).trim();

    const parts = [degree, school].filter(Boolean).join(" — ");
    const full = [parts, years].filter(Boolean).join(" | ").trim();

    if (!full) continue;

    if (out.length) {
      for (let g = 0; g < gap; g++) out.push(""); // gaps reales
    }

    out.push(full);
  }

  return out;
}

/**
 * Devuelve un ARRAY de secciones para loop en PPT:
 * sidebar_sections = [{title,line,body}, ...]
 *
 * IMPORTANTE:
 * - title: "Educación" (no mayúsculas)
 * - line: "──────────..." (una sola línea, NO underline tipográfico)
 * - body: texto con saltos (bullets incluidos)
 */
function buildSidebarSections(
  data,
  {
    underlineMin = 20,     // largo mínimo de la línea
    underlineExtra = 20,   // extra para que “pase” el título
    educationGapLines = 1, // separación entre educaciones
  } = {}
) {
  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 20), t.length + Number(underlineExtra || 20));
    return "─".repeat(len);
  }

  // ✅ Sidebar alternativo (SIN Educación)
data.sidebar_sections_noedu = buildSidebarSectionsNoEdu(data, {
  underlineMin: 20,
  underlineExtra: 20,
});

  const sections = [];

  function pushSection(title, lines, { preserveGaps = false } = {}) {
    const t = capitalizeFirst(title);

    if (preserveGaps) {
      // ✅ mantiene líneas vacías intencionales
      const raw = linesPreserveGaps(lines);

      // elimina vacíos al inicio/fin, pero conserva los del medio
      while (raw.length && !safeStr(raw[0]).trim()) raw.shift();
      while (raw.length && !safeStr(raw[raw.length - 1]).trim()) raw.pop();

      // si quedó todo vacío, no agrega sección
      if (!raw.some((x) => safeStr(x).trim())) return;

      sections.push({
        title: t,
        line: underlineForTitle(t),
        body: raw.join("\n"),
      });
      return;
    }

    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    sections.push({
      title: t,
      line: underlineForTitle(t),
      body: clean.join("\n"),
    });
  }

  // Educación (sin bullets) ✅ preserva gaps
  pushSection(
    "Educación",
    buildEducationLines(data, { gapLines: educationGapLines }),
    { preserveGaps: true }
  );

  // Cursos (bullets)
  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos", asBulletLines(cursos));

  // Informática (bullets)
  const it = [];
  for (let i = 1; i <= 6; i++) it.push(data[`it_${i}`]);
  pushSection("Informática", asBulletLines(it));

  // Idiomas (bullets)
  const idiomas = [];
  for (let i = 1; i <= 3; i++) idiomas.push(data[`idioma_${i}`]);
  pushSection("Idiomas", asBulletLines(idiomas));

  // Competencias (bullets)
  const skills = [];
  for (let i = 1; i <= 7; i++) skills.push(data[`skill_${i}`]);
  pushSection("Competencias", asBulletLines(skills));

  return sections;
}

// ✅ Variante: sidebar sin Educación/Formación (para plantillas que ya la tienen en el cuerpo)
function buildSidebarSectionsNoEdu(data, opts = {}) {
  // Reutiliza tu lógica pero salteando Educación
  const {
    underlineMin = 20,
    underlineExtra = 20,
  } = opts;

  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 20), t.length + Number(underlineExtra || 20));
    return "─".repeat(len);
  }

  const sections = [];

  function pushSection(title, lines) {
    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    const t = capitalizeFirst(title);

    sections.push({
      title: t,
      line: underlineForTitle(t),
      body: clean.join("\n"),
    });
  }

  // ❌ NO Educación

  // Cursos (bullets)
  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos", asBulletLines(cursos));

  // Informática (bullets)
  const it = [];
  for (let i = 1; i <= 6; i++) it.push(data[`it_${i}`]);
  pushSection("Informática", asBulletLines(it));

  // Idiomas (bullets)
  const idiomas = [];
  for (let i = 1; i <= 3; i++) idiomas.push(data[`idioma_${i}`]);
  pushSection("Idiomas", asBulletLines(idiomas));

  // Competencias (bullets)
  const skills = [];
  for (let i = 1; i <= 7; i++) skills.push(data[`skill_${i}`]);
  pushSection("Competencias", asBulletLines(skills));

  return sections;
}



/* =========================================================================
   2) PERFIL POR PLANTILLA (ABOUT + FOTO)
   ========================================================================= */

const TEMPLATE_PROFILES = {
  1: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  2: { about: { maxChars: 6000 }, photoSize: [420, 420] },
  3: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  4: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  5: { about: { maxChars: 6000 }, photoSize: [210, 210] },
  6: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  7: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  8: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  9: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  10: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  11: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  12: { about: { maxChars: 6000 }, photoSize: [420, 420] },
  13: { about: { maxChars: 6000 }, photoSize: [415, 415] },
  14: { about: { maxChars: 6000 }, photoSize: [520, 520] },
};

function getProfile(templateId) {
  const id = Number(templateId);
  return TEMPLATE_PROFILES[id] || TEMPLATE_PROFILES[1];
}

/* =========================================================================
   3) COLOR: parse del primer color + mapping pastel + aplicar REEMPLAZO XML
   ========================================================================= */

// Paleta pastel razonable (evita chillones)
const COLOR_MAP = {
  "violeta clarito": "E5E0EA",
  "violeta claro": "D2C9DB",
  violeta: "7E6597",
  lila: "E5E0EA",
  lavanda: "D2C9DB",
  "gama violeta": "D2C9DB",

  "rosa clarito": "F6E8ED",
  "rosa claro": "F6E8ED",
  "rosa pastel": "F6E8ED",
  "rosa pálido": "F6E8ED",
  "rosa palido": "F6E8ED",
  "rosa claro/pastel": "F6E8ED",
  "rosa oscuro": "D79DB3",

  "celeste clarito": "D5DEE6",
  "celeste claro": "608ABF",
  "celeste claro estilo pastel": "D5DEE6",
  "azul clarito": "D5DEE6",
  "azul claro": "608ABF",
  "azul suave": "D5DEE6",
  "gama de los azules": "323B4C",
  "gama azul": "002D6A",
  azules: "002D6A",
  azul: "002D6A",
  "azules - grises": "323B4C",
  "azul 0092ff": "608ABF",

  "verde clarito": "D2E0E1",
  "verde claro": "D2E0E1",
  "verde pastel": "D2E0E1",
  "verde sobrio": "44867B",
  "verde intermedio": "44867B",
  "verde oscuro": "2B554E",

  "gris clarito": "C7C8CA",
  "gris claro": "C7C8CA",
  gris: "696969",
  "azules - grises ": "323B4C",
  negro: "062446",
  "negro - gris - azul": "062446",
  "negro gris azul": "062446",

  beige: "B8A797",
  baige: "B8A797",
  "beige claro": "B8A797",
  "baige claro": "B8A797",
  crema: "D5DEE6",
  ocre: "B8A797",
  "tonos naranja": "A62C46",
  naranja: "A62C46",
  "naranja pastel": "F6E8ED",
  marron: "323B4C",
  marrones: "323B4C",

  "azul oscuro": "062446",
  "azul medio": "002D6A",
  "azul grisaceo": "323B4C",
  "gris oscuro": "696969",
  rojo: "A62C46",
};

function extractColorKeyword(raw) {
  const s = safeStr(raw).toLowerCase();
  const checks = [
    { k: "naranja", v: "naranja" },
    { k: "ocre", v: "ocre" },
    { k: "marron", v: "marron" },
    { k: "marrón", v: "marron" },
    { k: "beige", v: "beige" },
    { k: "baige", v: "baige" },
    { k: "crema", v: "crema" },
    { k: "rosa", v: "rosa claro" },
    { k: "celeste", v: "celeste claro" },
    { k: "azul", v: "azul" },
    { k: "verde", v: "verde sobrio" },
    { k: "gris", v: "gris" },
    { k: "negro", v: "negro" },
    { k: "pastel", v: "celeste claro" },
  ];
  for (const c of checks) {
    if (s.includes(c.k)) return c.v;
  }
  return "";
}

function pickFirstColorRaw(coloresRaw) {
  const raw = safeStr(coloresRaw).trim();
  if (!raw) return "";
  const mHex = raw.match(/#?([0-9A-Fa-f]{6})/);
  if (mHex?.[1]) return mHex[1].toUpperCase();
  const first =
    raw
      .split(/\s+(?:o|or)\s+|[\/,;\n-]+/i)
      .map((x) => x.trim())
      .filter(Boolean)[0] || "";
  return first;
}

function normalizeHexColor(s) {
  const t = safeStr(s).trim();
  if (!t) return "";
  const m = t.match(/^#?([0-9A-Fa-f]{6})$/);
  return m ? m[1].toUpperCase() : "";
}

function resolveAccentHex(coloresRawOrFirst) {
  const first = pickFirstColorRaw(coloresRawOrFirst);
  const hex = normalizeHexColor(first);
  if (hex) return hex;

  const lowFirst = safeStr(first).toLowerCase();
  if (lowFirst.includes("azul") || lowFirst.includes("blue")) return "1E3A8A";
  if (lowFirst.includes("verde") || lowFirst.includes("green") || lowFirst.includes("teal"))
    return "0F766E";

  const key1 = lowFirst.replace(/\s+/g, " ").trim();
  if (COLOR_MAP[key1]) return COLOR_MAP[key1];

  const kw = extractColorKeyword(coloresRawOrFirst);
  if (kw) {
    const key2 = safeStr(kw).toLowerCase().replace(/\s+/g, " ").trim();
    if (COLOR_MAP[key2]) return COLOR_MAP[key2];
    if (key2.includes("azul")) return "1E3A8A";
    if (key2.includes("verde")) return "0F766E";
  }
  return "";
}

function hexToRgb(hex6) {
  const h = hex6.replace("#", "").toUpperCase();
  return {
    r: parseInt(h.slice(0, 2), 16),
    g: parseInt(h.slice(2, 4), 16),
    b: parseInt(h.slice(4, 6), 16),
  };
}

function relLuminance({ r, g, b }) {
  const srgb = [r, g, b]
    .map((v) => v / 255)
    .map((v) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)));
  return 0.2126 * srgb[0] + 0.7152 * srgb[1] + 0.0722 * srgb[2];
}

function rgbToHex({ r, g, b }) {
  const to2 = (n) => n.toString(16).padStart(2, "0").toUpperCase();
  return `${to2(r)}${to2(g)}${to2(b)}`;
}

function clamp(n, a, b) {
  return Math.max(a, Math.min(b, n));
}

function darken(hex6, amount) {
  const { r, g, b } = hexToRgb(hex6);
  const rr = clamp(Math.round(r * (1 - amount)), 0, 255);
  const gg = clamp(Math.round(g * (1 - amount)), 0, 255);
  const bb = clamp(Math.round(b * (1 - amount)), 0, 255);
  return rgbToHex({ r: rr, g: gg, b: bb });
}

function pickSidebarColorForWhiteText(hex6) {
  const L = relLuminance(hexToRgb(hex6));
  if (L > 0.6) {
    for (const amt of [0.18, 0.28, 0.38, 0.48, 0.58]) {
      const out = darken(hex6, amt);
      const L2 = relLuminance(hexToRgb(out));
      if (L2 <= 0.5) return out;
    }
    return darken(hex6, 0.6);
  }
  return hex6;
}

function rgbToHsv({ r, g, b }) {
  const rr = r / 255,
    gg = g / 255,
    bb = b / 255;
  const max = Math.max(rr, gg, bb),
    min = Math.min(rr, gg, bb);
  const d = max - min;

  let h = 0;
  if (d !== 0) {
    if (max === rr) h = ((gg - bb) / d) % 6;
    else if (max === gg) h = (bb - rr) / d + 2;
    else h = (rr - gg) / d + 4;
    h *= 60;
    if (h < 0) h += 360;
  }
  const s = max === 0 ? 0 : d / max;
  const v = max;
  return { h, s, v };
}

function shouldForceWhiteText(bgHex6, accentRaw) {
  const raw = safeStr(accentRaw).toLowerCase();
  if (raw.includes("naranja") || raw.includes("ocre") || raw.includes("marron") || raw.includes("marrón")) {
    return true;
  }
  const hsv = rgbToHsv(hexToRgb(bgHex6));
  if (hsv.s >= 0.22 && hsv.v >= 0.3 && hsv.h >= 15 && hsv.h <= 55) return true;
  return false;
}

function pickTextColorForSidebar(bgHex6, accentRaw = "") {
  if (shouldForceWhiteText(bgHex6, accentRaw)) return "FFFFFF";
  const L = relLuminance(hexToRgb(bgHex6));
  const contrastWhite = 1.05 / (L + 0.05);
  const contrastBlack = (L + 0.05) / 0.05;
  return contrastBlack >= contrastWhite ? "000000" : "FFFFFF";
}

function replaceColorInAllXml(pptxBuffer, fromHex6, toHex6) {
  const FROM = normalizeHexColor(fromHex6);
  const TO = normalizeHexColor(toHex6);
  if (!FROM || !TO) return { buffer: pptxBuffer, touchedFiles: 0, replacements: 0 };

  const zip = new PizZip(pptxBuffer);
  const files = zip.file(/\.xml$/) || [];

  let touchedFiles = 0;
  let replacements = 0;
  const touchedNames = [];

  const attrNames = ["val", "rgb", "lastClr", "fill", "color", "bgColor", "fgColor"];
  const attrGroup = attrNames.join("|");

  const reAttr = new RegExp(`(${attrGroup})="${FROM}"`, "gi");
  const reBare = new RegExp(`>${FROM}<`, "gi");

  for (const f of files) {
    const name = f.name;
    let xml = f.asText();

    const hits1 = (xml.match(reAttr) || []).length;
    const hits2 = (xml.match(reBare) || []).length;
    const hits = hits1 + hits2;
    if (!hits) continue;

    xml = xml.replace(reAttr, `$1="${TO}"`);
    xml = xml.replace(reBare, `>${TO}<`);

    zip.file(name, xml);

    touchedFiles += 1;
    replacements += hits;
    if (DEBUG_COLOR) touchedNames.push(name);
  }

  if (DEBUG_COLOR && touchedNames.length) {
    console.log(`[COLOR][DEBUG] touched: ${touchedNames.join(" | ")}`);
  }

  return {
    buffer: zip.generate({ type: "nodebuffer" }),
    touchedFiles,
    replacements,
  };
}

/* =========================================================================
   4) IMÁGENES (FIX: remover fondo blanco -> transparencia + círculo)
   ========================================================================= */

function decodeBase64Image(base64) {
  const s = safeStr(base64).trim();
  if (!s) return null;

  const m = s.match(/^data:(.+);base64,(.*)$/);
  const payload = m ? m[2] : s;

  try {
    return Buffer.from(payload, "base64");
  } catch {
    return null;
  }
}

function normalizeGoogleDriveUrl(url) {
  const u = safeStr(url).trim();
  if (!u) return "";

  const m1 = u.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1?.[1]) return `https://drive.google.com/uc?export=download&id=${m1[1]}`;

  const m2 = u.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (m2?.[1]) return `https://drive.google.com/uc?export=download&id=${m2[1]}`;

  const idMatch = u.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (u.includes("drive.google.com/uc") && idMatch?.[1]) {
    return `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;
  }

  return u;
}

function fetchBufferFromUrl(url) {
  return new Promise((resolve, reject) => {
    const finalUrl = normalizeGoogleDriveUrl(url);
    const lib = finalUrl.startsWith("https") ? https : http;

    const req = lib.get(
      finalUrl,
      { headers: { "User-Agent": "Mozilla/5.0 (CV-Generator)", Accept: "*/*" } },
      (resp) => {
        const code = resp.statusCode || 0;

        if (code >= 300 && code < 400 && resp.headers.location) {
          return resolve(fetchBufferFromUrl(resp.headers.location));
        }

        if (code !== 200) return reject(new Error(`No pude descargar imagen. HTTP ${code}`));

        const chunks = [];
        resp.on("data", (d) => chunks.push(d));
        resp.on("end", () => resolve(Buffer.concat(chunks)));
      }
    );

    req.on("error", reject);
  });
}

async function sampleCornerColor(buf, sampleSize = 8) {
  const { data, info } = await sharp(buf).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  const W = info.width,
    H = info.height;

  const pts = [
    { x0: 0, y0: 0 },
    { x0: Math.max(0, W - sampleSize), y0: 0 },
    { x0: 0, y0: Math.max(0, H - sampleSize) },
    { x0: Math.max(0, W - sampleSize), y0: Math.max(0, H - sampleSize) },
  ];

  let rSum = 0,
    gSum = 0,
    bSum = 0,
    n = 0;

  for (const p of pts) {
    for (let y = p.y0; y < Math.min(H, p.y0 + sampleSize); y++) {
      for (let x = p.x0; x < Math.min(W, p.x0 + sampleSize); x++) {
        const idx = (y * W + x) * 4;
        const a = data[idx + 3];
        if (a < 10) continue;
        rSum += data[idx];
        gSum += data[idx + 1];
        bSum += data[idx + 2];
        n++;
      }
    }
  }

  if (!n) return { r: 255, g: 255, b: 255 };
  return {
    r: Math.round(rSum / n),
    g: Math.round(gSum / n),
    b: Math.round(bSum / n),
  };
}

function colorDist(c1, c2) {
  const dr = c1.r - c2.r;
  const dg = c1.g - c2.g;
  const db = c1.b - c2.b;
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

async function keyOutBackgroundToAlpha(
  inputBuf,
  { backgroundColor, threshold = 55, softness = 0.35 } = {}
) {
  const bg = backgroundColor || (await sampleCornerColor(inputBuf, 8));

  const { data, info } = await sharp(inputBuf).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  const out = Buffer.from(data);

  const t = Number(threshold || 55);
  const soft = Math.max(0, Math.min(1, Number(softness || 0.35)));

  const t0 = t * (1 - soft);
  const t1 = t;

  for (let i = 0; i < out.length; i += 4) {
    const px = { r: out[i], g: out[i + 1], b: out[i + 2] };
    const d = colorDist(px, bg);

    if (d <= t0) {
      out[i + 3] = 0;
    } else if (d < t1) {
      const k = (d - t0) / Math.max(1e-6, t1 - t0);
      const alpha = out[i + 3];
      out[i + 3] = Math.round(alpha * k);
    }
  }

  return sharp(out, { raw: { width: info.width, height: info.height, channels: 4 } })
    .png()
    .toBuffer();
}

async function forceCircleTransparentOutside(inputBuf, { width, height, padding = 0 } = {}) {
  const W = Number(width || 520);
  const H = Number(height || 520);

  const size = Math.min(W, H);
  const r = Math.max(1, Math.floor(size / 2) - Number(padding || 0));
  const cx = Math.floor(W / 2);
  const cy = Math.floor(H / 2);

  const svgMask = `
  <svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="black"/>
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="white"/>
  </svg>`;

  return await sharp(inputBuf)
    .resize(W, H, { fit: "cover", position: "centre" })
    .ensureAlpha()
    .composite([{ input: Buffer.from(svgMask), blend: "dest-in" }])
    .png()
    .toBuffer();
}

async function buildFinalPhotoPng(photoBuf, { W, H } = {}) {
  const w = Number(W || 520);
  const h = Number(H || 520);

  let base = await sharp(photoBuf).resize(w, h, { fit: "cover", position: "centre" }).png().toBuffer();

  const bg = await sampleCornerColor(base, 10);
  const bgIsLight = (bg.r + bg.g + bg.b) / 3 >= 210;

  if (bgIsLight) {
    base = await keyOutBackgroundToAlpha(base, {
      backgroundColor: bg,
      threshold: 35,
      softness: 0.4,
    });
  }

  const circ = await forceCircleTransparentOutside(base, { width: w, height: h, padding: 0 });
  return circ;
}

/* =========================================================================
   5) TEMPLATES (MAPPING 1..14) + VARIANTES (v1..v5)
   ========================================================================= */

/**
 * ✅ Ahora podés tener (por template_id) hasta 5 variantes:
 *   v1, v2, v3, v4, v5
 *
 * ✅ Selección automática según expCount:
 *   expCount <= 1 => v1
 *   expCount == 2 => v2
 *   expCount == 3 => v3
 *   expCount == 4 => v4
 *   expCount >= 5 => v5
 *
 * ✅ Si falta un archivo (no existe esa variante), hace fallback “al más cercano disponible”.
 *
 * IMPORTANT: podés ir creando archivos de a poco.
 * Si hoy solo tenés v3 y v5, funciona igual.
 */
const TEMPLATE_VARIANTS = {
  1: { v1: "Plantilla_oficial_1_verde_v1.pptx",  v2: "Plantilla_oficial_1_verde_v2.pptx",  v3: "Plantilla_oficial_1_verde_v3.pptx",  v4: "Plantilla_oficial_1_verde_v4.pptx",  v5: "Plantilla_oficial_1_verde_v5.pptx" },
  2: { v1: "Plantilla_oficial_2_v1.pptx",  v2: "Plantilla_oficial_2_v2.pptx",  v3: "Plantilla_oficial_2_v3.pptx",  v4: "Plantilla_oficial_2_v4.pptx",  v5: "Plantilla_oficial_2_v5.pptx" },
  3: { v1: "Plantilla_oficial_3_v1.pptx",  v2: "Plantilla_oficial_3_v2.pptx",  v3: "Plantilla_oficial_3_v3.pptx",  v4: "Plantilla_oficial_3_v4.pptx",  v5: "Plantilla_oficial_3_v5.pptx" },
  4: { v1: "Plantilla_oficial_4_v1.pptx",  v2: "Plantilla_oficial_4_v2.pptx",  v3: "Plantilla_oficial_4_v3.pptx",  v4: "Plantilla_oficial_4_v4.pptx",  v5: "Plantilla_oficial_4_v5.pptx" },
  5: { v1: "Plantilla_oficial_5_v1.pptx",  v2: "Plantilla_oficial_5_v2.pptx",  v3: "Plantilla_oficial_5_v3.pptx",  v4: "Plantilla_oficial_5_v4.pptx",  v5: "Plantilla_oficial_5_v5.pptx" },
  6: { v1: "Plantilla_oficial_6_v1.pptx",  v2: "Plantilla_oficial_6_v2.pptx",  v3: "Plantilla_oficial_6_v3.pptx",  v4: "Plantilla_oficial_6_v4.pptx",  v5: "Plantilla_oficial_6_v5.pptx" },
  7: { v1: "Plantilla_oficial_7_v1.pptx",  v2: "Plantilla_oficial_7_v2.pptx",  v3: "Plantilla_oficial_7_v3.pptx",  v4: "Plantilla_oficial_7_v4.pptx",  v5: "Plantilla_oficial_7_v5.pptx" },
  8: { v1: "Plantilla_oficial_8_v1.pptx",  v2: "Plantilla_oficial_8_v2.pptx",  v3: "Plantilla_oficial_8_v3.pptx",  v4: "Plantilla_oficial_8_v4.pptx",  v5: "Plantilla_oficial_8_v5.pptx" },
  9: { v1: "Plantilla_oficial_9_v1.pptx",  v2: "Plantilla_oficial_9_v2.pptx",  v3: "Plantilla_oficial_9_v3.pptx",  v4: "Plantilla_oficial_9_v4.pptx",  v5: "Plantilla_oficial_9_v5.pptx" },
  10:{ v1: "Plantilla_oficial_10_v1.pptx", v2: "Plantilla_oficial_10_v2.pptx", v3: "Plantilla_oficial_10_v3.pptx", v4: "Plantilla_oficial_10_v4.pptx", v5: "Plantilla_oficial_10_v5.pptx" },
  11:{ v1: "Plantilla_oficial_11_v1.pptx", v2: "Plantilla_oficial_11_v2.pptx", v3: "Plantilla_oficial_11_v3.pptx", v4: "Plantilla_oficial_11_v4.pptx", v5: "Plantilla_oficial_11_v5.pptx" },
  12:{ v1: "Plantilla_oficial_12_v1.pptx", v2: "Plantilla_oficial_12_v2.pptx", v3: "Plantilla_oficial_12_v3.pptx", v4: "Plantilla_oficial_12_v4.pptx", v5: "Plantilla_oficial_12_v5.pptx" },
  13:{ v1: "Plantilla_oficial_13_v1.pptx", v2: "Plantilla_oficial_13_v2.pptx", v3: "Plantilla_oficial_13_v3.pptx", v4: "Plantilla_oficial_13_v4.pptx", v5: "Plantilla_oficial_13_v5.pptx" },
  14:{ v1: "Plantilla_oficial_14_v1.pptx", v2: "Plantilla_oficial_14_v2.pptx", v3: "Plantilla_oficial_14_v3.pptx", v4: "Plantilla_oficial_14_v4.pptx", v5: "Plantilla_oficial_14_v5.pptx" },
};


// Fallback “sin variantes” (compatibilidad)
const TEMPLATE_MAP = Object.fromEntries(Object.entries(TEMPLATE_VARIANTS).map(([id, v]) => [Number(id), v.v3]));

function normalizeTemplateId(templateId) {
  const raw = (templateId || DEFAULT_TEMPLATE_ID || "").toString().trim();
  if (!raw) throw new Error("Falta template_id y no hay DEFAULT_TEMPLATE_ID");
  const id = Number(raw);
  if (!Number.isFinite(id) || id < 1 || id > 14) {
    throw new Error(`template_id inválido: "${raw}". Debe ser un número 1..14.`);
  }
  return id;
}

function countExperiencesFromSrc(src, max = 7) {
  let count = 0;
  for (let n = 1; n <= max; n++) {
    const dates = safeStr(src?.[`exp_${n}_dates`]).trim();
    const company = safeStr(src?.[`exp_${n}_company`]).trim();
    const role = safeStr(src?.[`exp_${n}_role`]).trim();
    const anyBullet =
      safeStr(src?.[`exp_${n}_b1`]).trim() ||
      safeStr(src?.[`exp_${n}_b2`]).trim() ||
      safeStr(src?.[`exp_${n}_b3`]).trim() ||
      safeStr(src?.[`exp_${n}_b4`]).trim() ||
      safeStr(src?.[`exp_${n}_b5`]).trim();

    if (dates || company || role || anyBullet) count++;
  }
  return count;
}

function wantedVariantKeyByExpCount(expCount) {
  if (!Number.isFinite(expCount) || expCount <= 0) return "v1";
  if (expCount <= 1) return "v1";
  if (expCount === 2) return "v2";
  if (expCount === 3) return "v3";
  if (expCount === 4) return "v4";
  return "v5";
}

// fallback “cercano”: intenta el pedido, si no existe baja/sube a lo disponible
function pickVariantFileName(templateId, expCount) {
  const variants = TEMPLATE_VARIANTS[templateId];
  if (!variants) return TEMPLATE_MAP[templateId];

  const order = ["v1", "v2", "v3", "v4", "v5"];
  const want = wantedVariantKeyByExpCount(expCount);
  const wantIdx = order.indexOf(want);

  // estrategia: primero intenta el “want”, luego busca cerca (hacia abajo y arriba)
  const tries = [];
  if (wantIdx >= 0) {
    tries.push(order[wantIdx]);
    for (let step = 1; step < order.length; step++) {
      const down = wantIdx - step;
      const up = wantIdx + step;
      if (down >= 0) tries.push(order[down]);
      if (up < order.length) tries.push(order[up]);
    }
  }

  // si no hay wantIdx (raro), probar de v3/v5
  if (!tries.length) tries.push("v3", "v5", "v4", "v2", "v1");

  for (const k of tries) {
    const fname = variants[k];
    if (!fname) continue;
    const p = path.join(TEMPLATES_DIR, fname);
    if (fs.existsSync(p)) return fname;
  }

  // último fallback: v3 si existe, si no, cualquier cosa que exista
  const v3 = variants.v3 && fs.existsSync(path.join(TEMPLATES_DIR, variants.v3)) ? variants.v3 : "";
  if (v3) return v3;

  // buscar cualquiera existente dentro de variants
  for (const k of Object.keys(variants)) {
    const fname = variants[k];
    if (fname && fs.existsSync(path.join(TEMPLATES_DIR, fname))) return fname;
  }

  return TEMPLATE_MAP[templateId];
}

function getTemplatePath(templateId, expCountForVariant = 0) {
  const id = normalizeTemplateId(templateId);

  const fileName =
    expCountForVariant >= 0
      ? pickVariantFileName(id, expCountForVariant)
      : TEMPLATE_MAP[id] || TEMPLATE_VARIANTS[id]?.v3;

  if (!fileName) throw new Error(`No hay mapping para template_id=${id}. Revisá TEMPLATE_VARIANTS.`);

  const templatePath = path.join(TEMPLATES_DIR, fileName);

  if (!fs.existsSync(templatePath)) {
    let list = [];
    try {
      list = fs.readdirSync(TEMPLATES_DIR).filter((f) => f.toLowerCase().endsWith(".pptx"));
    } catch (_) {}

    throw new Error(
      `No encuentro la plantilla: ${fileName} en ${TEMPLATES_DIR}. ` +
        `Disponibles: ${list.join(", ")}.`
    );
  }

  return templatePath;
}

/* =========================================================================
   6) BULLETS: backup de segmentación (por si la IA no separa)
   ========================================================================= */

function segmentBulletsIfNeeded(data, expIndex) {
  const b1 = safeStr(data[`exp_${expIndex}_b1`]).trim();
  if (!b1) return;

  const b2 = safeStr(data[`exp_${expIndex}_b2`]).trim();
  const b3 = safeStr(data[`exp_${expIndex}_b3`]).trim();
  const b4 = safeStr(data[`exp_${expIndex}_b4`]).trim();
  const b5 = safeStr(data[`exp_${expIndex}_b5`]).trim();

  const emptyRest = !b2 && !b3 && !b4 && !b5;
  const commaCount = (b1.match(/,/g) || []).length;

  if (!emptyRest || commaCount < 2) return;

  const parts = b1
    .split(/[,.;•\-]+/g)
    .map((x) => x.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  if (parts.length <= 1) return;

  data[`exp_${expIndex}_b1`] = parts[0] || "";
  data[`exp_${expIndex}_b2`] = parts[1] || "";
  data[`exp_${expIndex}_b3`] = parts[2] || "";
  data[`exp_${expIndex}_b4`] = parts[3] || "";
  data[`exp_${expIndex}_b5`] = parts[4] || "";
}

/* =========================================================================
   7) DATA MAPPING (hasta 7 experiencias) + sidebar_block
   ========================================================================= */

function flattenToTemplateData(body) {
  const src =
    body?.data && typeof body.data === "object"
      ? body.data
      : body?.fields && typeof body.fields === "object"
      ? body.fields
      : body || {};

  const data = {};

  // Base
  data.template_id = safeStr(getAny(src, ["template_id", "template", "templateId"]));
  data.photo_url = safeStr(getAny(src, ["photo_url", "photoUrl"]));
  data.photo_base64 = safeStr(getAny(src, ["photo_base64", "photoBase64", "photo"]));

  // Colores
  data.accent_color_raw = safeStr(getAny(src, ["accent_color_raw", "colores_raw", "colors_raw", "colores", "colors"], ""));

  data.name = maybeClampPlain(getAny(src, ["name", "nombre"]), LIMITS.NAME_MAX_CHARS);
  data.title = maybeClampPlain(getAny(src, ["title", "titulo"]), LIMITS.TITLE_MAX_CHARS);

  // ABOUT
  const templateId = Number(data.template_id || body.template_id || DEFAULT_TEMPLATE_ID || 1);
  const profile = getProfile(templateId);
  const aboutMax = profile?.about?.maxChars || LIMITS.ABOUT_MAX_CHARS_DEFAULT;
  data.about = maybeClampPlain(getAny(src, ["about", "objective", "objetivo"]), aboutMax);

  data.contact_phone = maybeClampPlain(getAny(src, ["contact_phone", "phone", "telefono"]), LIMITS.CONTACT_PHONE_MAX);
  data.contact_email = maybeClampPlain(getAny(src, ["contact_email", "email"]), LIMITS.CONTACT_EMAIL_MAX);
  data.contact_location = maybeClampPlain(getAny(src, ["contact_location", "location", "ubicacion"]), LIMITS.CONTACT_LOCATION_MAX);
  data.contact_website = maybeClampPlain(getAny(src, ["contact_website", "website", "web"]), LIMITS.CONTACT_WEBSITE_MAX);

  // Experiencia (hasta 7) con bullets hasta 5
  for (let n = 1; n <= 7; n++) {
    data[`exp_${n}_dates`] = maybeClampPlain(getAny(src, [`exp_${n}_dates`]), LIMITS.EXP_DATES_MAX);
    data[`exp_${n}_company`] = maybeClampPlain(getAny(src, [`exp_${n}_company`]), LIMITS.EXP_COMPANY_MAX);
    data[`exp_${n}_role`] = maybeClampPlain(getAny(src, [`exp_${n}_role`]), LIMITS.EXP_ROLE_MAX);

    for (let b = 1; b <= 5; b++) {
      data[`exp_${n}_b${b}`] = maybeClampPlain(getAny(src, [`exp_${n}_b${b}`]), LIMITS.EXP_BULLET_MAX);
    }

    segmentBulletsIfNeeded(data, n);
    data[`exp_${n}_bullets_block`] = buildBulletsBlock(data, n);
  }

  // Skills (7)
  for (let i = 1; i <= 7; i++) {
    data[`skill_${i}`] = maybeClampPlain(getAny(src, [`skill_${i}`]), LIMITS.SKILL_MAX);
  }

  // Educación (3)
  for (let i = 1; i <= 3; i++) {
    data[`edu_${i}_school`] = maybeClampPlain(getAny(src, [`edu_${i}_school`]), LIMITS.EDU_SCHOOL_MAX);
    data[`edu_${i}_degree`] = maybeClampPlain(getAny(src, [`edu_${i}_degree`]), LIMITS.EDU_DEGREE_MAX);
    data[`edu_${i}_years`] = maybeClampPlain(getAny(src, [`edu_${i}_years`]), LIMITS.EDU_YEARS_MAX);
  }

  // Idiomas / IT / Cursos (fallback)
  const idiomasRaw = safeStr(getAny(src, ["idiomas_raw", "idiomas", "languages_raw"], ""));
  const itRaw = safeStr(getAny(src, ["it_raw", "it", "informatica_raw"], ""));
  const cursosRaw = safeStr(getAny(src, ["cursos_raw", "cursos", "courses_raw"], ""));

  const idiomasParts = splitByCommonDelimiters(idiomasRaw);
  const itParts = splitByCommonDelimiters(itRaw);
  const cursoParts = splitByCommonDelimiters(cursosRaw);

  for (let i = 1; i <= 3; i++) {
    data[`idioma_${i}`] = maybeClampPlain(getAny(src, [`idioma_${i}`], idiomasParts[i - 1] || ""), LIMITS.ITEM_MAX);
  }
  for (let i = 1; i <= 6; i++) {
    data[`it_${i}`] = maybeClampPlain(getAny(src, [`it_${i}`], itParts[i - 1] || ""), LIMITS.ITEM_MAX);
  }
  for (let i = 1; i <= 6; i++) {
    data[`curso_${i}`] = maybeClampPlain(getAny(src, [`curso_${i}`], cursoParts[i - 1] || ""), LIMITS.ITEM_MAX);
  }




// ✅ Sidebar dinámico con loop (respeta tipografías)
// ✅ Sidebar dinámico con loop (respeta tipografías)
// ✅ Sidebar dinámico con loop (respeta tipografías)
data.sidebar_sections = buildSidebarSections(data, {
  underlineMin: 20,
  underlineExtra: 20,
  educationGapLines: 1,
});

// ✅ Sidebar alternativo (SIN Educación)
data.sidebar_sections_noedu = buildSidebarSectionsNoEdu(data, {
  underlineMin: 20,
  underlineExtra: 20,
});



  return data;
}






/* =========================================================================
   8) RENDER PPTX
   ========================================================================= */

function renderPptxFromTemplate(templateBuf, data) {
  const templateId = Number(data.template_id || DEFAULT_TEMPLATE_ID || 1);
  const profile = getProfile(templateId);
  console.log("[PHOTO][DEBUG] templateId=", templateId, "profile.photoSize=", profile?.photoSize);

  const zip = new PizZip(templateBuf);

  const imageModule = new ImageModule({
    centered: false,

    getImage: (tagValue, tagName) => {
      if (tagName !== "photo") return null;

      if (Buffer.isBuffer(tagValue)) return tagValue;

      if (typeof tagValue === "string" && tagValue.trim()) {
        const b = decodeBase64Image(tagValue);
        return b;
      }

      return null;
    },

    getSize: (img, tagValue, tagName) => {
      if (tagName !== "photo") return [0, 0];

      const ps = profile?.photoSize;
      const w = Array.isArray(ps) && Number.isFinite(ps[0]) ? ps[0] : 520;
      const h = Array.isArray(ps) && Number.isFinite(ps[1]) ? ps[1] : 520;
      return [w, h];
    },
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    modules: [imageModule],
    nullGetter: () => "",
  });

  try {
    doc.setData(data);
    doc.render();
  } catch (e) {
    throw new Error(`Docxtemplater render failed: ${e?.message || e}`);
  }

  let pptxBuf = doc.getZip().generate({ type: "nodebuffer" });

  // ✅ Aplicar color de sidebar (SENTINEL_HEX) + texto auto (TEXT_SENTINEL_HEX)
  let accentHex = resolveAccentHex(data.accent_color_raw);

  if (accentHex) {
    const adjusted = pickSidebarColorForWhiteText(accentHex);
    const sidebarHex = adjusted;

    // 1) Fondo/sidebar
    const r1 = replaceColorInAllXml(pptxBuf, SENTINEL_HEX, sidebarHex);
    pptxBuf = r1.buffer;

    // 2) Texto sidebar auto
    const textHex = pickTextColorForSidebar(sidebarHex, data.accent_color_raw);
    const r2 = replaceColorInAllXml(pptxBuf, TEXT_SENTINEL_HEX, textHex);
    pptxBuf = r2.buffer;
  }

  return pptxBuf;
}

/* =========================================================================
   9) PPTX -> PDF
   ========================================================================= */

function convertPptxToPdf(pptxPath, outDir) {
  return new Promise((resolve, reject) => {
    const args = [
      "--headless",
      "--nologo",
      "--nofirststartwizard",
      "--norestore",
      "--convert-to",
      "pdf",
      "--outdir",
      outDir,
      pptxPath,
    ];

    execFile(SOFFICE_PATH, args, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        return reject(
          new Error(
            `Error convirtiendo a PDF.\nsofficePath: ${SOFFICE_PATH}\nstderr: ${stderr}\nstdout: ${stdout}`
          )
        );
      }

      const pdfPath = pptxPath.replace(/\.pptx$/i, ".pdf");
      if (!fs.existsSync(pdfPath)) {
        return reject(new Error(`LibreOffice no generó el PDF esperado: ${pdfPath}`));
      }
      resolve(pdfPath);
    });
  });
}

/* =========================================================================
   10) ENDPOINT
   ========================================================================= */

app.post("/generate-pdf", async (req, res) => {
  try {
    const body = req.body || {};

    // Fuente real de campos
    const src =
      body?.data && typeof body.data === "object"
        ? body.data
        : body?.fields && typeof body.fields === "object"
        ? body.fields
        : body || {};

    const templateId = body.template_id || body.template || src.template_id || DEFAULT_TEMPLATE_ID;

    // ✅ Elegir variante por cantidad de experiencias detectadas en el input
    const expCount = countExperiencesFromSrc(src, 7);
    const templatePath = getTemplatePath(templateId, expCount);
    const templateBuf = fs.readFileSync(templatePath);

    const data = flattenToTemplateData(body);

    console.log(
      `[TEMPLATE] template_id=${normalizeTemplateId(templateId)} expCount=${expCount} file=${path.basename(templatePath)}`
    );

    // ✅ FOTO (FIX REAL)
    let photoBuf = null;

    if (data.photo_base64) {
      photoBuf = decodeBase64Image(data.photo_base64);
    } else if (data.photo_url) {
      photoBuf = await fetchBufferFromUrl(data.photo_url);
    }

    if (photoBuf && photoBuf.length) {
      try {
        const tId = Number(data.template_id || templateId || DEFAULT_TEMPLATE_ID || 1);
        const profile = getProfile(tId);
        const [W, H] = profile.photoSize || [520, 520];

        const finalPng = await buildFinalPhotoPng(photoBuf, { W, H });

        const meta = await sharp(finalPng).metadata();
        console.log(`[PHOTO] OK tId=${tId} w=${W} h=${H} hasAlpha=${!!meta.hasAlpha} format=${meta.format}`);

        data.photo = "data:image/png;base64," + finalPng.toString("base64");
      } catch (e) {
        console.warn("[PHOTO] processing failed, fallback:", e?.message || e);
        try {
          const fallbackPng = await sharp(photoBuf).png().toBuffer();
          data.photo = "data:image/png;base64," + fallbackPng.toString("base64");
        } catch (_) {
          data.photo = "";
        }
      }
    } else {
      data.photo = "";
    }

    const pptxBuf = renderPptxFromTemplate(templateBuf, data);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "cv-"));
    const id = crypto.randomBytes(8).toString("hex");
    const pptxPath = path.join(tmpDir, `cv-${id}.pptx`);
    fs.writeFileSync(pptxPath, pptxBuf);

    const pdfPath = await convertPptxToPdf(pptxPath, tmpDir);
    const pdfBuf = fs.readFileSync(pdfPath);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'attachment; filename="cv.pdf"');
    res.status(200).send(pdfBuf);
  } catch (err) {
    console.error(err);
    res.status(500).json({
      error: String(err?.message || err),
      stack: String(err?.stack || ""),
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CV API OK en http://127.0.0.1:${PORT}`);
  console.log(`Templates dir: ${TEMPLATES_DIR}`);
  console.log(`LibreOffice: ${SOFFICE_PATH}`);
  console.log(`ENABLE_CLAMP: ${LIMITS.ENABLE_CLAMP}`);
  console.log(`SENTINEL_HEX: ${SENTINEL_HEX}`);
  console.log(`TEXT_SENTINEL_HEX: ${TEXT_SENTINEL_HEX}`);
  console.log(`DEBUG_COLOR: ${DEBUG_COLOR}`);
});
