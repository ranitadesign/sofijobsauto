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
// ✅ Texto sidebar auto (blanco/negro): reemplaza TEXT_SENTINEL_HEX
//
// ✅ FIX FOTO DEFINITIVO (tu caso):
// - Si Gemini devuelve un PNG con fondo blanco pegado, hacemos keying de fondo y alpha
// - Luego recortamos a círculo con alpha afuera
// - Nunca flatten en la foto final
//
// ✅ FIX SIDEBAR (overflow):
// - Parchar shapes por marker en slides + layouts + masters
// - Soportar <a:bodyPr .../> self-closing
// - Forzar wrap=1 + horzOverflow="clip"
// - (Opcional) borrar el marker para que no se vea

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
const archiver = require("archiver");

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

const ICONS_DIR = path.join(__dirname, "assets", "icons");

// cache simple para no leer disco todo el tiempo
const ICON_CACHE = new Map();

function loadIconBuffer(theme, fileName) {
  const t = theme === "light" ? "light" : "dark";
  const key = `${t}/${fileName}`;

  if (ICON_CACHE.has(key)) return ICON_CACHE.get(key);

  const p = path.join(ICONS_DIR, t, fileName);
  const buf = fs.readFileSync(p);
  ICON_CACHE.set(key, buf);
  return buf;
}

// Template default (si no mandás template_id)
const DEFAULT_TEMPLATE_ID = process.env.DEFAULT_TEMPLATE_ID || "1";

/**
 * ✅ SENTINELA de color de FONDO (sidebar / acentos)
 * IMPORTANTE: en LibreOffice poné EXACTAMENTE este color en sidebar/títulos/lineas/etc.
 */
const SENTINEL_HEX = (process.env.SENTINEL_HEX || "c0504d").replace("#", "").toUpperCase();

/**
 * ✅ SENTINELA de color para TEXTO del sidebar (auto blanco/negro)
 * En tus PPTX: poné este color a TODO el texto del sidebar que quieras auto-contraste
 */
const TEXT_SENTINEL_HEX = (process.env.TEXT_SENTINEL_HEX || "543F3F").replace("#", "").toUpperCase();

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

function normalizeProficiencyParen(s) {
  let t = safeStr(s).trim();
  if (!t) return "";

  // incluye "promedio" como nivel
  t = t.replace(/\(\s*(avanzado|intermedio|promedio|basico|básico|alto)\s*\)/gi, (m, w) => `(${w.toLowerCase()})`);
  t = t.replace(/\b(Avanzado|Intermedio|Promedio|Basico|Básico|alto)\b/g, (w) => w.toLowerCase());

  return t;
}


function normalizeStatusWords(s, { style = "lower" } = {}) {
  let t = safeStr(s).trim();
  if (!t) return "";

  const mapLower = (w) => w.toLowerCase();

  // Normaliza variantes
  t = t.replace(/\b(en\s*curso)\b/gi, "en curso");
  t = t.replace(/\b(finalizado|finalizada)\b/gi, (m) => m.toLowerCase());
  t = t.replace(/\b(incompleto|incompleta)\b/gi, (m) => m.toLowerCase());

  if (style === "title") {
    // Si alguna vez querés Title Case en otro lado:
    t = t.replace(/\ben curso\b/g, "En curso");
    t = t.replace(/\bfinalizado\b/g, "Finalizado");
    t = t.replace(/\bfinalizada\b/g, "Finalizada");
    t = t.replace(/\bincompleto\b/g, "Incompleto");
    t = t.replace(/\bincompleta\b/g, "Incompleta");
  }

  return t;
}


function normalizeNivelWord(s) {
  let t = safeStr(s).trim();
  if (!t) return "";
  return t.replace(/\bnivel\b/gi, "nivel");
}


function normalizeRomanNumerals(s) {
  let t = safeStr(s);
  if (!t) return "";
  // i, ii, iii, iv, v, vi, vii, viii, ix, x => ROMAN
  return t.replace(/\b(i{1,3}|iv|v|vi{0,3}|ix|x)\b/gi, (m) => m.toUpperCase());
}


function normalizeCourseLine(s) {
  let t = safeStr(s).replace(/\r/g, "").trim();
  if (!t) return "";

  // Colapsa espacios
  t = t.replace(/\s+/g, " ");

  // 1) Extraer "(...)" final como fechas/estado si existe
  let dates = "";
  const mPar = t.match(/\(([^)]{2,})\)\s*$/);
  if (mPar) {
    dates = mPar[1].trim();
    t = t.slice(0, mPar.index).trim();
  }

  // 2) Split principal: preferimos "|"
  let title = "";
  let place = "";

  if (t.includes("|")) {
    const parts = t
      .split("|")
      .map((x) => safeStr(x).trim())
      .filter(Boolean);

    title = parts[0] || "";
    place = parts.slice(1).join(" "); // si venían muchos pipes, los juntamos
  } else {
    // Fallbacks comunes si no hay pipe:
    // "Título - Lugar", "Título — Lugar", "Título – Lugar"
    const mDash = t.match(/^(.+?)\s*(?:-|\u2013|\u2014)\s*(.+)$/);
    if (mDash) {
      title = mDash[1].trim();
      place = mDash[2].trim();
    } else {
      title = t.trim();
      place = "";
    }
  }

  // 3) Limpieza interna: NO dejamos pipes dentro de cada campo
  title = stripAllPipes(title);
  place = stripAllPipes(place);
  dates = stripAllPipes(dates);
    title = normalizeRomanNumerals(title);
  place = normalizeRomanNumerals(place);
  dates = normalizeRomanNumerals(dates);


  // 4) Normalizaciones (Power BI, Finalizado, etc.)
    // 4) Normalizaciones (tech + instituciones)
  title = normalizeTechNames(title);
  place = normalizeTechNames(place);
place = canonicalizeInstitutionRobust(place); // ✅ nuevo

  place = normalizeInstitutionField(place); // ✅ NUEVO


  // Fechas: normaliza estados (Finalizado/En curso) y paréntesis de nivel si viniera
  dates = normalizeStatusWords(dates);
  dates = normalizeProficiencyParen(dates);

  // 5) Reconstrucción: "Título | Lugar (Fechas)"
  let out = title;

  if (place) out += ` | ${place}`;

  if (dates) {
    // si dates ya trae paréntesis por algún motivo, no duplicamos
    const d = dates.trim();
    out += d.startsWith("(") ? ` ${d}` : ` (${d})`;
  }

  return out.trim();
}



function normalizeTechNames(s) {
  let t = safeStr(s).trim();
  if (!t) return "";

  t = normalizeStatusWords(t);
  t = normalizeProficiencyParen(t);

  const repl = [
    // ✅ UX/UI siempre así (cubre ux/ui y ui/ux, con cualquier mezcla de mayúsculas)
[/\bui\s*\/\s*ux\b/gi, "UX/UI"],
[/\bux\s*\/\s*ui\b/gi, "UX/UI"],
[/\bui\s*(?:\/|&|y|-|–|—|\s)\s*ux\b/gi, "UX/UI"],
[/\bux\s*(?:\/|&|y|-|–|—|\s)\s*ui\b/gi, "UX/UI"],
    // ✅ acrónimos / mayúsculas obligatorias
    [/\bb2b\b/gi, "B2B"],
    [/\bb2c\b/gi, "B2C"],

    // ✅ razón social: S.A. siempre en mayúscula (cubre "sa", "s.a", "s. a.", etc.)
    [/\bS\s*\.?\s*A\s*\.?\b/gi, "S.A."],



    [/power\s*bi/gi, "Power BI"],
    [/\bexcel\b/gi, "Excel"],
    [/google\s*sheets/gi, "Google Sheets"],
    [/google\s*docs/gi, "Google Docs"],
    [/javascript/gi, "JavaScript"],
    [/typescript/gi, "TypeScript"],
    [/\bnode\s*js\b|\bnodejs\b/gi, "Node.js"],
    [/\breact\s*js\b|\breactjs\b/gi, "React"],
    [/\bnext\s*js\b|\bnextjs\b/gi, "Next.js"],
    [/\bpostgres\b|\bpostgresql\b/gi, "PostgreSQL"],
    [/\bmysql\b/gi, "MySQL"],
    [/\bmongodb\b/gi, "MongoDB"],
    [/\baws\b/gi, "AWS"],
    [/\bazure\b/gi, "Azure"],
    [/\bgcp\b/gi, "GCP"],
    [/\bapi\b/gi, "API"],
    [/\bapis\b/gi, "APIs"],
    [/\bsap\b/gi, "SAP"],
    [/\bkyc\b/gi, "KYC"],
    [/\buif\b/gi, "UIF"],
    [/\bong\b/gi, "ONG"],
    [/\bddjj\b/gi, "DDJJ"],
    [/\bafip\b/gi, "AFIP"],
    [/\barba\b/gi, "ARBA"],
    [/\bagip\b/gi, "AGIP"],
        // ✅ acrónimos / mayúsculas obligatorias
    [/\bespsyc\s*\(\s*a-?1327\s*\)\b/gi, "ESPSYC (A-1327)"],
    [/\bespsyc\b/gi, "ESPSYC"],
    [/\ba-?1327\b/gi, "A-1327"],
    [/\bsql\b/gi, "SQL"],
    [/\brrhh\b/gi, "RRHH"],
    [/\baysa\b/gi, "AYSA"],

  ];

  for (const [re, to] of repl) t = t.replace(re, to);
  return t;
}


function normalizeSeniorityWords(s) {
  let t = safeStr(s);
  if (!t) return "";
  return t.replace(/\bS[éeê]nior\b/g, "Senior");
}

function normalizeRetailWord(s) {
  let t = safeStr(s);
  if (!t) return "";
  return t.replace(/\bretail\b/gi, "Retail");
}

function normalizeDegreesInText(s) {
  let t = safeStr(s);
  if (!t) return "";

  t = t.replace(/\blicenciatura en\b/gi, "Licenciatura en");
  t = t.replace(/\bingenier[ií]a en\b/gi, "Ingeniería en");
  t = t.replace(/\btecnicatura en\b/gi, "Tecnicatura en");

  return t;
}



/* =========================================================================
   PROPER NOUNS / INSTITUTIONS (diccionario + title case controlado)
   ========================================================================= */

const CANONICAL_INSTITUTIONS = {
  /* =========================
     UNIVERSIDADES NACIONALES
     ========================= */

  // UBA
  "universidad de buenos aires": "UBA",
  "uba": "UBA",

  // UNLP
  "universidad nacional de la plata": "UNLP",
  "unlp": "UNLP",

  // UNC (Córdoba) — muy estándar
  "universidad nacional de cordoba": "UNC",
  "universidad nacional de córdoba": "UNC",
  "unc": "UNC",

  // UNR
  "universidad nacional de rosario": "UNR",
  "universidad nacional rosario": "UNR",
  "unr": "UNR",

  // UNSAM
  "universidad nacional de san martin": "UNSAM",
  "universidad nacional de san martín": "UNSAM",
  "unsam": "UNSAM",

  // UNLaM
  "universidad nacional de la matanza": "UNLaM",
  "unlam": "UNLaM",

  // UNQ
  "universidad nacional de quilmes": "UNQ",
  "unq": "UNQ",

  // UNL
  "universidad nacional del litoral": "UNL",
  "unl": "UNL",

  // UNMDP
  "universidad nacional de mar del plata": "UNMDP",
  "unmdp": "UNMDP",

  // UNNE
  "universidad nacional del nordeste": "UNNE",
  "unne": "UNNE",

  // UNNOBA
  "universidad nacional del noroeste de la provincia de buenos aires": "UNNOBA",
  "universidad nacional del noroeste de la provincia de buenos aires unno": "UNNOBA",
  "unnoba": "UNNOBA",

  // UNLu
  "universidad nacional de lujan": "UNLu",
  "universidad nacional de luján": "UNLu",
  "unlu": "UNLu",

  // UNaHur
  "universidad nacional de hurlingham": "UNaHur",
  "unahur": "UNaHur",

  // UNAJ
  "universidad nacional arturo jauretche": "UNAJ",
  "universidad nacional arturo jauretche unaj": "UNAJ",
  "unaj": "UNAJ",

  // UNGS
  "universidad nacional de general sarmiento": "UNGS",
  "ungs": "UNGS",

  // UNTREF
  "universidad nacional de tres de febrero": "UNTREF",
  "untref": "UNTREF",

  // UNVIME
  "universidad nacional de villa mercedes": "UNViMe",
  "unvime": "UNViMe",

  // UNSL
  "universidad nacional de san luis": "UNSL",
  "unsl": "UNSL",

  // UNSJ
  "universidad nacional de san juan": "UNSJ",
  "unsj": "UNSJ",

  // UNCA
  "universidad nacional de catamarca": "UNCA",
  "unca": "UNCA",

  // UNSE
  "universidad nacional de santiago del estero": "UNSE",
  "unse": "UNSE",

  // UNSa (Salta)
  "universidad nacional de salta": "UNSa",
  "unsa": "UNSa",

  // UNER
  "universidad nacional de entre rios": "UNER",
  "universidad nacional de entre ríos": "UNER",
  "uner": "UNER",

  // UNICEN
  "universidad nacional del centro de la provincia de buenos aires":
    "UNICEN",
  "universidad nacional del centro de la provincia de buenos aires unicen":
    "UNICEN",
  "unicen": "UNICEN",

  // UNPA (Patagonia Austral)
  "universidad nacional de la patagonia austral": "UNPA",
  "unpa": "UNPA",

  // UNPSJB (Patagonia San Juan Bosco)
  "universidad nacional de la patagonia san juan bosco": "UNPSJB",
  "unpsjb": "UNPSJB",

  // UNRN
  "universidad nacional de rio negro": "UNRN",
  "universidad nacional de río negro": "UNRN",
  "unrn": "UNRN",

  // UNCO (Comahue)
  "universidad nacional del comahue": "UNCo",
  "unco": "UNCo",

  // UNS (Sur)
  "universidad nacional del sur": "UNS",
  "uns": "UNS",

  // UNAF (Formosa)
  "universidad nacional de formosa": "UNaF",
  "unaf": "UNaF",

  // UNLa (Lanús)
  "universidad nacional de lanus": "UNLa",
  "universidad nacional de lanús": "UNLa",
  "unla": "UNLa",

  // UNdAv
  "universidad nacional de avellaneda": "UNDAV",
  "undav": "UNDAV",

  // UNVM
  "universidad nacional de villa maria": "UNVM",
  "universidad nacional de villa maría": "UNVM",
  "unvm": "UNVM",

  // UNRC
  "universidad nacional de rio cuarto": "UNRC",
  "universidad nacional de río cuarto": "UNRC",
  "unrc": "UNRC",

  // UNPAZ
  "universidad nacional de jose c paz": "UNPAZ",
  "universidad nacional de josé c paz": "UNPAZ",
  "unpaz": "UNPAZ",

  // UNIC — (no agregar, es ambiguo)

  /* =========================
     UTN (Universidad Tecnológica Nacional)
     ========================= */

  "universidad tecnologica nacional": "UTN",
  "universidad tecnológica nacional": "UTN",
  "utn": "UTN",

  // Regionales más comunes (si aparece “UTN FRBA”, etc.)
  "utn frba": "UTN FRBA",
  "utn regional buenos aires": "UTN FRBA",
  "utn facultad regional buenos aires": "UTN FRBA",

  "utn frlp": "UTN FRLP",
  "utn regional la plata": "UTN FRLP",
  "utn facultad regional la plata": "UTN FRLP",

  "utn frro": "UTN FRRO",
  "utn regional rosario": "UTN FRRO",
  "utn facultad regional rosario": "UTN FRRO",

  "utn frc": "UTN FRC",
  "utn regional cordoba": "UTN FRC",
  "utn regional córdoba": "UTN FRC",
  "utn facultad regional cordoba": "UTN FRC",
  "utn facultad regional córdoba": "UTN FRC",

  "utn frsf": "UTN FRSF",
  "utn regional santa fe": "UTN FRSF",
  "utn facultad regional santa fe": "UTN FRSF",

  "utn frm": "UTN FRM",
  "utn regional mendoza": "UTN FRM",
  "utn facultad regional mendoza": "UTN FRM",

  "utn frt": "UTN FRT",
  "utn regional tucuman": "UTN FRT",
  "utn regional tucumán": "UTN FRT",
  "utn facultad regional tucuman": "UTN FRT",
  "utn facultad regional tucumán": "UTN FRT",

  /* =========================
     UNIVERSIDADES PRIVADAS FRECUENTES
     ========================= */

  // UADE
  "universidad argentina de la empresa": "UADE",
  "uade": "UADE",

  // UCA
  "universidad catolica argentina": "UCA",
  "universidad católica argentina": "UCA",
  "uca": "UCA",

  // UCES
  "universidad de ciencias empresariales y sociales": "UCES",
  "uces": "UCES",

  // UAI
  "universidad abierta interamericana": "UAI",
  "uai": "UAI",

  // UP
  "universidad de palermo": "UP",
  "up": "UP",

  // USAL
  "universidad del salvador": "USAL",
  "usal": "USAL",

  // UB (Belgrano)
  "universidad de belgrano": "UB",
  "ub": "UB",

  // UDESA
  "universidad de san andres": "UDESA",
  "universidad de san andrés": "UDESA",
  "udesa": "UDESA",

  // UCEMA
  "universidad del cema": "UCEMA",
  "ucema": "UCEMA",

  // UCAECE
  "universidad caece": "CAECE",
  "caece": "CAECE",

  // UM (Morón)
  "universidad de moron": "Universidad de Morón",
  "universidad de morón": "Universidad de Morón",

  // UCA (ya)
  // UNSTA (privada, Tucumán)
  "universidad del norte santo tomas de aquino": "UNSTA",
  "universidad del norte santo tomás de aquino": "UNSTA",
  "unsta": "UNSTA",

  // UBP (Blas Pascal)
  "universidad blas pascal": "Universidad Blas Pascal",
  "ubp": "Universidad Blas Pascal",

  // Siglo 21
  "universidad siglo 21": "Universidad Siglo 21",
  "universidad siglo xxi": "Universidad Siglo 21",
  "siglo 21": "Universidad Siglo 21",
  "siglo xxi": "Universidad Siglo 21",

  // Austral
  "universidad austral": "Universidad Austral",

  // Maimónides
  "universidad maimonides": "Universidad Maimónides",
  "universidad maimónides": "Universidad Maimónides",

  // Kennedy
  "universidad argentina john f kennedy": "Universidad Kennedy",
  "universidad kennedy": "Universidad Kennedy",

  // Católica de Córdoba
  "universidad catolica de cordoba": "Universidad Católica de Córdoba",
  "universidad católica de córdoba": "Universidad Católica de Córdoba",

  // Católica de Salta
  "universidad catolica de salta": "Universidad Católica de Salta",
  "universidad católica de salta": "Universidad Católica de Salta",

  // Católica de Santa Fe
  "universidad catolica de santa fe": "Universidad Católica de Santa Fe",
  "universidad católica de santa fe": "Universidad Católica de Santa Fe",

  /* =========================
     ORGANISMOS / ESTADO (muy comunes en CVs)
     ========================= */

  "ministerio de cultura de la nacion": "Ministerio de Cultura de la Nación",
  "ministerio de cultura de la nación": "Ministerio de Cultura de la Nación",

  "ministerio de trabajo empleo y seguridad social": "Ministerio de Trabajo, Empleo y Seguridad Social",
  "ministerio de trabajo, empleo y seguridad social": "Ministerio de Trabajo, Empleo y Seguridad Social",

  "ministerio de educacion": "Ministerio de Educación",
  "ministerio de educación": "Ministerio de Educación",

  "ministerio de salud": "Ministerio de Salud",

  "anses": "ANSES",
  "afip": "AFIP",
  "renaper": "RENAPER",
  "pami": "PAMI",
  "senasa": "SENASA",
  "inta": "INTA",
  "inti": "INTI",
  "conicet": "CONICET",

  "gcba": "GCBA",
  "gobierno de la ciudad de buenos aires": "GCBA",

  "gobierno de la provincia de buenos aires": "Gobierno de la Provincia de Buenos Aires",

  "amia": "AMIA",

  "universidad tecnologica nacional": "UTN",
"universidad tecnológica nacional": "UTN",
"universidad del museo social argentino": "UMSA",
"umsa": "UMSA",


  /* =========================
     LUGARES / MARCAS (casos típicos CVs)
     ========================= */

  "alvear palace hotel": "Alvear Palace Hotel",
  "boston us": "Boston US",
};


const SMALL_WORDS_ES = new Set([
  "de","del","la","las","el","los","y","e","en","a","al","por","para","con","sin","o","u"
]);

function normalizeSpaces(s) {
  return safeStr(s).replace(/\s+/g, " ").trim();
}

function stripEdgePunct(token) {
  // Mantiene puntuación final tipo "," "." ")" etc.
  const m = token.match(/^(.+?)([.,;:!?)]*)$/);
  return m ? { core: m[1], punct: m[2] } : { core: token, punct: "" };
}

function isAllCapsAcronym(w) {
  return /^[A-Z0-9]{2,}$/.test(w);
}

function looksLikeAcronym(w) {
  // "uade" "amia" "unr" etc.
  return /^[a-z]{2,6}$/.test(w);
}

function toTitleCaseEs(text) {
  const t = normalizeSpaces(text);
  if (!t) return "";

  const parts = t.split(" ");
  const out = [];

  for (let i = 0; i < parts.length; i++) {
    const { core, punct } = stripEdgePunct(parts[i]);
    if (!core) continue;

    // Si ya viene como sigla en mayúsculas, mantener
    if (isAllCapsAcronym(core)) {
      out.push(core + punct);
      continue;
    }

    const low = core.toLocaleLowerCase("es-AR");

    // conectores en minúscula (excepto si es primera palabra)
    if (i !== 0 && SMALL_WORDS_ES.has(low)) {
      out.push(low + punct);
      continue;
    }

    // Capitaliza palabra
    const cap = low.charAt(0).toLocaleUpperCase("es-AR") + low.slice(1);
    out.push(cap + punct);
  }

  return out.join(" ");
}

function canonicalizeInstitution(text) {
  const t = normalizeSpaces(text);
  if (!t) return "";

  const key = t.toLocaleLowerCase("es-AR");
  if (CANONICAL_INSTITUTIONS[key]) return CANONICAL_INSTITUTIONS[key];

  // Caso: si es un “acronym-like” conocido y está en diccionario, lo levanta arriba.
  // Si no está, no lo fuerces a sigla: solo Title Case.
  return toTitleCaseEs(t);
}

function canonicalizeInstitutionRobust(text) {
  let t = normalizeSpaces(text);
  if (!t) return "";

  // saca paréntesis típicos tipo "(frba, utn)" o "(UTN)"
  // pero preserva el contenido útil agregándolo afuera como tokens
  // Ej: "Universidad Tecnológica Nacional (frba, UTN)" -> "UTN FRBA"
  const parens = [];
  t = t.replace(/\(([^)]{1,80})\)/g, (_, inside) => {
    parens.push(inside);
    return " ";
  });

  t = normalizeSpaces(t);

  // tokens extra desde paréntesis (separadores comunes)
  const extras = parens
    .join(" ")
    .split(/[,;/|]+|\s+/g)
    .map((x) => stripAccents(x).toLowerCase().trim())
    .filter(Boolean);

  // si aparece "utn" + "frba", priorizamos UTN FRBA
  const hasUtn = extras.includes("utn") || stripAccents(t).toLowerCase().includes("utn");
  const regional = extras.find((x) => /^fr[a-z]{2,4}$/.test(x)); // frba, frlp, frro, etc.

  if (hasUtn && regional) {
    return `UTN ${regional.toUpperCase()}`;
  }

  // si el texto base es una institución conocida
  const base = canonicalizeInstitution(t);

  // si extras trae una sigla conocida sola, úsala
  for (const ex of extras) {
    if (CANONICAL_INSTITUTIONS[ex]) return CANONICAL_INSTITUTIONS[ex];
  }

  return base;
}


/**
 * Aplica diccionario + TitleCase controlado SOLO donde corresponde:
 * - instituciones/empresas/organismos/lugares
 * - NO usar en emails, websites, ni textos largos.
 */
function normalizeInstitutionField(s) {
  const t = normalizeSpaces(s);
  if (!t) return "";
  // Evitar tocar emails/urls por si entra algo raro
  if (/@/.test(t) || /\bhttps?:\/\//i.test(t) || /\bwww\./i.test(t)) return t;
  return canonicalizeInstitution(t);
}


function dropExperiencesFrom(data, fromIndex = 6, toIndex = 7) {
  for (let n = fromIndex; n <= toIndex; n++) {
    data[`exp_${n}_dates`] = "";
    data[`exp_${n}_company`] = "";
    data[`exp_${n}_role`] = "";
    for (let b = 1; b <= 5; b++) data[`exp_${n}_b${b}`] = "";
    data[`exp_${n}_bullets_block`] = "";
  }
}

function normalizePhone(phoneRaw, { defaultCountry = "54" } = {}) {
  let s = safeStr(phoneRaw).trim();
  if (!s) return "";

  // Si viene tipo 00... lo pasamos a +
  s = s.replace(/^\s*00/, "+");

  // Dejamos solo números y +
  s = s.replace(/[^\d+]/g, "");

  // Si empieza con + lo sacamos para normalizar
  if (s.startsWith("+")) s = s.slice(1);

  // Si quedó vacío
  if (!s) return "";

  // Si NO empieza con código país, asumimos defaultCountry (Argentina 54)
  // Ej: "1161839587" -> "54" + "1161839587"
  if (!s.startsWith(defaultCountry)) {
    s = defaultCountry + s;
  }

  // Siempre devolvemos con +
  return `+${s}`;
}



function stripAccents(s) {
  return safeStr(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function toSafeFilename(s, fallback = "CV") {
  const base = stripAccents(s)
    .replace(/[^a-zA-Z0-9\s._-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const out = base.length >= 2 ? base : fallback;
  return out.replace(/\s+/g, "_");
}

function maybeClampPlain(s, maxChars) {
  const text = safeStr(s).replace(/\s+/g, " ").trim();
  if (!text) return "";
  if (!LIMITS.ENABLE_CLAMP) return text;
  if (!maxChars || maxChars <= 0) return text;
  return text.length > maxChars ? text.slice(0, Math.max(0, maxChars - 1)).trimEnd() + "…" : text;
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

function shouldDropITItem(s) {
  const t = stripAccents(safeStr(s)).toLowerCase().replace(/\s+/g, " ").trim();
  if (!t) return true;

  // Variantes típicas que NO queremos
  if (t === "internet") return true;
  if (t === "email") return true;
  if (t === "e-mail") return true;
  if (t === "correo electronico") return true;
  if (t === "correo electrónico") return true;

  // combinaciones tipo "email e internet", "internet y email", etc.
  if (/\b(email|e-mail|correo electronico|correo electrónico)\b/.test(t) && /\binternet\b/.test(t)) return true;

  // también frases tipo "manejo de internet", "uso de email"
  if (/\b(uso|manejo|gestion|gestión)\s+de\s+(internet|email|e-mail|correo electronico|correo electrónico)\b/.test(t))
    return true;

  return false;
}


function shouldDropLanguageItem(s) {
  const t = stripAccents(safeStr(s))
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

  if (!t) return true;

  // Si el idioma base es español/castellano, lo descartamos aunque venga con nivel/formato
  if (/\b(espanol|español|castellano)\b/.test(t)) return true;

  // Casos donde viene solo el nivel
  if (t === "nativo" || t === "nativa") return true;

  // Opcional: si a veces viene "idioma: español"
  if (/^\s*idioma\s*:\s*(espanol|español|castellano)\b/.test(t)) return true;

  return false;
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

function linesPreserveGaps(arr) {
  return (arr || []).map((x) => safeStr(x).replace(/\r/g, "").trimEnd());
}

function capitalizeFirst(s) {
  const t = safeStr(s).trim();
  if (!t) return "";
  return t.toLocaleUpperCase("es-AR");
}

function asBulletLines(items) {
  return nonEmptyLines(items).map((x) => "• " + x);
}

function buildEducationLines(data, { gapLines = 1 } = {}) {
  const out = [];
  const gap = Math.max(0, Number(gapLines || 0));

  for (let i = 1; i <= 3; i++) {
    const degree = stripAllPipes(data[`edu_${i}_degree`]);   // ✅ limpia cualquier |
const school = stripAllPipes(data[`edu_${i}_school`]);   // ✅ limpia cualquier |
const schoolNorm = normalizeInstitutionField(school);

let years = stripAllPipes(data[`edu_${i}_years`]);       // ✅ limpia cualquier |


    // Normaliza mayúsculas básicas dentro de years (Finalizado, En curso, etc.)
    years = normalizeStatusWords(years);

    // ✅ Formato objetivo:
    // DEGREE | SCHOOL (YEARS)
    let line = "";

    if (degree && schoolNorm) line = `${degree} | ${schoolNorm}`;
else line = degree || schoolNorm || "";


    if (years) {
      // si ya viene con paréntesis, no los duplica
      const y = years.trim();
      line += y.startsWith("(") ? ` ${y}` : ` (${y})`;
    }

    if (!line.trim()) continue;

    if (out.length) for (let g = 0; g < gap; g++) out.push("");
    out.push(line.trim());
  }

  return out;
}



function stripPipes(s) {
  return safeStr(s)
    .replace(/\|{2,}/g, "|")
    .replace(/^\s*\|\s*/g, "")
    .replace(/\s*\|\s*$/g, "")
    .trim();
}

function stripPipesFromBullets(s) {
  // En bullets NUNCA queremos pipes: los convertimos a coma o espacio
  return safeStr(s)
    .replace(/\|+/g, " / ")     // o ", " si preferís
    .replace(/\s+/g, " ")
    .trim();
}


function stripAllPipes(s) {
  return safeStr(s)
    .replace(/\|/g, " ")        // cualquier pipe => espacio
    .replace(/\s+/g, " ")       // colapsa espacios
    .trim();
}


function normalizeEducationVisualGroup(data) {
  for (let i = 1; i <= 3; i++) {
    let degree = stripPipes(data[`edu_${i}_degree`]);
    let school = stripPipes(data[`edu_${i}_school`]);
    let years  = stripPipes(data[`edu_${i}_years`]);

    // Limpia pipes “colgando”
    degree = degree.replace(/\s*\|\s*$/g, "").trim();
    school = school.replace(/\s*\|\s*$/g, "").trim();
    years  = years.replace(/\s*\|\s*$/g, "").trim();

    const hasDegree = !!degree;
    const hasSchool = !!school;
    const hasYears  = !!years;

    // ✅ El separador debe quedar pegado a DEGREE (bold)
    // En PPTX: degree suele estar en un run y school en otro.
    // Entonces ponemos el pipe al inicio de school (run normal).
    if (hasDegree && hasSchool) {
      school = `| ${school}`;
    } else if (hasDegree && !hasSchool && hasYears) {
      years = `| ${years}`;
    }

    // Espaciado prolijo si years no trae paréntesis
    // (no tocamos si ya viene con "(...)")
    // Si tu plantilla ya pone paréntesis, esto no molesta.
    data[`edu_${i}_degree`] = degree;
    data[`edu_${i}_school`] = school;
    data[`edu_${i}_years`] = years;
  }
}


function normalizeExperienceVisualGroup(data) {
  for (let n = 1; n <= 7; n++) {
    let role = stripPipes(data[`exp_${n}_role`]);
    let company = stripPipes(data[`exp_${n}_company`]);
    let dates = stripPipes(data[`exp_${n}_dates`]);

    role = role.replace(/\s*\|\s*$/g, "").trim();
    company = company.replace(/\s*\|\s*$/g, "").trim();
    dates = dates.replace(/\s*\|\s*$/g, "").trim();

    const hasRole = !!role;
    const hasCompany = !!company;
    const hasDates = !!dates;

    // ✅ El separador va en "company" (normal), no en "role" (bold)
    if (hasRole && hasCompany) {
      company = `| ${company}`;
    } else if (hasRole && !hasCompany && hasDates) {
      // si no hay company, el pipe queda antes de dates
      dates = `| ${dates}`;
    }

    // Espaciado antes de dates si corresponde
    if (hasDates && !dates.startsWith(" ") && (hasRole || hasCompany)) {
      dates = " " + dates;
    }

    data[`exp_${n}_role`] = role;
    data[`exp_${n}_company`] = company;
    data[`exp_${n}_dates`] = dates;
  }
}




function buildSidebarSections(
  data,
  { underlineMin = 5, underlineExtra = 0, educationGapLines = 0 } = {}
) {
  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 5), t.length + Number(underlineExtra || 0));
    return "─".repeat(len);
  }

  data.sidebar_sections_noedu = buildSidebarSectionsNoEdu(data, {
    underlineMin: 5,
    underlineExtra: 0,
  });

  const sections = [];

  function pushSection(title, lines, { preserveGaps = false } = {}) {
    const t = capitalizeFirst(title);

    if (preserveGaps) {
      const raw = linesPreserveGaps(lines);
      while (raw.length && !safeStr(raw[0]).trim()) raw.shift();
      while (raw.length && !safeStr(raw[raw.length - 1]).trim()) raw.pop();
      if (!raw.some((x) => safeStr(x).trim())) return;

      sections.push({ title: t, line: underlineForTitle(t), body: raw.join("\n") });
      return;
    }

    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    sections.push({ title: t, line: underlineForTitle(t), body: clean.join("\n") });
  }

  pushSection("Educación", asBulletLines(buildEducationLines(data, { gapLines: 0 })));

  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos", asBulletLines(cursos));

  const it = [];
  for (let i = 1; i <= 6; i++) it.push(data[`it_${i}`]);
  pushSection("Informática", asBulletLines(it));

  const idiomas = [];
  for (let i = 1; i <= 3; i++) idiomas.push(data[`idioma_${i}`]);
  pushSection("Idiomas", asBulletLines(idiomas));

  const skills = [];
  for (let i = 1; i <= 7; i++) skills.push(data[`skill_${i}`]);
  pushSection("Competencias", asBulletLines(skills));

  return sections;
}

function buildSidebarSectionsNoEdu(data, opts = {}) {
  const { underlineMin = 5, underlineExtra = 0 } = opts;

  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 5), t.length + Number(underlineExtra || 0));
    return "─".repeat(len);
  }

  const sections = [];

  function pushSection(title, lines) {
    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    const t = capitalizeFirst(title);
    sections.push({ title: t, line: underlineForTitle(t), body: clean.join("\n") });
  }

  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos", asBulletLines(cursos));

  const it = [];
  for (let i = 1; i <= 6; i++) it.push(data[`it_${i}`]);
  pushSection("Informática", asBulletLines(it));

  const idiomas = [];
  for (let i = 1; i <= 3; i++) idiomas.push(data[`idioma_${i}`]);
  pushSection("Idiomas", asBulletLines(idiomas));

  const skills = [];
  for (let i = 1; i <= 7; i++) skills.push(data[`skill_${i}`]);
  pushSection("Competencias", asBulletLines(skills));

  return sections;
}

function buildSidebarSectionsEduCursosOnly(
  data,
  { underlineMin = 22, underlineExtra = 10, educationGapLines = 1 } = {}
) {
  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 5), t.length + Number(underlineExtra || 0));
    return "─".repeat(len);
  }

  const sections = [];

  function pushSection(title, lines, { preserveGaps = false } = {}) {
    const t = safeStr(title).trim();

    if (preserveGaps) {
      const raw = linesPreserveGaps(lines);
      while (raw.length && !safeStr(raw[0]).trim()) raw.shift();
      while (raw.length && !safeStr(raw[raw.length - 1]).trim()) raw.pop();
      if (!raw.some((x) => safeStr(x).trim())) return;

      sections.push({ title: t, line: underlineForTitle(t), body: raw.join("\n") });
      return;
    }

    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    sections.push({ title: t, line: underlineForTitle(t), body: clean.join("\n") });
  }

  pushSection(
  "Formación Académica",
  asBulletLines(buildEducationLines(data, { gapLines: 0 }))
);
 
  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos y Capacitaciones", asBulletLines(cursos));

  return sections;
}

function buildSidebarContactSection(data, { underlineMin = 22, underlineExtra = 10 } = {}) {
  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 5), t.length + Number(underlineExtra || 0));
    return "─".repeat(len);
  }

  const rows = [];

  const phone = safeStr(data.contact_phone).trim();
  const email = safeStr(data.contact_email).trim();
  const loc = safeStr(data.contact_location).trim();

  const licenciaRaw = safeStr(data.licencia).trim();
  const licenciaNorm = licenciaRaw.toLowerCase();
  const hasLicense = licenciaNorm === "si" || licenciaNorm === "sí" || licenciaNorm.includes("licencia");

  if (phone) rows.push({ icon: "☎", text: phone });
  if (email) rows.push({ icon: "✉", text: email });
  if (loc) rows.push({ icon: "📍", text: loc });
  if (hasLicense) rows.push({ icon: "🚗", text: "Licencia de conducir" });

  if (!rows.length) return null;

  const title = "Contacto";

  return {
    title,
    line: underlineForTitle(title),
    icons_block: rows.map((r) => r.icon).join("\n"),
    text_block: rows.map((r) => r.text).join("\n"),
  };
}

function buildContactRowsWithIcons(data, { iconTheme = "dark" } = {}) {
  const rows = [];

  const phone = safeStr(data.contact_phone).trim();
  const email = safeStr(data.contact_email).trim();
  const loc = safeStr(data.contact_location).trim();

  const licenciaNorm = safeStr(data.licencia).trim().toLowerCase();
  const hasLicense = licenciaNorm === "si" || licenciaNorm === "sí";

  if (phone) rows.push({ icon: loadIconBuffer(iconTheme, "phone.png"), text: phone });
  if (email) rows.push({ icon: loadIconBuffer(iconTheme, "mail.png"), text: email });
  if (loc) rows.push({ icon: loadIconBuffer(iconTheme, "location.png"), text: loc });
  if (hasLicense) rows.push({ icon: loadIconBuffer(iconTheme, "car.png"), text: "Licencia de conducir" });

  return rows;
}

function buildSidebarSectionsFull(
  data,
  { underlineMin = 5, underlineExtra = 0, educationGapLines = 0 } = {}
) {
  function underlineForTitle(title) {
    const t = safeStr(title).trim();
    const len = Math.max(Number(underlineMin || 5), t.length + Number(underlineExtra || 0));
    return "─".repeat(len);
  }

  const sections = [];

  function pushSection(title, lines, { preserveGaps = false } = {}) {
    const t = capitalizeFirst(title);

    if (preserveGaps) {
      const raw = linesPreserveGaps(lines);
      while (raw.length && !safeStr(raw[0]).trim()) raw.shift();
      while (raw.length && !safeStr(raw[raw.length - 1]).trim()) raw.pop();
      if (!raw.some((x) => safeStr(x).trim())) return;

      sections.push({ title: t, line: underlineForTitle(t), body: raw.join("\n") });
      return;
    }

    const clean = nonEmptyLines(lines);
    if (!clean.length) return;

    sections.push({ title: t, line: underlineForTitle(t), body: clean.join("\n") });
  }

  const contact = [];
  const phone = safeStr(data.contact_phone).trim();
  const email = safeStr(data.contact_email).trim();
  const loc = safeStr(data.contact_location).trim();

  if (phone) contact.push(phone);
  if (email) contact.push(email);
  if (loc) contact.push(loc);

  pushSection("Contacto", contact);

  pushSection("Educación", asBulletLines(buildEducationLines(data, { gapLines: 0 })));

  const cursos = [];
  for (let i = 1; i <= 6; i++) cursos.push(data[`curso_${i}`]);
  pushSection("Cursos", asBulletLines(cursos));

  const it = [];
  for (let i = 1; i <= 6; i++) it.push(data[`it_${i}`]);
  pushSection("Informática", asBulletLines(it));

  const idiomas = [];
  for (let i = 1; i <= 3; i++) idiomas.push(data[`idioma_${i}`]);
  pushSection("Idiomas", asBulletLines(idiomas));

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
  5: { about: { maxChars: 6000 }, photoSize: [200, 200] },
  6: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  7: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  8: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  9: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  10: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  11: { about: { maxChars: 6000 }, photoSize: [520, 520] },
  12: { about: { maxChars: 6000 }, photoSize: [390, 390] },
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

const COLOR_MAP = {
  "violeta clarito": "E5E0EA",
  "violeta claro": "D2C9DB",
  violeta: "7E6597",
  lila: "E5E0EA",
  lavanda: "D2C9DB",
  malva: "D2C9DB",
  purpura: "7E6597",
  púrpura: "7E6597",

  "rosa clarito": "F6E8ED",
  "rosa claro": "F6E8ED",
  "rosa pastel": "F6E8ED",
  "rosa pálido": "F6E8ED",
  "rosa palido": "F6E8ED",
  "rosa viejo": "E8B4B8",
  "rosa palo": "E8B4B8",
  "rosa nude": "EED6D3",
  rosado: "F6E8ED",
  salmon: "EED6D3",
  salmón: "EED6D3",

  "celeste clarito": "D5DEE6",
  "celeste claro": "D5DEE6",
  "celeste pastel": "D5DEE6",
  "celeste bebé": "D5DEE6",
  "azul clarito": "D5DEE6",
  "azul claro": "608ABF",
  "azul pastel": "D5DEE6",
  "azul grisaceo": "323B4C",
  "azul grisáceo": "323B4C",
  "azul marino": "062446",
  "azul oscuro": "062446",
  "azul noche": "062446",
  "azul petróleo": "1F3A5F",
  "azul petroleo": "1F3A5F",
  azul: "002D6A",

  "verde clarito": "D2E0E1",
  "verde claro": "D2E0E1",
  "verde pastel": "D2E0E1",
  "verde menta": "D2E0E1",
  "verde agua": "D2E0E1",
  "verde seco": "44867B",
  "verde sobrio": "44867B",
  "verde oliva": "657757",
  "verde musgo": "5F6F65",
  "verde esmeralda": "0F766E",
  "verde oscuro": "2B554E",
  verde: "44867B",

  "gris clarito": "C7C8CA",
  "gris claro": "C7C8CA",
  "gris perla": "C7C8CA",
  "gris humo": "C7C8CA",
  "gris grafito": "3F3F3F",
  "gris oscuro": "696969",
  gris: "696969",
  negro: "062446",

  beige: "B8A797",
  arena: "B8A797",
  nude: "EED6D3",
  camel: "B8A797",
  crema: "D5DEE6",
  vison: "B8A797",
  visón: "B8A797",
  ocre: "B8A797",
  marron: "323B4C",
  marrón: "323B4C",
  terracota: "A62C46",
};

function extractColorKeyword(raw) {
  const s = safeStr(raw).toLowerCase();

  const checks = [
    ["petroleo", "azul petróleo"],
    ["marino", "azul marino"],
    ["oliva", "verde oliva"],
    ["musgo", "verde musgo"],
    ["menta", "verde menta"],
    ["agua", "verde agua"],
    ["celeste", "celeste claro"],
    ["azul", "azul"],
    ["verde", "verde sobrio"],
    ["violeta", "violeta claro"],
    ["lila", "lila"],
    ["rosa", "rosa claro"],
    ["beige", "beige"],
    ["arena", "beige"],
    ["nude", "rosa nude"],
    ["gris", "gris"],
    ["negro", "negro"],
  ];

  for (const [k, v] of checks) {
    if (s.includes(k)) return v;
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

function wantsNoColor(raw) {
  const s = safeStr(raw).toLowerCase();
  return (
    s.includes("sin color") ||
    s.includes("blanco y negro") ||
    s.includes("ats") ||
    s.includes("harvard") ||
    s.includes("solo negro") ||
    s.includes("formato clasico") ||
    s.includes("formato clásico")
  );
}

function resolveAccentHex(coloresRawOrFirst) {
  if (wantsNoColor(coloresRawOrFirst)) return "";
  const first = pickFirstColorRaw(coloresRawOrFirst);
  const hex = normalizeHexColor(first);
  if (hex) return isTooVibrantHex(hex) ? softenVibrantHex(hex) : hex;

  const lowFirst = safeStr(first).toLowerCase();
  if (lowFirst.includes("azul") || lowFirst.includes("blue")) return "1E3A8A";
  if (lowFirst.includes("verde") || lowFirst.includes("green") || lowFirst.includes("teal")) return "0F766E";

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
  return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16) };
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

function hsvToRgb({ h, s, v }) {
  const C = v * s;
  const X = C * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = v - C;

  let r1 = 0,
    g1 = 0,
    b1 = 0;

  if (h >= 0 && h < 60) {
    r1 = C;
    g1 = X;
    b1 = 0;
  } else if (h < 120) {
    r1 = X;
    g1 = C;
    b1 = 0;
  } else if (h < 180) {
    r1 = 0;
    g1 = C;
    b1 = X;
  } else if (h < 240) {
    r1 = 0;
    g1 = X;
    b1 = C;
  } else if (h < 300) {
    r1 = X;
    g1 = 0;
    b1 = C;
  } else {
    r1 = C;
    g1 = 0;
    b1 = X;
  }

  return {
    r: Math.round((r1 + m) * 255),
    g: Math.round((g1 + m) * 255),
    b: Math.round((b1 + m) * 255),
  };
}

function isTooVibrantHex(hex6) {
  const hsv = rgbToHsv(hexToRgb(hex6));
  const L = relLuminance(hexToRgb(hex6));

  if (hsv.v >= 0.72 && hsv.s >= 0.40) return true;

  const hue = hsv.h;
  const isCyanish = hue >= 160 && hue <= 210;
  if (isCyanish && L >= 0.50 && hsv.v >= 0.62 && hsv.s >= 0.30) return true;

  return false;
}

function softenVibrantHex(hex6, { maxS = 0.30, maxV = 0.70 } = {}) {
  const hsv = rgbToHsv(hexToRgb(hex6));
  const s2 = Math.min(hsv.s, maxS);
  const v2 = Math.min(hsv.v, maxV);
  if (s2 === hsv.s && v2 === hsv.v) return hex6;
  const rgb2 = hsvToRgb({ h: hsv.h, s: s2, v: v2 });
  return rgbToHex(rgb2);
}

function shouldForceWhiteText(bgHex6, accentRaw) {
  const raw = safeStr(accentRaw).toLowerCase();
  if (raw.includes("naranja") || raw.includes("ocre") || raw.includes("marron") || raw.includes("marrón")) return true;
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

  return { buffer: zip.generate({ type: "nodebuffer" }), touchedFiles, replacements };
}

function removeEmptyBulletedParagraphs(pptxBuffer) {
  const zip = new PizZip(pptxBuffer);

  const targets = [
    ...(zip.file(/^ppt\/slides\/slide\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/) || []),
  ];

  for (const f of targets) {
    let xml = f.asText();

    xml = xml.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (p) => {
      const hasBullet = /<a:buChar\b|<a:buAutoNum\b/.test(p);
      if (!hasBullet) return p;

      const texts = [...p.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]);
      const joined = texts.join("").replace(/&nbsp;/g, " ").replace(/\s+/g, "").trim();

      if (!joined) return "";
      return p;
    });

    zip.file(f.name, xml);
  }

  return zip.generate({ type: "nodebuffer" });
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
  return { r: Math.round(rSum / n), g: Math.round(gSum / n), b: Math.round(bSum / n) };
}

function colorDist(c1, c2) {
  const dr = c1.r - c2.r;
  const dg = c1.g - c2.g;
  const db = c1.b - c2.b;
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

async function keyOutBackgroundToAlpha(inputBuf, { backgroundColor, threshold = 55, softness = 0.35 } = {}) {
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

    if (d <= t0) out[i + 3] = 0;
    else if (d < t1) {
      const k = (d - t0) / Math.max(1e-6, t1 - t0);
      const alpha = out[i + 3];
      out[i + 3] = Math.round(alpha * k);
    }
  }

  return sharp(out, { raw: { width: info.width, height: info.height, channels: 4 } }).png().toBuffer();
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

  let base = await sharp(photoBuf)
    .rotate()
    .trim({ threshold: 10 })
    .resize(w, h, { fit: "cover", position: "centre" })
    .ensureAlpha()
    .png()
    .toBuffer();

  const bg = await sampleCornerColor(base, 10);
  const bgIsLight = (bg.r + bg.g + bg.b) / 3 >= 210;

  if (bgIsLight) {
    base = await keyOutBackgroundToAlpha(base, { backgroundColor: bg, threshold: 35, softness: 0.4 });
  }

  base = await sharp(base).resize(w, h, { fit: "fill" }).ensureAlpha().png().toBuffer();
  return await forceCircleTransparentOutside(base, { width: w, height: h, padding: 0 });
}

/* =========================================================================
   5) TEMPLATES (MAPPING 1..14) + VARIANTES (v1..v7)
   ========================================================================= */

const TEMPLATE_VARIANTS = {
  1: {
    v1: "Plantilla_oficial_1_verde_v1.pptx",
    v2: "Plantilla_oficial_1_verde_v2.pptx",
    v3: "Plantilla_oficial_1_verde.pptx",
    v4: "Plantilla_oficial_1_verde_v4.pptx",
    v5: "Plantilla_oficial_1_verde_v5.pptx",
    v6: "Plantilla_oficial_1_verde_v6.pptx",
    v7: "Plantilla_oficial_1_verde_v7.pptx",
  },
  2: {
    v1: "Plantilla_oficial_2_v1.pptx",
    v2: "Plantilla_oficial_2_v2.pptx",
    v3: "Plantilla_oficial_2.pptx",
    v4: "Plantilla_oficial_2_v4.pptx",
    v5: "Plantilla_oficial_2_v5.pptx",
    v6: "Plantilla_oficial_2_v6.pptx",
    v7: "Plantilla_oficial_2_v7.pptx",
  },
  3: {
    v1: "Plantilla_oficial_3_v1.pptx",
    v2: "Plantilla_oficial_3_v2.pptx",
    v3: "Plantilla_oficial_3.pptx",
    v4: "Plantilla_oficial_3_v4.pptx",
    v5: "Plantilla_oficial_3_v5.pptx",
    v6: "Plantilla_oficial_3_v6.pptx",
    v7: "Plantilla_oficial_3_v7.pptx",
  },
  4: {
    v1: "Plantilla_oficial_4_v1.pptx",
    v2: "Plantilla_oficial_4_v2.pptx",
    v3: "Plantilla_oficial_4.pptx",
    v4: "Plantilla_oficial_4_v4.pptx",
    v5: "Plantilla_oficial_4_v5.pptx",
    v6: "Plantilla_oficial_4_v6.pptx",
    v7: "Plantilla_oficial_4_v7.pptx",
  },
  5: {
    v1: "Plantilla_oficial_5_v1.pptx",
    v2: "Plantilla_oficial_5_v2.pptx",
    v3: "Plantilla_oficial_5.pptx",
    v4: "Plantilla_oficial_5_v4.pptx",
    v5: "Plantilla_oficial_5_v5.pptx",
    v6: "Plantilla_oficial_5_v6.pptx",
    v7: "Plantilla_oficial_5_v7.pptx",
  },
  6: {
    v1: "Plantilla_oficial_6_v1.pptx",
    v2: "Plantilla_oficial_6_v2.pptx",
    v3: "Plantilla_oficial_6.pptx",
    v4: "Plantilla_oficial_6_v4.pptx",
    v5: "Plantilla_oficial_6_v5.pptx",
    v6: "Plantilla_oficial_6_v6.pptx",
    v7: "Plantilla_oficial_6_v7.pptx",
  },
  7: {
    v1: "Plantilla_oficial_7_v1.pptx",
    v2: "Plantilla_oficial_7_v2.pptx",
    v3: "Plantilla_oficial_7.pptx",
    v4: "Plantilla_oficial_7_v4.pptx",
    v5: "Plantilla_oficial_7_v5.pptx",
    v6: "Plantilla_oficial_7_v6.pptx",
    v7: "Plantilla_oficial_7_v7.pptx",
  },
  8: {
    v1: "Plantilla_oficial_8_v1.pptx",
    v2: "Plantilla_oficial_8_v2.pptx",
    v3: "Plantilla_oficial_8.pptx",
    v4: "Plantilla_oficial_8_v4.pptx",
    v5: "Plantilla_oficial_8_v5.pptx",
    v6: "Plantilla_oficial_8_v6.pptx",
    v7: "Plantilla_oficial_8_v7.pptx",
  },
  9: {
    v1: "Plantilla_oficial_9_v1.pptx",
    v2: "Plantilla_oficial_9_v2.pptx",
    v3: "Plantilla_oficial_9.pptx",
    v4: "Plantilla_oficial_9_v4.pptx",
    v5: "Plantilla_oficial_9_v5.pptx",
    v6: "Plantilla_oficial_9_v6.pptx",
    v7: "Plantilla_oficial_9_v7.pptx",
  },
  10: {
    v1: "Plantilla_oficial_10_v1.pptx",
    v2: "Plantilla_oficial_10_v2.pptx",
    v3: "Plantilla_oficial_10.pptx",
    v4: "Plantilla_oficial_10_v4.pptx",
    v5: "Plantilla_oficial_10_v5.pptx",
    v6: "Plantilla_oficial_10_v6.pptx",
    v7: "Plantilla_oficial_10_v7.pptx",
  },
  11: {
    v1: "Plantilla_oficial_11_v1.pptx",
    v2: "Plantilla_oficial_11_v2.pptx",
    v3: "Plantilla_oficial_11.pptx",
    v4: "Plantilla_oficial_11_v4.pptx",
    v5: "Plantilla_oficial_11_v5.pptx",
    v6: "Plantilla_oficial_11_v6.pptx",
    v7: "Plantilla_oficial_11_v7.pptx",
  },
  12: {
    v1: "Plantilla_oficial_12_v1.pptx",
    v2: "Plantilla_oficial_12_v2.pptx",
    v3: "Plantilla_oficial_12.pptx",
    v4: "Plantilla_oficial_12_v4.pptx",
    v5: "Plantilla_oficial_12_v5.pptx",
    v6: "Plantilla_oficial_12_v6.pptx",
    v7: "Plantilla_oficial_12_v7.pptx",
  },
  13: {
    v1: "Plantilla_oficial_13_v1.pptx",
    v2: "Plantilla_oficial_13_v2.pptx",
    v3: "Plantilla_oficial_13.pptx",
    v4: "Plantilla_oficial_13_v4.pptx",
    v5: "Plantilla_oficial_13_v5.pptx",
    v6: "Plantilla_oficial_13_v6.pptx",
    v7: "Plantilla_oficial_13_v7.pptx",
  },
  14: {
    v1: "Plantilla_oficial_14_v1.pptx",
    v2: "Plantilla_oficial_14_v2.pptx",
    v3: "Plantilla_oficial_14.pptx",
    v4: "Plantilla_oficial_14_v4.pptx",
    v5: "Plantilla_oficial_14_v5.pptx",
    v6: "Plantilla_oficial_14_v6.pptx",
    v7: "Plantilla_oficial_14_v7.pptx",
  },
};


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

const TWO_PAGE_TEMPLATES = new Set([9, 11, 12, 13]);

function capExpCountForTemplate(templateId, expCount) {
  const id = Number(templateId);
  const c = Number(expCount) || 0;
  // Si el template NO es de 2 hojas, nunca pedimos v6/v7 (cap a 5)
  if (!TWO_PAGE_TEMPLATES.has(id)) return Math.min(c, 5);
  return c; // 9/11/12/13 permiten 6-7
}


function wantedVariantKeyByExpCount(expCount) {
  if (!Number.isFinite(expCount) || expCount <= 0) return "v1";
  if (expCount <= 1) return "v1";
  if (expCount === 2) return "v2";
  if (expCount === 3) return "v3";
  if (expCount === 4) return "v4";
  if (expCount === 5) return "v5";
  if (expCount === 6) return "v6";
  return "v7"; // 7 o más
}


function pickVariantFileName(templateId, expCount) {
  const variants = TEMPLATE_VARIANTS[templateId];
  if (!variants) return TEMPLATE_MAP[templateId];

  const order = ["v1", "v2", "v3", "v4", "v5", "v6", "v7"];
  const want = wantedVariantKeyByExpCount(expCount);
  const wantIdx = order.indexOf(want);

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

  if (!tries.length) tries.push("v3", "v7", "v6", "v5", "v4", "v2", "v1");


  for (const k of tries) {
    const fname = variants[k];
    if (!fname) continue;
    const p = path.join(TEMPLATES_DIR, fname);
    if (fs.existsSync(p)) return fname;
  }

  const v3 = variants.v3 && fs.existsSync(path.join(TEMPLATES_DIR, variants.v3)) ? variants.v3 : "";
  if (v3) return v3;

  for (const k of Object.keys(variants)) {
    const fname = variants[k];
    if (fname && fs.existsSync(path.join(TEMPLATES_DIR, fname))) return fname;
  }

  return TEMPLATE_MAP[templateId];
}

function getTemplatePath(templateId, expCountForVariant = 0) {
  const id = normalizeTemplateId(templateId);

  const fileName =
    expCountForVariant >= 0 ? pickVariantFileName(id, expCountForVariant) : TEMPLATE_MAP[id] || TEMPLATE_VARIANTS[id]?.v3;

  if (!fileName) throw new Error(`No hay mapping para template_id=${id}. Revisá TEMPLATE_VARIANTS.`);

  const templatePath = path.join(TEMPLATES_DIR, fileName);

  if (!fs.existsSync(templatePath)) {
    let list = [];
    try {
      list = fs.readdirSync(TEMPLATES_DIR).filter((f) => f.toLowerCase().endsWith(".pptx"));
    } catch (_) {}

    throw new Error(
      `No encuentro la plantilla: ${fileName} en ${TEMPLATES_DIR}. ` + `Disponibles: ${list.join(", ")}.`
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

function ensureDefaultBulletIfEmpty(data, expIndex, text = "El usuario no le puso info") {
  const role = safeStr(data[`exp_${expIndex}_role`]).trim();
  const company = safeStr(data[`exp_${expIndex}_company`]).trim();
  const dates = safeStr(data[`exp_${expIndex}_dates`]).trim();

  const hasExperience = !!(role || company || dates);
  if (!hasExperience) return;

  const anyBullet =
    safeStr(data[`exp_${expIndex}_b1`]).trim() ||
    safeStr(data[`exp_${expIndex}_b2`]).trim() ||
    safeStr(data[`exp_${expIndex}_b3`]).trim() ||
    safeStr(data[`exp_${expIndex}_b4`]).trim() ||
    safeStr(data[`exp_${expIndex}_b5`]).trim();

  if (!anyBullet) {
    data[`exp_${expIndex}_b1`] = text;
  }
}



function removeCommonRedundancies(text) {
  let t = safeStr(text);

  t = t.replace(/\bpara la empresa\b/gi, "");
  t = t.replace(/\bde la empresa\b/gi, "");
  t = t.replace(/\ben la empresa\b/gi, "");

  return t.replace(/\s+/g, " ").trim();
}

function replaceForbiddenWords(s) {
  let t = safeStr(s);
  if (!t) return "";
  // "empleada" prohibida
  t = t.replace(/\bempleada\b/gi, "Asistente");
  return t;
}



function removeDuplicateBullets(data) {
  for (let n = 1; n <= 7; n++) {
    const seen = new Set();

    for (let b = 1; b <= 5; b++) {
      const key = `exp_${n}_b${b}`;
      let bullet = safeStr(data[key]).trim();
      if (!bullet) continue;

      const norm = bullet.toLowerCase();

      if (seen.has(norm)) {
        data[key] = "";
      } else {
        seen.add(norm);
      }
    }
  }
}


/* =========================================================================
   7) DATA MAPPING (hasta 7 experiencias) + sidebar_block
   ========================================================================= */

function removeCompanyOvermentions(bullet, company, { maxMentions = 1 } = {}) {
  let t = safeStr(bullet);
  const c = safeStr(company).trim();
  if (!t || !c) return t;

  const re = new RegExp(c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
  const matches = t.match(re) || [];
  if (matches.length <= maxMentions) return t;

  let seen = 0;
  t = t.replace(re, (m) => {
    seen++;
    return seen <= maxMentions ? m : "";
  });

  return t.replace(/\s+/g, " ").replace(/\(\s*\)/g, "").trim();
}

function normalizeCompanyForMatch(s) {
  let t = stripAccents(safeStr(s)).toLowerCase();
  t = t.replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();

  // Sufijos legales comunes
  t = t
    .replace(/\b(s\.?a\.?|sa)\b/g, "")
    .replace(/\b(s\.?r\.?l\.?|srl)\b/g, "")
    .replace(/\b(s\.?a\.?s\.?|sas)\b/g, "")
    .replace(/\b(s\.?a\.?u\.?|sau)\b/g, "")
    .replace(/\b(ltd\.?|llc|inc\.?)\b/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return t;
}

function removeCompanyMentionsInBullet(bullet, company) {
  let b = safeStr(bullet);
  const cRaw = safeStr(company).trim();
  if (!b || !cRaw) return b;

  const cNorm = normalizeCompanyForMatch(cRaw);
  if (!cNorm || cNorm.length < 4) return b;

  // Match tolerante: palabras de la empresa (sin legales) en orden, con espacios flexibles
  const parts = cNorm.split(" ").filter(Boolean);
  if (!parts.length) return b;

  const pattern = parts.map(escapeRegExp).join("\\s+");
  const reCompany = new RegExp(`\\b${pattern}\\b`, "gi");

  // Borra la empresa si aparece
  b = b.replace(reCompany, "").trim();

  // Borra colas típicas: "en", "para", "de" colgando al final o dobles espacios
  b = b
    .replace(/\b(en|para|de)\s*$/gi, "")
    .replace(/\s+/g, " ")
    .trim();

  // Si quedó algo tipo ": ." o " ,", lo limpia
  b = b.replace(/\s+([,.;:])/g, "$1").trim();

  return b;
}


function cleanExperienceRedundancy(data) {
  for (let n = 1; n <= 7; n++) {
    const company = safeStr(data[`exp_${n}_company`]).trim();
    if (!company) continue;

    for (let b = 1; b <= 5; b++) {
      let bullet = safeStr(data[`exp_${n}_b${b}`]).trim();
      if (!bullet) continue;

      // Limpiezas generales
      bullet = removeCommonRedundancies(bullet);

      // ✅ Cortar "en/para/de + Empresa" y cualquier mención directa
      bullet = removeCompanyMentionsInBullet(bullet, company);

      // Limpiar dobles espacios
      bullet = bullet.replace(/\s+/g, " ").trim();

      data[`exp_${n}_b${b}`] = bullet;
    }
  }
}


// =========================================================================
// FOTO: preferencia del usuario (SIN FOTO)
// =========================================================================

// Ajustá estas 2 frases EXACTAS a lo que tengas en tu formulario
const NO_PHOTO_ANSWERS = {
  NO: "No quisiera utilizar foto así que no es necesario.",
  ADD_LATER: "Agrego luego la foto para no perder tiempo, quisiera que me lo entreguen sin la imagen.",
};

function normalizeNoPhotoAnswer(raw) {
  const s = safeStr(raw).trim();
  if (!s) return "";

  const norm = stripAccents(s).toLowerCase().replace(/\s+/g, " ").trim();

  const no1 = stripAccents(NO_PHOTO_ANSWERS.NO).toLowerCase().replace(/\s+/g, " ").trim();
  const no2 = stripAccents(NO_PHOTO_ANSWERS.ADD_LATER).toLowerCase().replace(/\s+/g, " ").trim();

  if (norm === no1) return "no_photo";
  if (norm === no2) return "no_photo";

  // fallback por keywords (por si cambia un poco el texto)
  if (
    norm.includes("sin foto") ||
    norm.includes("sin la imagen") ||
    norm.includes("no quisiera utilizar foto") ||
    (norm.includes("agrego luego") && norm.includes("foto"))
  ) {
    return "no_photo";
  }

  return "";
}

/**
 * Busca la respuesta "sin foto" aunque cambie la key del formulario.
 * - Primero intenta keys comunes
 * - Si no, busca por texto en el nombre del campo
 */
function getNoPhotoAnswerFromSrc(src) {
  if (!src || typeof src !== "object") return "";

  // 1) keys comunes (si querés, agregá las tuyas acá)
  const direct = getAny(
    src,
    [
      "photo_preference",
      "include_photo",
      "wants_photo",
      "usar_foto",
      "foto",
      "prefiere_foto",
      "foto_cv",
    ],
    ""
  );
  if (direct) return direct;

  // 2) fallback: buscar por texto parecido en el nombre del campo/pregunta
  const keys = Object.keys(src);
  const patterns = [
    "foto",
    "utilizar foto",
    "incluir foto",
    "sin foto",
    "sin la imagen",
  ];

  for (const k of keys) {
    const kn = stripAccents(String(k)).toLowerCase();
    if (patterns.some((p) => kn.includes(stripAccents(p).toLowerCase()))) {
      const v = src[k];
      if (v !== undefined && v !== null && String(v).trim() !== "") return v;
    }
  }

  return "";
}


function flattenToTemplateData(body) {
  const src =
    body?.data && typeof body.data === "object"
      ? body.data
      : body?.fields && typeof body.fields === "object"
      ? body.fields
      : body || {};

  const data = {};

    // Preferencia CV 2 páginas (v5 vs v6/v7)
  data.two_pages_answer_raw = safeStr(getTwoPagesAnswerFromSrc(src));
  data.two_pages_pref = normalizeTwoPagesAnswer(data.two_pages_answer_raw);


  data.template_id = safeStr(getAny(src, ["template_id", "template", "templateId"]));
  data.photo_url = safeStr(getAny(src, ["photo_url", "photoUrl"]));
  data.photo_base64 = safeStr(getAny(src, ["photo_base64", "photoBase64", "photo"]));

    // ✅ Preferencia SIN FOTO (desde formulario)
  data.no_photo_answer_raw = safeStr(getNoPhotoAnswerFromSrc(src));
  data.no_photo_pref = normalizeNoPhotoAnswer(data.no_photo_answer_raw);

  // wants_photo = true por defecto, salvo que el usuario pida explícitamente sin foto
  data.wants_photo = data.no_photo_pref !== "no_photo";

  data.accent_color_raw = safeStr(getAny(src, ["accent_color_raw", "colores_raw", "colors_raw", "colores", "colors"], ""));

  data.name = maybeClampPlain(getAny(src, ["name", "nombre"]), LIMITS.NAME_MAX_CHARS);
  data.title = normalizeSeniorityWords(
  maybeClampPlain(getAny(src, ["title", "titulo"]), LIMITS.TITLE_MAX_CHARS)
);


  const templateId = Number(data.template_id || body.template_id || DEFAULT_TEMPLATE_ID || 1);
  const profile = getProfile(templateId);
  const aboutMax = profile?.about?.maxChars || LIMITS.ABOUT_MAX_CHARS_DEFAULT;
  data.about = normalizeDegreesInText(
  normalizeRetailWord(
    normalizeSeniorityWords(
      maybeClampPlain(getAny(src, ["about", "objective", "objetivo"]), aboutMax)
    )
  )
);


  data.contact_phone = maybeClampPlain(
  normalizePhone(getAny(src, ["contact_phone", "phone", "telefono"])),
  LIMITS.CONTACT_PHONE_MAX
);

  data.contact_email = maybeClampPlain(getAny(src, ["contact_email", "email"]), LIMITS.CONTACT_EMAIL_MAX);
  data.contact_location = maybeClampPlain(getAny(src, ["contact_location", "location", "ubicacion"]), LIMITS.CONTACT_LOCATION_MAX);
  data.contact_website = maybeClampPlain(getAny(src, ["contact_website", "website", "web"]), LIMITS.CONTACT_WEBSITE_MAX);
  data.licencia = safeStr(getAny(src, ["licencia", "licencia_conducir", "driver_license"], ""));

  for (let n = 1; n <= 7; n++) {
    data[`exp_${n}_dates`] = maybeClampPlain(getAny(src, [`exp_${n}_dates`]), LIMITS.EXP_DATES_MAX);
    data[`exp_${n}_company`] = maybeClampPlain(
  normalizeTechNames(canonicalizeInstitutionRobust(getAny(src, [`exp_${n}_company`]))),
  LIMITS.EXP_COMPANY_MAX
);


    
data[`exp_${n}_role`] = normalizeSeniorityWords(
  normalizeTechNames(
    replaceForbiddenWords(
      maybeClampPlain(getAny(src, [`exp_${n}_role`]), LIMITS.EXP_ROLE_MAX)
    )
  )
);



    for (let b = 1; b <= 5; b++) {
      data[`exp_${n}_b${b}`] = maybeClampPlain(
  normalizeTechNames(
    removeCommonRedundancies(
      stripPipesFromBullets(replaceForbiddenWords(getAny(src, [`exp_${n}_b${b}`])))

    )
  ),
  LIMITS.EXP_BULLET_MAX
);


    }

    segmentBulletsIfNeeded(data, n);
    ensureDefaultBulletIfEmpty(data, n, "El usuario no coloco Info");
    data[`exp_${n}_bullets_block`] = buildBulletsBlock(data, n);
  }

  
normalizeExperienceVisualGroup(data);
cleanExperienceRedundancy(data);
removeDuplicateBullets(data);


  for (let i = 1; i <= 7; i++) {
    data[`skill_${i}`] = maybeClampPlain(getAny(src, [`skill_${i}`]), LIMITS.SKILL_MAX);
  }

   // Educación (3 items)
  for (let i = 1; i <= 3; i++) {
    data[`edu_${i}_school`] = maybeClampPlain(
      canonicalizeInstitutionRobust(getAny(src, [`edu_${i}_school`])),
      LIMITS.EDU_SCHOOL_MAX
    );

    data[`edu_${i}_degree`] = maybeClampPlain(
      getAny(src, [`edu_${i}_degree`]),
      LIMITS.EDU_DEGREE_MAX
    );

    data[`edu_${i}_years`] = maybeClampPlain(
      getAny(src, [`edu_${i}_years`]),
      LIMITS.EDU_YEARS_MAX
    );
  }

  // ✅ esto lo habías sacado: dejalo acá
  normalizeEducationVisualGroup(data);

  // Idiomas / IT / Cursos raw -> campos (lo que ya tenías más abajo en tu versión original)
  const idiomasRaw = safeStr(getAny(src, ["idiomas_raw", "idiomas", "languages_raw"], ""));
  const itRaw = safeStr(getAny(src, ["it_raw", "it", "informatica_raw"], ""));
  const cursosRaw = safeStr(getAny(src, ["cursos_raw", "cursos", "courses_raw"], ""));

  const idiomasParts = splitByCommonDelimiters(idiomasRaw);
  const itParts = splitByCommonDelimiters(itRaw);
  const cursoParts = splitByCommonDelimiters(cursosRaw);

  



  for (let i = 1; i <= 3; i++) {
  const v0 = getAny(src, [`idioma_${i}`], idiomasParts[i - 1] || "");
  if (shouldDropLanguageItem(v0)) {
    data[`idioma_${i}`] = "";
    continue;
  }

  data[`idioma_${i}`] = maybeClampPlain(
    normalizeStatusWords(
      normalizeNivelWord(
        enforceLevelFormat(normalizeTechNames(v0), "idioma")
      )
    ),
    LIMITS.ITEM_MAX
  );
}


  for (let i = 1; i <= 6; i++) {
  const v0 = getAny(src, [`it_${i}`], itParts[i - 1] || "");
  if (shouldDropITItem(v0)) {
    data[`it_${i}`] = "";
    continue;
  }

  
 data[`it_${i}`] = maybeClampPlain(
  normalizeNivelWord(
    normalizeITNoParens(
      enforceLevelFormat(normalizePipes(normalizeTechNames(v0)), "it")
    )
  ),
  LIMITS.ITEM_MAX
);
}

// ✅ MICROSOFT OFFICE: promedio de niveles a partir de herramientas Office
(function collapseMicrosoftOfficeAvg() {
  const LEVEL_SCORE = (lvl) => {
    const x = stripAccents(safeStr(lvl)).toLowerCase().trim();
    if (!x) return null;

    if (x.includes("avanz")) return 3;
    if (x.includes("inter")) return 2;
    if (x.includes("promedio")) return 1.5;
    if (x.includes("basico") || x.includes("básico")) return 1;

    // equivalentes por si entran
    if (x === "alto") return 3;
    if (x === "medio") return 2;
    if (x === "bajo") return 1;

    return null;
  };

  const SCORE_LEVEL = (score) => {
    if (!Number.isFinite(score)) return "";
    if (score < 1.5) return "básico";
    if (score < 2.5) return "intermedio";
    return "avanzado";
  };

  const parseItLine = (line) => {
    const s = safeStr(line).replace(/\s+/g, " ").trim();
    if (!s) return { name: "", level: "" };

    // "X | nivel intermedio"
    let m = s.match(/^(.*?)\s*\|\s*nivel\s*([a-z0-9áéíóúüñ]+)\s*$/i);
    if (m) return { name: m[1].trim(), level: m[2].trim() };

    // "X (nivel intermedio)"
    m = s.match(/^(.*?)\s*\(\s*nivel\s*([^)]+)\s*\)\s*$/i);
    if (m) return { name: m[1].trim(), level: m[2].trim() };

    // "X (intermedio)" (asumimos nivel)
    m = s.match(/^(.*?)\s*\(\s*(avanzado|intermedio|promedio|basico|básico|alto|medio|bajo)\s*\)\s*$/i);
    if (m) return { name: m[1].trim(), level: m[2].trim() };

    // sin nivel
    return { name: s, level: "" };
  };

  const isOfficeTool = (name) => {
    const n = stripAccents(safeStr(name)).toLowerCase().replace(/\s+/g, " ").trim();
    if (!n) return false;

    // Incluímos Office apps típicas (con o sin "Microsoft")
    const keys = [
      "word",
      "excel",
      "powerpoint",
      "outlook",
      "access",
      "teams",
      "onenote",
      "publisher",
      "visio",
      "project",
      "microsoft word",
      "microsoft excel",
      "microsoft powerpoint",
      "microsoft outlook",
      "microsoft office",
      "office",
    ];

    return keys.some((k) => n === k || n.endsWith(" " + k) || n.includes(k));
  };

  // levantar items actuales
  const items = [];
  for (let i = 1; i <= 6; i++) {
    const v = safeStr(data[`it_${i}`]).trim();
    if (v) items.push(v);
  }
  if (!items.length) return;

  // detectar office tools y niveles
  const officeScores = [];
  const keep = [];

  for (const line of items) {
    const { name, level } = parseItLine(line);
    if (isOfficeTool(name)) {
      const sc = LEVEL_SCORE(level);
      if (sc !== null) officeScores.push(sc);
      // NO lo guardamos en keep (lo vamos a “colapsar” en Microsoft Office)
      continue;
    }
    keep.push(line);
  }

  // si había al menos 2 herramientas Office o 1 herramienta con nivel -> armamos Office
  if (officeScores.length >= 1) {
    const avg = officeScores.reduce((a, b) => a + b, 0) / officeScores.length;
    const lvl = SCORE_LEVEL(avg);
    const officeLine = lvl ? `Microsoft Office | nivel ${lvl}` : `Microsoft Office | nivel`;

    // Office arriba
    const final = [officeLine, ...keep].slice(0, 6);

    // reescribir it_1..it_6
    for (let i = 1; i <= 6; i++) data[`it_${i}`] = final[i - 1] || "";
  } else {
    // si no pudimos calcular, al menos normalizamos "Microsoft Office" al formato pipe
    // (por si quedó "Microsoft Office (nivel)")
    const final = items
      .map((x) => normalizeNivelWord(enforceLevelFormat(x, "it")))
      .slice(0, 6);
    for (let i = 1; i <= 6; i++) data[`it_${i}`] = final[i - 1] || "";
  }
})();



  for (let i = 1; i <= 6; i++) {
    const v = getAny(src, [`curso_${i}`], cursoParts[i - 1] || "");
    data[`curso_${i}`] = maybeClampPlain(normalizeCourseLine(v), LIMITS.ITEM_MAX);
  }

  data.sidebar_sections = buildSidebarSections(data, { underlineMin: 22, underlineExtra: 10, educationGapLines: 0 });
  data.sidebar_contact = buildSidebarContactSection(data, { underlineMin: 22, underlineExtra: 10 });
  data.sidebar_sections_noedu = buildSidebarSectionsNoEdu(data, { underlineMin: 22, underlineExtra: 15 });
  data.sidebar_sections_eduycursos = buildSidebarSectionsEduCursosOnly(data, {
    underlineMin: 22,
    underlineExtra: 10,
    educationGapLines: 0,
  });
  data.sidebar_sections_full = buildSidebarSectionsFull(data, { underlineMin: 22, underlineExtra: 15, educationGapLines: 0 });

  // Post normalizaciones de consistencia
data.title = normalizeRetailWord(normalizeSeniorityWords(data.title));
data.about = normalizeDegreesInText(normalizeRetailWord(normalizeSeniorityWords(data.about)));

// Asegura estado en minúscula (si algo se coló)
for (let i = 1; i <= 3; i++) data[`edu_${i}_years`] = normalizeStatusWords(data[`edu_${i}_years`]);
for (let i = 1; i <= 6; i++) data[`curso_${i}`] = normalizeStatusWords(data[`curso_${i}`]);


  return data;
} // ✅ CERRAR flattenToTemplateData ACA


function normalizePipes(v) {
  let s = safeStr(v);

  // normaliza espacios “raros” que a veces vienen desde formularios/IA
  s = s.replace(/\u00A0/g, " "); // nbsp

  // colapsa cualquier repetición de pipes con o sin espacios: "||", "| |", "|   |" => "|"
  // lo hacemos en loop porque a veces viene "|||"
  while (/\|\s*\|/.test(s)) s = s.replace(/\|\s*\|/g, "|");

  // normaliza espacios alrededor de pipe: "a|b" / "a |b" / "a| b" => "a | b"
  s = s.replace(/\s*\|\s*/g, " | ");

  // limpia pipe al inicio/fin (por si queda algo tipo "| nivel intermedio")
  s = s.replace(/^\s*\|\s*/g, "").replace(/\s*\|\s*$/g, "");

  // colapsa espacios
  s = s.replace(/\s{2,}/g, " ").trim();

  return s;
}

function normalizeITNoParens(value) {
  let v = safeStr(value).replace(/\s+/g, " ").trim();
  if (!v) return "";

  // ✅ Solo convertimos paréntesis SI son de nivel
  // "(nivel intermedio)" -> " | nivel intermedio"
  v = v.replace(/\(\s*nivel\s*([^)]+?)\s*\)\s*$/i, (_, lvl) => ` | nivel ${lvl.trim()}`);

  // ✅ "(intermedio)" -> " | nivel intermedio" (solo si parece nivel)
  v = v.replace(/\(\s*(avanzado|intermedio|promedio|basico|básico|alto|medio|bajo|a1|a2|b1|b2|c1|c2|c3|nativo|nativa)\s*\)\s*$/i,
    (_, lvl) => ` | nivel ${lvl.trim()}`
  );

  // ❌ NO convertir "(profesional)" ni otros paréntesis a pipe
  // ❌ NO borrar paréntesis internos

  // normaliza espacios alrededor de pipe
  v = v.replace(/\s*\|\s*/g, " | ").trim();

  // limpia pipe al inicio/fin si quedó raro
  v = v.replace(/^\s*\|\s*/g, "").replace(/\s*\|\s*$/g, "").trim();

  return v;
}


function enforceLevelFormat(value, type = "it") {
  
  let v = normalizePipes(value);
  if (!v) return "";

  v = v.replace(/\bNivel\b/g, "nivel");

  const LVL =
    "(avanzado|intermedio|promedio|basico|básico|alto|medio|bajo|a1|a2|b1|b2|c1|c2|c3|nativo|nativa|native)";

  function mapGenericLvl(lvlRaw) {
    const x = safeStr(lvlRaw).trim().toLowerCase();

    if (x === "alto") return "avanzado";
    if (x === "medio") return "intermedio";
    if (x === "bajo") return "básico";

    if (x === "nativa" || x === "native") return "nativo";

    if (x === "a1" || x === "a2") return "básico";
    if (x === "b1" || x === "b2") return "intermedio";
    if (x === "c1" || x === "c2" || x === "c3") return "avanzado";

    if (x === "basico" || x === "básico") return "básico";
    if (x === "promedio") return "promedio";
    if (x === "intermedio") return "intermedio";
    if (x === "avanzado") return "avanzado";

    return x;
  }


if (type === "it") {
  // 1) "Word - Nivel intermedio" → "Word | nivel intermedio"
  v = v.replace(
    new RegExp(`\\s*(?:-|–|—)\\s*nivel\\s*(${LVL})\\b`, "i"),
    (_, lvl) => ` | nivel ${mapGenericLvl(lvl)}`
  ).trim();

  // 2) "(nivel intermedio)" → "| nivel intermedio"
  v = v.replace(
    new RegExp(`\\(\\s*nivel\\s*(${LVL})\\s*\\)`, "i"),
    (_, lvl) => ` | nivel ${mapGenericLvl(lvl)}`
  ).trim();

  // 3) "(intermedio)" → "| nivel intermedio"
  v = v.replace(
    new RegExp(`\\(\\s*(${LVL})\\s*\\)\\s*$`, "i"),
    (_, lvl) => ` | nivel ${mapGenericLvl(lvl)}`
  ).trim();

  // 4) "Excel intermedio" → "Excel | nivel intermedio"
  const mTail = v.match(new RegExp(`^(.*?)(?:\\s+)(${LVL})\\s*$`, "i"));
  if (mTail && !/\|\s*nivel\b/i.test(v) && !/\(nivel\b/i.test(v)) {
    const name = mTail[1].trim();
    const lvl = mapGenericLvl(mTail[2]);
    if (name) return `${name} | nivel ${lvl}`;
  }

  // 5) Si ya viene con "| nivel ..." lo dejamos
  if (/\|\s*nivel\b/i.test(v)) return v;

  // --- SANITIZE FINAL (evita "| |" y bullets duplicados) ---

// a) si viene con bullet, lo sacamos acá (el bullet lo agrega el formateador final)
v = v.replace(/^\s*[•\-]\s+/, "").trim();

// b) normaliza pipes repetidos: "| |" o "||" => "|"
v = v.replace(/\|\s*\|+/g, "|").replace(/\s*\|\s*/g, " | ").trim();

// c) si quedó " | nivel ..." repetido dos veces, nos quedamos con una sola
// Ej: "Adobe | nivel avanzado | nivel avanzado" => "Adobe | nivel avanzado"
v = v.replace(
  /\s*\|\s*nivel\s+(basico|básico|intermedio|avanzado)\b(?:\s*\|\s*nivel\s+\1\b)+/i,
  " | nivel $1"
).trim();

// d) si por algún motivo quedó " | | nivel ..." lo arreglamos
v = v.replace(/\|\s*\|\s*(?=nivel\b)/i, "| ").trim();

  // 6) Si no tiene nivel explícito → lo dejamos SIN agregar nada
  return v;

}


  // =========================
  // ✅ IDIOMAS
  // =========================
  if (type === "idioma") {

    // 🔥 BORRAR cualquier estado tipo (en curso), (finalizado), etc.
    v = v.replace(/\(\s*(en\s*curso|finalizado|finalizada|incompleto|incompleta)\s*\)/gi, "").trim();

    // Evitar duplicado "| nivel"
    const occNivel = (v.match(/\|\s*nivel\b/gi) || []).length;
    if (occNivel >= 2 && /\|\s*nivel\s*$/i.test(v)) {
      v = v.replace(/\s*\|\s*nivel\s*$/i, "").trim();
    }

    // Inglés | intermedio
    let mPipeNoNivel = v.match(new RegExp(`^(.*?)\\s*\\|\\s*(${LVL})\\s*$`, "i"));
    if (mPipeNoNivel) {
      const name = mPipeNoNivel[1].trim();
      const lvl = mapGenericLvl(mPipeNoNivel[2]);
      return `${name} | nivel ${lvl}`;
    }

    // Inglés | nivel intermedio
    let mPipeNivel = v.match(new RegExp(`^(.*?)\\s*\\|\\s*nivel\\s*(${LVL})\\s*$`, "i"));
    if (mPipeNivel) {
      const name = mPipeNivel[1].trim();
      const lvl = mapGenericLvl(mPipeNivel[2]);
      return `${name} | nivel ${lvl}`;
    }

    // Inglés (b2)
    let mPar = v.match(new RegExp(`^(.*?)\\s*\\(\\s*(${LVL})\\s*\\)\\s*$`, "i"));
    if (mPar) {
      const name = mPar[1].trim();
      const lvl = mapGenericLvl(mPar[2]);
      return `${name} | nivel ${lvl}`;
    }

    // Inglés intermedio
    let mTail = v.match(new RegExp(`^(.*?)(?:\\s+)(${LVL})\\s*$`, "i"));
    if (mTail) {
      const name = mTail[1].trim();
      const lvl = mapGenericLvl(mTail[2]);
      return `${name} | nivel ${lvl}`;
    }

    if (/\|\s*nivel\b/i.test(v)) return v;

    return `${v} | nivel`;
  }

  // =========================
  // ✅ INFORMÁTICA
  // =========================

    // Normalizar paréntesis tipo "(Profesional)" => "(profesional)"
  v = v.replace(/\(\s*([^)]{2,40})\s*\)\s*$/i, (m, content) => {
    const clean = content.toLowerCase().trim();
    return `(${clean})`;
  });

  // Caso roto: "(profesional) (nivel)" => dejar solo "(profesional)"
  v = v.replace(/\)\s*\(\s*nivel\s*\)\s*$/i, ")").trim();

  // Si ya tiene un paréntesis final y NO es "(nivel ...)", no agregamos nada
  const hasAnyParen = /\(\s*[^)]{2,40}\s*\)\s*$/.test(v);
  const hasNivelInParen = /\(\s*nivel\b[^)]*\)\s*$/i.test(v);

  if (hasAnyParen && !hasNivelInParen) {
    return v;
  }


  if (new RegExp(`\\(\\s*nivel(\\s+${LVL})?\\s*\\)`, "i").test(v)) return v;
  if (new RegExp(`\\|\\s*nivel(\\s+${LVL})?\\b`, "i").test(v)) return v;
  if (new RegExp(`\\bnivel\\s+${LVL}\\b`, "i").test(v)) return v;

  if (new RegExp(`\\(\\s*${LVL}\\s*\\)`, "i").test(v)) return v;

  const m = v.match(new RegExp(`^(.*?)(?:\\s+)(${LVL})\\s*$`, "i"));
  if (m) {
    const name = m[1].trim();
    const lvl = mapGenericLvl(m[2]);
    if (!name) return v;
    return `${name} (nivel ${lvl})`;
  }

  return type === "it" ? `${v} | nivel` : `${v} (nivel)`;
}







/* =========================================================================
   7.1) PATCH SIDEBAR (markers) — FIX DEFINITIVO
   ========================================================================= */


function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function patchBodyPr(spXml, patchFn) {
  // ✅ soporta:
  // - <a:bodyPr ...>...</a:bodyPr>
  // - <a:bodyPr .../>
  if (/<a:bodyPr[^>]*>[\s\S]*?<\/a:bodyPr>/.test(spXml)) {
    return spXml.replace(/<a:bodyPr([^>]*)>([\s\S]*?)<\/a:bodyPr>/, (m, attrs, inner) => {
      const { attrsOut, innerOut } = patchFn(attrs, inner);
      return `<a:bodyPr${attrsOut}>${innerOut}</a:bodyPr>`;
    });
  }

  return spXml.replace(/<a:bodyPr([^>]*)\/>/, (m, attrs) => {
    const { attrsOut, innerOut } = patchFn(attrs, "");
    return `<a:bodyPr${attrsOut}>${innerOut}</a:bodyPr>`;
  });
}

function ensureWrap(attrs) {
  let out = attrs || "";

  // ✅ PPTX válido: wrap="square" (wrap="1" es inválido)
  if (!/\bwrap=/.test(out)) out += ` wrap="square"`;
  out = out.replace(/\bwrap="[^"]*"/g, `wrap="square"`);

  // ✅ clave: corta overflow horizontal (evita invadir hacia la derecha)
  if (!/\bhorzOverflow=/.test(out)) out += ` horzOverflow="clip"`;
  else out = out.replace(/\bhorzOverflow="[^"]*"/g, `horzOverflow="clip"`);

  // (opcional pero ayuda) evita que LO recorte raro verticalmente
  if (!/\bvertOverflow=/.test(out)) out += ` vertOverflow="overflow"`;
  else out = out.replace(/\bvertOverflow="[^"]*"/g, `vertOverflow="overflow"`);

  return out;
}


function removeAutofitChildren(inner) {
  return (inner || "").replace(/<a:(noAutofit|normAutofit|spAutoFit)\b[^\/]*\/>/g, "");
}

function forceNoAutofit(spXml) {
  return patchBodyPr(spXml, (attrs, inner) => {
    const attrsOut = ensureWrap(attrs);
    let innerOut = removeAutofitChildren(inner);
    if (!innerOut.includes("<a:noAutofit/>")) innerOut = `<a:noAutofit/>` + innerOut;
    return { attrsOut, innerOut };
  });
}

function forceSidebarBody(spXml) {
  return patchBodyPr(spXml, (attrs, inner) => {
    const attrsOut = ensureWrap(attrs);

    // para BODY: sacamos autofit y dejamos noAutofit
    let innerOut = removeAutofitChildren(inner);
    if (!innerOut.includes("<a:noAutofit/>")) innerOut = `<a:noAutofit/>` + innerOut;

    return { attrsOut, innerOut };
  });
}


function stripMarkerFromShape(spXml, marker) {
  const re = new RegExp(`<a:t>${escapeRegExp(marker)}</a:t>`, "g");
  return spXml.replace(re, "<a:t></a:t>");
}

function patchShapesByMarker(pptxBuffer, rules) {
  // rules: [{ marker, fn(spXml)->spXml }]
  const zip = new PizZip(pptxBuffer);

  // ✅ clave: también layouts y masters
  const targets = [
    ...(zip.file(/^ppt\/slides\/slide\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/) || []),
  ];

  for (const f of targets) {
    let xml = f.asText();
    let changed = false;

    xml = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (sp) => {
      for (const r of rules) {
        if (sp.includes(r.marker)) {
          changed = true;
          const cleaned = stripMarkerFromShape(sp, r.marker);
          return r.fn(cleaned);
        }
      }
      return sp;
    });

    if (changed) zip.file(f.name, xml);
  }

  return zip.generate({ type: "nodebuffer" });
}

// ✅ Compacta interlineado/espaciado de TODOS los <a:p> dentro de un shape
function forceCompactParagraphSpacingInShape(spXml, { lnPct = 90000, spcAft = 0, spcBef = 0 } = {}) {
  return spXml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, (p) => {
    // Asegura <a:pPr>
    if (!/<a:pPr\b/.test(p)) {
      p = p.replace(/<a:p>/, `<a:p><a:pPr></a:pPr>`);
    }

    // Limpia previos para no duplicar
    p = p
      .replace(/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/g, "")
      .replace(/<a:spcAft>[\s\S]*?<\/a:spcAft>/g, "")
      .replace(/<a:spcBef>[\s\S]*?<\/a:spcBef>/g, "");

    // Inserta compact spacing dentro de <a:pPr ...>
    p = p.replace(/<a:pPr([^>]*)>/, (m, attrs) => {
      return `<a:pPr${attrs}>` +
        `<a:spcBef><a:spcPts val="${spcBef}"/></a:spcBef>` +
        `<a:spcAft><a:spcPts val="${spcAft}"/></a:spcAft>` +
        `<a:lnSpc><a:spcPct val="${lnPct}"/></a:lnSpc>`;
    });

    return p;
  });
}

// ✅ Parchea shapes que contengan cualquiera de los placeholders indicados
function patchShapesByAnyPlaceholder(pptxBuffer, placeholders, fn) {
  const zip = new PizZip(pptxBuffer);

  const targets = [
    ...(zip.file(/^ppt\/slides\/slide\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/) || []),
    ...(zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/) || []),
  ];

  const needles = (placeholders || []).map((s) => String(s)).filter(Boolean);

  for (const f of targets) {
    let xml = f.asText();
    let changed = false;

    xml = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (sp) => {
      const hit = needles.some((n) => sp.includes(n));
      if (!hit) return sp;
      changed = true;
      return fn(sp);
    });

    if (changed) zip.file(f.name, xml);
  }

  return zip.generate({ type: "nodebuffer" });
}


/* =========================================================================
   8) RENDER PPTX
   ========================================================================= */

function renderPptxFromTemplate(templateBuf, data) {
  const templateId = Number(data.template_id || DEFAULT_TEMPLATE_ID || 1);
  const profile = getProfile(templateId);
  console.log("[PHOTO][DEBUG] templateId=", templateId, "profile.photoSize=", profile?.photoSize);

  const zip = new PizZip(templateBuf);

  let accentHex = resolveAccentHex(data.accent_color_raw);

  const DEFAULT_SIDEBAR_HEX = "C7C8CA";
  const DEFAULT_TEXT_HEX = "000000";

  const hasUserColor = !!accentHex;
  const sidebarHex = hasUserColor ? pickSidebarColorForWhiteText(accentHex) : DEFAULT_SIDEBAR_HEX;
  const textHex = hasUserColor ? pickTextColorForSidebar(sidebarHex, data.accent_color_raw) : DEFAULT_TEXT_HEX;

  data._sidebarHex = sidebarHex;
  data._sidebarTextHex = textHex;

  const iconTheme = textHex === "FFFFFF" ? "light" : "dark";
  data.contact_rows = buildContactRowsWithIcons(data, { iconTheme });

  const imageModule = new ImageModule({
    centered: false,

    getImage: (tagValue, tagName) => {
      if (tagName !== "photo" && !tagName.endsWith("_icon") && tagName !== "icon") return null;

      if (Buffer.isBuffer(tagValue)) return tagValue;

      if (typeof tagValue === "string" && tagValue.trim()) {
        const b = decodeBase64Image(tagValue);
        return b;
      }

      return null;
    },

    getSize: (img, tagValue, tagName) => {
      if (tagName === "photo") {
        const ps = profile?.photoSize;
        const w = Array.isArray(ps) && Number.isFinite(ps[0]) ? ps[0] : 520;
        const h = Array.isArray(ps) && Number.isFinite(ps[1]) ? ps[1] : 520;
        return [w, h];
      }

      if (tagName === "icon" || tagName.endsWith("_icon")) {
        return [24, 24];
      }

      return [0, 0];
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

  pptxBuf = replaceColorInAllXml(pptxBuf, SENTINEL_HEX, sidebarHex).buffer;
  pptxBuf = replaceColorInAllXml(pptxBuf, TEXT_SENTINEL_HEX, textHex).buffer;

  // ✅ FIX: noAutofit + wrap + horzOverflow clip (no invade)
  pptxBuf = patchShapesByMarker(pptxBuf, [
    { marker: "__SB_TITLE__", fn: forceNoAutofit },
    { marker: "__SB_BODY__", fn: forceNoAutofit }, // <- más seguro que "keep autofit"
  ]);

    // ✅ EDUCACIÓN: compactar interlineado en el shape que contiene placeholders edu_*
  pptxBuf = patchShapesByAnyPlaceholder(
    pptxBuf,
    ["{{edu_1_degree}}", "{{edu_1_school}}", "{{edu_1_years}}", "{{edu_2_degree}}", "{{edu_3_degree}}"],
    (sp) => forceCompactParagraphSpacingInShape(sp, { lnPct: 90000, spcAft: 0, spcBef: 0 })
  );


  pptxBuf = removeEmptyBulletedParagraphs(pptxBuf);

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
          new Error(`Error convirtiendo a PDF.\nsofficePath: ${SOFFICE_PATH}\nstderr: ${stderr}\nstdout: ${stdout}`)
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
   9) PPTX -> png
   ========================================================================= */

function convertPptxToPng(pptxPath, outDir) {
  return new Promise((resolve, reject) => {
    const args = [
      "--headless",
      "--nologo",
      "--nofirststartwizard",
      "--norestore",
      "--convert-to",
      "png",
      "--outdir",
      outDir,
      pptxPath,
    ];

    execFile(SOFFICE_PATH, args, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        return reject(
          new Error(`Error convirtiendo a PNG.\nsofficePath: ${SOFFICE_PATH}\nstderr: ${stderr}\nstdout: ${stdout}`)
        );
      }

      const pngs = fs
        .readdirSync(outDir)
        .filter((f) => f.toLowerCase().endsWith(".png"))
        .map((f) => path.join(outDir, f));

      if (!pngs.length) {
        return reject(new Error(`LibreOffice no generó PNGs en: ${outDir}`));
      }

      resolve(pngs.sort()[0]);
    });
  });
}

async function applyTiledWatermarkToJpg(
  inputPngPath,
  outputJpgPath,
  { text = "SOFIJOBS", color = "#9CA3AF", opacity = 0.22, fontSize = 44, angle = -28, stepX = 420, stepY = 240, fontFamily = "Arial" } = {}
) {
  const inputBuf = fs.readFileSync(inputPngPath);

  const img = sharp(inputBuf).rotate();
  const meta = await img.metadata();
  const W = Math.round(Number(meta.width));
  const H = Math.round(Number(meta.height));

  if (!Number.isFinite(W) || !Number.isFinite(H) || W <= 0 || H <= 0) {
    throw new Error(`No pude leer width/height del PNG. meta.width=${meta.width} meta.height=${meta.height} format=${meta.format}`);
  }

  const safeText = String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

  const bigW = W * 3;
  const bigH = H * 3;
  const cx = W / 2;
  const cy = H / 2;

  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs>
      <pattern id="wm" patternUnits="userSpaceOnUse" width="${stepX}" height="${stepY}">
        <text
          x="40"
          y="${Math.round(stepY * 0.70)}"
          fill="${color}"
          fill-opacity="${opacity}"
          font-family="${fontFamily}"
          font-size="${fontSize}"
          font-weight="700"
        >${safeText}</text>
      </pattern>
    </defs>

    <g transform="rotate(${angle} ${cx} ${cy})">
      <rect
        x="${-W}"
        y="${-H}"
        width="${bigW}"
        height="${bigH}"
        fill="url(#wm)"
      />
    </g>
  </svg>`;

  await img.composite([{ input: Buffer.from(svg) }]).jpeg({ quality: 92, mozjpeg: true }).toFile(outputJpgPath);
}

/* =========================================================================
   10) HELPERS NUEVOS: pipeline común para endpoints
   ========================================================================= */

function getSrcFromBody(body) {
  return body?.data && typeof body.data === "object"
    ? body.data
    : body?.fields && typeof body.fields === "object"
    ? body.fields
    : body || {};
}


// =========================================================================
// 5.X) CV 2 PÁGINAS: preferencia desde formulario -> control de variantes
// =========================================================================

const TWO_PAGES_ANSWERS = {
  ONE_PAGE: "No, por favor que sea de una página",
  PREFER_TWO: "Si, prefiero que esté bien desarrollado toda mi experiencia",
  ONLY_IF_NEEDED: "Sólo si lo consideran estrictamente necesario utilizar dos páginas",
};

function normalizeTwoPagesAnswer(raw) {
  const s = safeStr(raw).trim();
  if (!s) return "";

  // Normalización suave por si vienen tildes distintas o espacios
  const norm = stripAccents(s).toLowerCase().replace(/\s+/g, " ").trim();

  const one = stripAccents(TWO_PAGES_ANSWERS.ONE_PAGE).toLowerCase();
  const pref = stripAccents(TWO_PAGES_ANSWERS.PREFER_TWO).toLowerCase();
  const only = stripAccents(TWO_PAGES_ANSWERS.ONLY_IF_NEEDED).toLowerCase();

  if (norm === one) return "one_page";
  if (norm === pref) return "prefer_two_pages";
  if (norm === only) return "only_if_needed";

  return ""; // desconocido -> no forzamos
}

/**
 * Intenta encontrar la respuesta del formulario aunque cambie el key.
 * - Busca por keys comunes
 * - Si no, busca por texto de la pregunta dentro de las keys
 */
function getTwoPagesAnswerFromSrc(src) {
  if (!src || typeof src !== "object") return "";

  // 1) keys "probables"
  const direct = getAny(src, [
    "two_pages",
    "two_pages_preference",
    "cv_two_pages",
    "cv_pages",
    "dos_paginas",
    "dos_páginas",
    "preferencia_dos_paginas",
  ], "");
  if (direct) return direct;

  // 2) fallback: buscar por texto similar en el nombre de la pregunta/columna
  const keys = Object.keys(src);
  const patterns = [
    "te gustaria que tu cv tenga dos paginas",
    "te gustaría que tu cv tenga dos páginas",
    "cv tenga dos paginas",
    "cv tenga dos páginas",
    "dos paginas",
    "dos páginas",
  ];

  for (const k of keys) {
    const kn = stripAccents(String(k)).toLowerCase();
    if (patterns.some((p) => kn.includes(stripAccents(p).toLowerCase()))) {
      const v = src[k];
      if (v !== undefined && v !== null && String(v).trim() !== "") return v;
    }
  }

  return "";
}


async function buildPptxAndTmpFiles(body) {
  const src = getSrcFromBody(body);

  const templateId = body.template_id || body.template || src.template_id || DEFAULT_TEMPLATE_ID;

  const expCountRaw = countExperiencesFromSrc(src, 7);

  // Preferencia del formulario (2 páginas)
  const twoPagesAnswerRaw = getTwoPagesAnswerFromSrc(src);
  const twoPagesPref = normalizeTwoPagesAnswer(twoPagesAnswerRaw);

  const tIdNum = Number(templateId || DEFAULT_TEMPLATE_ID || 1);
  const templateAllows2Pages = TWO_PAGE_TEMPLATES.has(tIdNum);

  // Regla:
  // - "No..." => v5
  // - "Sólo si lo consideran necesario" => v5
  // - "Sí, prefiero..." => v6/v7 según expCount (si template lo soporta)
  let expCountForVariant;

  if (twoPagesPref === "prefer_two_pages") {
    // deja que llegue a v6/v7 si hay 6/7 experiencias y el template soporta 2 páginas
    expCountForVariant = templateAllows2Pages ? expCountRaw : Math.min(expCountRaw, 5);
  } else if (twoPagesPref === "one_page" || twoPagesPref === "only_if_needed") {
    // FORZAR v5
    expCountForVariant = 5;
  } else {
    // fallback: lógica actual (según expCount + cap)
    expCountForVariant = capExpCountForTemplate(templateId, expCountRaw);
  }

  // Aplicar cap final por template (por seguridad)
  expCountForVariant = capExpCountForTemplate(templateId, expCountForVariant);

  const templatePath = getTemplatePath(templateId, expCountForVariant);


  const templateBuf = fs.readFileSync(templatePath);

   const data = flattenToTemplateData(body);

  const tIdNum2 = Number(templateId || data.template_id || DEFAULT_TEMPLATE_ID || 1);
  const allows2 = TWO_PAGE_TEMPLATES.has(tIdNum2);

  // Si NO permite 2 hojas, truncamos a 5 siempre
  if (!allows2) {
    dropExperiencesFrom(data, 6, 7);
  } else {
    // Si el usuario forzó 1 página (o only_if_needed), también truncamos a 5
    const pref = data.two_pages_pref || normalizeTwoPagesAnswer(getTwoPagesAnswerFromSrc(src));
    if (pref === "one_page" || pref === "only_if_needed") {
      dropExperiencesFrom(data, 6, 7);
    }
  }



   console.log(
    `[TEMPLATE] template_id=${normalizeTemplateId(templateId)} ` +
    `expCountRaw=${expCountRaw} expCountVariant=${expCountForVariant} ` +
    `twoPagesPref=${twoPagesPref || data.two_pages_pref || ""} ` +
    `file=${path.basename(templatePath)}`
  );


  // ============================================================
  // ✅ FOTO: respetar preferencia del usuario (SIN FOTO)
  // ============================================================

  // Si el usuario pidió explícitamente sin foto, forzamos vacío
  if (data.wants_photo === false) {
    data.photo = "";
  } else {
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
  }

  const pptxBuf = renderPptxFromTemplate(templateBuf, data);

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "cv-"));
  const id = crypto.randomBytes(8).toString("hex");
  const pptxPath = path.join(tmpDir, `cv-${id}.pptx`);
  fs.writeFileSync(pptxPath, pptxBuf);

  const userNameRaw = data?.name || src?.name || src?.nombre || "";
  const fileBase = toSafeFilename(userNameRaw, "CV");

  return { tmpDir, pptxPath, data, src, fileBase };
}

async function buildPdfFromBody(body) {
  const { tmpDir, pptxPath, fileBase } = await buildPptxAndTmpFiles(body);
  const pdfPath = await convertPptxToPdf(pptxPath, tmpDir);
  const pdfBuf = fs.readFileSync(pdfPath);
  return { pdfBuf, fileBase };
}

async function buildJpgFromBody(body) {
  const { tmpDir, pptxPath, fileBase } = await buildPptxAndTmpFiles(body);

  const pngPath = await convertPptxToPng(pptxPath, tmpDir);
  const jpgOut = path.join(tmpDir, `${fileBase}.jpg`);

  await applyTiledWatermarkToJpg(pngPath, jpgOut, { text: "SOFIJOBS", color: "#9CA3AF", opacity: 0.22 });

  const jpgBuf = fs.readFileSync(jpgOut);
  return { jpgBuf, fileBase };
}

/* =========================================================================
   11) ENDPOINTS
   ========================================================================= */

   app.post("/generate-pdf", async (req, res) => {
    try {
      const body = req.body || {};
  
      // Genera PPTX (templating + foto + color + parches)
      const { pptxPath, fileBase } = await buildPptxAndTmpFiles(body);
  
      const pptxBuf = fs.readFileSync(pptxPath);
  
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${fileBase}.pptx"`
      );
  
      return res.send(pptxBuf);
    } catch (err) {
      console.error(err);
      res.status(500).json({
        error: String(err?.message || err),
        stack: String(err?.stack || ""),
      });
    }
  });

app.post("/generate-only-pdf", async (req, res) => {
  try {
    const { pdfBuf, fileBase } = await buildPdfFromBody(req.body || {});
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fileBase}.pdf"`);
    return res.send(pdfBuf);
  } catch (err) {
    console.error(err);
    res.status(500).json({
      error: String(err?.message || err),
      stack: String(err?.stack || ""),
    });
  }
});

app.post("/generate-only-jpg", async (req, res) => {
  try {
    const { jpgBuf, fileBase } = await buildJpgFromBody(req.body || {});
    res.setHeader("Content-Type", "image/jpeg");
    res.setHeader("Content-Disposition", `attachment; filename="${fileBase}.jpg"`);
    return res.send(jpgBuf);
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
