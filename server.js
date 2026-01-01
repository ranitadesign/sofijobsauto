// server.js
// ✅ Genera PDF desde PPTX con Docxtemplater + ImageModule
// ✅ FOTO: en el PPTX usar placeholder EXACTO: {{%photo}}
// ✅ Texto: usar {{name}}, {{title}}, etc.
// ✅ Acepta foto por: photo_base64 (recomendado) o photo_url

const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { execFile } = require("child_process");
const http = require("http");
const https = require("https");

const express = require("express");
const cors = require("cors");

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
    ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    : "soffice";

const SOFFICE_PATH = process.env.SOFFICE_PATH || DEFAULT_SOFFICE;

// Carpeta de plantillas PPTX
const TEMPLATES_DIR = path.join(__dirname, "templates");

// Template default (si no mandás template_id)
const DEFAULT_TEMPLATE_ID = process.env.DEFAULT_TEMPLATE_ID || "1";

/* =========================================================================
   1) LIMITES GLOBALES (modo “no cortar por código”)
   ========================================================================= */

const LIMITS = {
  // ✅ si está en false: el server NO recorta nada (sólo limpia/normaliza espacios)
  ENABLE_CLAMP: false,

  // Valores altos por si activás ENABLE_CLAMP = true
  NAME_MAX_CHARS: 200,
  TITLE_MAX_CHARS: 240,
  ABOUT_MAX_CHARS_DEFAULT: 4000,

  CONTACT_EMAIL_MAX: 180,
  CONTACT_PHONE_MAX: 80,
  CONTACT_LOCATION_MAX: 240,
  CONTACT_WEBSITE_MAX: 220,

  EXP_ROLE_MAX: 400,
  EXP_COMPANY_MAX: 400,
  EXP_DATES_MAX: 120,
  EXP_BULLET_MAX: 800,

  EDU_SCHOOL_MAX: 380,
  EDU_DEGREE_MAX: 380,
  EDU_YEARS_MAX: 120,

  SKILL_MAX: 200,
  ITEM_MAX: 260, // idiomas/it/cursos
};

function safeStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize("NFC");
}

// “plain”: no mete \n (mejor para que el textbox haga wrap natural)
function maybeClampPlain(s, maxChars) {
  const text = safeStr(s).replace(/\s+/g, " ").trim();
  if (!text) return "";
  if (!LIMITS.ENABLE_CLAMP) return text;
  if (!maxChars || maxChars <= 0) return text;
  return text.length > maxChars ? text.slice(0, Math.max(0, maxChars - 1)).trimEnd() + "…" : text;
}

// “lines”: sólo si querés forzar líneas (yo lo evito para ABOUT)
function maybeClampLines(s, maxCharsPerLine, maxLines) {
  const text = safeStr(s).replace(/\s+/g, " ").trim();
  if (!text) return "";
  if (!LIMITS.ENABLE_CLAMP) return text; // sin clamp: devolvemos sin forzar \n
  const words = text.split(" ");
  const lines = [];
  let line = "";

  for (const w of words) {
    const next = line ? line + " " + w : w;
    if (next.length <= maxCharsPerLine) {
      line = next;
    } else {
      if (line) lines.push(line);
      line = w;
      if (lines.length >= maxLines) break;
    }
  }
  if (lines.length < maxLines && line) lines.push(line);

  let out = lines.join("\n");
  const rebuilt = lines.join(" ");
  if (rebuilt.length < text.length) out = out.trimEnd() + "…";
  return out;
}

function getAny(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

function toArrayFromFlat(obj, prefix, count) {
  const out = [];
  for (let i = 1; i <= count; i++) out.push(safeStr(obj?.[`${prefix}${i}`]).trim());
  return out.filter(Boolean);
}

function splitByCommonDelimiters(s) {
  const raw = safeStr(s).trim();
  if (!raw) return [];
  return raw
    .split(/[\n,;•]+/g)
    .map((x) => x.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

/* =========================================================================
   2) PERFIL POR PLANTILLA (ABOUT)
   ========================================================================= */

const TEMPLATE_PROFILES = {
  // Por ahora: todos iguales (alto). Después ajustamos 1 por 1.
  1: { about: { mode: "plain", maxChars: 4000 } },
  2: { about: { mode: "plain", maxChars: 4000 } },
  3: { about: { mode: "plain", maxChars: 4000 } },
  4: { about: { mode: "plain", maxChars: 4000 } },
  5: { about: { mode: "plain", maxChars: 4000 } },
  6: { about: { mode: "plain", maxChars: 4000 } },
  7: { about: { mode: "plain", maxChars: 4000 } },
  8: { about: { mode: "plain", maxChars: 4000 } },
  9: { about: { mode: "plain", maxChars: 4000 } },
  10: { about: { mode: "plain", maxChars: 4000 } },
  11: { about: { mode: "plain", maxChars: 4000 } },
  12: { about: { mode: "plain", maxChars: 4000 } },
  13: { about: { mode: "plain", maxChars: 4000 } },
  14: { about: { mode: "plain", maxChars: 4000 } },
  default: { about: { mode: "plain", maxChars: LIMITS.ABOUT_MAX_CHARS_DEFAULT } },
};

function getProfile(templateId) {
  const id = Number((templateId ?? DEFAULT_TEMPLATE_ID).toString().trim());
  return TEMPLATE_PROFILES[id] || TEMPLATE_PROFILES.default;
}

/* =========================================================================
   3) IMÁGENES
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
    const u = safeStr(url).trim();
    if (!u) return resolve(null);

    const finalUrl = normalizeGoogleDriveUrl(u);
    const lib = finalUrl.startsWith("https://") ? https : http;

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

/* =========================================================================
   4) TEMPLATES
   ========================================================================= */

const TEMPLATE_MAP = {
  1: "Plantilla_oficial_1_verde.pptx",
  2: "Template_2_moderno.pptx",
  3: "Template_3_oficial.pptx",
  4: "PONER_NOMBRE_REAL_4.pptx",
  5: "PONER_NOMBRE_REAL_5.pptx",
  6: "PONER_NOMBRE_REAL_6.pptx",
  7: "PONER_NOMBRE_REAL_7.pptx",
  8: "PONER_NOMBRE_REAL_8.pptx",
  9: "PONER_NOMBRE_REAL_9.pptx",
  10: "Currículum Vitae Cv de Marketing Minimalista Beige (2).pptx",
  11: "PONER_NOMBRE_REAL_11.pptx",
  12: "PONER_NOMBRE_REAL_12.pptx",
  13: "PONER_NOMBRE_REAL_13.pptx",
  14: "PONER_NOMBRE_REAL_14.pptx",
};

function getTemplatePath(templateId) {
  const raw = (templateId || DEFAULT_TEMPLATE_ID || "").toString().trim();
  if (!raw) throw new Error("Falta template_id y no hay DEFAULT_TEMPLATE_ID");

  const id = Number(raw);
  if (!Number.isFinite(id) || id < 1 || id > 14) {
    throw new Error(`template_id inválido: "${raw}". Debe ser un número 1..14.`);
  }

  const fileName = TEMPLATE_MAP[id];
  if (!fileName) throw new Error(`No hay mapping para template_id=${id}. Revisá TEMPLATE_MAP.`);

  const templatePath = path.join(TEMPLATES_DIR, fileName);

  if (!fs.existsSync(templatePath)) {
    let list = [];
    try {
      list = fs.readdirSync(TEMPLATES_DIR).filter((f) => f.toLowerCase().endsWith(".pptx"));
    } catch (_) {}
    throw new Error(
      `No encuentro la plantilla: ${fileName} en ${TEMPLATES_DIR}. Disponibles: ${list.join(", ")}`
    );
  }

  return templatePath;
}

/* =========================================================================
   5) ANTI-UNDEFINED: completar claves faltantes con ""
   ========================================================================= */

function fillMissingKeys(data) {
  // Base
  const baseKeys = [
    "template_id",
    "name",
    "title",
    "about",
    "contact_phone",
    "contact_email",
    "contact_location",
    "contact_website",
    "idiomas_raw",
    "it_raw",
    "cursos_raw",
    "photo_base64",
    "photo_url",
  ];
  for (const k of baseKeys) if (!(k in data)) data[k] = "";

  // Skills 1..7
  for (let i = 1; i <= 7; i++) if (!(`skill_${i}` in data)) data[`skill_${i}`] = "";

  // Idiomas/IT/Cursos 1..5
  for (let i = 1; i <= 5; i++) {
    if (!(`idioma_${i}` in data)) data[`idioma_${i}`] = "";
    if (!(`it_${i}` in data)) data[`it_${i}`] = "";
    if (!(`curso_${i}` in data)) data[`curso_${i}`] = "";
  }

  // Educación 1..3
  for (let n = 1; n <= 3; n++) {
    for (const k of [`edu_${n}_school`, `edu_${n}_degree`, `edu_${n}_years`]) {
      if (!(k in data)) data[k] = "";
    }
  }

  // Experiencia 1..3 + bullets b1..b9 (para evitar undefined en templates “largos”)
  for (let n = 1; n <= 3; n++) {
    for (const k of [`exp_${n}_company`, `exp_${n}_role`, `exp_${n}_dates`]) {
      if (!(k in data)) data[k] = "";
    }
    for (let b = 1; b <= 9; b++) {
      const kb = `exp_${n}_b${b}`;
      if (!(kb in data)) data[kb] = "";
    }
  }

  // Foto (tag imagen)
  if (!("photo" in data)) data.photo = null;

  return data;
}

/* =========================================================================
   6) DATA MAPPING
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
  data.photo_url = safeStr(getAny(src, ["photo_url"])); // IMPORTANT: no tocar
  data.photo_base64 = safeStr(getAny(src, ["photo_base64"]));
  data.photo = null;

  // Nombre / title (sin recorte por default; si activás ENABLE_CLAMP, corta alto)
  data.name = maybeClampPlain(getAny(src, ["name"]), LIMITS.NAME_MAX_CHARS);
  data.title = maybeClampPlain(getAny(src, ["title"]), LIMITS.TITLE_MAX_CHARS);

  // ✅ ABOUT por plantilla (preferimos plain para que haga wrap natural)
  const templateId = getAny(src, ["template_id", "template", "templateId"], DEFAULT_TEMPLATE_ID);
  const profile = getProfile(templateId);
  const aboutCfg = profile?.about || TEMPLATE_PROFILES.default.about;

  if (aboutCfg.mode === "lines") {
    data.about = maybeClampLines(
      getAny(src, ["about"]),
      aboutCfg.maxCharsPerLine || 120,
      aboutCfg.maxLines || 12
    );
  } else {
    data.about = maybeClampPlain(getAny(src, ["about"]), aboutCfg.maxChars || LIMITS.ABOUT_MAX_CHARS_DEFAULT);
  }

  data.contact_phone = maybeClampPlain(getAny(src, ["contact_phone"]), LIMITS.CONTACT_PHONE_MAX);
  data.contact_email = maybeClampPlain(getAny(src, ["contact_email"]), LIMITS.CONTACT_EMAIL_MAX);
  data.contact_location = maybeClampPlain(getAny(src, ["contact_location"]), LIMITS.CONTACT_LOCATION_MAX);
  data.contact_website = maybeClampPlain(getAny(src, ["contact_website"]), LIMITS.CONTACT_WEBSITE_MAX);

  // Experiencia 1..3
  for (let n = 1; n <= 3; n++) {
    data[`exp_${n}_company`] = maybeClampPlain(getAny(src, [`exp_${n}_company`]), LIMITS.EXP_COMPANY_MAX);
    data[`exp_${n}_role`] = maybeClampPlain(getAny(src, [`exp_${n}_role`]), LIMITS.EXP_ROLE_MAX);
    data[`exp_${n}_dates`] = maybeClampPlain(getAny(src, [`exp_${n}_dates`]), LIMITS.EXP_DATES_MAX);

    // Bullets b1..b9 (si no vienen, quedan vacíos por fillMissingKeys)
    for (let b = 1; b <= 9; b++) {
      data[`exp_${n}_b${b}`] = maybeClampPlain(getAny(src, [`exp_${n}_b${b}`]), LIMITS.EXP_BULLET_MAX);
    }
  }

  // Skills 1..7
  let skills = [];
  if (Array.isArray(src.skills)) skills = src.skills.map((x) => safeStr(x).trim()).filter(Boolean);
  if (!skills.length) skills = toArrayFromFlat(src, "skill_", 7);
  for (let i = 0; i < 7; i++) data[`skill_${i + 1}`] = maybeClampPlain(skills[i] || "", LIMITS.SKILL_MAX);

  // Educación 1..3
  for (let n = 1; n <= 3; n++) {
    data[`edu_${n}_school`] = maybeClampPlain(getAny(src, [`edu_${n}_school`]), LIMITS.EDU_SCHOOL_MAX);
    data[`edu_${n}_degree`] = maybeClampPlain(getAny(src, [`edu_${n}_degree`]), LIMITS.EDU_DEGREE_MAX);
    data[`edu_${n}_years`] = maybeClampPlain(getAny(src, [`edu_${n}_years`]), LIMITS.EDU_YEARS_MAX);
  }

  // RAW + PARTES (idiomas/it/cursos)
  data.idiomas_raw = maybeClampPlain(getAny(src, ["idiomas_raw"]), 5000);
  data.it_raw = maybeClampPlain(getAny(src, ["it_raw"]), 5000);
  data.cursos_raw = maybeClampPlain(getAny(src, ["cursos_raw"]), 5000);

  // Si ya vienen items, los respetamos; si no, derivamos del raw
  const anyIdioma = safeStr(getAny(src, ["idioma_1", "idioma_2", "idioma_3"])).trim();
  const anyIt = safeStr(getAny(src, ["it_1", "it_2", "it_3"])).trim();
  const anyCurso = safeStr(getAny(src, ["curso_1", "curso_2", "curso_3"])).trim();

  const idiomasParts = splitByCommonDelimiters(anyIdioma ? "" : data.idiomas_raw);
  const itParts = splitByCommonDelimiters(anyIt ? "" : data.it_raw);
  const cursoParts = splitByCommonDelimiters(anyCurso ? "" : data.cursos_raw);

  for (let i = 1; i <= 5; i++) {
    data[`idioma_${i}`] = maybeClampPlain(getAny(src, [`idioma_${i}`], idiomasParts[i - 1] || ""), LIMITS.ITEM_MAX);
    data[`it_${i}`] = maybeClampPlain(getAny(src, [`it_${i}`], itParts[i - 1] || ""), LIMITS.ITEM_MAX);
    data[`curso_${i}`] = maybeClampPlain(getAny(src, [`curso_${i}`], cursoParts[i - 1] || ""), LIMITS.ITEM_MAX);
  }

  return fillMissingKeys(data);
}

/* =========================================================================
   7) RENDER PPTX
   ========================================================================= */

function renderPptxFromTemplate(templateBuf, data) {
  const zip = new PizZip(templateBuf);

  const imageModule = new ImageModule({
    centered: false,
    getImage: (tagValue, tagName) => {
      if (tagName !== "photo") return null;
      if (Buffer.isBuffer(tagValue)) return tagValue;
      if (typeof tagValue === "string" && tagValue.trim()) return decodeBase64Image(tagValue);
      return null;
    },
    getSize: () => [520, 520],
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    modules: [imageModule],

    // ✅ CLAVE: si falta un placeholder → vacío (nunca "undefined")
    nullGetter: () => "",
  });

  doc.setData(data);
  doc.render();
  return doc.getZip().generate({ type: "nodebuffer" });
}

/* =========================================================================
   8) PPTX -> PDF
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
   9) ENDPOINT
   ========================================================================= */

app.post("/generate-pdf", async (req, res) => {
  try {
    const body = req.body || {};

    const templateId = body.template_id || body.template || DEFAULT_TEMPLATE_ID;
    const templatePath = getTemplatePath(templateId);
    const templateBuf = fs.readFileSync(templatePath);

    const data = flattenToTemplateData(body);

    // ✅ FOTO: prioridad base64, si no URL
    if (data.photo_base64) {
      data.photo = data.photo_base64;
    } else if (data.photo_url) {
      const buf = await fetchBufferFromUrl(data.photo_url);
      if (buf && buf.length) data.photo = buf;
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
});
