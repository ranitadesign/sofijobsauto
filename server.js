// server.js
// ✅ Genera PDF desde PPTX con Docxtemplater + ImageModule
// ✅ FOTO: en el PPTX usar placeholder EXACTO: {{%photo}}
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

// Template default (si no mandás template_id) -> ahora es numérico
const DEFAULT_TEMPLATE_ID = process.env.DEFAULT_TEMPLATE_ID || "1";

/* ---------------- utils texto ---------------- */

function safeStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize("NFC");
}

function clamp(s, max) {
  s = safeStr(s).trim();
  if (!s) return "";
  return s.length > max ? s.slice(0, Math.max(0, max - 1)).trimEnd() + "…" : s;
}

function clampLines(s, maxCharsPerLine, maxLines) {
  s = safeStr(s).replace(/\s+/g, " ").trim();
  if (!s) return "";
  const words = s.split(" ");
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
  if (rebuilt.length < s.length) out = out.trimEnd() + "…";
  return out;
}

function asBullets(arr, maxItems, maxCharsEach) {
  const a = Array.isArray(arr) ? arr : [];
  return a
    .filter(Boolean)
    .slice(0, maxItems)
    .map((x) => clampLines(x, maxCharsEach, 2));
}

function getAny(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

function toArrayFromFlat(obj, prefix, maxN) {
  const out = [];
  for (let i = 1; i <= maxN; i++) {
    const v = obj?.[`${prefix}${i}`];
    if (v !== undefined && v !== null && String(v).trim() !== "") out.push(v);
  }
  return out;
}

/* ---------------- utils imagen ---------------- */

function decodeBase64Image(b64) {
  if (!b64) return null;
  const s = String(b64).trim();

  // data:image/jpeg;base64,....
  const m = s.match(/^data:(image\/\w+);base64,(.+)$/i);
  if (m) return Buffer.from(m[2], "base64");

  // base64 puro
  return Buffer.from(s, "base64");
}

// Convierte links de Drive a descarga directa (para photo_url)
function normalizeGoogleDriveUrl(url) {
  if (!url) return "";
  const u = String(url).trim();

  // file/d/ID
  const m1 = u.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1?.[1]) return `https://drive.google.com/uc?export=download&id=${m1[1]}`;

  // open?id=ID
  const m2 = u.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (m2?.[1]) return `https://drive.google.com/uc?export=download&id=${m2[1]}`;

  // cualquier URL con ?id=ID
  const idMatch = u.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch?.[1]) return `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;

  // ya es uc
  if (u.includes("drive.google.com/uc")) return u;

  return u;
}


function fetchBufferFromUrl(url) {
  return new Promise((resolve, reject) => {
    if (!url) return resolve(null);

    const finalUrl = normalizeGoogleDriveUrl(String(url).trim());
    const lib = finalUrl.startsWith("https://") ? https : http;

    const req = lib.get(
      finalUrl,
      { headers: { "User-Agent": "Mozilla/5.0 (CV-Generator)", Accept: "*/*" } },
      (resp) => {
        const code = resp.statusCode || 0;

        // redirects
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

/* ---------------- templates (MAPPING) ---------------- */

// 🔥 ACÁ ES DONDE DEFINÍS QUÉ ARCHIVO USA CADA MODELO
// Poné EXACTAMENTE el nombre del .pptx tal como está dentro de /templates
const TEMPLATE_MAP = {
  1: "Plantilla_oficial_1_verde.pptx",
  2: "Template_2_moderno.pptx",

  // 👇 completá estos con tus nombres reales
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
  if (!fileName) {
    throw new Error(`No hay mapping para template_id=${id}. Revisá TEMPLATE_MAP.`);
  }

  const templatePath = path.join(TEMPLATES_DIR, fileName);

  if (!fs.existsSync(templatePath)) {
    let list = [];
    try {
      list = fs
        .readdirSync(TEMPLATES_DIR)
        .filter((f) => f.toLowerCase().endsWith(".pptx"));
    } catch (_) {}
    throw new Error(
      `No encuentro la plantilla: ${fileName} en ${TEMPLATES_DIR}. Disponibles: ${list.join(", ")}`
    );
  }

  return templatePath;
}

/* ---------------- data mapping ---------------- */

function flattenToTemplateData(body) {
  const src =
    body?.data && typeof body.data === "object"
      ? body.data
      : body?.fields && typeof body.fields === "object"
      ? body.fields
      : body || {};

  const data = {};

  // OJO: template_id acá queda en data solo por si lo querés loguear o usar en pptx
  data.template_id = getAny(src, ["template_id", "template", "Plantilla de CV"], "");

  data.name = clampLines(getAny(src, ["name", "Nombre completo", "full_name"]), 22, 2);
  data.title = clampLines(getAny(src, ["title", "Objetivo / rol buscado", "Puesto"]), 28, 2);
  data.about = clampLines(getAny(src, ["about", "Resumen profesional"]), 80, 6);

  data.contact_phone = clamp(getAny(src, ["contact_phone", "Telefono", "Teléfono"]), 22);
  data.contact_email = clamp(getAny(src, ["contact_email", "Email"]), 60);
  data.contact_location = clampLines(getAny(src, ["contact_location", "Ubicacion", "Ubicación"]), 40, 2);
  data.contact_website = clamp(getAny(src, ["contact_website", "Linkedin", "Portfolio", "GITHUB"]), 80);

  // Experiencia 1/2
  for (let n = 1; n <= 2; n++) {
    data[`exp_${n}_company`] = clampLines(getAny(src, [`exp_${n}_company`]), 26, 2);
    data[`exp_${n}_role`] = clampLines(getAny(src, [`exp_${n}_role`]), 30, 2);
    data[`exp_${n}_dates`] = clamp(getAny(src, [`exp_${n}_dates`]), 30);

    data[`exp_${n}_b1`] = clampLines(getAny(src, [`exp_${n}_b1`]), 42, 2);
    data[`exp_${n}_b2`] = clampLines(getAny(src, [`exp_${n}_b2`]), 42, 2);
    data[`exp_${n}_b3`] = clampLines(getAny(src, [`exp_${n}_b3`]), 42, 2);
  }

  // Skills 1..7
  let skills = Array.isArray(src.skills) ? src.skills : [];
  if (!skills.length) skills = toArrayFromFlat(src, "skill_", 7);
  for (let i = 0; i < 7; i++) data[`skill_${i + 1}`] = clamp(skills[i], 26);

  // Educación 1/2
  for (let n = 1; n <= 2; n++) {
    data[`edu_${n}_school`] = clampLines(getAny(src, [`edu_${n}_school`]), 34, 2);
    data[`edu_${n}_degree`] = clampLines(getAny(src, [`edu_${n}_degree`]), 34, 2);
    data[`edu_${n}_years`] = clamp(getAny(src, [`edu_${n}_years`]), 40);
  }

  // Foto inputs
  data.photo_url = getAny(src, ["photo_url", "photo" , "archivos_main"], "");
  data.photo_base64 = getAny(src, ["photo_base64"], "");

  // clave del tag de imagen
  data.photo = null;

  return data;
}

/* ---------------- render PPTX ---------------- */

function renderPptxFromTemplate(templateBuf, data) {
  const zip = new PizZip(templateBuf);

  const imageModule = new ImageModule({
    centered: false,
    getImage: (tagValue, tagName) => {
      if (tagName !== "photo") return null;
      if (Buffer.isBuffer(tagValue)) return tagValue;
      if (typeof tagValue === "string" && tagValue.trim()) {
        return decodeBase64Image(tagValue);
      }
      return null;
    },
    getSize: () => [220, 220],
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    // ⚠️ IMPORTANTE:
    // Si usás placeholders {{name}} para texto, entonces la imagen debe ser {{%photo}}
    delimiters: { start: "{{", end: "}}" },
    modules: [imageModule],
  });

  doc.setData(data);
  doc.render();
  return doc.getZip().generate({ type: "nodebuffer" });
}

/* ---------------- PPTX -> PDF ---------------- */

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

/* ---------------- endpoint ---------------- */

app.post("/generate-pdf", async (req, res) => {
  try {
    const body = req.body || {};

    // ahora template_id esperado: "1".."14"
    const templateId = body.template_id || body.template || DEFAULT_TEMPLATE_ID;
    const templatePath = getTemplatePath(templateId);
    const templateBuf = fs.readFileSync(templatePath);

    const data = flattenToTemplateData(body);

    // ✅ FOTO (prioridad base64, que es tu flujo actual)
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
    res.status(500).json({ error: String(err.message || err) });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CV API OK en http://127.0.0.1:${PORT}`);
  console.log(`Templates dir: ${TEMPLATES_DIR}`);
  console.log(`LibreOffice: ${SOFFICE_PATH}`);
});
