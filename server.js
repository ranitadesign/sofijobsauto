// server.js
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

// ✅ módulo de imágenes
const ImageModule = require("docxtemplater-image-module-free");

const app = express();
app.use(cors());
app.use(express.json({ limit: "40mb" })); // subimos por imágenes base64

app.get("/health", (_, res) => res.json({ ok: true }));

/**
 * LibreOffice / soffice:
 * - En Windows: ruta completa suele ser esa
 * - En Linux/Render/Railway: normalmente es "soffice"
 */
const DEFAULT_SOFFICE =
  process.platform === "win32"
    ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    : "soffice";

const SOFFICE_PATH = process.env.SOFFICE_PATH || DEFAULT_SOFFICE;

// Carpeta de plantillas
const TEMPLATES_DIR = path.join(__dirname, "templates");

// Template fallback (si no mandás template_id)
const DEFAULT_TEMPLATE_ID = process.env.DEFAULT_TEMPLATE_ID || "Template_1_clasico";

/* ---------------- utils texto ---------------- */

function safeStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize("NFC");
}

function clamp(s, max) {
  s = safeStr(s);
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

/** Lee un valor desde varias keys alternativas */
function getAny(body, keys, fallback = "") {
  for (const k of keys) {
    const v = body?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

/** Construye array desde campos planos tipo skill_1..skill_7 */
function toArrayFromFlat(body, prefix, maxN) {
  const out = [];
  for (let i = 1; i <= maxN; i++) {
    const v = body?.[`${prefix}${i}`];
    if (v !== undefined && v !== null && String(v).trim() !== "") out.push(v);
  }
  return out;
}

/* ---------------- utils imagen ---------------- */

/**
 * Acepta:
 * - data:image/png;base64,....
 * - base64 puro
 */
function decodeBase64Image(b64) {
  if (!b64) return null;
  const s = String(b64).trim();

  const m = s.match(/^data:(image\/\w+);base64,(.+)$/i);
  if (m) return Buffer.from(m[2], "base64");

  // base64 puro
  return Buffer.from(s, "base64");
}

function fetchBufferFromUrl(url) {
  return new Promise((resolve, reject) => {
    if (!url) return resolve(null);
    const u = String(url).trim();
    const lib = u.startsWith("https://") ? https : http;

    lib
      .get(u, (resp) => {
        const code = resp.statusCode || 0;
        if (code >= 300 && code < 400 && resp.headers.location) {
          // redirect
          return resolve(fetchBufferFromUrl(resp.headers.location));
        }
        if (code !== 200) {
          return reject(new Error(`No pude descargar imagen. HTTP ${code}`));
        }
        const chunks = [];
        resp.on("data", (d) => chunks.push(d));
        resp.on("end", () => resolve(Buffer.concat(chunks)));
      })
      .on("error", reject);
  });
}

/* ---------------- templates ---------------- */

function getTemplatePath(templateId) {
  const id = (templateId || DEFAULT_TEMPLATE_ID || "").trim();
  if (!id) throw new Error("Falta template_id y no hay DEFAULT_TEMPLATE_ID");

  // permitimos que te manden "Template_1_clasico" o "Template_1_clasico.pptx"
  const fileName = id.toLowerCase().endsWith(".pptx") ? id : `${id}.pptx`;
  const templatePath = path.join(TEMPLATES_DIR, fileName);

  if (!fs.existsSync(templatePath)) {
    // Debug rápido de qué hay en templates/
    let list = [];
    try {
      list = fs.readdirSync(TEMPLATES_DIR);
    } catch (_) {}
    throw new Error(
      `No encuentro la plantilla: ${fileName} en ${TEMPLATES_DIR}. Disponibles: ${list.join(", ")}`
    );
  }
  return templatePath;
}

/* ---------------- mapeo data ---------------- */

/**
 * ✅ Soporta:
 * A) ANIDADO: { contact:{...}, experience:[...], skills:[...], education:[...], photo_url/photo_base64 }
 * B) PLANO:   { contact_email, exp_1_company, skill_1, edu_1_school, photo_url/photo_base64 }
 * C) ENVUELTO:{ data: {...} }
 */
function flattenToTemplateData(body) {
  const src = body?.data && typeof body.data === "object" ? body.data : body || {};
  const data = {};

  // Header
  data.name = clampLines(getAny(src, ["name", "Nombre completo", "nombre", "full_name"]), 22, 2);
  data.title = clampLines(getAny(src, ["title", "Objetivo / rol buscado", "Puesto", "rol"]), 28, 2);

  // 👇 PERFIL: acá es donde tocás los límites (antes 52,4)
  // Recomendación: subir de a poco. Un valor "estable" suele ser 80 chars/linea y 6 lineas
  data.about = clampLines(getAny(src, ["about", "Resumen profesional", "resumen"]), 80, 6);

  // Contacto
  const c = src.contact && typeof src.contact === "object" ? src.contact : {};
  data.contact_phone = clamp(
    getAny(src, ["contact_phone", "Telefono", "Teléfono", "phone"], getAny(c, ["phone"])),
    22
  );
  data.contact_email = clamp(
    getAny(src, ["contact_email", "Email", "email"], getAny(c, ["email"])),
    32
  );
  data.contact_location = clampLines(
    getAny(src, ["contact_location", "Ubicacion", "Ubicación", "location"], getAny(c, ["location"])),
    40,
    2
  );
  data.contact_website = clamp(
    getAny(src, ["contact_website", "Linkedin", "Portfolio", "GITHUB", "website"], getAny(c, ["website"])),
    40
  );

  // Educación (2)
  const education = Array.isArray(src.education) ? src.education : [];
  for (let i = 0; i < 2; i++) {
    const n = i + 1;
    const row = education[i] || {};

    const years = getAny(src, [`edu_${n}_years`], row.years || row.dates || row.period);
    const school = getAny(src, [`edu_${n}_school`], row.institution || row.school || row.institucion);
    const degree = getAny(src, [`edu_${n}_degree`], row.degree || row.area || row.career || row.carrera);

    data[`edu_${n}_years`] = clamp(years, 40);
    data[`edu_${n}_school`] = clampLines(school, 34, 2);
    data[`edu_${n}_degree`] = clampLines(degree, 34, 2);
  }

  // Skills (7)
  let skills = Array.isArray(src.skills) ? src.skills : [];
  if (!skills.length) skills = toArrayFromFlat(src, "skill_", 7);
  for (let i = 0; i < 7; i++) data[`skill_${i + 1}`] = clamp(skills[i], 26);

  // Experiencia (2, 3 bullets)
  const exp = Array.isArray(src.experience) ? src.experience : [];
  for (let i = 0; i < 2; i++) {
    const n = i + 1;
    const e = exp[i] || {};

    const company = getAny(src, [`exp_${n}_company`], e.company);
    const role = getAny(src, [`exp_${n}_role`], e.role);
    const dates = getAny(src, [`exp_${n}_dates`], e.dates);

    data[`exp_${n}_company`] = clampLines(company, 26, 2);
    data[`exp_${n}_role`] = clampLines(role, 30, 2);
    data[`exp_${n}_dates`] = clamp(dates, 30);

    let bullets = Array.isArray(e.bullets) ? e.bullets : [];
    if (!bullets.length) {
      bullets = [src[`exp_${n}_b1`], src[`exp_${n}_b2`], src[`exp_${n}_b3`]].filter(Boolean);
    }

    const b = asBullets(bullets, 3, 42);
    for (let j = 0; j < 3; j++) data[`exp_${n}_b${j + 1}`] = b[j] || "";
  }

  // ✅ FOTO (lo importante)
  // Tu PPTX debe tener {{photo}} en un textbox donde va la imagen.
  data.photo_url = getAny(src, ["photo_url", "Foto URL", "foto_url"]);
  data.photo_base64 = getAny(src, ["photo_base64", "Foto Base64", "foto_base64"]);

  // El tag real que usa el PPTX:
  // {{photo}}
  // El ImageModule llama a getImage con (tagValue, tagName)
  // acá le pasamos un objeto para que getImage tenga todo.
  data.photo = {
    url: data.photo_url || "",
    base64: data.photo_base64 || "",
  };

  return data;
}

/* ---------------- render docxtemplater + imagen ---------------- */

function renderPptxFromTemplate(templateBuf, data) {
  const zip = new PizZip(templateBuf);

  // ImageModule config
  const imageModule = new ImageModule({
    centered: false,
    // tagValue: lo que le pasamos en data.photo
    getImage: (tagValue, tagName) => {
      if (tagName !== "photo") return null;

      const b = tagValue?.base64 ? decodeBase64Image(tagValue.base64) : null;
      if (b && b.length) return b;

      // Si no hay base64, intentamos URL (pero ojo: esto debería ser sync en este módulo)
      // Para mantenerlo simple y robusto: SOLO base64.
      // Si querés URL sí o sí, lo resolvemos antes (en /generate-pdf) y lo metemos como base64.
      return null;
    },
    // tamaño de la foto final (px). Ajustalo a tu plantilla.
    // Si tu foto queda chica/grande, cambiá estos números.
    getSize: () => [220, 220],
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    modules: [imageModule],
  });

  doc.setData(data);

  try {
    doc.render();
  } catch (err) {
    const details =
      err?.properties?.errors?.map((e) => ({
        id: e.id,
        explanation: e?.properties?.explanation,
        tag: e?.properties?.xtag,
      })) || [];
    throw new Error(
      `Error de template (PPTX): ${err.message}\nDETAILS: ${JSON.stringify(details, null, 2)}`
    );
  }

  return doc.getZip().generate({ type: "nodebuffer" });
}

/* ---------------- libreoffice convert ---------------- */

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

    // 1) elegir plantilla
    const templateId = body.template_id || body.template || DEFAULT_TEMPLATE_ID;
    const templatePath = getTemplatePath(templateId);
    const templateBuf = fs.readFileSync(templatePath);

    // 2) armar data
    const data = flattenToTemplateData(body);

    // 3) si vino photo_url, la convertimos a base64 para que ImageModule sea robusto
    if (!data.photo?.base64 && data.photo?.url) {
      const buf = await fetchBufferFromUrl(data.photo.url);
      if (buf && buf.length) {
        data.photo.base64 = buf.toString("base64");
      }
    }

    // 4) render PPTX
    const pptxBuf = renderPptxFromTemplate(templateBuf, data);

    // 5) escribir temp + convertir
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
