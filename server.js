// server.js
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { execFile } = require("child_process");

const express = require("express");
const cors = require("cors");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

app.get("/health", (_, res) => res.json({ ok: true }));

const DEFAULT_SOFFICE =
  process.platform === "win32"
    ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    : "/usr/bin/soffice";

const SOFFICE_PATH = process.env.SOFFICE_PATH || DEFAULT_SOFFICE;


// ✅ Default template (fallback)
const DEFAULT_TEMPLATE_PATH = path.join(__dirname, "template.pptx");

// ✅ Folder donde guardás las 15 plantillas
const TEMPLATES_DIR = path.join(__dirname, "templates");

// (Opcional) allowlist para evitar que te pidan cualquier archivo
// Si querés, lo dejás vacío y solo validás por regex.
const ALLOWED_TEMPLATES = new Set([
  // "template_01_clasico",
  // "template_02_moderno",
  // "template_03_minimal",
]);

function safeStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize("NFC");
}

function ensureFileExists(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`No encuentro ${label} en: ${filePath}`);
  }
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

function getAny(body, keys, fallback = "") {
  for (const k of keys) {
    const v = body?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

function toArrayFromFlat(body, prefix, maxN) {
  const out = [];
  for (let i = 1; i <= maxN; i++) {
    const v = body?.[`${prefix}${i}`];
    if (v !== undefined && v !== null && String(v).trim() !== "") out.push(v);
  }
  return out;
}

/**
 * ✅ Lee template_id desde:
 * - body.template_id
 * - body.template_name
 * - body["Plantilla de CV"] (del Form/Sheet)
 *
 * Y devuelve una ruta segura.
 */
function resolveTemplatePath(body) {
  const src =
    body?.data && typeof body.data === "object" ? body.data : body || {};

  const raw = getAny(src, ["template_id", "template_name", "Plantilla de CV"], "");
  const templateId = safeStr(raw).trim();

  // Si no mandaron nada => default
  if (!templateId) {
    ensureFileExists(DEFAULT_TEMPLATE_PATH, "template.pptx (default)");
    return DEFAULT_TEMPLATE_PATH;
  }

  // Sanitización fuerte: solo letras/números/_/-
  if (!/^[a-zA-Z0-9_-]+$/.test(templateId)) {
    throw new Error(`template_id inválido: "${templateId}". Usá solo letras, números, _ o -`);
  }

  // (Opcional) allowlist
  if (ALLOWED_TEMPLATES.size > 0 && !ALLOWED_TEMPLATES.has(templateId)) {
    throw new Error(`template_id no permitido: "${templateId}"`);
  }

  const candidate = path.join(TEMPLATES_DIR, `${templateId}.pptx`);
  ensureFileExists(candidate, `template "${templateId}.pptx"`);
  return candidate;
}

/**
 * A) ANIDADO: { contact:{...}, experience:[...], skills:[...], education:[...] }
 * B) PLANO:   { contact_email, exp_1_company, skill_1, edu_1_school, ... }
 * C) ENVUELTO:{ data: {...} }
 */
function flattenToTemplateData(body) {
  const src =
    body?.data && typeof body.data === "object" ? body.data : body || {};

  const data = {};

  // Header
  data.name = clampLines(
    getAny(src, ["name", "Nombre completo", "nombre", "full_name"]),
    22,
    2
  );
  data.title = clampLines(
    getAny(src, ["title", "Objetivo / rol buscado", "Puesto", "rol"]),
    28,
    2
  );

  // ✅ PERFIL (esto es lo que estabas tocando)
  // Ajustalo desde acá:
  data.about = clampLines(
    getAny(src, ["about", "Resumen profesional", "resumen"]),
    80, // maxCharsPerLine
    6   // maxLines
  );

  // Contacto
  const c = src.contact && typeof src.contact === "object" ? src.contact : {};
  data.contact_phone = clamp(
    getAny(src, ["contact_phone", "Telefono", "Teléfono", "phone"], getAny(c, ["phone"])),
    22
  );
  data.contact_email = clamp(
    getAny(src, ["contact_email", "Email", "email"], getAny(c, ["email"])),
    28
  );
  data.contact_location = clampLines(
    getAny(
      src,
      ["contact_location", "Ubicacion", "Ubicación", "location"],
      getAny(c, ["location"])
    ),
    40,
    2
  );
  data.contact_website = clamp(
    getAny(src, ["contact_website", "Linkedin", "Portfolio", "GITHUB", "website"], getAny(c, ["website"])),
    40
  );

  // Educación
  const education = Array.isArray(src.education) ? src.education : [];
  for (let i = 0; i < 2; i++) {
    const n = i + 1;
    const row = education[i] || {};

    const years = getAny(src, [`edu_${n}_years`], row.years || row.dates || row.period);
    const school = getAny(
      src,
      [`edu_${n}_school`],
      row.institution || row.school || row.institucion || src["Institucion"]
    );
    const degree = getAny(
      src,
      [`edu_${n}_degree`],
      row.degree || row.area || row.career || row.carrera || src["Carreara"]
    );

    data[`edu_${n}_years`] = clamp(years || src["Inicio y final"], 40);
    data[`edu_${n}_school`] = clampLines(school, 30, 2);
    data[`edu_${n}_degree`] = clampLines(degree, 30, 2);
  }

  // Skills 
  let skills = Array.isArray(src.skills) ? src.skills : [];
  if (!skills.length) skills = toArrayFromFlat(src, "skill_", 7);
  for (let i = 0; i < 7; i++) data[`skill_${i + 1}`] = clamp(skills[i], 22);

  // Experiencia
  const exp = Array.isArray(src.experience) ? src.experience : [];
  for (let i = 0; i < 2; i++) {
    const n = i + 1;
    const e = exp[i] || {};

    const company = getAny(src, [`exp_${n}_company`], e.company || src["Empresa"]);
    const role = getAny(src, [`exp_${n}_role`], e.role || src["Puesto"]);
    const dates = getAny(src, [`exp_${n}_dates`], e.dates || "");

    data[`exp_${n}_company`] = clampLines(company, 22, 2);
    data[`exp_${n}_role`] = clampLines(role, 26, 2);
    data[`exp_${n}_dates`] = clamp(dates, 30);

    let bullets = Array.isArray(e.bullets) ? e.bullets : [];
    if (!bullets.length) {
      bullets = [src[`exp_${n}_b1`], src[`exp_${n}_b2`], src[`exp_${n}_b3`]].filter(Boolean);
    }

    const b = asBullets(bullets, 3, 34);
    for (let j = 0; j < 3; j++) data[`exp_${n}_b${j + 1}`] = b[j] || "";
  }

  // Referencias
  const refs = Array.isArray(src.reference) ? src.reference : [];
  for (let i = 0; i < 2; i++) {
    const n = i + 1;
    const r = refs[i] || {};

    data[`ref_${n}_name`] = clampLines(getAny(src, [`ref_${n}_name`], r.name), 16, 2);
    data[`ref_${n}_role`] = clampLines(getAny(src, [`ref_${n}_role`], r.role), 16, 2);
    data[`ref_${n}_phone`] = clamp(getAny(src, [`ref_${n}_phone`], r.phone), 16);
  }

  return data;
}

function renderPptxFromTemplate(templateBuf, data) {
  const zip = new PizZip(templateBuf);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
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

function fileUrlFromPath(p) {
  // convierte C:\algo\carpeta en file:///C:/algo/carpeta (para LibreOffice)
  const norm = p.replace(/\\/g, "/");
  return `file:///${norm}`;
}

function convertPptxToPdf(pptxPath, outDir, profileDir) {
  return new Promise((resolve, reject) => {
    // IMPORTANTÍSIMO: perfil aislado para evitar locks
    const userInstall = `-env:UserInstallation=${fileUrlFromPath(profileDir)}`;

    const args = [
      "--headless",
      "--nologo",
      "--nofirststartwizard",
      "--norestore",
      "--invisible",
      userInstall,
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
            `Error convirtiendo a PDF.\n` +
            `sofficePath: ${SOFFICE_PATH}\n` +
            `args: ${JSON.stringify(args)}\n` +
            `stderr: ${stderr}\nstdout: ${stdout}\n` +
            `error: ${error.message}`
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


app.post("/generate-pdf", async (req, res) => {
  try {
    // ✅ Elegimos plantilla según el body
    const templatePath = resolveTemplatePath(req.body || {});
    const templateBuf = fs.readFileSync(templatePath);

    const data = flattenToTemplateData(req.body || {});
    const pptxBuf = renderPptxFromTemplate(templateBuf, data);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "cv-"));
    const id = crypto.randomBytes(8).toString("hex");
    const pptxPath = path.join(tmpDir, `cv-${id}.pptx`);
    fs.writeFileSync(pptxPath, pptxBuf);

    const loProfileDir = path.join(tmpDir, "lo-profile");
fs.mkdirSync(loProfileDir, { recursive: true });

const pdfPath = await convertPptxToPdf(pptxPath, tmpDir, loProfileDir);

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
  console.log(`Default Template: ${DEFAULT_TEMPLATE_PATH}`);
  console.log(`Templates Dir: ${TEMPLATES_DIR}`);
  console.log(`LibreOffice: ${SOFFICE_PATH}`);
});
