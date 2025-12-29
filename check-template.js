const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");

const pptxPath = path.join(__dirname, "template.pptx");
const buf = fs.readFileSync(pptxPath);
const zip = new PizZip(buf);

const slide = zip.files["ppt/slides/slide1.xml"]?.asText();
if (!slide) {
  console.error("No encontré ppt/slides/slide1.xml");
  process.exit(1);
}

const textRuns = [...slide.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map(m => m[1]);

// buscamos tags rotos: aparece "{{" y antes de cerrar "}}" aparece otro "{{"
let open = 0;
let errors = [];

for (let i = 0; i < textRuns.length; i++) {
  const t = textRuns[i];
  const opens = (t.match(/\{\{/g) || []).length;
  const closes = (t.match(/\}\}/g) || []).length;

  open += opens;
  open -= closes;

  if (open > 1) {
    errors.push({ i, t });
    open = 0; // reseteo para no spamear infinito
  }
}

console.log("Total text runs:", textRuns.length);
console.log("Runs con '{{' o '}}':", textRuns.filter(t => t.includes("{{") || t.includes("}}")).length);

if (!errors.length) {
  console.log("OK: no detecté doble-apertura obvia. Si Docxtemplater sigue fallando, el problema es un placeholder partido entre runs.");
  console.log("Te muestro los runs con llaves para inspección:");
  textRuns
    .map((t, idx) => ({ idx, t }))
    .filter(x => x.t.includes("{{") || x.t.includes("}}"))
    .forEach(x => console.log(`[${x.idx}]`, x.t));
  process.exit(0);
}

console.log("Posibles tags rotos (doble '{{' sin cerrar):");
errors.forEach(e => console.log(`[run ${e.i}]`, e.t));
