/**
 * Convierte DOCUMENTACION_FLUJO_CV_PPTX_API.md a HTML para abrir en navegador
 * y guardar como PDF (Archivo > Imprimir > Guardar como PDF).
 */
const fs = require("fs");
const path = require("path");
const { marked } = require("marked");

const root = path.join(__dirname, "..");
const mdPath = path.join(root, "docs", "DOCUMENTACION_FLUJO_CV_PPTX_API.md");
const htmlPath = path.join(root, "docs", "DOCUMENTACION_FLUJO_CV_PPTX_API.html");

const md = fs.readFileSync(mdPath, "utf8");
const body = marked.parse(md);

const html = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Documentación CV PPTX API</title>
  <style>
    body { font-family: 'Segoe UI', system-ui, sans-serif; max-width: 900px; margin: 0 auto; padding: 2rem; line-height: 1.5; color: #1a1a1a; }
    h1 { border-bottom: 2px solid #c0504d; padding-bottom: 0.3em; }
    h2 { margin-top: 1.8em; border-bottom: 1px solid #ddd; padding-bottom: 0.2em; }
    h3 { margin-top: 1.2em; }
    code { background: #f4f4f4; padding: 0.15em 0.4em; border-radius: 4px; font-size: 0.9em; }
    pre { background: #f4f4f4; padding: 1rem; overflow-x: auto; border-radius: 6px; }
    pre code { background: none; padding: 0; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    th, td { border: 1px solid #ddd; padding: 0.5em 0.75em; text-align: left; }
    th { background: #f0f0f0; }
    hr { border: none; border-top: 1px solid #ddd; margin: 2em 0; }
    @media print { body { max-width: none; } }
  </style>
</head>
<body>
${body}
</body>
</html>
`;

fs.writeFileSync(htmlPath, html, "utf8");
console.log("Generado:", htmlPath);
console.log("Abrí este archivo en el navegador y usá Archivo > Imprimir > Guardar como PDF para obtener el .pdf");
