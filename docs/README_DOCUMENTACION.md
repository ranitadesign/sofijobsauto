# Documentación del flujo CV PPTX API

En esta carpeta están la documentación completa del flujo y cómo obtener el PDF.

## Archivos

- **DOCUMENTACION_FLUJO_CV_PPTX_API.md** — Documentación en Markdown (completa y detallada).
- **DOCUMENTACION_FLUJO_CV_PPTX_API.html** — Misma documentación en HTML para leer en el navegador.

## Cómo obtener el PDF

### Opción 1: Desde el HTML (recomendado)

1. Abrí **DOCUMENTACION_FLUJO_CV_PPTX_API.html** en Chrome o Edge.
2. Menú **Archivo → Imprimir** (o Ctrl+P).
3. En destino elegí **Guardar como PDF**.
4. Guardá el archivo (por ejemplo `DOCUMENTACION_FLUJO_CV_PPTX_API.pdf`).

### Opción 2: Generar el PDF con Node

Desde la raíz del proyecto:

```bash
npm run docs:pdf
```

La primera vez puede tardar varios minutos (descarga de Chromium). El PDF se genera en esta misma carpeta con el nombre `DOCUMENTACION_FLUJO_CV_PPTX_API.pdf`.

### Regenerar el HTML

Si editás el Markdown, podés regenerar el HTML con:

```bash
npm run docs:html
```
