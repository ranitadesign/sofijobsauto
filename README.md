# Sofijobs automation

## Postproceso de experiencias (PPTX)

Algunas plantillas soportan **separaci?n autom?tica del bloque de experiencia laboral**. Tras generar el PPTX, un script en Python busca textboxes que contengan el marker exacto `[[EXPERIENCE_BLOCK]]`, parsea el contenido (encabezados tipo `ROL | EMPRESA | FECHA` y bullets) y los reemplaza por varios textboxes (rol, empresa, fecha y bullets por experiencia).

- **Si la plantilla no tiene** el marker `[[EXPERIENCE_BLOCK]]` en ning?n textbox, el flujo sigue igual y se devuelve el PPTX sin cambios.
- **Requisito:** Python instalado y dependencias: `pip install -r scripts/requirements.txt` (en Windows: `python -m pip install -r scripts/requirements.txt` si `pip` no está en el PATH).
- **Variable de entorno:** `PYTHON_BIN` permite forzar el intérprete (opcional). El servidor resuelve automáticamente: (1) `PYTHON_BIN` si está definido, (2) `python3`, (3) `python`. En Render/Linux suele usarse `python3` sin configurar nada.
- **Render:** En el servicio, añadir en Build Command: `pip install -r scripts/requirements.txt` para que el postproceso tenga `python-pptx`. Opcional: `PYTHON_BIN=python3` si el runtime usa otro nombre.
- **Render con Docker:** La imagen ya incluye Python 3 y python-pptx (vía `requirements.txt` en la raíz). En el servicio, setear **`PYTHON_BIN=python3`** para que el postproceso use el binario correcto.

El postproceso se aplica al endpoint que devuelve PPTX (`POST /generate-pdf`). No cambia endpoints ni el contrato con n8n.

## Postproceso del sidebar (PPTX)

Algunas plantillas soportan **separaci�n del bloque lateral (sidebar)** en textboxes por secci�n. En el textbox que contiene todo el sidebar (cursos, inform�tica, idiomas, etc.) se debe colocar el marker exacto `[[SIDEBAR_BLOCK]]`. El script `scripts/split_sidebar_blocks.py` lo detecta, parsea secciones (t�tulo, l�nea de subrayado, body) y recrea el sidebar en textboxes independientes. Si no hay marker, se copia el archivo sin cambios. El pipeline es: 1) generaci�n PPTX, 2) split_experience_blocks.py, 3) split_sidebar_blocks.py, 4) respuesta.
