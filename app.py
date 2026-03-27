import os
import json
import uuid
import tempfile
from pathlib import Path
from typing import Optional

import anthropic
import openpyxl
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

load_dotenv(override=True)

app = FastAPI(title="OrderEAT Product Upload Agent")
static_dir = Path(__file__).parent / "static"
static_dir.mkdir(exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

UPLOAD_DIR = Path(tempfile.gettempdir()) / "ordereat_uploads"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "ordereat_outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_COLUMNS = ["TÍTULO", "DESCRIPCIÓN", "CATEGORÍA", "PRECIO", "FECHAS DISPONIBLES"]

SYSTEM_PROMPT = """Eres un asistente especializado en procesar listados de productos para la plataforma OrderEAT.
Tu tarea es tomar texto desordenado (menús, listas de productos, chats de WhatsApp, contenido de PDFs/Excels)
y extraer una lista estructurada de productos.

REGLAS:

1. Cada producto debe tener: título, descripción, categoría, precio y fechas disponibles.

2. TÍTULO: Nombre limpio y claro del producto. Primera letra en mayúscula. Sin símbolos raros.
   - Normalizar nombres: "Coca Cola 600ml" → correcto. "coca 600", "COCA COLA 600ML" → incorrecto.
   - Si es un menú del día, incluir el día: "Menú Lunes - Milanesa con puré".

3. DESCRIPCIÓN: Obligatoria siempre.
   - Si hay información adicional relevante (ingredientes, guarnición, incluye postre, tamaño, etc.), incluirla.
   - Si NO hay información adicional, colocar "-".

4. CATEGORÍA:
   Primero verificar si el texto de entrada ya define categorías o secciones (por ejemplo: encabezados como "Bebidas", "Platos del día", agrupaciones explícitas, etc.). Si las categorías ya están definidas en el archivo, respetar esas categorías tal cual aparecen.

   Solo si NO hay categorías definidas en el texto de entrada, inferir usando estas cuatro categorías:
   - "Menú": Productos que tienen una fecha específica asociada (menú del día, menú semanal con fechas).
   - "Comidas": Platos de almuerzo (hamburguesas, milanesas, pizzas, empanadas, sandwiches, ensaladas, etc.).
   - "Snacks": Productos chicos o de merienda/receso (galletas, alfajores, papas fritas, cereales, frutos secos, golosinas, barras de cereal, etc.).
   - "Bebidas": Cualquier bebida (agua, gaseosas, jugos, café, té, leche, etc.).

   Si no es posible determinar la categoría ni del texto ni por inferencia, dejar vacío.

5. PRECIO: Solo el número, sin símbolos de moneda ("$", "UYU", "$U", "CLP", etc.).
   - Si no hay precio, poner "REVISAR".

6. FECHAS DISPONIBLES: Si se menciona una fecha específica para el producto (ej: "lunes 3 de marzo"), ponerla en formato DD/MM/AAAA.
   - Si no hay fecha, dejar vacío.

7. Ignorar: saludos, comentarios, conversaciones, emojis, texto no relacionado con productos.

8. Evitar duplicados exactos (mismo nombre y precio).

9. Si hay varios precios para el mismo producto (ej: por tamaño), crear una fila por cada variante.

FORMATO DE RESPUESTA:
Devuelve ÚNICAMENTE un JSON array. Cada elemento tiene estas keys exactas:
- "titulo": string
- "descripcion": string ("-" si no hay info adicional)
- "categoria": string (vacío si no se puede determinar)
- "precio": string (número o "REVISAR")
- "fechas_disponibles": string (vacío si no aplica)

Ejemplo:
[
  {"titulo": "Hamburguesa Simple", "descripcion": "-", "categoria": "Comidas", "precio": "250", "fechas_disponibles": ""},
  {"titulo": "Coca Cola 600ml", "descripcion": "-", "categoria": "Bebidas", "precio": "150", "fechas_disponibles": ""},
  {"titulo": "Menú Lunes - Milanesa con Puré", "descripcion": "Incluye postre: fruta de estación", "categoria": "Menú", "precio": "4300", "fechas_disponibles": "03/03/2026"}
]

NO incluyas explicaciones, solo el JSON array."""


def extract_text_from_pdf(file_path: str) -> str:
    text_parts = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_parts.append(text)
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row:
                        text_parts.append(" | ".join(str(cell) for cell in row if cell))

    combined = "\n".join(text_parts).strip()

    # Fallback: if pdfplumber got nothing, try OCR
    if not combined:
        try:
            images = convert_from_path(file_path, dpi=300)
            ocr_parts = []
            for img in images:
                ocr_text = pytesseract.image_to_string(img, lang="spa")
                if ocr_text and ocr_text.strip():
                    ocr_parts.append(ocr_text.strip())
            combined = "\n".join(ocr_parts)
        except Exception as e:
            combined = f"[Error en OCR: {str(e)}]"

    return combined


def extract_text_from_excel(file_path: str) -> str:
    text_parts = []
    wb = openpyxl.load_workbook(file_path, data_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        text_parts.append(f"--- Hoja: {sheet_name} ---")
        for row in ws.iter_rows(values_only=True):
            cells = [str(cell) for cell in row if cell is not None]
            if cells:
                text_parts.append(" | ".join(cells))
    return "\n".join(text_parts)


def extract_text_from_csv(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def extract_content_from_file(file_path: str, filename: str) -> str:
    ext = Path(filename).suffix.lower()
    try:
        if ext == ".pdf":
            return extract_text_from_pdf(file_path)
        elif ext in (".xlsx", ".xls"):
            return extract_text_from_excel(file_path)
        elif ext in (".csv", ".tsv"):
            return extract_text_from_csv(file_path)
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                return f.read()
        else:
            return f"[Archivo no soportado: {filename}]"
    except Exception as e:
        return f"[Error procesando {filename}: {str(e)}]"


def generate_excel(products: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"

    # Header styling matching the original template
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    header_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Data rows
    data_font = Font(name="Arial", size=11)
    for row_idx, product in enumerate(products, 2):
        values = [
            product.get("titulo", ""),
            product.get("descripcion", ""),
            product.get("categoria", ""),
            product.get("precio", ""),
            product.get("fechas_disponibles", ""),
        ]
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border
            if col_idx == 4 and value and value != "REVISAR":
                try:
                    cell.value = float(value)
                    cell.number_format = "#,##0"
                except ValueError:
                    pass

    # Column widths
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 22

    wb.save(output_path)


def call_claude(text: str, business_type: str = "") -> list[dict]:
    client = anthropic.Anthropic()

    user_message = f"""Procesá el siguiente contenido y extraé todos los productos con sus precios.

{f'Tipo de comercio: {business_type}' if business_type else ''}

CONTENIDO A PROCESAR:
{text}"""

    response = client.messages.create(
        model="claude-opus-4-20250514",
        max_tokens=8192,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )

    response_text = response.content[0].text.strip()

    # Extract JSON from response (handle markdown code blocks)
    if response_text.startswith("```"):
        lines = response_text.split("\n")
        json_lines = []
        in_block = False
        for line in lines:
            if line.startswith("```") and not in_block:
                in_block = True
                continue
            elif line.startswith("```") and in_block:
                break
            elif in_block:
                json_lines.append(line)
        response_text = "\n".join(json_lines)

    products = json.loads(response_text)
    return products


@app.post("/process")
async def process_products(
    files: list[UploadFile] = File(default=[]),
    text: str = Form(default=""),
    business_type: str = Form(default=""),
):
    all_text_parts = []

    # Process uploaded files
    for file in files:
        if not file.filename or file.size == 0:
            continue
        file_id = uuid.uuid4().hex
        ext = Path(file.filename).suffix
        temp_path = UPLOAD_DIR / f"{file_id}{ext}"
        content = await file.read()
        with open(temp_path, "wb") as f:
            f.write(content)

        extracted = extract_content_from_file(str(temp_path), file.filename)
        all_text_parts.append(f"--- Archivo: {file.filename} ---\n{extracted}")
        temp_path.unlink(missing_ok=True)

    # Add free text
    if text.strip():
        all_text_parts.append(f"--- Texto ingresado ---\n{text.strip()}")

    if not all_text_parts:
        return {"error": "No se proporcionó ningún contenido para procesar."}

    combined_text = "\n\n".join(all_text_parts)

    # Truncate if too long (Claude context limit safety)
    if len(combined_text) > 80000:
        combined_text = combined_text[:80000] + "\n[... contenido truncado ...]"

    try:
        products = call_claude(combined_text, business_type)
    except json.JSONDecodeError as e:
        return {"error": f"Error al parsear la respuesta del modelo. Intentá de nuevo. Detalle: {str(e)}"}
    except anthropic.APIError as e:
        return {"error": f"Error de API: {str(e)}"}
    except Exception as e:
        return {"error": f"Error inesperado: {str(e)}"}

    # Generate Excel
    output_id = uuid.uuid4().hex
    output_filename = f"carga_masiva_productos_{output_id}.xlsx"
    output_path = OUTPUT_DIR / output_filename

    generate_excel(products, str(output_path))

    return {
        "success": True,
        "product_count": len(products),
        "products": products,
        "download_url": f"/download/{output_filename}",
    }


@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return {"error": "Archivo no encontrado"}
    return FileResponse(
        path=str(file_path),
        filename="carga_masiva_productos.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


HTML_PAGE = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chepi</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f7fa; color: #333; min-height: 100vh; }
        .container { max-width: 720px; margin: 0 auto; padding: 24px 16px; }
        .header { text-align: center; margin-bottom: 32px; }
        .header h1 { font-size: 24px; color: #1a1a2e; margin-bottom: 4px; }
        .header p { color: #666; font-size: 14px; }
        .card { background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 16px; }
        .card h2 { font-size: 16px; margin-bottom: 12px; color: #1a1a2e; }
        .file-upload { border: 2px dashed #d0d5dd; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; transition: all 0.2s; position: relative; }
        .file-upload:hover { border-color: #FF3553; background: #fff5f6; }
        .file-upload input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
        .file-upload .icon { font-size: 32px; margin-bottom: 8px; }
        .file-upload p { color: #666; font-size: 14px; }
        .file-list { margin-top: 12px; }
        .file-item { display: flex; align-items: center; justify-content: space-between; padding: 8px 12px; background: #f9fafb; border-radius: 6px; margin-bottom: 4px; font-size: 13px; }
        .file-item .remove { color: #e53e3e; cursor: pointer; font-weight: bold; padding: 2px 6px; }
        textarea { width: 100%; min-height: 120px; border: 1px solid #d0d5dd; border-radius: 8px; padding: 12px; font-family: inherit; font-size: 14px; resize: vertical; }
        textarea:focus { outline: none; border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
        select { width: 100%; padding: 10px 12px; border: 1px solid #d0d5dd; border-radius: 8px; font-size: 14px; background: #fff; }
        select:focus { outline: none; border-color: #2563eb; }
        .btn { width: 100%; padding: 14px; background: #2563eb; color: #fff; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
        .btn:hover { background: #1d4ed8; }
        .btn:disabled { background: #93b4f6; cursor: not-allowed; }
        .result { display: none; }
        .result.show { display: block; }
        .result-success { background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 16px; }
        .result-error { background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; padding: 16px; color: #dc2626; }
        .download-btn { display: inline-block; margin-top: 12px; padding: 10px 20px; background: #16a34a; color: #fff; border-radius: 8px; text-decoration: none; font-weight: 600; }
        .download-btn:hover { background: #15803d; }
        .spinner { display: none; text-align: center; padding: 24px; }
        .spinner.show { display: block; }
        .spinner .dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; background: #2563eb; margin: 0 4px; animation: bounce 1.4s infinite ease-in-out both; }
        .spinner .dot:nth-child(1) { animation-delay: -0.32s; }
        .spinner .dot:nth-child(2) { animation-delay: -0.16s; }
        @keyframes bounce { 0%, 80%, 100% { transform: scale(0); } 40% { transform: scale(1); } }
        .preview-table { width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 13px; }
        .preview-table th { background: #FF3553; color: #fff; padding: 8px; text-align: left; }
        .preview-table td { padding: 6px 8px; border-bottom: 1px solid #e5e7eb; }
        .preview-table tr:nth-child(even) { background: #f9fafb; }
        .preview-table .revisar { color: #dc2626; font-weight: 600; }
        .stats { display: flex; gap: 12px; margin-bottom: 12px; }
        .stat { flex: 1; text-align: center; padding: 12px; background: #fff5f6; border-radius: 8px; }
        .stat .num { font-size: 24px; font-weight: 700; color: #FF3553; }
        .stat .label { font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <img src="/static/Gemini_Generated_Image_qmcft7qmcft7qmcf.png" alt="Chepi" style="width: 200px; height: 200px; margin-bottom: 12px; object-fit: contain;">
            <h1>Chepi</h1>
            <p>Subí archivos o pegá texto para generar el Excel de carga masiva</p>
        </div>

        <div class="card">
            <h2>Archivos del cliente</h2>
            <div class="file-upload" id="dropZone">
                <input type="file" id="fileInput" multiple accept=".pdf,.xlsx,.xls,.csv,.txt">
                <div class="icon">📄</div>
                <p>Arrastrá archivos o hacé clic para subir<br><small>PDF, Excel, CSV, TXT</small></p>
            </div>
            <div class="file-list" id="fileList"></div>
        </div>

        <div class="card">
            <h2>Texto libre</h2>
            <textarea id="textInput" placeholder="Pegá acá el texto del cliente: mensajes de WhatsApp, listas de productos, mails, etc."></textarea>
        </div>

        <button class="btn" id="processBtn" onclick="processProducts()">Procesar</button>

        <div class="spinner" id="spinner">
            <div class="dot"></div><div class="dot"></div><div class="dot"></div>
            <p style="margin-top: 12px; color: #666;">Procesando productos con IA...</p>
        </div>

        <div class="result" id="result"></div>
    </div>

    <script>
        let selectedFiles = [];

        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const dropZone = document.getElementById('dropZone');

        fileInput.addEventListener('change', (e) => {
            for (const file of e.target.files) {
                selectedFiles.push(file);
            }
            renderFileList();
            fileInput.value = '';
        });

        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.style.borderColor = '#FF3553'; });
        dropZone.addEventListener('dragleave', () => { dropZone.style.borderColor = '#d0d5dd'; });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#d0d5dd';
            for (const file of e.dataTransfer.files) {
                selectedFiles.push(file);
            }
            renderFileList();
        });

        function renderFileList() {
            fileList.innerHTML = selectedFiles.map((f, i) =>
                `<div class="file-item"><span>${f.name} (${(f.size/1024).toFixed(1)} KB)</span><span class="remove" onclick="removeFile(${i})">✕</span></div>`
            ).join('');
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            renderFileList();
        }

        async function processProducts() {
            const text = document.getElementById('textInput').value;

            if (selectedFiles.length === 0 && !text.trim()) {
                alert('Subí al menos un archivo o pegá texto para procesar.');
                return;
            }

            const formData = new FormData();
            for (const file of selectedFiles) {
                formData.append('files', file);
            }
            formData.append('text', text);

            document.getElementById('processBtn').disabled = true;
            document.getElementById('spinner').classList.add('show');
            document.getElementById('result').classList.remove('show');

            try {
                const response = await fetch('/process', { method: 'POST', body: formData });
                const data = await response.json();

                if (data.error) {
                    showError(data.error);
                } else {
                    showSuccess(data);
                }
            } catch (err) {
                showError('Error de conexión. Verificá que el servidor esté corriendo.');
            } finally {
                document.getElementById('processBtn').disabled = false;
                document.getElementById('spinner').classList.remove('show');
            }
        }

        function showError(message) {
            const result = document.getElementById('result');
            result.innerHTML = `<div class="result-error"><strong>Error:</strong> ${message}</div>`;
            result.classList.add('show');
        }

        function showSuccess(data) {
            const result = document.getElementById('result');
            const revisar = data.products.filter(p => p.precio === 'REVISAR').length;

            let html = `<div class="result-success">
                <div class="stats">
                    <div class="stat"><div class="num">${data.product_count}</div><div class="label">Productos</div></div>
                    <div class="stat"><div class="num">${revisar}</div><div class="label">Para revisar</div></div>
                </div>
                <a href="${data.download_url}" class="download-btn">Descargar Excel</a>
            </div>
            <div class="card" style="margin-top:16px; overflow-x:auto;">
                <h2>Vista previa</h2>
                <table class="preview-table">
                    <thead><tr><th>Título</th><th>Descripción</th><th>Categoría</th><th>Precio</th><th>Fechas</th></tr></thead>
                    <tbody>`;

            for (const p of data.products) {
                const precioClass = p.precio === 'REVISAR' ? ' class="revisar"' : '';
                html += `<tr>
                    <td>${esc(p.titulo)}</td>
                    <td>${esc(p.descripcion)}</td>
                    <td>${esc(p.categoria)}</td>
                    <td${precioClass}>${esc(p.precio)}</td>
                    <td>${esc(p.fechas_disponibles)}</td>
                </tr>`;
            }

            html += '</tbody></table></div>';
            result.innerHTML = html;
            result.classList.add('show');
        }

        function esc(str) {
            const div = document.createElement('div');
            div.textContent = str || '';
            return div.innerHTML;
        }
    </script>
</body>
</html>"""


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
