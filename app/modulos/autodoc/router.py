from fastapi import APIRouter, Request, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import Dict, Any, List
from app.modulos.auth.dependencies import login_required

# --- CONFIGURACIÓN DEL ROUTER (Es como el controlador principal del módulo) ---
# Usamos APIRouter para organizar nuestras rutas.
# 'prefix="/autodoc"' significa que todas las URLs de este archivo empezarán con /autodoc
# 'tags' sirve para agrupar endpoints en la documentación automática (Swagger UI)
router = APIRouter(prefix="/autodoc", tags=["AutoDoc"], dependencies=[Depends(login_required)])


@router.get("/", response_class=HTMLResponse)
async def autodoc_home(request: Request):
    return templates.TemplateResponse("autodoc/index.html", {"request": request})


# --- CONFIGURACIÓN DE PLANTILLAS (JINJA2) ---
# Aquí le decimos a FastAPI dónde buscar los archivos .html.
# MUY IMPORTANTE: Incluimos carpetas compartidas ('inicio/templates' y 'app/templates')
# porque nuestro index.html usa {% extends "base.html" %}, y ese archivo base vive ahí.
from app.core.config import settings

templates_dir = settings.get_template_path("app/modulos/autodoc/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates_global = settings.get_template_path("app/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base, templates_global])


# --- MODELOS DE DATOS (Pydantic) ---
# Definimos la estructura exacta que esperamos recibir del Frontend.
# Esto es una "buena práctica" para validar datos automáticamente.
# Si el JS nos manda algo mal, FastAPI rechazará la petición antes de que rompa nuestro código.
class GenerateRequest(BaseModel):
    filename: str  # Nombre del archivo plantilla (ej: Esquela.docx)
    data: Dict[str, Any]  # Diccionario con los datos del formulario (nombre, fecha, etc.)


# --- RUTAS / ENDPOINTS ---

from io import BytesIO
import zipfile
from fastapi.responses import Response

# Textos repetidos
_AUTODOC = "AUTODOC"
_IMAGEN_FIRMA = "imagen_firma"
_MESSAGE = "message"
_PLANTILLAS = "PLANTILLAS"
_PLANTILLA_BASE_DOCX = "Plantilla_Base.docx"
_SICODE = "SICODE"


# --- HELPERS PARA IMÁGENES DINÁMICAS ---
def prepare_context_for_doc(doc, row: Dict[str, Any]) -> Dict[str, Any]:
    from docxtpl import InlineImage
    from docx.shared import Mm
    import os
    import datetime

    ctx = row.copy()

    # Inyectar fecha de hoy formateada
    meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    ]
    h = datetime.date.today()
    ctx["fechaActual"] = f"{h.day:02d} de {meses[h.month - 1]} de {h.year}"
    firma_nombre = ctx.get("firma", "")
    if firma_nombre:
        # Busca el nombre del archivo PNG correspondiente al campo firma
        firma_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, "FIRMAS", f"{firma_nombre}.png".upper())
        # En caso de no ser mayúscula el archivo, verificamos también normal
        if not os.path.exists(firma_path):
            firma_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, "FIRMAS", f"{firma_nombre}.png")

        if os.path.exists(firma_path):
            # Carga la imagen al documento Word con el tamaño especificado
            ctx[_IMAGEN_FIRMA] = InlineImage(doc, firma_path, width=Mm(50))
        else:
            ctx[_IMAGEN_FIRMA] = ""  # Archivo no encontrado
    else:
        ctx[_IMAGEN_FIRMA] = ""

    return ctx


# 3. API: GENERAR DOCUMENTOS ZIP
@router.post("/api/generate")
async def generate_documents(rows: List[Dict[str, Any]]):
    from docxtpl import DocxTemplate
    import os

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for i, row in enumerate(rows):
            tipo_plantilla = row.get("tipoPlantilla", "").strip()
            subtipo_plantilla = row.get("subtipoPlantilla", "").strip()

            # Priorizar el subtipo para el nombre del archivo (ej: PAGO CON ERROR ATENDIDO.docx)
            nombre_base = subtipo_plantilla or tipo_plantilla
            template_name = f"{nombre_base}.docx" if nombre_base else _PLANTILLA_BASE_DOCX
            template_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, _PLANTILLAS, template_name)

            if not os.path.exists(template_path):
                # Fallback a Plantilla_Base si no existe
                template_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, _PLANTILLAS, _PLANTILLA_BASE_DOCX)
            if not os.path.exists(template_path):
                continue

            # Crear documento desde plantilla
            doc = DocxTemplate(template_path)
            # Preparamos el contexto (procesa firma dinámica)
            context = prepare_context_for_doc(doc, row)
            doc.render(context)

            # Guardar en buffer
            doc_io = BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            # Nombre del archivo
            filename = f"{row.get('ruc', 'doc')}_{row.get('numExpediente', i)}.docx"
            zip_file.writestr(filename, doc_io.read())

    zip_buffer.seek(0)

    return Response(
        content=zip_buffer.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=documentos.zip"},
    )


@router.post("/api/preview")
async def preview_document(row: Dict[str, Any]):
    from docxtpl import DocxTemplate
    import os
    import tempfile
    import uuid
    import pythoncom

    tipo_plantilla = row.get("tipoPlantilla", "").strip()
    subtipo_plantilla = row.get("subtipoPlantilla", "").strip()

    # Priorizar el subtipo para el nombre del archivo
    nombre_base = subtipo_plantilla or tipo_plantilla
    template_name = f"{nombre_base}.docx" if nombre_base else _PLANTILLA_BASE_DOCX
    template_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, _PLANTILLAS, template_name)

    if not os.path.exists(template_path):
        template_path = os.path.join("D:\\", _SICODE, "APPS", _AUTODOC, _PLANTILLAS, _PLANTILLA_BASE_DOCX)

    if not os.path.exists(template_path):
        return JSONResponse(
            status_code=404,
            content={_MESSAGE: f"Plantilla no encontrada: {template_name} en D:\\SICODE\\APPS\\AUTODOC\\PLANTILLAS"},
        )

    doc = DocxTemplate(template_path)
    # Llenar con datos y firmas
    context = prepare_context_for_doc(doc, row)
    doc.render(context)

    # Generate unique filenames directly in temp dir
    temp_dir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())
    temp_docx_path = os.path.join(temp_dir, f"tmp_{unique_id}.docx")
    temp_pdf_path = os.path.join(temp_dir, f"tmp_{unique_id}.pdf")

    # Save the rendered template
    doc.save(temp_docx_path)

    # Convertir a pdf usando win32com en lugar de docx2pdf para ocultar UI
    pythoncom.CoInitialize()
    try:
        import win32com.client

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False

        # Necesita rutas absolutas de backslash para COM en Windows
        abs_docx = os.path.abspath(temp_docx_path).replace("/", "\\")
        abs_pdf = os.path.abspath(temp_pdf_path).replace("/", "\\")

        word_doc = word.Documents.Open(abs_docx)
        word_doc.SaveAs(abs_pdf, FileFormat=17)  # 17 = wdFormatPDF
        word_doc.Close()
        word.Quit()
    except Exception as e:
        pythoncom.CoUninitialize()
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        import traceback

        error_details = traceback.format_exc()
        return JSONResponse(
            status_code=500,
            content={_MESSAGE: f"Error al generar vista previa en Word COM: {e!s} \nDetails: {error_details}"},
        )
    finally:
        pythoncom.CoUninitialize()

    # Leer el PDF devuelto
    if os.path.exists(temp_pdf_path):
        with open(temp_pdf_path, "rb") as f:
            pdf_bytes = f.read()
    else:
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)
        return JSONResponse(status_code=500, content={_MESSAGE: "No se pudo generar el archivo PDF"})

    # Limpiar
    if os.path.exists(temp_docx_path):
        os.remove(temp_docx_path)
    if os.path.exists(temp_pdf_path):
        os.remove(temp_pdf_path)

    return Response(content=pdf_bytes, media_type="application/pdf")
