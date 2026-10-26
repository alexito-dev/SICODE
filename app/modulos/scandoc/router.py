from fastapi import APIRouter, Request, WebSocket, WebSocketDisconnect, BackgroundTasks
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Dict
import uuid

from app.core.config import settings
from app.modulos.scandoc.service import process_ocr_batch, generate_qr_base64

router = APIRouter(prefix="/scandoc", tags=["ScanDoc"])

# Configuración de plantillas
templates_dir = settings.get_template_path("app/modulos/scandoc/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates_global = settings.get_template_path("app/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base, templates_global])


# --- Modelos ---
class OCRRequest(BaseModel):
    scan_id: str
    images_base64: List[str]


# --- WebSocket Manager ---
class ConnectionManager:
    def __init__(self):
        self.active_connections: Dict[str, List[WebSocket]] = {}

    async def connect(self, websocket: WebSocket, scan_id: str):
        await websocket.accept()
        if scan_id not in self.active_connections:
            self.active_connections[scan_id] = []
        self.active_connections[scan_id].append(websocket)

    def disconnect(self, websocket: WebSocket, scan_id: str):
        if scan_id in self.active_connections:
            self.active_connections[scan_id].remove(websocket)
            if not self.active_connections[scan_id]:
                del self.active_connections[scan_id]

    async def send_personal_message(self, message: str, websocket: WebSocket):
        await websocket.send_text(message)

    async def broadcast(self, scan_id: str, message: dict):
        if scan_id in self.active_connections:
            for connection in self.active_connections[scan_id]:
                await connection.send_json(message)


manager = ConnectionManager()

# --- Rutas ---


@router.get("/", response_class=HTMLResponse)
async def scandoc_home(request: Request):
    scan_id = str(uuid.uuid4())[:8]
    # For simplicity, we use the request URL to build the mobile link
    base_url = str(request.base_url).rstrip("/")
    mobile_url = f"{base_url}/scandoc/m/{scan_id}"
    qr_code = generate_qr_base64(mobile_url)

    return templates.TemplateResponse(
        "scandoc/index.html",
        {"request": request, "scan_id": scan_id, "mobile_url": mobile_url, "qr_code_base64": qr_code},
    )


@router.get("/m/{scan_id}", response_class=HTMLResponse)
async def scandoc_mobile(request: Request, scan_id: str):
    return templates.TemplateResponse("scandoc/mobile.html", {"request": request, "scan_id": scan_id})


@router.post("/api/ocr")
async def api_ocr(payload: OCRRequest, background_tasks: BackgroundTasks):
    # Procesar en background o esperar? Flask esperaba.
    # Para mantener la UX de "ver los cambios", procesamos y luego emitimos.
    result = await process_ocr_batch(payload.images_base64)

    # Emitir via WebSocket
    await manager.broadcast(
        payload.scan_id,
        {
            "type": "new_scan",
            "data": result["fields"],
            "raw_text": result["raw_text"],
            "processed_images": result["processed_images"],
        },
    )

    return {"ok": True, "item": {"fields": result["fields"]}}


# --- WebSocket Endpoint ---


@router.websocket("/ws/{scan_id}")
async def websocket_endpoint(websocket: WebSocket, scan_id: str):
    await manager.connect(websocket, scan_id)
    try:
        while True:
            # Mantener la conexión abierta
            await websocket.receive_text()
            # Podríamos manejar mensajes del dashboard aquí si fuera necesario
    except WebSocketDisconnect:
        manager.disconnect(websocket, scan_id)
