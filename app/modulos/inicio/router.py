from fastapi import APIRouter, Request, Depends
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates
from app.modulos.auth.dependencies import login_required
from app.core.database import SessionLocal
from app.modulos.auth.models import User
import json

# Se inicializa el router.
router = APIRouter(dependencies=[Depends(login_required)])

from app.core.config import settings

# Textos repetidos
_ERROR = "error"

templates_dir = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir])


@router.get("/", name="inicio.index")
async def pagina_inicio(request: Request):
    """Renderiza la página de inicio del sistema."""
    # Obtener orden guardado del usuario
    app_order = None
    if hasattr(request.state, "user") and request.state.user:
        db = SessionLocal()
        try:
            user = db.query(User).filter(User.id == request.state.user.id).first()
            if user and user.app_order:
                app_order = user.app_order  # Ya es string JSON
        finally:
            db.close()

    return templates.TemplateResponse(
        "inicio/index.html", {"request": request, "titulo": "Panel Principal", "app_order": app_order or "null"}
    )


@router.post("/api/app-order")
async def save_app_order(request: Request):
    """Guarda el orden personalizado de apps del usuario."""
    body = await request.json()
    order = body.get("order", [])

    if not order or not isinstance(order, list):
        return JSONResponse(status_code=400, content={_ERROR: "Orden inválido"})

    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == request.state.user.id).first()
        if user:
            user.app_order = json.dumps(order)
            db.commit()
            return {"status": "ok"}
        return JSONResponse(status_code=404, content={_ERROR: "Usuario no encontrado"})
    finally:
        db.close()


@router.delete("/api/app-order")
async def reset_app_order(request: Request):
    """Resetea el orden de apps del usuario."""
    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == request.state.user.id).first()
        if user:
            user.app_order = None
            db.commit()
            return {"status": "ok"}
        return JSONResponse(status_code=404, content={_ERROR: "Usuario no encontrado"})
    finally:
        db.close()
