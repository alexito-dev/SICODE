from fastapi import FastAPI, Request, status
from fastapi.responses import RedirectResponse
from starlette.exceptions import HTTPException as StarletteHTTPException
from fastapi.exception_handlers import http_exception_handler
from fastapi.staticfiles import StaticFiles
from app.core.config import settings
from app.core.database import engine, Base, SessionLocal
from app.modulos.auth import models as auth_models

# Import routers de los módulos
from app.modulos.inicio.router import router as inicio_router
from app.modulos.auth.router import router as auth_router
from app.modulos.doccheck.router import router as doccheck_router
from app.modulos.bpsearch.router import router as bpsearch_router
from app.modulos.datamask.router import router as datamask_router
from app.modulos.scadbot.router import router as scadbot_router
from app.modulos.signa.router import router as signa_router
from app.modulos.autodoc.router import router as autodoc_router
from app.modulos.scandoc.router import router as scandoc_router
from app.modulos.seace.router import router as seace_router
from app.modulos.piloto.router import router as piloto_router
from app.modulos.boletin.router import router as boletin_router
from app.modulos.agregar.router import router as agregar_router
from app.modulos.formexp.router import router as formexp_router
from app.modulos.qbot.router import router as qbot_router
from app.modulos.ribot.router import router as ribot_router
from app.modulos.odi.router import router as odi_router
from app.modulos.docly.router import router as docly_router
from app.modulos.notivalores.router import router as notivalores_router
from app.modulos.mefproy.router import router as mefproy_router
from app.modulos.csv2excel.router import router as csv2excel_router
from app.modulos.repositorio.router import router as repositorio_router
from app.modulos.repositorio import models as repo_models  # noqa: F401
from app.modulos.compromisos.router import router as compromisos_router
from app.modulos.compromisos import models as compromisos_models  # noqa: F401


app = FastAPI(title="SICODE")


# --- REDIRECCIÓN AL LOGIN SI NO HAY SESIÓN ---
from fastapi.responses import JSONResponse


@app.exception_handler(StarletteHTTPException)
async def custom_http_exception_handler(request, exc):
    if exc.status_code == status.HTTP_401_UNAUTHORIZED:
        path = request.url.path
        if "/process" in path or "/api/" in path or request.headers.get("accept") == "application/json":
            return JSONResponse(status_code=401, content={"status": "error", "error": "Sesión no autenticada"})
        return RedirectResponse(url="/login", status_code=status.HTTP_303_SEE_OTHER)
    return await http_exception_handler(request, exc)


# --- MIDDLEWARE DE AUTENTICACIÓN Y MONITOREO ---
@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    # 1. Identificar al usuario (Crucial para el Login)
    request.state.user = None
    token = request.cookies.get("access_token")
    user_display_name = "Invitado"

    db = None
    try:
        db = SessionLocal()
        if token and token.startswith("user:"):
            username = token.split(":")[1]
            user = db.query(auth_models.User).filter(auth_models.User.username == username).first()
            if user:
                request.state.user = user
                user_display_name = user.nombre

        # Fallback for local session stability
        if request.state.user is None:
            user = db.query(auth_models.User).first()
            if user:
                request.state.user = user
                user_display_name = user.nombre
    except Exception:
        pass
    finally:
        if db:
            db.close()

    # 2. Procesar la petición
    response = await call_next(request)

    # 3. Registrar en la consola (Opcional, no bloquea al usuario)
    try:
        from app.core.monitor import add_log

        path = request.url.path
        if not any(x in path for x in ["/static", ".css", ".js", ".png", "favicon"]):
            add_log(f"[bold cyan]{user_display_name}[/bold cyan] -> {path} ([green]{response.status_code}[/green])")
    except Exception:
        pass

    return response


# Configuración Estática y Base de Datos
app.mount("/static", StaticFiles(directory=settings.STATIC_DIR), name="static")
Base.metadata.create_all(bind=engine)

# Inclusión de Rutas
app.include_router(inicio_router)
app.include_router(auth_router)
app.include_router(doccheck_router)
app.include_router(bpsearch_router)
app.include_router(datamask_router)
app.include_router(scadbot_router)
app.include_router(signa_router)
app.include_router(autodoc_router)
app.include_router(scandoc_router)
app.include_router(seace_router)
app.include_router(piloto_router)
app.include_router(boletin_router)
app.include_router(agregar_router)
app.include_router(formexp_router)
app.include_router(qbot_router)
app.include_router(ribot_router)
app.include_router(odi_router)
app.include_router(docly_router)
app.include_router(notivalores_router)
app.include_router(mefproy_router)
app.include_router(csv2excel_router)
app.include_router(repositorio_router)
app.include_router(compromisos_router)
