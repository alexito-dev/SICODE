from fastapi import APIRouter, Request, Form, Depends
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse
from app.modulos.auth.dependencies import login_required

router = APIRouter(prefix="/scadbot", tags=["scadbot"], dependencies=[Depends(login_required)])

from app.core.config import settings

templates_dir = settings.get_template_path("app/modulos/scadbot/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get("/")
async def scadbot_home(request: Request):
    return templates.TemplateResponse("scadbot/index.html", {"request": request})


@router.post("/verify")
async def verify_folders(month: str = Form(...), year: str = Form(...)):
    # Mock simulation of long running process
    # In reality this would check folders
    import random

    results = []

    # Simulating some checks
    pedidos_mock = range(242448, 242460)

    for p in pedidos_mock:
        status = random.choice(["ok", "missing_folder", "missing_txt"])  # NOSONAR (datos demo)
        results.append(
            {
                "pedido": str(p),
                "status": status,
                "message": f"Pedido {p}: "
                + ("Carpeta y TXT correctos" if status == "ok" else "Falta carpeta o archivo"),
            }
        )

    return JSONResponse(
        {
            "status": "success",
            "data": results,
            "summary": {
                "total": len(pedidos_mock),
                "ok": len([r for r in results if r["status"] == "ok"]),
                "errors": len([r for r in results if r["status"] != "ok"]),
            },
        }
    )
