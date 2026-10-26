from fastapi import APIRouter, Request, Depends
from fastapi.templating import Jinja2Templates
from app.modulos.auth.dependencies import login_required
from app.core.config import settings

router = APIRouter(prefix="/qbot", tags=["qbot"], dependencies=[Depends(login_required)])

templates_dir = settings.get_template_path("app/modulos/qbot/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get("/")
async def qbot_home(request: Request):
    return templates.TemplateResponse("qbot/index.html", {"request": request, "modulo": "QBot"})
