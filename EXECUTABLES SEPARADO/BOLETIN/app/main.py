from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse

from app.core_config import STATIC_DIR
from app.boletin.router import router as boletin_router

app = FastAPI(title="Boletín Oficial - SICODE")

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.include_router(boletin_router)


@app.get("/")
async def root():
    return RedirectResponse(url="/boletin/")
