"""
RepoSUNAT - repositorio minimalista sin login.
Una sola web en la red local: http://<IP>:8095
Los archivos se guardan en D:\\RepoSUNAT
"""
import os
import re
import socket
import sys
import webbrowser
import uuid
import urllib.parse
from pathlib import Path

import uvicorn
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

import identidad

PORT = 8095
BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
STORE = Path(os.environ.get("REPOSUNAT_DIR", r"D:\RepoSUNAT"))
STORE.mkdir(parents=True, exist_ok=True)

META = ".uploader"

app = FastAPI(title="RepoSUNAT", docs_url=None, redoc_url=None)


def clean_name(name: str) -> str:
    name = os.path.basename((name or "").replace("\\", "/"))
    name = re.sub(r'[\\/*?:"<>|\x00-\x1f]', "_", name).strip(" .")
    return name or "archivo"


def find(fid: str):
    if not re.fullmatch(r"[0-9a-f]{10}", fid):
        raise HTTPException(404, "No existe")
    folder = STORE / fid
    files = [p for p in folder.iterdir() if p.is_file() and p.name != META] if folder.is_dir() else []
    if not files:
        raise HTTPException(404, "No existe")
    return folder, files[0]


@app.get("/api/files")
def list_files():
    out = []
    for folder in STORE.iterdir():
        if not folder.is_dir() or not re.fullmatch(r"[0-9a-f]{10}", folder.name):
            continue
        for p in folder.iterdir():
            if p.is_file() and p.name != META:
                st = p.stat()
                meta = folder / META
                by = meta.read_text(encoding="utf-8").strip() if meta.exists() else ""
                out.append({"id": folder.name, "name": p.name, "size": st.st_size,
                            "date": st.st_mtime, "by": by})
    out.sort(key=lambda f: f["date"], reverse=True)
    return out


@app.post("/api/upload")
async def upload(request: Request, files: list[UploadFile] = File(...)):
    saved = []
    user = identidad.get_current_user(request)["nombre"]
    for f in files:
        name = clean_name(f.filename)
        fid = uuid.uuid4().hex[:10]
        folder = STORE / fid
        folder.mkdir()
        dest = folder / name
        with open(dest, "wb") as out:
            while chunk := await f.read(1024 * 1024):
                out.write(chunk)
        (folder / META).write_text(user, encoding="utf-8")
        saved.append({"id": fid, "name": name})
    return saved


@app.get("/api/me")
def me(request: Request):
    if identidad.NTLM_ENABLED and not request.cookies.get(identidad.COOKIE):
        challenge = identidad.ntlm_challenge(request)
        if challenge:
            return Response(status_code=401, headers={"WWW-Authenticate": challenge})
    user = identidad.get_current_user(request)
    resp = JSONResponse(user)
    if user["origen"] != "cookie":
        resp.set_cookie(identidad.COOKIE, quote(f"{user['login']}|{user['nombre']}"),
                        max_age=8 * 3600, httponly=True, samesite="lax")
    return resp


@app.get("/d/{fid}")
def download(fid: str):
    _, p = find(fid)
    q = urllib.parse.quote(p.name)
    return FileResponse(p, media_type="application/octet-stream",
                        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{q}"})


@app.delete("/api/files/{fid}")
def delete(fid: str):
    folder, p = find(fid)
    p.unlink()
    (folder / META).unlink(missing_ok=True)
    try:
        folder.rmdir()
    except OSError:
        pass
    return {"ok": True}


app.mount("/", StaticFiles(directory=BASE / "static", html=True), name="static")


def lan_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    print(f"\n  RepoSUNAT en: http://{lan_ip()}:{PORT}\n  Archivos en:  {STORE}\n")
    webbrowser.open(f"http://localhost:{PORT}")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="warning")
