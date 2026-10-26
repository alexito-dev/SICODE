import json
import os
import secrets
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from fastapi import APIRouter, Request, Response, Form, Depends, status, HTTPException
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from app.core.config import settings
from app.core.database import get_db
from app.modulos.auth.models import User, PasswordResetToken
from datetime import timedelta, timezone

# Textos repetidos
_ADMINISTRADOR = "Administrador"
_ADMIN_USERS = "/admin/users"
_AUTH_LOGIN_HTML = "auth/login.html"
_AUTH_SETTINGS_HTML = "auth/settings.html"
_AUTH_USERS_HTML = "auth/users.html"
_COLABORADOR = "Colaborador"
_COLABORADOR_PLUS = "Colaborador Plus"
_ERROR = "error"
_LOGIN = "/login"
_REQUEST = "request"
_ROLES = "roles"
_SUCCESS = "success"
_SUPERVISOR = "Supervisor"
_USERS = "users"

router = APIRouter()
templates_dir = settings.get_template_path("app/modulos/auth/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get(_LOGIN, name="auth.login")
async def login_page(request: Request):
    """Muestra la página de inicio de sesión."""
    success = request.query_params.get(_SUCCESS)
    return templates.TemplateResponse(_AUTH_LOGIN_HTML, {_REQUEST: request, _SUCCESS: success})


@router.post(_LOGIN, name="auth.login_post")
async def login_submit(
    request: Request,
    response: Response,
    username: str = Form(...),
    password: str = Form(...),
    remember: bool = Form(False),
    db: Session = Depends(get_db),
):
    """Procesa el inicio de sesión con soporte para Recordar sesión y BD."""

    # 1. Validación de Credenciales via BD
    user = db.query(User).filter(User.username == username).first()

    # Simple check (Plain text). In production compare hashes!
    if not user or user.password != password:
        return templates.TemplateResponse(
            _AUTH_LOGIN_HTML, {_REQUEST: request, _ERROR: "Usuario o contraseña incorrectos"}
        )

    if not user.is_active:
        return templates.TemplateResponse(
            _AUTH_LOGIN_HTML, {_REQUEST: request, _ERROR: "Usuario desactivado. Contacte al administrador."}
        )

    # 2. Configuración de Sesión (Cookie)
    max_age = 60 * 60 * 24 * 30 if remember else None
    expires = datetime.now(timezone.utc) + timedelta(days=30) if remember else None

    # Redireccionar al Dashboard o módulo principal
    redirect_url = "/"  # Ir al inicio
    auth_response = RedirectResponse(url=redirect_url, status_code=status.HTTP_303_SEE_OTHER)

    # 3. Establecer Cookie Segura
    auth_response.set_cookie(
        key="access_token",
        value=f"user:{username}",  # Token simple "user:admin"
        httponly=True,
        path="/",
        max_age=max_age,
        expires=expires,
        samesite="lax",
    )

    return auth_response


@router.get("/logout", name="auth.logout")
async def logout():
    """Cierra la sesión eliminando la cookie y previniendo caché."""
    response = RedirectResponse(url=_LOGIN, status_code=status.HTTP_303_SEE_OTHER)

    # IMPORTANTE: delete_cookie debe usar los mismos parámetros que set_cookie
    response.delete_cookie(key="access_token", path="/", samesite="lax")

    # Headers de seguridad para prevenir caché de la respuesta
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    return response


# --- Recuperación de Contraseña ---


@router.get("/forgot-password", name="auth.forgot_password")
async def forgot_password_page(request: Request):
    return templates.TemplateResponse("auth/forgot_password.html", {_REQUEST: request})


# --- Gestión de Usuarios (Admin/Supervisor Only) ---


def send_password_reset_email(email_to: str, user_name: str, reset_link: str):
    """Envía un correo real usando SMTP."""
    if not settings.SMTP_USER or not settings.SMTP_PASSWORD:
        print("AVISO: SMTP no configurado. El correo se verá solo en el terminal.")
        return False

    try:
        msg = MIMEMultipart()
        msg["From"] = settings.SMTP_FROM
        msg["To"] = email_to
        msg["Subject"] = "Recuperación de Contraseña - SICODE"

        body = f"""
        Hola {user_name},

        Has solicitado restablecer tu contraseña en el portal SICODE.
        Para continuar, haz clic en el siguiente enlace:

        {reset_link}

        Este enlace vencerá en 1 hora. Si no realizaste esta solicitud, puedes ignorar este correo.

        Atentamente,
        El equipo de SICODE
        """
        msg.attach(MIMEText(body, "plain"))

        server = smtplib.SMTP(settings.SMTP_HOST, settings.SMTP_PORT)
        server.starttls()
        server.login(settings.SMTP_USER, settings.SMTP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando correo: {str(e)}")
        return False


@router.post("/forgot-password", name="auth.forgot_password_post")
async def forgot_password_submit(request: Request, email: str = Form(...), db: Session = Depends(get_db)):
    """Genera un token de recuperación y envía el correo real."""
    user = db.query(User).filter(User.correo == email.strip()).first()

    if user:
        # 1. Generar Token único
        token = secrets.token_urlsafe(32)
        expires_at = datetime.datetime.now() + datetime.timedelta(hours=1)

        # 2. Guardar en BD
        reset_token = PasswordResetToken(user_id=user.id, token=token, expires_at=expires_at)
        db.add(reset_token)
        db.commit()

        # 3. Envío de Email Real
        reset_link = f"{request.base_url}reset-password?token={token}"
        sent = send_password_reset_email(user.correo, user.nombre, reset_link)

        if not sent:
            # Fallback al terminal si falla o no está configurado
            print(f"\n{'=' * 50}")
            print(f"FALLO DE ENVÍO O NO CONFIGURADO. ENLACE: {reset_link}")
            print(f"{'=' * 50}\n")

    return templates.TemplateResponse(
        "auth/forgot_password.html", {_REQUEST: request, "test_link": reset_link if user else None}
    )


@router.get("/reset-password", name="auth.reset_password")
async def reset_password_page(request: Request, token: str, db: Session = Depends(get_db)):
    """Valida el token y muestra el formulario de nueva clave."""
    token_rec = (
        db.query(PasswordResetToken).filter(PasswordResetToken.token == token, not PasswordResetToken.used).first()
    )

    if not token_rec or token_rec.is_expired:
        return RedirectResponse(url="/forgot-password?error=Token+inválido+o+expirado")

    return templates.TemplateResponse(
        "auth/reset_password.html", {_REQUEST: request, "token": token, "username": token_rec.user.username}
    )


@router.post("/reset-password", name="auth.reset_password_post")
async def reset_password_submit(
    request: Request,
    token: str = Form(...),
    password: str = Form(...),
    confirm_password: str = Form(...),
    db: Session = Depends(get_db),
):
    """Procesa el cambio de clave post-recuperación."""
    if password != confirm_password:
        return templates.TemplateResponse(
            "auth/reset_password.html", {_REQUEST: request, "token": token, _ERROR: "Las contraseñas no coinciden"}
        )

    token_rec = (
        db.query(PasswordResetToken).filter(PasswordResetToken.token == token, not PasswordResetToken.used).first()
    )

    if not token_rec or token_rec.is_expired:
        return RedirectResponse(
            url="/forgot-password?error=Reseteo+fallido+por+token", status_code=status.HTTP_303_SEE_OTHER
        )

    # Realizar el cambio
    user = token_rec.user
    user.password = password
    token_rec.used = True
    db.commit()

    return RedirectResponse(
        url="/login?success=Contraseña+actualizada.+Ya+puedes+entrar.", status_code=status.HTTP_303_SEE_OTHER
    )


# --- Gestión de Perfil ---


@router.get("/profile", name="auth.profile")
async def profile_page(request: Request):
    """Muestra la página de perfil del usuario."""
    if not request.state.user:
        return RedirectResponse(url=_LOGIN)
    return templates.TemplateResponse("auth/profile.html", {_REQUEST: request})


@router.post("/profile", name="auth.profile_post")
async def profile_update(
    request: Request,
    nombre: str = Form(...),
    apellido: str = Form(...),
    codigo: str = Form(...),
    correo: str = Form(...),
    db: Session = Depends(get_db),
):
    """Actualiza la información del perfil del usuario."""
    if not request.state.user:
        return RedirectResponse(url=_LOGIN)

    user = db.query(User).filter(User.id == request.state.user.id).first()
    if user:
        user.nombre = nombre
        user.apellido = apellido
        user.codigo = codigo
        user.correo = correo
        db.commit()

    return templates.TemplateResponse(
        "auth/profile.html", {_REQUEST: request, _SUCCESS: "Perfil actualizado correctamente"}
    )


# --- Ajustes y Seguridad ---


@router.get("/settings", name="auth.settings")
async def settings_page(request: Request):
    """Muestra la página de ajustes."""
    if not request.state.user:
        return RedirectResponse(url=_LOGIN)
    return templates.TemplateResponse(_AUTH_SETTINGS_HTML, {_REQUEST: request})


@router.post("/settings/change-password", name="auth.change_password")
async def change_password(
    request: Request,
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...),
    db: Session = Depends(get_db),
):
    """Procesa el cambio de contraseña."""
    if not request.state.user:
        return RedirectResponse(url=_LOGIN)

    user = db.query(User).filter(User.id == request.state.user.id).first()

    if user.password != current_password:
        return templates.TemplateResponse(
            _AUTH_SETTINGS_HTML, {_REQUEST: request, _ERROR: "La contraseña actual es incorrecta"}
        )

    if new_password != confirm_password:
        return templates.TemplateResponse(
            _AUTH_SETTINGS_HTML, {_REQUEST: request, _ERROR: "Las nuevas contraseñas no coinciden"}
        )

    user.password = new_password
    db.commit()

    return templates.TemplateResponse(
        _AUTH_SETTINGS_HTML, {_REQUEST: request, _SUCCESS: "Contraseña actualizada correctamente"}
    )


# --- Gestión de Usuarios (Admin/Supervisor Only) ---


def load_roles():
    roles = []
    for path in [os.path.join(r"D:\SICODE\DB", "roles.json"), os.path.join(os.path.abspath("."), "roles.json")]:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list) and data:
                        roles = data
                        break
            except Exception:
                pass
    if not roles:
        roles = [_ADMINISTRADOR, _SUPERVISOR, _COLABORADOR_PLUS, _COLABORADOR]

    if _COLABORADOR_PLUS not in roles:
        if _COLABORADOR in roles:
            idx = roles.index(_COLABORADOR)
            roles.insert(idx, _COLABORADOR_PLUS)
        else:
            roles.append(_COLABORADOR_PLUS)
    return roles


@router.get(_ADMIN_USERS, name="auth.admin_users")
async def admin_users_page(request: Request, db: Session = Depends(get_db)):
    """Muestra la página de gestión de usuarios."""
    if not request.state.user or request.state.user.role not in [_ADMINISTRADOR, _SUPERVISOR]:
        return RedirectResponse(url="/")

    users = db.query(User).all()
    roles = load_roles()
    return templates.TemplateResponse(_AUTH_USERS_HTML, {_REQUEST: request, _USERS: users, _ROLES: roles})


@router.post(_ADMIN_USERS, name="auth.admin_users_post")
async def admin_save_user(
    request: Request,
    user_id: str = Form(None),
    codigo: str = Form(None),
    nombre: str = Form(...),
    apellido: str = Form(...),
    correo: str = Form(...),
    username: str = Form(None),
    role: str = Form(...),
    password: str = Form(None),
    db: Session = Depends(get_db),
):
    """Crea o actualiza un usuario."""
    if not request.state.user or request.state.user.role not in [_ADMINISTRADOR, _SUPERVISOR]:
        raise HTTPException(status_code=403, detail="No autorizado")

    # Lógica de Username: Priorizar input manual, sino generar automático
    if not username:
        username = correo.split("@")[0]

    try:
        if user_id:
            # EDITAR
            user = db.query(User).filter(User.id == int(user_id)).first()
            if user:
                user.codigo = codigo
                user.nombre = nombre
                user.apellido = apellido
                user.correo = correo
                user.role = role
                user.username = username
                if password:
                    user.password = password
        else:
            # CREAR NUEVO
            existing = db.query(User).filter(User.username == username).first()
            if existing:
                # En la web, mejor devolver un error que ignorar
                users = db.query(User).all()
                roles = load_roles()
                return templates.TemplateResponse(
                    _AUTH_USERS_HTML,
                    {
                        _REQUEST: request,
                        _USERS: users,
                        _ROLES: roles,
                        _ERROR: f"El usuario login '{username}' ya está en uso.",
                    },
                )

            new_user = User(
                codigo=codigo,
                nombre=nombre,
                apellido=apellido,
                correo=correo,
                role=role,
                username=username,
                password=password or "123456",
                is_active=True,
            )
            db.add(new_user)

        db.commit()
        return RedirectResponse(url=_ADMIN_USERS, status_code=status.HTTP_303_SEE_OTHER)

    except Exception as e:
        db.rollback()
        users = db.query(User).all()
        roles = load_roles()
        return templates.TemplateResponse(
            _AUTH_USERS_HTML,
            {_REQUEST: request, _USERS: users, _ROLES: roles, _ERROR: f"Error al guardar: {str(e)}"},
        )


@router.post("/admin/users/{del_id}/delete", name="auth.admin_users_delete")
async def admin_delete_user(request: Request, del_id: int, db: Session = Depends(get_db)):
    """Elimina un usuario."""
    if not request.state.user or request.state.user.role not in [_ADMINISTRADOR, _SUPERVISOR]:
        raise HTTPException(status_code=403, detail="No autorizado")

    try:
        db.query(User).filter(User.id == del_id).delete()
        db.commit()
    except Exception:
        db.rollback()

    return RedirectResponse(url=_ADMIN_USERS, status_code=status.HTTP_303_SEE_OTHER)
