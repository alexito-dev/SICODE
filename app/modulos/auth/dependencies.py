from fastapi import Request, HTTPException, status


async def login_required(request: Request):
    """
    Dependencia para rutas que requieren autenticación.
    Si no hay usuario en request.state (puesto por el middleware), lanza un error 401.
    Este error será capturado en main.py para redirigir al login.
    """
    if not hasattr(request.state, "user") or request.state.user is None:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="No autenticado")
    return request.state.user
