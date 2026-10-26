"""

permisos.py

-----------

Este archivo define las reglas de ACCESO y SEGURIDAD específicas del módulo.



Responsabilidades:

- Definir funciones de dependencia (Depends) para verificar roles.

- Comprobar si el usuario tiene permiso para ejecutar una acción.



Ejemplo:

from fastapi import Depends, HTTPException



def solo_admin(usuario: dict = Depends(obtener_usuario_actual)):

    if usuario['rol'] != 'admin':

        raise HTTPException(status_code=403, detail="No autorizado")

"""


def verificar_acceso_inicio():

    # Lógica de verificación

    return True
