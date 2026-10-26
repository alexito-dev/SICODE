"""

esquemas.py

-----------

Este archivo define las validaciones de datos usando PYDANTIC.



Responsabilidades:

- Definir qué datos se envían y reciben en las peticiones (Request/Response Models).

- Validar tipos de datos, longitudes y formatos.

- Serializar objetos ORM a JSON.



Ejemplo:

from pydantic import BaseModel



class MensajeBase(BaseModel):

    titulo: str

    cuerpo: str

"""

from pydantic import BaseModel


class MensajeBienvenida(BaseModel):
    mensaje: str

    version: str
