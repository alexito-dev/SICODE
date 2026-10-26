from sqlalchemy import Column, Integer, String, Boolean, DateTime, ForeignKey, Text
from sqlalchemy.orm import relationship
from app.core.database import Base
import datetime


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)  # Usuario para Login (ej: prac-azorrilla)
    password = Column(String)

    # Datos Personales
    codigo = Column(String)  # Ej: TE78
    nombre = Column(String)
    apellido = Column(String)
    correo = Column(String)  # Ej: prac-azorrilla@sunat.gob.pe

    role = Column(String, default="Usuario")
    is_active = Column(Boolean, default=True)

    # Preferencias del usuario
    app_order = Column(Text, nullable=True)  # JSON: orden personalizado de apps en el inicio

    @property
    def full_name(self):
        return f"{self.nombre} {self.apellido}"


class PasswordResetToken(Base):
    __tablename__ = "password_reset_tokens"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    token = Column(String, unique=True, index=True, nullable=False)
    expires_at = Column(DateTime, nullable=False)
    used = Column(Boolean, default=False)

    user = relationship("User")

    @property
    def is_expired(self):
        return datetime.datetime.now() > self.expires_at
