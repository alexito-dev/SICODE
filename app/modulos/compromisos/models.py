import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime, Date, Float, ForeignKey, Text
from app.core.database import Base

# Textos repetidos
_PENDIENTE = "Pendiente"

RESULTADOS = ["Contestó", "No contestó", "Número errado", "Se negó", "Dejó mensaje"]
ESTADOS = [_PENDIENTE, "Cumplido", "Incumplido"]


class Llamada(Base):
    """Registro de una llamada a un contribuyente y su compromiso de pago."""

    __tablename__ = "compromisos_llamadas"

    id = Column(Integer, primary_key=True, index=True)
    ruc = Column(String(11), nullable=False, index=True)
    razon_social = Column(String(255), default="")
    telefono = Column(String(50), default="")
    contacto = Column(String(120), default="")

    fecha_llamada = Column(DateTime, default=datetime.datetime.now, index=True)
    resultado = Column(String(30), default="Contestó", index=True)

    fecha_compromiso = Column(Date, nullable=True, index=True)
    monto_compromiso = Column(Float, nullable=True)
    estado = Column(String(20), default=_PENDIENTE, index=True)

    proxima_llamada = Column(Date, nullable=True)
    observaciones = Column(Text, default="")

    usuario_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    usuario_nombre = Column(String(100), default="")
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    is_active = Column(Boolean, default=True)

    @property
    def estado_efectivo(self):
        """Pendiente con fecha pasada se muestra como Vencido."""
        if not self.fecha_compromiso:
            return "Sin compromiso"
        if self.estado == _PENDIENTE and self.fecha_compromiso < datetime.date.today():
            return "Vencido"
        return self.estado

    def to_dict(self):
        fmt = lambda d: d.strftime("%Y-%m-%d") if d else ""
        return {
            "id": self.id,
            "ruc": self.ruc,
            "razon_social": self.razon_social or "",
            "telefono": self.telefono or "",
            "contacto": self.contacto or "",
            "fecha_llamada": self.fecha_llamada.strftime("%Y-%m-%dT%H:%M") if self.fecha_llamada else "",
            "resultado": self.resultado,
            "fecha_compromiso": fmt(self.fecha_compromiso),
            "monto_compromiso": self.monto_compromiso,
            "estado": self.estado,
            "estado_efectivo": self.estado_efectivo,
            "proxima_llamada": fmt(self.proxima_llamada),
            "observaciones": self.observaciones or "",
            "usuario_nombre": self.usuario_nombre or "",
        }
