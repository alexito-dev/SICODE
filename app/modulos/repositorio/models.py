import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime, ForeignKey, Text
from sqlalchemy.orm import relationship
from app.core.database import Base


class RepoArchivo(Base):
    """
    Modelo ORM para almacenar el registro de archivos subidos al Repositorio Compartido.
    Permite acceso común a todos los usuarios del sistema.
    """
    __tablename__ = "repositorio_archivos"

    id = Column(Integer, primary_key=True, index=True)
    filename_original = Column(String(255), nullable=False, index=True)
    filename_storage = Column(String(255), nullable=False, unique=True, index=True)
    file_path = Column(String(500), nullable=False)
    file_size = Column(Integer, nullable=False, default=0)
    file_size_formatted = Column(String(50), default="0 B")
    extension = Column(String(20), default="", index=True)
    category = Column(String(50), default="General", index=True)
    description = Column(Text, nullable=True)
    
    # Usuario que subió el archivo
    uploader_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploader_name = Column(String(100), default="Invitado")
    uploader_username = Column(String(100), nullable=True)
    
    # Métricas y auditoría
    downloads_count = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    is_active = Column(Boolean, default=True)

    # Relación con usuario si existe
    uploader = relationship("User", foreign_keys=[uploader_id], lazy="joined")

    def to_dict(self):
        return {
            "id": self.id,
            "filename_original": self.filename_original,
            "filename_storage": self.filename_storage,
            "file_size": self.file_size,
            "file_size_formatted": self.file_size_formatted,
            "extension": self.extension,
            "category": self.category,
            "description": self.description or "",
            "uploader_name": self.uploader_name or "Invitado",
            "uploader_username": self.uploader_username or "",
            "downloads_count": self.downloads_count or 0,
            "created_at": self.created_at.strftime("%d/%m/%Y %H:%M") if self.created_at else "",
            "is_active": self.is_active
        }
