from sqlalchemy import Column, Integer, String, LargeBinary, DateTime, ForeignKey
from sqlalchemy.orm import relationship
from app.core.database import Base
import datetime


class DocCheckFile(Base):
    """
    Almacena el archivo Excel de cada usuario en la base de datos.
    - Se guarda como bytes (LargeBinary) → sin archivos en disco.
    - Se actualiza cada vez que el usuario guarda un cambio.
    - Se carga automáticamente al abrir DocCheck.
    """

    __tablename__ = "doccheck_files"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), unique=True, nullable=False)
    filename = Column(String, nullable=False)  # Nombre original del archivo
    file_data = Column(LargeBinary, nullable=False)  # Bytes del Excel
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.datetime.utcnow, onupdate=datetime.datetime.utcnow)

    user = relationship("User")
