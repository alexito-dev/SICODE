from sqlalchemy import create_engine, event
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from app.core.config import settings

# Crear el motor de base de datos SQLite
# connect_args={"check_same_thread": False} es necesario solo para SQLite
engine = create_engine(
    settings.DATABASE_URL, 
    connect_args={"check_same_thread": False, "timeout": 15}
)

# Crear una clase SessionLocal configurada
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Clase base para nuestros modelos
Base = declarative_base()

def get_db():
    """
    Generador de dependencias para obtener una sesión de base de datos.
    Asegura que la sesión se cierre después de cada petición.
    """
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
