from app.core.database import SessionLocal
from app.modulos.auth.models import User, PasswordResetToken
import secrets
import datetime

def test_flow():
    db = SessionLocal()
    email = "ranco@sunat.gob.pe"

    # 1. Simular solicitud de recuperación
    user = db.query(User).filter(User.correo == email).first()
    if not user:
        print("Error: Usuario de prueba no encontrado.")
        return

    # Generar token manual (lo que haría el router)
    token = secrets.token_urlsafe(32)
    new_token = PasswordResetToken(
        user_id=user.id,
        token=token,
        expires_at=datetime.datetime.now() + datetime.timedelta(hours=1)
    )
    db.add(new_token)
    db.commit()
    print(f"Token generado: {token}")

    # 2. Simular cambio de clave via reset
    token_rec = db.query(PasswordResetToken).filter(PasswordResetToken.token == token).first()
    if token_rec and not token_rec.is_expired:
        u = token_rec.user
        u.password = "NuevaClave123"
        token_rec.used = True
        db.commit()
        print("Clave actualizada exitosamente en la simulación.")
    else:
        print("Error: Token inválido en la simulación.")

    # 3. Verificar en BD
    updated_user = db.query(User).filter(User.correo == email).first()
    if updated_user.password == "NuevaClave123":
        print("VERIFICACIÓN EXITOSA: La base de datos tiene la nueva clave.")
    else:
        print("VERIFICACIÓN FALLIDA: La clave no cambió.")

    db.close()

if __name__ == "__main__":
    test_flow()
