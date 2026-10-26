from pydantic import BaseModel


class User(BaseModel):
    id: int
    username: str
    email: str
    full_name: str = ""
    disabled: bool = False


# Mock Database
class UserDB:
    users = {
        "admin": {
            "id": 1,
            "username": "admin",
            "password": "admin123",  # In production use hashed passwords!
            "email": "admin@sicode.com",
            "full_name": "Administrador del Sistema",
        },
        "usuario": {
            "id": 2,
            "username": "usuario",
            "password": "user123",
            "email": "user@sicode.com",
            "full_name": "Operador de Sistema",
        },
    }

    @staticmethod
    def get_user(username: str):
        return UserDB.users.get(username)
