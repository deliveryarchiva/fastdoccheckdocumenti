import hashlib, json, os, uuid
from pathlib import Path
from typing import Optional
from fastapi import Header, HTTPException

DATA_DIR   = Path(os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "data"))
USERS_FILE = DATA_DIR / "users.json"
DATA_DIR.mkdir(parents=True, exist_ok=True)

sessions: dict[str, dict] = {}

def hash_password(plain: str) -> str:
    return hashlib.sha256(plain.encode()).hexdigest()

def verify_password(plain: str, hashed: str) -> bool:
    return hash_password(plain) == hashed

def load_users() -> dict:
    try:
        if USERS_FILE.exists():
            return json.loads(USERS_FILE.read_text("utf-8"))
    except Exception:
        pass
    return {}

def save_user(user: dict):
    users = load_users()
    users[user["username"].lower()] = user
    tmp = str(USERS_FILE) + ".tmp"
    Path(tmp).write_text(json.dumps(users, indent=2, ensure_ascii=False), "utf-8")
    Path(tmp).rename(USERS_FILE)

def delete_user(username: str):
    users = load_users()
    users.pop(username.lower(), None)
    tmp = str(USERS_FILE) + ".tmp"
    Path(tmp).write_text(json.dumps(users, indent=2, ensure_ascii=False), "utf-8")
    Path(tmp).rename(USERS_FILE)

def seed_users():
    if load_users():
        return
    pwd = hash_password(os.getenv("DEFAULT_PASSWORD", "archiva2026"))
    for u in [
        {"nome": "Marco Pastore",     "username": "marco.pastore",     "ruolo": "admin", "role": "Head of Project Delivery"},
        {"nome": "Paolo Gandini",      "username": "paolo.gandini",      "ruolo": "admin", "role": "Delivery & Customer Service Director"},
        {"nome": "Chiara Pettenuzzo", "username": "chiara.pettenuzzo", "ruolo": "admin", "role": "Service Delivery Manager"},
    ]:
        save_user({**u, "password": pwd})

def create_session(user: dict) -> str:
    token = str(uuid.uuid4())
    sessions[token] = {
        "username": user["username"], "nome": user["nome"],
        "ruolo": user["ruolo"], "role": user.get("role", "")
    }
    return token

def get_current_user(authorization: Optional[str] = Header(default=None)) -> dict:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Non autenticato")
    user = sessions.get(authorization.removeprefix("Bearer ").strip())
    if not user:
        raise HTTPException(status_code=401, detail="Sessione scaduta")
    return user

def require_admin(user: dict):
    if user.get("ruolo") != "admin":
        raise HTTPException(status_code=403, detail="Accesso riservato agli amministratori")
