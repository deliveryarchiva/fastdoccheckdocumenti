import logging
import os
import shutil
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, File, Header, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates

from auth import (
    create_session, delete_user, get_current_user, hash_password,
    load_users, require_admin, save_user, seed_users, sessions,
    verify_password,
)
from data_loader import compute_stats, load_data, search_records

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="Archiva File DB")
templates = Jinja2Templates(directory="templates")

DATA_DIR   = Path(os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_PATH = DATA_DIR / "database.xlsx"

# ── In-memory database ──────────────────────────────────────────────────────
db_records: list = []
db_stats:   dict = {}
db_loaded:  bool = False


def reload_db():
    global db_records, db_stats, db_loaded
    if EXCEL_PATH.exists():
        logger.info(f"Loading Excel from {EXCEL_PATH} …")
        db_records = load_data(str(EXCEL_PATH))
        db_stats   = compute_stats(db_records)
        db_loaded  = True
        logger.info(f"Done – {len(db_records):,} records")
    else:
        logger.warning("Excel file not found – search will be unavailable")
        db_records = []
        db_stats   = {}
        db_loaded  = False


@app.on_event("startup")
async def startup():
    seed_users()
    reload_db()


# ── HTML pages ───────────────────────────────────────────────────────────────
@app.get("/login", response_class=HTMLResponse)
def page_login(request: Request):
    return templates.TemplateResponse(request=request, name="login.html")

@app.get("/", response_class=HTMLResponse)
def page_index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/admin", response_class=HTMLResponse)
def page_admin(request: Request):
    return templates.TemplateResponse(request=request, name="admin.html")


# ── Auth endpoints ───────────────────────────────────────────────────────────
@app.post("/api/auth/login")
async def api_login(request: Request):
    body     = await request.json()
    username = (body.get("username") or "").strip().lower()
    password = body.get("password") or ""
    users    = load_users()
    user     = users.get(username)
    if not user or not verify_password(password, user["password"]):
        raise HTTPException(401, "Credenziali non valide")
    token = create_session(user)
    return {"ok": True, "token": token, "user": {
        "username": user["username"], "nome": user["nome"],
        "ruolo": user["ruolo"], "role": user.get("role", ""),
    }}

@app.post("/api/auth/logout")
def api_logout(request: Request):
    auth = request.headers.get("authorization", "")
    if auth.startswith("Bearer "):
        sessions.pop(auth.removeprefix("Bearer ").strip(), None)
    return {"ok": True}

@app.get("/api/auth/me")
def api_me(user=Depends(get_current_user)):
    return {"ok": True, "user": user}

@app.post("/api/auth/change-password")
async def api_change_password(request: Request, user=Depends(get_current_user)):
    body    = await request.json()
    users   = load_users()
    db_user = users.get(user["username"].lower())
    if not verify_password(body.get("currentPassword", ""), db_user["password"]):
        raise HTTPException(401, "Password attuale non corretta")
    db_user["password"] = hash_password(body["newPassword"])
    save_user(db_user)
    return {"ok": True}


# ── Admin: users ─────────────────────────────────────────────────────────────
@app.get("/api/admin/users")
def api_get_users(user=Depends(get_current_user)):
    require_admin(user)
    return {k: {kk: vv for kk, vv in v.items() if kk != "password"}
            for k, v in load_users().items()}

@app.post("/api/admin/users")
async def api_create_user(request: Request, user=Depends(get_current_user)):
    require_admin(user)
    body = await request.json()
    save_user({
        "username": body["username"].lower(),
        "nome":     body["nome"],
        "ruolo":    body.get("ruolo", "user"),
        "role":     body.get("role", ""),
        "password": hash_password(body["password"]),
    })
    return {"ok": True}

@app.put("/api/admin/users/{username}")
async def api_update_user(username: str, request: Request, user=Depends(get_current_user)):
    require_admin(user)
    body  = await request.json()
    users = load_users()
    u     = users.get(username.lower())
    if not u:
        raise HTTPException(404, "Utente non trovato")
    for field in ("nome", "ruolo", "role"):
        if field in body:
            u[field] = body[field]
    if body.get("password"):
        u["password"] = hash_password(body["password"])
    save_user(u)
    return {"ok": True}

@app.delete("/api/admin/users/{username}")
def api_delete_user(username: str, user=Depends(get_current_user)):
    require_admin(user)
    if username.lower() == user["username"].lower():
        raise HTTPException(400, "Non puoi eliminare te stesso")
    delete_user(username.lower())
    return {"ok": True}


# ── Admin: Excel upload & reload ─────────────────────────────────────────────
@app.post("/api/admin/upload")
async def api_upload(
    request: Request,
    file: UploadFile = File(...),
    user=Depends(get_current_user),
):
    require_admin(user)
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Sono accettati solo file .xlsx")

    tmp = DATA_DIR / "database_tmp.xlsx"
    with open(tmp, "wb") as f:
        shutil.copyfileobj(file.file, f)
    tmp.rename(EXCEL_PATH)

    reload_db()
    return {
        "ok":      True,
        "message": f"File caricato correttamente. {len(db_records):,} record elaborati.",
        "stats":   db_stats,
    }

@app.post("/api/admin/reload")
def api_reload(user=Depends(get_current_user)):
    require_admin(user)
    reload_db()
    return {
        "ok":      True,
        "message": f"Dati ricaricati. {len(db_records):,} record.",
        "stats":   db_stats,
    }

@app.get("/api/admin/db-status")
def api_db_status(user=Depends(get_current_user)):
    require_admin(user)
    return {
        "loaded":     db_loaded,
        "file_exists": EXCEL_PATH.exists(),
        "stats":      db_stats,
    }


# ── Search ───────────────────────────────────────────────────────────────────
@app.get("/api/search")
def api_search(
    ragione_sociale: str = "",
    piva:            str = "",
    nome_file:       str = "",
    file_match:      str = "partial",
    date_from:       str = "",
    date_to:         str = "",
    anno:            str = "",
    mese:            str = "",
    page:            int = 1,
    per_page:        int = 50,
    user=Depends(get_current_user),
):
    has_filter = any([ragione_sociale, piva, nome_file, date_from, date_to, anno, mese])

    if not db_loaded:
        return {
            "ok": False,
            "error": "Database non caricato. Caricare il file Excel dalla sezione Admin.",
            "total": 0, "page": 1, "pages": 0, "per_page": per_page,
            "results": [], "stats": db_stats,
        }

    if not has_filter:
        return {
            "ok": True,
            "no_filter": True,
            "message": "Inserire almeno un criterio di ricerca.",
            "total": 0, "page": 1, "pages": 0, "per_page": per_page,
            "results": [], "stats": db_stats,
        }

    results = search_records(
        db_records,
        ragione_sociale=ragione_sociale or None,
        piva=piva or None,
        nome_file=nome_file or None,
        file_match=file_match,
        date_from=date_from or None,
        date_to=date_to or None,
        anno=anno or None,
        mese=mese or None,
    )

    total    = len(results)
    per_page = max(10, min(per_page, 200))
    pages    = max(1, (total + per_page - 1) // per_page)
    page     = max(1, min(page, pages))
    start    = (page - 1) * per_page

    return {
        "ok":       True,
        "total":    total,
        "page":     page,
        "pages":    pages,
        "per_page": per_page,
        "results":  results[start: start + per_page],
        "stats":    db_stats,
    }

@app.get("/api/filters")
def api_filters(user=Depends(get_current_user)):
    """Return sorted list of unique years and months present in the dataset."""
    years  = sorted({r["data_doc"][:4] for r in db_records if r["data_doc"]}, reverse=True)
    months = sorted({r["data_doc"][5:7] for r in db_records if r["data_doc"]})
    return {"years": years, "months": months}

@app.get("/api/stats")
def api_stats(user=Depends(get_current_user)):
    return db_stats
