from fastapi import FastAPI, UploadFile, File, HTTPException, Body, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.sessions import SessionMiddleware
import uvicorn, shutil, uuid
from pathlib import Path
from .parser import parse_pdf_punch, parse_xlsx_meta
from .processor import build_attendance_workbook, rebuild_from_json, parse_generated_register
from .database import (
    init_db, freeze_register, list_registers, get_register_detail, get_frozen_file,
    submit_register_for_approval, list_pending_registers, get_pending_register,
    get_pending_file, mark_pending_status, list_users, set_user_role,
)
from .auth import (
    APP_BASE_URL, SESSION_SECRET, oauth, oauth_configured, login_user,
    current_user_from_session, require_user, require_admin,
)

BASE       = Path(__file__).parent.parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Krishna Engineering Attendance Tool")
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    same_site="lax",
    https_only=APP_BASE_URL.startswith("https://"),
)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.mount("/static", StaticFiles(directory=str(BASE / "frontend" / "static")), name="static")


@app.on_event("startup")
def startup():
    init_db()


@app.get("/")
def index(request: Request):
    if not current_user_from_session(request):
        return RedirectResponse("/login")
    return FileResponse(str(BASE / "frontend" / "index.html"))


@app.get("/login")
def login_page():
    return FileResponse(str(BASE / "frontend" / "login.html"))


@app.get("/auth/login")
async def auth_login(request: Request):
    if not oauth_configured():
        raise HTTPException(500, "Google OAuth is not configured")
    redirect_uri = f"{APP_BASE_URL}/auth/callback"
    return await oauth.google.authorize_redirect(request, redirect_uri)


@app.get("/auth/signup")
async def auth_signup(request: Request):
    return await auth_login(request)


@app.get("/auth/callback")
async def auth_callback(request: Request):
    if not oauth_configured():
        raise HTTPException(500, "Google OAuth is not configured")
    token = await oauth.google.authorize_access_token(request)
    profile = token.get("userinfo") or await oauth.google.parse_id_token(request, token)
    login_user(request, profile)
    return RedirectResponse("/")


@app.get("/auth/logout")
def auth_logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login")


@app.get("/api/auth/me")
def auth_me(request: Request):
    user = current_user_from_session(request)
    return JSONResponse({"authenticated": bool(user), "user": user})


# ── Process uploaded files ─────────────────────────────────────────────────────

def _import_register_response(xlsx_path: Path, out_path: Path, session_id: str):
    summary = parse_generated_register(str(xlsx_path))
    shutil.copy(str(xlsx_path), str(out_path))
    return JSONResponse({
        "session_id":   session_id,
        "summary":      summary,
        "download_url": f"/api/download/{session_id}",
    })


@app.post("/api/process")
async def process_files(
    request: Request,
    punch_pdf:      UploadFile | None = File(None),
    leave_xlsx:     UploadFile | None = File(None),
    register_xlsx:  UploadFile | None = File(None),
):
    require_user(request)
    session_id = str(uuid.uuid4())[:8]
    pdf_path   = UPLOAD_DIR / f"{session_id}_punch.pdf"
    xlsx_path  = UPLOAD_DIR / f"{session_id}_leave.xlsx"
    reg_path   = UPLOAD_DIR / f"{session_id}_register.xlsx"
    out_path   = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"

    try:
        if register_xlsx is not None:
            with open(reg_path, "wb") as f:
                shutil.copyfileobj(register_xlsx.file, f)
            return _import_register_response(reg_path, out_path, session_id)

        if punch_pdf is None or leave_xlsx is None:
            raise HTTPException(
                status_code=400,
                detail="Upload either a generated register workbook, or both punch PDF and leave XLSX.",
            )

        with open(pdf_path,  "wb") as f: shutil.copyfileobj(punch_pdf.file,  f)
        with open(xlsx_path, "wb") as f: shutil.copyfileobj(leave_xlsx.file, f)

        punch_data, name_to_code, year, month = parse_pdf_punch(str(pdf_path))
        meta    = parse_xlsx_meta(str(xlsx_path))
        summary = build_attendance_workbook(punch_data, name_to_code, meta, str(out_path),
                                            year=year, month=month)

        return JSONResponse({
            "session_id":   session_id,
            "summary":      summary,
            "download_url": f"/api/download/{session_id}",
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        pdf_path.unlink(missing_ok=True)
        xlsx_path.unlink(missing_ok=True)
        reg_path.unlink(missing_ok=True)


@app.post("/api/import-register")
@app.post("/api/import_register")
@app.post("/api/process-register")
async def import_generated_register(request: Request, register_xlsx: UploadFile = File(...)):
    require_user(request)
    session_id = str(uuid.uuid4())[:8]
    xlsx_path  = UPLOAD_DIR / f"{session_id}_register.xlsx"
    out_path   = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"

    try:
        with open(xlsx_path, "wb") as f:
            shutil.copyfileobj(register_xlsx.file, f)

        return _import_register_response(xlsx_path, out_path, session_id)
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        xlsx_path.unlink(missing_ok=True)


# ── Rebuild from frontend edits ────────────────────────────────────────────────

@app.post("/api/rebuild/{session_id}")
async def rebuild_workbook(request: Request, session_id: str, payload: dict = Body(...)):
    require_user(request)
    out_path = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"
    try:
        from datetime import date as _date
        employees = payload["employees"]
        year      = int(payload["year"])
        month     = int(payload["month_num"])
        holidays  = [_date.fromisoformat(h) for h in payload.get("holidays", [])]
        rebuild_from_json(employees, holidays, str(out_path), year, month)
        return JSONResponse({"download_url": f"/api/download/{session_id}"})
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ── Freeze register to local database ─────────────────────────────────────────

@app.post("/api/freeze/{session_id}")
async def freeze_register_endpoint(request: Request, session_id: str, payload: dict = Body(...)):
    require_admin(request)
    out_path = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"
    try:
        from datetime import date as _date
        employees = payload["employees"]
        year      = int(payload["year"])
        month     = int(payload["month_num"])
        holidays  = [_date.fromisoformat(h) for h in payload.get("holidays", [])]

        # Always rebuild with latest edits before freezing
        rebuild_from_json(employees, holidays, str(out_path), year, month)

        summary = {
            "month":     payload.get("month", f"{year}-{month:02d}"),
            "year":      year,
            "month_num": month,
        }

        reg_id = freeze_register(summary, employees, str(out_path), session_id)

        return JSONResponse({
            "register_id":   reg_id,
            "download_url":  f"/api/download/{session_id}",
            "message":       "Register frozen successfully",
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/submit/{session_id}")
async def submit_for_approval(request: Request, session_id: str, payload: dict = Body(...)):
    user = require_user(request)
    out_path = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"
    try:
        from datetime import date as _date
        employees = payload["employees"]
        year      = int(payload["year"])
        month     = int(payload["month_num"])
        holidays  = [_date.fromisoformat(h) for h in payload.get("holidays", [])]

        rebuild_from_json(employees, holidays, str(out_path), year, month)
        summary = {
            "month":     payload.get("month", f"{year}-{month:02d}"),
            "year":      year,
            "month_num": month,
        }
        pending_id = submit_register_for_approval(
            summary, employees, str(out_path), session_id, user["email"]
        )
        return JSONResponse({
            "pending_id": pending_id,
            "message": "Register submitted for admin approval",
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/pending")
def list_pending_endpoint(request: Request):
    require_admin(request)
    return JSONResponse(list_pending_registers())


@app.get("/api/pending/{pending_id}")
def get_pending_endpoint(request: Request, pending_id: int):
    require_admin(request)
    item = get_pending_register(pending_id)
    if not item:
        raise HTTPException(404, "Pending register not found")
    return JSONResponse(item)


@app.get("/api/pending/{pending_id}/download")
def download_pending(request: Request, pending_id: int):
    require_admin(request)
    path = get_pending_file(pending_id)
    if not path:
        raise HTTPException(404, "Pending file not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Attendance_Register_pending.xlsx",
    )


@app.post("/api/pending/{pending_id}/freeze")
def freeze_pending_endpoint(request: Request, pending_id: int):
    require_admin(request)
    item = get_pending_register(pending_id)
    path = get_pending_file(pending_id)
    if not item or not path:
        raise HTTPException(404, "Pending register not found")
    if item["status"] != "pending":
        raise HTTPException(400, "Pending register has already been handled")
    reg_id = freeze_register(item["summary"], item["employees"], path, item["session_id"])
    mark_pending_status(pending_id, "approved")
    return JSONResponse({
        "register_id": reg_id,
        "message": "Register approved and frozen",
    })


@app.get("/api/users")
def list_users_endpoint(request: Request):
    require_admin(request)
    return JSONResponse(list_users())


@app.post("/api/users/{email}/role")
def set_user_role_endpoint(request: Request, email: str, payload: dict = Body(...)):
    require_admin(request)
    user = set_user_role(email, payload.get("role"))
    if not user:
        raise HTTPException(404, "User not found")
    return JSONResponse(user)


# ── Previous registers ─────────────────────────────────────────────────────────

@app.get("/api/registers")
def list_registers_endpoint(request: Request):
    require_user(request)
    return JSONResponse(list_registers())


@app.get("/api/registers/{register_id}")
def get_register_endpoint(request: Request, register_id: int):
    require_user(request)
    reg = get_register_detail(register_id)
    if not reg:
        raise HTTPException(404, "Register not found")
    return JSONResponse(reg)


@app.get("/api/registers/{register_id}/download")
def download_frozen(request: Request, register_id: int):
    require_user(request)
    path = get_frozen_file(register_id)
    if not path:
        raise HTTPException(404, "Frozen file not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Attendance_Register_frozen.xlsx",
    )


# ── Session download ───────────────────────────────────────────────────────────

@app.get("/api/download/{session_id}")
def download(request: Request, session_id: str):
    require_user(request)
    path = OUTPUT_DIR / f"Attendance_Register_{session_id}.xlsx"
    if not path.exists():
        raise HTTPException(404, "File not found or expired")
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Attendance_Register.xlsx",
    )


if __name__ == "__main__":
    uvicorn.run("backend.main:app", host="0.0.0.0", port=8000, reload=True)
