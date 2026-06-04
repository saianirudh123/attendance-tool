import os
from functools import wraps
from typing import Optional

from authlib.integrations.starlette_client import OAuth
from dotenv import load_dotenv
from fastapi import HTTPException, Request

from .database import get_user, upsert_user

load_dotenv()

SESSION_SECRET = os.getenv("SESSION_SECRET", "change-me-in-production")
GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
APP_BASE_URL = os.getenv("APP_BASE_URL", "http://127.0.0.1:8000").rstrip("/")
ALLOWED_GOOGLE_DOMAINS = {
    d.strip().lower()
    for d in os.getenv("ALLOWED_GOOGLE_DOMAINS", "").split(",")
    if d.strip()
}
ADMIN_EMAILS = {
    e.strip().lower()
    for e in os.getenv("ADMIN_EMAILS", "").split(",")
    if e.strip()
}

oauth = OAuth()
if GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET:
    oauth.register(
        name="google",
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
        client_kwargs={"scope": "openid email profile"},
    )


def oauth_configured() -> bool:
    return bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET)


def is_allowed_email(email: str) -> bool:
    email = (email or "").strip().lower()
    if not email or "@" not in email:
        return False
    domain = email.split("@", 1)[1]
    return not ALLOWED_GOOGLE_DOMAINS or domain in ALLOWED_GOOGLE_DOMAINS


def current_user_from_session(request: Request) -> Optional[dict]:
    email = request.session.get("user_email")
    if not email:
        return None
    return get_user(email)


def require_user(request: Request) -> dict:
    user = current_user_from_session(request)
    if not user:
        raise HTTPException(status_code=401, detail="Login required")
    return user


def require_admin(request: Request) -> dict:
    user = require_user(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return user


def login_user(request: Request, profile: dict) -> dict:
    email = (profile.get("email") or "").strip().lower()
    if not is_allowed_email(email):
        raise HTTPException(status_code=403, detail="Email domain is not allowed")
    user = upsert_user(
        email=email,
        name=profile.get("name") or email,
        picture=profile.get("picture") or "",
        admin_emails=ADMIN_EMAILS,
    )
    request.session["user_email"] = user["email"]
    return user
