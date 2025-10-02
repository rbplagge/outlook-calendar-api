import os
import datetime
import collections
import httpx
from fastapi import FastAPI, HTTPException, Header, Request
from msal import ConfidentialClientApplication

# -------------------- App --------------------
app = FastAPI(title="Outlook Calendar API", version="1.0")

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]  # Application permissions

REQUIRED_ENV_KEYS = [
    "AZURE_CLIENT_ID",
    "AZURE_CLIENT_SECRET",
    "AZURE_TENANT_ID",
    "API_KEY",
    "TARGET_USER",
]

# -------------------- Helpers --------------------
def get_env():
    cfg = {k: os.getenv(k) for k in REQUIRED_ENV_KEYS}
    missing = [k for k, v in cfg.items() if not v]
    if missing:
        raise HTTPException(status_code=500, detail=f"Missing required env vars: {', '.join(missing)}")
    return cfg

def make_msal(cfg):
    authority = f"https://login.microsoftonline.com/{cfg['AZURE_TENANT_ID']}"
    return ConfidentialClientApplication(
        client_id=cfg["AZURE_CLIENT_ID"],
        client_credential=cfg["AZURE_CLIENT_SECRET"],
        authority=authority
    )

def get_token(msal_app):
    result = msal_app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or "Unknown MSAL error"
        raise HTTPException(status_code=500, detail=f"Could not obtain access token: {err}")
    return result["access_token"]

async def gget(path, params=None):
    cfg = get_env()
    msal_client = make_msal(cfg)
    token = get_token(msal_client)
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(GRAPH + path, headers={"Authorization": f"Bearer {token}"}, params=params or {})
    if r.status_code >= 400:
        raise HTTPException(status_code=r.status_code, detail=r.text)
    return r.json()

def _mask(s: str) -> str:
    if not s:
        return ""
    return (s[:3] + "..." + s[-3:]) if len(s) > 6 else "***"

def require_key(x_api_key: str = Header(..., alias="x-api-key")):
    env_key = os.getenv("API_KEY") or ""
    # Normalize/trim both sides to avoid invisible whitespace issues
    sent = (x_api_key or "").strip()
    kept = env_key.strip()
    if sent != kept:
        raise HTTPException(
            status_code=403,
            detail=f"Invalid or missing API Key (len_sent={len(sent)}, len_env={len(kept)})"
        )

# -------------------- Diagnostics --------------------
@app.get("/env-check")
async def env_check():
    present = {k: bool(os.getenv(k)) for k in REQUIRED_ENV_KEYS}
    return {"present": present}

@app.get("/ping")
async def ping(request: Request):
    return {"headers": dict(request.headers)}

@app.get("/debug-key")
async def debug_key():
    val = os.getenv("API_KEY") or ""
    return {
        "api_key_present": bool(val),
        "api_key_length": len(val.strip()),
        "api_key_preview": _mask(val.strip())
    }

@app.get("/key-compare")
async def key_compare(request: Request, x_api_key: str = Header(None, alias="x-api-key")):
    env_key = os.getenv("API_KEY") or ""
    sent = (x_api_key or "").strip()
    kept = env_key.strip()
    return {
        "received_header": bool(x_api_key),
        "match_after_trim": sent == kept,
        "sent_len": len(sent),
        "env_len": len(kept),
        "sent_preview": _mask(sent),
        "env_preview": _mask(kept)
    }

# -------------------- Application endpoints --------------------
@app.get("/profile")
async def profile(x_api_key: str = Header(..., alias="x-api-key")):
    require_key(x_api_key)
    cfg = get_env()
    user = cfg["TARGET_USER"]
    data = await gget(f"/users/{user}/mailboxSettings")
    return {"timeZone": data.get("timeZone"), "workingHours": data.get("workingHours", {})}

@app.get("/calendar/view")
async def view(start: str, end: str, x_api_key: str = Header(..., alias="x-api-key")):
    require_key(x_api_key)
    cfg = get_env()
    user = cfg["TARGET_USER"]
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "subject,start,end,isAllDay,showAs,categories,location,organizer"
    }
    return await gget(f"/users/{user}/calendarView", params)

@app.get("/stats")
async def stats(start: str, end: str, groupBy: str = "category",
                x_api_key: str = Header(..., alias="x-api-key")):
    require_key(x_api_key)
    cfg = get_env()
    user = cfg["TARGET_USER"]
    params = {"startDateTime": start, "endDateTime": end, "$select": "start,end,showAs,categories"}
    res = await gget(f"/users/{user}/calendarView", params)

    def hours(ev):
        s = datetime.datetime.fromisoformat(ev["start"]["dateTime"].replace("Z", "+00:00"))
        e = datetime.datetime.fromisoformat(ev["end"]["dateTime"].replace("Z", "+00:00"))
        return (e - s).total_seconds() / 3600

    bucket = collections.Counter()
    for ev in res.get("value", []):
        if groupBy == "showAs":
            bucket[(ev.get("showAs") or "unknown")] += hours(ev)
        else:
            cats = ev.get("categories") or ["Uncategorized"]
            bucket[cats[0]] += hours(ev)
    return dict(bucket)
