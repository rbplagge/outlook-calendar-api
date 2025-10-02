import os
import datetime
import collections
import httpx
from fastapi import FastAPI, HTTPException, Header
from msal import ConfidentialClientApplication

# -----------------------------------------------------
# Create FastAPI app (must come before any route!)
# -----------------------------------------------------
app = FastAPI(title="Outlook Calendar API", version="1.0")

# -----------------------------------------------------
# Config and constants
# -----------------------------------------------------
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]  # Application permissions

REQUIRED_ENV_KEYS = [
    "AZURE_CLIENT_ID",
    "AZURE_CLIENT_SECRET",
    "AZURE_TENANT_ID",
    "API_KEY",
    "TARGET_USER",  # userPrincipalName/email in your tenant
]

# -----------------------------------------------------
# Helpers: env vars, MSAL auth, API key
# -----------------------------------------------------
def get_env():
    """Load env vars each request so we don't crash at import time."""
    cfg = {k: os.getenv(k) for k in REQUIRED_ENV_KEYS}
    missing = [k for k, v in cfg.items() if not v]
    if missing:
        raise HTTPException(status_code=500,
                            detail=f"Missing required env vars: {', '.join(missing)}")
    return cfg


def make_msal(cfg):
    """Create MSAL client."""
    authority = f"https://login.microsoftonline.com/{cfg['AZURE_TENANT_ID']}"
    return ConfidentialClientApplication(
        client_id=cfg["AZURE_CLIENT_ID"],
        client_credential=cfg["AZURE_CLIENT_SECRET"],
        authority=authority
    )


def get_token(msal_app):
    """Acquire or refresh an access token via MSAL."""
    result = msal_app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or "Unknown MSAL error"
        raise HTTPException(status_code=500, detail=f"Could not obtain access token: {err}")
    return result["access_token"]


async def gget(path, params=None):
    """Helper to call Microsoft Graph with a valid token."""
    cfg = get_env()
    msal_app = make_msal(cfg)
    token = get_token(msal_app)
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(
            GRAPH + path,
            headers={"Authorization": f"Bearer {token}"},
            params=params or {}
        )
    if r.status_code >= 400:
        raise HTTPException(status_code=r.status_code, detail=r.text)
    return r.json()


def require_key(x_api_key: str = Header(..., alias="x-api-key")):
    """Check API key header against env var."""
    api_key = os.getenv("API_KEY")
    if x_api_key != api_key:
        raise HTTPException(status_code=403, detail="Invalid or missing API Key")

# -----------------------------------------------------
# Diagnostics
# -----------------------------------------------------
@app.get("/env-check")
async def env_check():
    """
    Returns which required env keys are present (True/False) without exposing values.
    """
    present = {k: bool(os.getenv(k)) for k in REQUIRED_ENV_KEYS}
    return {"present": present}

# -----------------------------------------------------
# Endpoints
# -----------------------------------------------------
@app.get("/profile")
async def profile(x_api_key: str = Header(..., alias="x-api-key")):
    require_key(x_api_key)
    cfg = get_env()
    user = cfg["TARGET_USER"]
    data = await gget(f"/users/{user}/mailboxSettings")
    return {
        "timeZone": data.get("timeZone"),
        "workingHours": data.get("workingHours", {})
    }


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
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "start,end,showAs,categories"
    }
    res = await gget(f"/users/{user}/calendarView", params)

    def hours(ev):
        s = datetime.datetime.fromisoformat(ev["start"]["dateTime"].replace("Z", "+00:00"))
        e = datetime.datetime.fromisoformat(ev["end"]["dateTime"].replace("Z", "+00:00"))
        return (e - s).total_seconds() / 3600

    bucket = collections.Counter()
    for ev in res.get("value", []):
        if groupBy == "showAs":
            key = ev.get("showAs") or "unknown"
            bucket[key] += hours(ev)
        else:
            cats = ev.get("categories") or ["Uncategorized"]
            bucket[cats[0]] += hours(ev)
    return dict(bucket)
