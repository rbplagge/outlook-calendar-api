import os, datetime, collections
import httpx
from fastapi import FastAPI, HTTPException, Security
from fastapi.security import APIKeyHeader
from msal import ConfidentialClientApplication

app = FastAPI()

# ---------------------------
# Config from environment vars
# ---------------------------
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
API_KEY = os.getenv("API_KEY")

if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, API_KEY]):
    raise RuntimeError("Missing one or more required environment variables.")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

GRAPH = "https://graph.microsoft.com/v1.0"

# ---------------------------
# MSAL client
# ---------------------------
msal_app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=AUTHORITY
)

def get_token():
    """Acquire or refresh an access token via MSAL."""
    result = msal_app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in result:
        error_msg = result.get("error_description", "Unknown error from MSAL")
        raise HTTPException(status_code=500, detail=f"Could not obtain access token: {error_msg}")
    return result["access_token"]

async def gget(path, params=None):
    """Helper to call Microsoft Graph with a valid token."""
    token = get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(GRAPH + path,
                             headers={"Authorization": f"Bearer {token}"},
                             params=params or {})
    if r.status_code >= 400:
        raise HTTPException(status_code=r.status_code, detail=r.text)
    return r.json()

# ---------------------------
# API Key protection
# ---------------------------
api_key_header = APIKeyHeader(name="x-api-key", auto_error=True)

def verify_api_key(x_api_key: str = Security(api_key_header)):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid or missing API Key")
    return True

# ---------------------------
# Endpoints
# ---------------------------

# Public (no API key)
@app.get("/profile")
async def profile():
    data = await gget("/me/mailboxSettings")
    return {
        "timeZone": data.get("timeZone"),
        "workingHours": data.get("workingHours", {})
    }

# Protected with API key
@app.get("/calendar/view")
async def view(start: str, end: str, authorized: bool = Security(verify_api_key)):
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "subject,start,end,isAllDay,showAs,categories,location"
    }
    return await gget("/me/calendarView", params)

# Protected with API key
@app.get("/stats")
async def stats(start: str, end: str, groupBy: str = "category", authorized: bool = Security(verify_api_key)):
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "start,end,showAs,categories"
    }
    res = await gget("/me/calendarView", params)

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
