import datetime, collections
import httpx
from fastapi import FastAPI, HTTPException
from msal import ConfidentialClientApplication
import os

app = FastAPI()

# Get credentials from environment variables
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Calendars.Read", "MailboxSettings.Read"]

# MSAL confidential client
msal_app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=AUTHORITY
)

GRAPH = "https://graph.microsoft.com/v1.0"


def get_token():
    """Acquire or refresh a token silently with client credentials."""
    result = msal_app.acquire_token_silent(SCOPE, account=None)
    if not result:
        result = msal_app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in result:
        raise HTTPException(status_code=500, detail="Could not obtain access token")
    return result["access_token"]


async def gget(path, params=None):
    """Helper to call Microsoft Graph with a valid token."""
    token = get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(
            GRAPH + path,
            headers={"Authorization": f"Bearer {token}"},
            params=params or {}
        )
    if r.status_code >= 400:
        raise HTTPException(r.status_code, r.text)
    return r.json()


@app.get("/profile")
async def profile():
    data = await gget("/me/mailboxSettings")
    return {
        "timeZone": data.get("timeZone"),
        "workingHours": data.get("workingHours", {})
    }


@app.get("/calendar/view")
async def view(start: str, end: str):
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "subject,start,end,isAllDay,showAs,categories,location"
    }
    return await gget("/me/calendarView", params)


@app.get("/stats")
async def stats(start: str, end: str, groupBy: str = "category"):
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
