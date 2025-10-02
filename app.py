from fastapi import FastAPI, Header, HTTPException
import httpx, datetime, collections

# Create the FastAPI app
app = FastAPI()

# Base URL for Microsoft Graph API
GRAPH = "https://graph.microsoft.com/v1.0"

# Helper function to call Microsoft Graph
async def gget(path, token, params=None):
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(GRAPH + path,
                             headers={"Authorization": f"Bearer {token}"},
                             params=params or {})
    if r.status_code >= 400:
        raise HTTPException(r.status_code, r.text)
    return r.json()

# Endpoint to return mailbox settings (timezone + working hours)
@app.get("/profile")
async def profile(authorization: str = Header(...)):
    token = authorization.split("Bearer ")[-1]
    data = await gget("/me/mailboxSettings", token)
    return {
        "timeZone": data.get("timeZone"),
        "workingHours": data.get("workingHours", {})
    }

# Endpoint to return calendar events between two dates
@app.get("/calendar/view")
async def view(start: str, end: str, authorization: str = Header(...)):
    token = authorization.split("Bearer ")[-1]
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "subject,start,end,isAllDay,showAs,categories,location,organizer"
    }
    return await gget("/me/calendarView", token, params)

# Endpoint to summarize time by category or by Busy/Free status
@app.get("/stats")
async def stats(start: str, end: str, groupBy: str = "category", authorization: str = Header(...)):
    token = authorization.split("Bearer ")[-1]
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "start,end,showAs,categories"
    }
    res = await gget("/me/calendarView", token, params)

    def hours(ev):
        s = datetime.datetime.fromisoformat(ev["start"]["dateTime"].replace("Z","+00:00"))
        e = datetime.datetime.fromisoformat(ev["end"]["dateTime"].replace("Z","+00:00"))
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
