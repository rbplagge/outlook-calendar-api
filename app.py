# ---------- Endpoints (all require API key) ----------

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
