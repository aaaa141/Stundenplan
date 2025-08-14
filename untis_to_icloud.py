#!/usr/bin/env python3
import datetime as dt, os, pytz, requests, sys
from icalendar import Calendar, Event, vText
from caldav import DAVClient
from caldav.lib import error as caldav_error

TZ = pytz.timezone("Europe/Berlin")

# --- ENV / Secrets ---
UNTIS_SERVER = os.getenv("UNTIS_SERVER")     # z.B. poly.webuntis.com
UNTIS_SCHOOL = os.getenv("UNTIS_SCHOOL")     # z.B. August-bebel-schule
UNTIS_USER   = os.getenv("UNTIS_USER")
UNTIS_PASS   = os.getenv("UNTIS_PASS")

ICLOUD_USER  = os.getenv("ICLOUD_USER")
ICLOUD_PASS  = os.getenv("ICLOUD_PASS")

ICLOUD_CAL   = os.getenv("ICLOUD_CAL", "Stundenplan")

DAYS_PAST       = int(os.getenv("DAYS_PAST", "1"))
DAYS_AHEAD      = int(os.getenv("DAYS_AHEAD", "35"))
DELETE_MISSING  = os.getenv("DELETE_MISSING", "true").lower() == "true"
MARK_CANCELLED  = os.getenv("MARK_CANCELLED", "true").lower() == "true"
# ----------------------

def require_env(name: str):
    if not os.getenv(name):
        print(f"[FEHLT] {name} ist nicht gesetzt (Secret).")
        sys.exit(2)

def ymd(d): return int(d.strftime("%Y%m%d"))

def to_local(date_int, time_int):
    s=str(date_int); y,m,d=int(s[:4]),int(s[4:6]),int(s[6:8])
    h,mi=time_int//100,time_int%100
    return TZ.localize(dt.datetime(y,m,d,h,mi))

def labels(item,key):
    arr=item.get(key,[]) or []; out=[]
    for a in arr:
        v=a.get("longname") or a.get("name") or ""
        if v: out.append(v)
    return ", ".join(out)

def uid_for(lesson): return f"untis-{lesson.get('id')}@{UNTIS_SERVER}"

def untis_rpc(s: requests.Session, url: str, method: str, params: dict):
    """JSON‑RPC Call mit Fallbacks auf Legacy-Methoden."""
    payload = {"id":"id","method":method,"jsonrpc":"2.0","params":params or {}}
    r = s.post(url, json=payload, timeout=25)
    if r.status_code >= 400:
        raise RuntimeError(f"{method}: HTTP {r.status_code}: {r.text[:800]}")
    try:
        data = r.json()
    except ValueError:
        raise RuntimeError(f"{method}: Ungültige JSON-Antwort: {r.text[:800]}")

    # Fallbacks auf ältere Methodennamen
    if "error" in data and isinstance(data["error"], dict) and data["error"].get("code") == -32601:
        # Method not found → auf Legacy ausweichen
        if method == "getUserData":
            return untis_rpc(s, url, "getUserData2017", params)
        if method == "getTimetable":
            return untis_rpc(s, url, "getTimetable2017", params)

    if "result" not in data:
        raise RuntimeError(f"{method} fehlgeschlagen: {data}")
    return data["result"]

def untis_login(s: requests.Session):
    base = f"https://{UNTIS_SERVER}/WebUntis"
    url  = f"{base}/jsonrpc.do?school={UNTIS_SCHOOL}"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "untis-icloud-sync/1.0",
        "Origin": base,
        "Referer": f"{base}/?school={UNTIS_SCHOOL}#/basic/login"
    }
    s.headers.update(headers)
    print(f"[DEBUG] using server='{UNTIS_SERVER}', school='{UNTIS_SCHOOL}'")
    print(f"[INFO] Login bei WebUntis …")
    # authenticate
    result = untis_rpc(s, url, "authenticate",
                       {"user": UNTIS_USER, "password": UNTIS_PASS, "client": "untis-icloud-sync"})
    sid = result["sessionId"]
    # Cookies korrekt setzen
    s.cookies.set("JSESSIONID", sid, domain=UNTIS_SERVER, path="/")
    s.cookies.set("schoolname", UNTIS_SCHOOL, domain=UNTIS_SERVER, path="/")
    return url

def untis_user(s, url):
    return untis_rpc(s, url, "getUserData", {})

def fetch_tt(s, url, pid, ptype, st, en):
    params = {"options":{"element":{"id":pid,"type":ptype},
                         "startDate":st,"endDate":en,"showStudentgroup": True}}
    return untis_rpc(s, url, "getTimetable", params)

def connect_caldav():
    c = DAVClient(url="https://caldav.icloud.com/", username=ICLOUD_USER, password=ICLOUD_PASS)
    p = c.principal()
    for cal in p.calendars():
        try:
            if cal.name == ICLOUD_CAL: return cal
        except Exception:
            pass
    try:
        return p.make_calendar(name=ICLOUD_CAL)
    except Exception:
        return p.calendars()[0]

def build_event(lesson):
    start=to_local(lesson["date"],lesson["startTime"])
    end=to_local(lesson["date"],lesson["endTime"])
    subject=labels(lesson,"su") or "Unterricht"
    teachers,rooms,klasse=labels(lesson,"te"),labels(lesson,"ro"),labels(lesson,"kl")
    desc="\n".join([f"Lehrer: {teachers}" if teachers else "",
                    f"Raum: {rooms}" if rooms else "",
                    f"Klasse: {klasse}" if klasse else ""]).strip()
    cancelled=(lesson.get("code")=="cancelled") or bool(lesson.get("cancelled"))

    cal=Calendar(); cal.add('prodid','-//Untis iCloud Sync//DE'); cal.add('version','2.0')
    ev=Event()
    ev.add('uid', uid_for(lesson))
    ev.add('summary', vText(subject))
    ev.add('dtstart', start); ev.add('dtend', end)
    ev.add('dtstamp', dt.datetime.utcnow())
    ev.add('location', vText(rooms or ""))
    ev.add('description', vText(desc))
    if cancelled and MARK_CANCELLED: ev.add('status','CANCELLED')
    cal.add_component(ev)
    return cal.to_ical().decode()

def existing_by_uid(cal, start, end):
    result={}
    try: objs=cal.date_search(start, end)
    except caldav_error.ReportError: objs=[]
    for o in objs:
        try: result[str(o.vobject_instance.vevent.uid.value)] = o
        except Exception: pass
    return result

def main():
    for n in ["UNTIS_SERVER","UNTIS_SCHOOL","UNTIS_USER","UNTIS_PASS","ICLOUD_USER","ICLOUD_PASS"]:
        require_env(n)

    s = requests.Session()
    s.headers.update({"Content-Type":"application/json"})

    url = untis_login(s)
    ud = untis_user(s, url)
    pid, ptype = ud["personId"], ud["personType"]

    today=dt.date.today()
    start=today-dt.timedelta(days=DAYS_PAST)
    end  =today+dt.timedelta(days=DAYS_AHEAD)

    print(f"[INFO] Stundenplan {start} bis {end} abrufen …")
    tt = fetch_tt(s, url, pid, ptype, ymd(start), ymd(end))
    print(f"[INFO] {len(tt)} Einträge erhalten.")

    wanted={uid_for(x):x for x in tt}

    cal=connect_caldav()
    start_dt = TZ.localize(dt.datetime.combine(start, dt.time.min))
    end_dt   = TZ.localize(dt.datetime.combine(end, dt.time.max))
    existing = existing_by_uid(cal, start_dt, end_dt)

    created=updated=deleted=0
    for uid, lesson in wanted.items():
        ics = build_event(lesson)
        if uid in existing:
            old = existing[uid].data.decode() if isinstance(existing[uid].data,bytes) else existing[uid].data
            if ics != old:
                existing[uid].data = ics
                existing[uid].save(); updated+=1
        else:
            cal.add_event(ics); created+=1

    if DELETE_MISSING:
        for uid,obj in list(existing.items()):
            if uid.startswith("untis-") and uid not in wanted:
                try: obj.delete(); deleted+=1
                except Exception: pass

    print(f"[ERGEBNIS] neu={created} geändert={updated} gelöscht={deleted} (Kalender: {ICLOUD_CAL})")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("[FEHLER]", e)
        sys.exit(1)
