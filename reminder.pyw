"""Work To-Do reminders: a small background helper with no window.

It starts with Windows (a shortcut in your Startup folder), reads todo-data.json every
30 seconds and shows Windows notifications for:
  * a morning summary of tasks due today and overdue tasks (once a day, at the time set in the app),
  * a heads-up shortly before a task's due time, and a "due now" alert at the due time.
It never changes your tasks. It remembers what it already showed in reminder-state.json.
Turn reminders off in the app (Settings), or delete "Work To-Do Reminders" from shell:startup.
Run with --test to show one test notification and exit. No installs: pure Python + PowerShell."""
import json, os, sys, time, math, threading, subprocess, datetime as dt
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from xml.sax.saxutils import escape, quoteattr

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.environ.get("TODO_DATA") or os.path.join(APP_DIR, "todo-data.json")
STATE_FILE = os.environ.get("TODO_REMINDER_STATE") or os.path.join(APP_DIR, "reminder-state.json")
LOG_FILE = os.path.join(APP_DIR, "reminder.log")
PORT = int(os.environ.get("TODO_REMINDER_PORT", "48232"))
OPEN_URL = f"http://127.0.0.1:{PORT}/open"
ICON_PNG = os.path.join(APP_DIR, "toast-icon.png")
AUMID = r"{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
CHECK_EVERY = 30
DEFAULTS = {"reminders": True, "morning": "09:00", "lead": 15}
VIEWS = ("today", "upcoming", "overdue", "notes")
RANK = {"urgent": 0, "high": 1, "normal": 2}
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
_status = {"started": dt.datetime.now().isoformat(timespec="seconds"), "last_check": None, "last_toast": None, "enabled": None}
_last_open = [0.0]


def log(msg):
    try:
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > 256 * 1024:
            os.replace(LOG_FILE, LOG_FILE + ".1")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except OSError:
        pass


def read_data():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else None
    except (OSError, ValueError):
        return None


def parse_hm(s):
    try:
        h, m = (int(x) for x in str(s).split(":")[:2])
        return dt.time(h, m)
    except (ValueError, TypeError):
        return None


def get_settings(data):
    s = dict(DEFAULTS)
    if isinstance(data.get("settings"), dict):
        s.update(data["settings"])
    s["reminders"] = s.get("reminders") is not False
    if not parse_hm(s.get("morning")):
        s["morning"] = DEFAULTS["morning"]
    try:
        s["lead"] = max(0, min(240, int(s.get("lead"))))
    except (TypeError, ValueError):
        s["lead"] = DEFAULTS["lead"]
    return s


def fmt_time(t):
    return t.strftime("%I:%M %p").lstrip("0")


def pythonw():
    w = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return w if os.path.exists(w) else sys.executable


def launch_app(view=""):
    args = [pythonw(), os.path.join(APP_DIR, "todo.pyw")]
    if view in VIEWS:
        args.append(f"--view={view}")
    subprocess.Popen(args, cwd=APP_DIR, creationflags=NO_WINDOW)


# ---------- notifications ----------
PS_SCRIPT = (
    "$ErrorActionPreference='Stop';"
    "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
    "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null;"
    "$d = New-Object Windows.Data.Xml.Dom.XmlDocument; $d.LoadXml($env:WT_TOAST_XML);"
    "$t = New-Object Windows.UI.Notifications.ToastNotification $d; $t.Tag = $env:WT_TOAST_TAG; $t.Group = 'worktodo';"
    "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($env:WT_TOAST_AUMID).Show($t)"
)


def build_xml(title, lines, view="today", sticky=False, snooze=False):
    launch = f"{OPEN_URL}?view={view}"
    img = ""
    if os.path.exists(ICON_PNG):
        img = f'<image placement="appLogoOverride" hint-crop="circle" src={quoteattr("file:///" + ICON_PNG.replace(os.sep, "/"))}/>'
    body = "".join(f"<text>{escape(l)}</text>" for l in [ "\n".join(lines) ] if l)
    actions = ""
    if snooze:
        actions += ('<input id="snooze" type="selection" defaultInput="10">'
                    '<selection id="5" content="5 minutes"/><selection id="10" content="10 minutes"/>'
                    '<selection id="30" content="30 minutes"/><selection id="60" content="1 hour"/></input>'
                    '<action activationType="system" arguments="snooze" hint-inputId="snooze" content=""/>')
    actions += f'<action content="Open Work To-Do" activationType="protocol" arguments={quoteattr(launch)}/>'
    actions += '<action activationType="system" arguments="dismiss" content=""/>'
    scen = ' scenario="reminder"' if sticky else ""
    return (f'<toast launch={quoteattr(launch)} activationType="protocol"{scen}>'
            f'<visual><binding template="ToastGeneric">{img}<text>{escape(title)}</text>{body}'
            f'<text placement="attribution">Work To-Do · Ctrl+Alt+T opens the app</text></binding></visual>'
            f'<actions>{actions}</actions><audio src="ms-winsoundevent:Notification.Reminder"/></toast>')


def fallback_box(title, text, wait=False):
    def run():
        try:
            import ctypes
            r = ctypes.windll.user32.MessageBoxW(0, text + "\n\nOpen Work To-Do now?", "Work To-Do: " + title, 0x4 | 0x40 | 0x1000 | 0x40000)
            if r == 6:
                launch_app("today")
        except Exception as e:
            log(f"fallback failed: {e}")
    th = threading.Thread(target=run, daemon=not wait)
    th.start()
    if wait:
        th.join()


def show_toast(title, lines, view="today", tag="wt", sticky=False, snooze=False, wait_fallback=False):
    xml = build_xml(title, lines, view, sticky, snooze)
    env = dict(os.environ, WT_TOAST_XML=xml, WT_TOAST_TAG=tag[:60], WT_TOAST_AUMID=AUMID)
    ok, err = False, ""
    try:
        r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS_SCRIPT],
                           env=env, capture_output=True, text=True, timeout=45, creationflags=NO_WINDOW)
        ok, err = r.returncode == 0, (r.stderr or "").strip()[:300]
    except Exception as e:
        err = str(e)[:300]
    _status["last_toast"] = dt.datetime.now().isoformat(timespec="seconds")
    log(f"toast {'ok' if ok else 'FAILED'}: {title}" + ("" if ok else f" | {err}"))
    if not ok:
        fallback_box(title, "\n".join(lines), wait=wait_fallback)
    return ok


# ---------- the check ----------
def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            st = json.load(f)
        if isinstance(st, dict) and isinstance(st.get("sent"), dict):
            return st
    except (OSError, ValueError):
        pass
    return {"sent": {}}


def save_state(st):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, indent=1)
    os.replace(tmp, STATE_FILE)


def check(now=None, notify=show_toast):
    data = read_data()
    if data is None:
        return
    s = get_settings(data)
    now = now or dt.datetime.now()
    _status["last_check"] = now.isoformat(timespec="seconds")
    _status["enabled"] = s["reminders"]
    if not s["reminders"]:
        return
    today = now.date().isoformat()
    st = load_state()
    sent = st["sent"]
    cutoff = (now - dt.timedelta(days=10)).isoformat()
    changed = False
    for k in [k for k, v in sent.items() if str(v) < cutoff]:
        del sent[k]
        changed = True
    stamp = now.isoformat(timespec="seconds")
    projects = {p.get("id"): p.get("name", "") for p in data.get("projects", []) if isinstance(p, dict)}
    tasks = [t for t in data.get("tasks", []) if isinstance(t, dict) and not t.get("done") and isinstance(t.get("due"), str) and t.get("due")]

    def first_line(t, n=70):
        x = (t.get("title") or "").strip().split("\n")[0]
        return x if len(x) <= n else x[: n - 1] + "…"

    def detail(t):
        bits = [(t.get("priority") or "normal").capitalize()]
        if projects.get(t.get("project_id")):
            bits.append(projects[t["project_id"]])
        return bits

    # 1) morning summary (once a day, at the set time or the first check after it)
    key = f"digest:{today}"
    if now.time() >= parse_hm(s["morning"]) and key not in sent:
        due_today = [t for t in tasks if t["due"] == today]
        overdue = [t for t in tasks if t["due"] < today]
        if due_today or overdue:
            parts = []
            if due_today:
                parts.append(f"{len(due_today)} task{'s' if len(due_today) != 1 else ''} due today")
            if overdue:
                parts.append(f"{len(overdue)} overdue")
            pick = sorted(overdue, key=lambda t: (RANK.get(t.get("priority"), 2), t["due"])) + \
                sorted(due_today, key=lambda t: (RANK.get(t.get("priority"), 2), t.get("due_time") or "99:99"))
            lines = []
            for t in pick[:3]:
                tag = "Overdue" if t["due"] < today else (fmt_time(parse_hm(t["due_time"])) if parse_hm(t.get("due_time")) else "Today")
                lines.append(f"• {first_line(t, 48)}  ({tag} · {' · '.join(detail(t))})")
            if len(pick) > 3:
                lines.append(f"+ {len(pick) - 3} more")
            notify(" and ".join(parts).capitalize(), lines, view="today", tag="digest")
        sent[key] = stamp
        changed = True

    # 2) due-time alerts for today's tasks that have a time
    for t in tasks:
        if t["due"] != today:
            continue
        tm = parse_hm(t.get("due_time"))
        if not tm:
            continue
        due_at = dt.datetime.combine(now.date(), tm)
        base = f"{t.get('id')}:{today}:{t.get('due_time')}"
        tag = "t" + str(t.get("id", ""))[:20]
        if now >= due_at:
            k = "due:" + base
            if k not in sent:
                late = (now - due_at).total_seconds()
                if late <= 12 * 3600:
                    head = "Due now" if late < 600 else f"Was due at {fmt_time(tm)}"
                    notify(f"{head}: {first_line(t)}", [" · ".join([f"Due {fmt_time(tm)}"] + detail(t))],
                           view="today", tag=tag, sticky=True, snooze=True)
                sent[k] = stamp
                sent.setdefault("soon:" + base, stamp)
                changed = True
        elif s["lead"] and now >= due_at - dt.timedelta(minutes=s["lead"]):
            k = "soon:" + base
            if k not in sent:
                mins = max(1, math.ceil((due_at - now).total_seconds() / 60))
                notify(f"Due in {mins} min: {first_line(t)}", [" · ".join([f"Due {fmt_time(tm)}"] + detail(t))],
                       view="today", tag=tag, sticky=True, snooze=True)
                sent[k] = stamp
                changed = True
    if changed:
        save_state(st)


# ---------- tiny local endpoint: lets a notification click open the app ----------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body=b"", ctype="application/json"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            self.wfile.write(body)

    def do_GET(self):
        if self.headers.get("Host", "") not in (f"127.0.0.1:{PORT}", f"localhost:{PORT}"):
            return self._send(403)
        u = urlparse(self.path)
        if u.path == "/status":
            return self._send(200, json.dumps(dict(_status, running=True, pid=os.getpid())).encode())
        if u.path == "/open":
            view = (parse_qs(u.query).get("view") or [""])[0]
            if time.time() - _last_open[0] > 3:
                _last_open[0] = time.time()
                try:
                    launch_app(view)
                except Exception as e:
                    log(f"open failed: {e}")
            page = ("<!doctype html><meta charset=utf-8><title>Work To-Do</title>"
                    "<body style='background:#0d1017;color:#e8ebf2;font:15px Segoe UI,sans-serif;display:grid;place-items:center;height:90vh'>"
                    "<div style='text-align:center'><div style='font-size:34px'>&#10003;</div><b>Opening Work To-Do…</b>"
                    "<p style='color:#8a93a8'>It opens in its own window. You can close this tab.</p></div>"
                    "<script>setTimeout(()=>window.close(),1200)</script>")
            return self._send(200, page.encode("utf-8"), "text/html; charset=utf-8")
        self._send(404)


def main():
    if "--test" in sys.argv:
        ok = show_toast("Test reminder from Work To-Do",
                        ["Reminders are working. Click here or \"Open Work To-Do\" to jump to the app."],
                        view="today", tag="test", wait_fallback=True)
        sys.exit(0 if ok else 1)
    try:
        server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    except OSError:
        return  # already running
    threading.Thread(target=server.serve_forever, daemon=True).start()
    log(f"started (pid {os.getpid()})")
    while True:
        try:
            check()
        except Exception as e:
            log(f"check error: {e!r}")
        time.sleep(CHECK_EVERY)


if __name__ == "__main__":
    main()
