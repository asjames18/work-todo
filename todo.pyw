"""Work To-Do: a local project to-do app with a modern look.
Runs a tiny private server on this PC (127.0.0.1 only) and opens the app in its own
Microsoft Edge window. Data (projects, tasks and notes) saves to todo-data.json next to this file. No installs.
AI suggestions (todo_ai.py) are built in but OFF until turned on in Reminders & settings.
Dashboard (v7, the landing view): today at a glance, with a built-in or AI brief and recent Live Transcriber meeting notes (read-only).
Email (todo_email.py + outlook_mail.ps1): a read-only inbox from classic Outlook, plus drag/paste of emails. Nothing is
added until you click Add, and no mail is ever sent, moved, deleted or marked read.
v8 (shareable): first-run setup wizard, "Get set up" checklist, light/dark themes + accents, global quick add (Ctrl+Alt+N,
owned by reminder.pyw), recurring tasks, focus timer with time logs, weekly review, backup/restore, shortcuts tour and
optional updates from GitHub Releases (wtd_update.py). Code and data locations come from wtd_paths.py."""
import json, os, re, sys, time, threading, subprocess, webbrowser, urllib.request, tempfile
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)
import wtd_paths as P
DATA_DIR = P.DATA_DIR
DATA_FILE = os.environ.get("TODO_DATA") or os.path.join(DATA_DIR, "todo-data.json")
VERSION = P.VERSION
PORT = int(os.environ.get("TODO_PORT", "48231"))
NO_WINDOW = os.environ.get("TODO_NO_WINDOW") == "1"
URL = f"http://127.0.0.1:{PORT}/"
REMINDER_PORT = int(os.environ.get("TODO_REMINDER_PORT", "48232"))
VIEWS = ("dash", "today", "upcoming", "overdue", "notes", "shortcuts", "settings", "ai", "email", "review", "setup")
_view = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--view=")), "")
APP_URL = URL + ("#" + _view if _view in VIEWS else "")
NO_WINDOW_FLAG = getattr(subprocess, "CREATE_NO_WINDOW", 0)
IDLE_LIMIT = 150  # seconds without a ping from the window before the server quits
_last_ping = [time.time()]
_data_lock = threading.RLock()   # PUT and AI apply both read-modify-write the data file
_voice_lock = threading.Lock()   # one voice clip at a time
try:
    import wtd_update
except Exception:
    wtd_update = None
try:
    import todo_ai
except Exception as _e:  # the app still works without the AI module
    todo_ai = None
try:
    import todo_email
except Exception as _e:  # ...and without the email module
    todo_email = None
_email_lock = threading.Lock()   # one Outlook read at a time
_email_job = {"running": False, "started": 0}
MAX_UPLOAD = 25 * 1024 * 1024
TX_DIR = os.environ.get("WORKTODO_TX_DIR") or os.path.join(os.path.dirname(APP_DIR), "transcriber")
TX_PY = os.path.join(TX_DIR, ".venv", "Scripts", "python.exe")
VOICE_HELPER = os.path.join(APP_DIR, "voice_helper.py")
MAX_AUDIO = 20 * 1024 * 1024


def load_data():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("not an object")
    except FileNotFoundError:
        data = {}
    except (ValueError, OSError):
        try:
            os.replace(DATA_FILE, DATA_FILE + ".bak")
        except OSError:
            pass
        data = {}
    data.setdefault("projects", [])
    data.setdefault("tasks", [])
    data.setdefault("notes", [])
    return data


def save_data(data):
    P.ensure_data_dir()
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, DATA_FILE)


def data_rev():
    try:
        return str(os.stat(DATA_FILE).st_mtime_ns)
    except OSError:
        return "0"


def migrate_data():
    """v8: existing data counts as already set up, so nobody who upgrades is sent through the setup wizard."""
    with _data_lock:
        data = load_data()
        s = data.get("settings") if isinstance(data.get("settings"), dict) else {}
        if isinstance(s.get("profile"), dict) and s["profile"].get("setup_done"):
            return False
        if not (data["projects"] or data["tasks"] or data["notes"] or s):
            return False   # brand-new: the wizard runs
        em = s.get("email") if isinstance(s.get("email"), dict) else {}
        nm = str(em.get("name") or "").strip()
        if "," in nm:
            nm = nm.split(",", 1)[1].strip()
        s = dict(s)
        s["profile"] = {"name": nm.split(" ")[0] if nm else "", "setup_done": True, "migrated": time.strftime("%Y-%m-%dT%H:%M:%S")}
        ob = dict(s.get("onboarding") or {})
        ob.setdefault("dismissed", True)    # the "Get set up" checklist is for new people; re-open it in Settings
        ob.setdefault("tour_seen", True)
        s["onboarding"] = ob
        s.setdefault("updates", P.IS_INSTALLED)
        data["settings"] = s
        save_data(data)
        return True


# ---------------------------------------------------------------- recurring tasks (same rules as the window's nextDue)
def next_due(rep, base):
    import datetime as _dt
    f = (rep or {}).get("freq")
    if f == "daily":
        return base + _dt.timedelta(days=1)
    if f == "weekdays":
        d = base + _dt.timedelta(days=1)
        while d.weekday() >= 5:
            d += _dt.timedelta(days=1)
        return d
    if f == "weekly":
        dow = int(rep.get("dow", base.isoweekday() % 7))   # 0 = Sunday ... 6 = Saturday (JS getDay)
        d = base + _dt.timedelta(days=1)
        while d.isoweekday() % 7 != dow:
            d += _dt.timedelta(days=1)
        return d
    if f == "monthly":
        import calendar
        dom = max(1, min(31, int(rep.get("dom", base.day))))
        y, m = (base.year + (base.month // 12), base.month % 12 + 1)
        return _dt.date(y, m, min(dom, calendar.monthrange(y, m)[1]))
    return None


def spawn_repeats(before, after):
    """After an AI apply: any repeating task that just became done gets its next occurrence (the window does this itself)."""
    import datetime as _dt
    was = {t.get("id"): t.get("done") for t in before.get("tasks", [])}
    today = _dt.date.today()
    for t in list(after.get("tasks", [])):
        if t.get("done") and not was.get(t.get("id")) and isinstance(t.get("repeat"), dict) and not t.get("repeat_next"):
            try:
                base = max(_dt.date.fromisoformat(t["due"]), today) if t.get("due") else today
            except ValueError:
                base = today
            nd = next_due(t["repeat"], base)
            if nd:
                n = dict(t, id=os.urandom(16).hex(), done=False, completed="", due=nd.isoformat(), created=time.strftime("%Y-%m-%dT%H:%M:%S"), focus_min=0)
                n.pop("repeat_next", None)
                t["repeat_next"] = n["id"]
                after["tasks"].append(n)


def clean_settings(settings):
    out = dict(settings)
    ai = out.get("ai")
    if ai is not None:  # only the known AI settings are kept, so an API key can never end up in the file
        if not isinstance(ai, dict):
            raise ValueError("bad ai settings")
        allowed = todo_ai.DEFAULT_AI if todo_ai else {"enabled": 0, "provider": 0, "base_url": 0, "model": 0, "api_key_env": 0, "voice_model": 0}
        out["ai"] = {k: v for k, v in ai.items() if k in allowed}
    em = out.get("email")
    if em is not None and todo_email is not None:
        out["email"] = todo_email.clean_settings(em)
    if "theme" in out and out["theme"] not in ("dark", "light"):
        out["theme"] = "dark"
    return out


def valid_shape(data):
    return isinstance(data, dict) and isinstance(data.get("projects"), list) and isinstance(data.get("tasks"), list)


def backup_file(tag):
    os.makedirs(P.BACKUP_DIR, exist_ok=True)
    dst = os.path.join(P.BACKUP_DIR, "todo-data." + time.strftime("%Y%m%d-%H%M%S") + "." + tag + ".json")
    with _data_lock:
        data = load_data()
        with open(dst, "w", encoding="utf-8") as f:
            json.dump(export_payload(data), f, indent=2)
    return dst


def export_payload(data):
    return {"app": "Work To-Do", "format": 1, "version": VERSION, "exported": time.strftime("%Y-%m-%dT%H:%M:%S"), "data": data}


def restore_data(body):
    payload = body.get("backup")
    if isinstance(payload, dict) and isinstance(payload.get("data"), dict):
        payload = payload["data"]
    if not valid_shape(payload):
        raise ValueError("that file isn't a Work To-Do backup")
    for k in ("notes", "focus_log"):
        if k in payload and not isinstance(payload[k], list):
            raise ValueError("bad " + k)
    saved = backup_file("pre-restore")   # always keep what was there first
    with _data_lock:
        cur = load_data()
        new = dict(payload)
        new.setdefault("notes", [])
        st = new.get("settings") if isinstance(new.get("settings"), dict) else {}
        st = clean_settings(st)
        prof = dict(st.get("profile") or {})
        prof["setup_done"] = True            # a restored copy is set up
        if not prof.get("name"):
            prof["name"] = ((cur.get("settings") or {}).get("profile") or {}).get("name", "")
        st["profile"] = prof
        new["settings"] = st
        save_data(new)
    return {"data": load_data(), "backup": saved}


def save_ai_key(key):
    """Saves the key ONLY in the user's own Windows environment (HKCU Environment, OPENROUTER_API_KEY), never in app files.
    Also sets it for this running app so AI works right away. Returns where it went; never returns the key."""
    key = str(key or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_\-.:]{20,200}", key):
        raise ValueError("that doesn't look like an API key")
    os.environ["OPENROUTER_API_KEY"] = key
    if os.name != "nt":
        return {"saved": "this session only (not Windows)"}
    import winreg, ctypes
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, "OPENROUTER_API_KEY", 0, winreg.REG_SZ, key)
    try:  # tell Windows (Explorer, new apps) the environment changed
        res = ctypes.c_ulong()
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 3000, ctypes.byref(res))
    except Exception:
        pass
    return {"saved": "your Windows user environment variable OPENROUTER_API_KEY"}


def quick_add(body):
    """A task from the global quick-add box (quickadd.pyw) while the app is running."""
    title = str(body.get("title") or "").strip()[:2000]
    if not title:
        raise ValueError("empty task")
    with _data_lock:
        data = load_data()
        pid = body.get("project_id")
        if not any(p.get("id") == pid for p in data["projects"]):
            raise ValueError("unknown project")
        pr = body.get("priority") if body.get("priority") in ("urgent", "high", "normal") else "normal"
        due = str(body.get("due") or "")
        due = due if re.fullmatch(r"\d{4}-\d{2}-\d{2}", due) else ""
        tm = str(body.get("due_time") or "")
        tm = tm if (due and re.fullmatch(r"\d{2}:\d{2}", tm)) else ""
        t = {"id": str(body.get("id") or os.urandom(16).hex())[:40], "project_id": pid, "title": title, "priority": pr, "due": due,
             "due_time": tm, "done": False, "created": time.strftime("%Y-%m-%dT%H:%M:%S"), "completed": "", "source": "quick-add"}
        data["tasks"].append(t)
        save_data(data)
    return {"task": t}


# ---------------------------------------------------------------- updates (GitHub Releases, wtd_update.py)
_update = {"last": None, "busy": False, "applied": None}


def updates_enabled(data=None):
    s = (data or load_data()).get("settings") or {}
    v = s.get("updates")
    return P.IS_INSTALLED if v is None else v is True


def update_status():
    st = wtd_update.load_state() if wtd_update else {}
    last = _update["last"] or st.get("last")
    return {"version": VERSION, "repo": P.UPDATE_REPO, "installed": P.IS_INSTALLED, "enabled": updates_enabled(),
            "can_apply": P.IS_INSTALLED and wtd_update is not None, "busy": _update["busy"], "last": last,
            "available": bool(last and last.get("available") and P.is_newer(last.get("tag", ""))), "data_dir": DATA_DIR}


def auto_update_check():
    """On launch: at most once a day, only when update checks are on. Quiet on 404/offline."""
    try:
        if wtd_update and updates_enabled() and wtd_update.due(wtd_update.load_state()):
            _update["last"] = wtd_update.check_and_store()
    except Exception:
        pass


def update_log(msg):
    try:
        os.makedirs(P.BACKUP_DIR, exist_ok=True)
        with open(os.path.join(DATA_DIR, "update.log"), "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
    except OSError:
        pass


def apply_update(server_ref):
    if not wtd_update:
        raise ValueError("the updater didn't load")
    info = _update["last"] or (wtd_update.load_state().get("last"))
    if not info or not info.get("available"):
        info = wtd_update.check_and_store()
        _update["last"] = info
    if _update["busy"]:
        raise ValueError("an update is already running")
    _update["busy"] = True
    try:
        res = wtd_update.apply(info, log=update_log)
    finally:
        _update["busy"] = False
    _update["applied"] = res
    # restart: stop the reminder helper, start a fresh server (it waits for this port) and a fresh helper, then quit
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{REMINDER_PORT}/quit", data=b"", headers={"X-WorkTodo": "1"}, method="POST")
        urllib.request.urlopen(req, timeout=2).close()
        had_helper = True
    except Exception:
        had_helper = False
    subprocess.Popen([pythonw(), os.path.join(APP_DIR, "todo.pyw"), "--after-update"], cwd=APP_DIR, creationflags=NO_WINDOW_FLAG)
    if had_helper:
        threading.Timer(1.5, start_reminder).start()
    threading.Timer(1.0, lambda: server_ref and server_ref.shutdown()).start()
    return res


_server = [None]


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _host_ok(self):
        return self.headers.get("Host", "") in (f"127.0.0.1:{PORT}", f"localhost:{PORT}")

    def _send(self, code, body=b"", ctype="application/json", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            self.wfile.write(body)

    def do_GET(self):
        if not self._host_ok():
            return self._send(403)
        _last_ping[0] = time.time()
        if self.path in ("/", "/index.html"):
            return self._send(200, HTML.replace("__WTD_VERSION__", VERSION).encode("utf-8"), "text/html; charset=utf-8")
        if self.path == "/api/rev":   # lets an open window notice outside changes (quick add, restore) and reload
            return self._send(200, json.dumps({"rev": data_rev()}).encode("utf-8"))
        if self.path == "/api/export":
            body = json.dumps(export_payload(load_data()), indent=2).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Disposition", 'attachment; filename="WorkToDo-backup-' + time.strftime("%Y-%m-%d") + '.json"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return self.wfile.write(body)
        if self.path == "/api/update/status":
            return self._send(200, json.dumps(update_status()).encode("utf-8"))
        if self.path == "/api/data":
            return self._send(200, json.dumps(load_data()).encode("utf-8"))
        if self.path == "/favicon.ico":
            try:
                with open(os.path.join(APP_DIR, "todo.ico"), "rb") as f:
                    return self._send(200, f.read(), "image/x-icon")
            except OSError:
                return self._send(404)
        if self.path == "/api/ping":
            return self._send(204)
        if self.path == "/api/reminder-status":
            return self._send(200, json.dumps(reminder_status()).encode("utf-8"))
        if self.path == "/api/ai/status":  # local only: reads settings and checks files, no network
            return self._send(200, json.dumps(ai_status()).encode("utf-8"))
        if self.path == "/api/email/inbox":
            return self._send(200, json.dumps(email_inbox()).encode("utf-8"))
        if self.path == "/api/meetings":  # read-only: Live Transcriber meeting notes
            return self._send(200, json.dumps(meetings_list()).encode("utf-8"))
        self._send(404)

    def _json_body(self):
        n = int(self.headers.get("Content-Length", "0"))
        if n > 2 * 1024 * 1024:
            raise ValueError("request too large")
        raw = self.rfile.read(n) if n else b"{}"
        body = json.loads(raw.decode("utf-8") or "{}")
        if not isinstance(body, dict):
            raise ValueError("bad request")
        return body

    def _big_json(self):
        n = int(self.headers.get("Content-Length", "0"))
        if n <= 0 or n > 50 * 1024 * 1024:
            raise ValueError("that backup is empty or too large")
        body = json.loads(self.rfile.read(n).decode("utf-8"))
        if not isinstance(body, dict):
            raise ValueError("bad request")
        return body

    def _ai(self):
        """Every /api/ai/* POST. Nothing runs (no provider, no network, no Whisper) unless settings.ai.enabled is true."""
        _last_ping[0] = time.time()
        if todo_ai is None:
            return self._send(500, json.dumps({"error": "the AI module didn't load"}).encode("utf-8"))
        route = self.path[len("/api/ai/"):]
        data = load_data()
        if not todo_ai.is_enabled(data):
            return self._send(409, json.dumps({"error": "AI suggestions are off. Turn them on in Reminders & settings."}).encode("utf-8"))
        try:
            if route == "transcribe":
                return self._send(200, json.dumps(transcribe_clip(self, data)).encode("utf-8"))
            body = self._json_body()
            if route == "recommend":
                out = todo_ai.recommend(data)
            elif route == "parse":
                out = todo_ai.parse_task(data, body.get("text", ""), body.get("project_hint"))
            elif route == "preview":
                acts = todo_ai.validate_actions(body.get("actions"), data)
                out = {"actions": acts, "changes": todo_ai.describe_all(acts, data)}
            elif route == "apply":
                with _data_lock:
                    data = load_data()
                    new, changes, undo = todo_ai.apply_actions(data, body.get("actions"))
                    spawn_repeats(data, new)
                    save_data(new)
                out = {"data": new, "changes": changes, "undo": undo}
            elif route == "test":
                out = todo_ai.get_provider(todo_ai.ai_settings(data)).test() if todo_ai.ai_settings(data)["provider"] == "openai_compatible" else {"models": 0, "rules": True}
            else:
                return self._send(404)
            return self._send(200, json.dumps(out).encode("utf-8"))
        except todo_ai.AIError as e:
            return self._send(400, json.dumps({"error": str(e)}).encode("utf-8"))
        except (ValueError, TypeError) as e:
            return self._send(400, json.dumps({"error": "bad request"}).encode("utf-8"))

    def do_POST(self):
        # custom header = browsers must preflight, so other websites can't trigger these
        if not self._host_ok() or self.headers.get("X-WorkTodo") != "1":
            return self._send(403)
        if self.path == "/api/reminder-test":
            start_reminder("--test")
            return self._send(202)
        if self.path == "/api/reminder-start":
            if not reminder_status().get("running"):
                start_reminder()
            return self._send(202)
        if self.path.startswith("/api/ai/"):
            return self._ai()
        if self.path.startswith("/api/email/"):
            return self._email()
        if self.path in ("/api/quick-add", "/api/restore", "/api/backup-now", "/api/setup/ai-key", "/api/review",
                         "/api/update/check", "/api/update/apply"):
            _last_ping[0] = time.time()
            try:
                body = self._json_body() if self.path != "/api/restore" else self._big_json()
                if self.path == "/api/quick-add":
                    out = quick_add(body)
                elif self.path == "/api/restore":
                    out = restore_data(body)
                elif self.path == "/api/backup-now":
                    out = {"path": backup_file("backup")}
                elif self.path == "/api/setup/ai-key":
                    out = save_ai_key(body.get("key"))
                elif self.path == "/api/review":
                    out = make_review(load_data(), body.get("start"), body.get("end"), bool(body.get("ai")))
                elif self.path == "/api/update/check":
                    if not wtd_update:
                        raise ValueError("the updater didn't load")
                    _update["last"] = wtd_update.check_and_store()
                    out = update_status()
                else:
                    out = apply_update(_server[0])
                headers = {"X-Rev": data_rev()}
                return self._send(200, json.dumps(out).encode("utf-8"), extra=headers)
            except Exception as e:
                msg = str(e) if isinstance(e, ValueError) or (wtd_update and isinstance(e, wtd_update.UpdateError)) else type(e).__name__
                return self._send(400, json.dumps({"error": msg[:300]}).encode("utf-8"))
        if self.path == "/api/brief":  # dashboard brief: built-in always; AI only if asked and AI is on
            _last_ping[0] = time.time()
            try:
                body = self._json_body()
                return self._send(200, json.dumps(make_brief(load_data(), bool(body.get("ai")))).encode("utf-8"))
            except Exception as e:
                return self._send(400, json.dumps({"error": "couldn't make the brief (" + type(e).__name__ + ")"}).encode("utf-8"))
        self._send(404)

    def _email(self):
        """Every /api/email/* POST. Outlook is only ever read; tasks are only added on /add (the confirm card's Add)."""
        _last_ping[0] = time.time()
        if todo_email is None:
            return self._send(500, json.dumps({"error": "the email module didn't load"}).encode("utf-8"))
        route = self.path[len("/api/email/"):]
        try:
            if route == "import":  # a dropped .eml / .msg / .txt file (raw bytes)
                return self._send(200, json.dumps(email_import_file(self)).encode("utf-8"))
            body = self._json_body()
            if route == "refresh":
                out = email_refresh(force=bool(body.get("force")))
            elif route == "suggest":
                out = email_suggest(body.get("id"), body.get("project_hint"))
            elif route == "import-text":
                rec = todo_email.parse_text(body.get("text", ""))
                out = dict(email_task_for(rec, body.get("project_hint")), email=todo_email.public_item(rec))
            elif route == "add":
                out = email_add(body)
            elif route == "state":
                with _data_lock:
                    data = load_data()
                    st = todo_email.set_state(data, body.get("id"), body.get("op"))
                    save_data(data)
                out = {"ok": True}
            elif route == "open":
                out = email_open(body.get("id"))
            elif route == "setup-outlook":
                out = email_setup_outlook()
            else:
                return self._send(404)
            return self._send(200, json.dumps(out).encode("utf-8"))
        except todo_email.EmailError as e:
            return self._send(400, json.dumps({"error": str(e)}).encode("utf-8"))
        except Exception as e:
            if todo_ai is not None and isinstance(e, todo_ai.AIError):
                return self._send(400, json.dumps({"error": str(e)}).encode("utf-8"))
            return self._send(400, json.dumps({"error": "bad request"}).encode("utf-8"))

    def do_PUT(self):
        if not self._host_ok() or self.path != "/api/data":
            return self._send(403)
        try:
            n = int(self.headers.get("Content-Length", "0"))
            data = json.loads(self.rfile.read(n).decode("utf-8"))
            if not (isinstance(data, dict) and isinstance(data.get("projects"), list) and isinstance(data.get("tasks"), list)):
                raise ValueError("bad shape")
            _data_lock.acquire()
            disk = load_data()
            out = dict(disk)  # keep anything else already saved
            out["projects"], out["tasks"] = data["projects"], data["tasks"]
            for key, typ in (("notes", list), ("settings", dict), ("focus_log", list)):
                val = data.get(key)
                if val is None:  # an older window that doesn't know this key: keep what's on disk
                    val = disk.get(key, [] if typ is list else {})
                if not isinstance(val, typ):
                    raise ValueError("bad " + key)
                out[key] = val
            out["focus_log"] = out["focus_log"][-3000:]
            out["settings"] = clean_settings(out["settings"])
            save_data(out)
            _last_ping[0] = time.time()
            self._send(204, extra={"X-Rev": data_rev()})
        except Exception as e:
            self._send(400, json.dumps({"error": str(e)}).encode("utf-8"))
        finally:
            try:
                _data_lock.release()
            except RuntimeError:
                pass


def pythonw():
    w = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return w if os.path.exists(w) else sys.executable


def start_reminder(*args):
    subprocess.Popen([pythonw(), os.path.join(APP_DIR, "reminder.pyw"), *args], cwd=APP_DIR, creationflags=NO_WINDOW_FLAG)


def reminder_status():
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{REMINDER_PORT}/status", timeout=1.5) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception:
        return {"running": False}


def voice_available():
    return os.path.exists(TX_PY) and os.path.exists(VOICE_HELPER) and os.path.isdir(os.path.join(TX_DIR, "models"))


def refresh_env_key(name="OPENROUTER_API_KEY"):
    """If the user set the variable in Windows after this app started, pick it up from HKCU Environment (value is never logged)."""
    if os.environ.get(name) or os.name != "nt" or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,63}", name or ""):
        return bool(os.environ.get(name))
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
            val = winreg.QueryValueEx(k, name)[0]
        if val:
            os.environ[name] = str(val)
    except OSError:
        pass
    return bool(os.environ.get(name))


def ai_status():
    data = load_data()
    cfg = todo_ai.ai_settings(data) if todo_ai else {"enabled": False, "provider": "rules"}
    if cfg.get("api_key_env"):
        refresh_env_key(cfg["api_key_env"])
    key_set = bool(cfg.get("api_key_env")) and bool(os.environ.get(cfg.get("api_key_env", ""), ""))
    return {"loaded": todo_ai is not None, "enabled": cfg.get("enabled", False), "provider": cfg.get("provider"),
            "voice_available": voice_available(), "voice_model": cfg.get("voice_model", "small.en"),
            "api_key_env_set": key_set, "openrouter_env_set": refresh_env_key("OPENROUTER_API_KEY")}  # whether set, never the value


def transcribe_clip(handler, data):
    """Save the posted clip to a temp file, run voice_helper.py with the transcriber's venv python, delete the clip."""
    if not voice_available():
        raise todo_ai.AIError("local Whisper isn't available (the Live Transcriber's .venv or models weren't found)")
    n = int(handler.headers.get("Content-Length", "0"))
    if n <= 0 or n > MAX_AUDIO:
        raise todo_ai.AIError("the recording is empty or too long")
    audio = handler.rfile.read(n)
    ext = ".wav" if "wav" in (handler.headers.get("X-Audio-Type") or "") else ".webm"
    cfg = todo_ai.ai_settings(data)
    model = cfg.get("voice_model", "small.en")
    prompt = ", ".join(p.get("name", "") for p in data.get("projects", []))[:500]
    tmpdir = os.path.join(DATA_DIR, "voice-tmp")
    os.makedirs(tmpdir, exist_ok=True)
    fd, path = tempfile.mkstemp(suffix=ext, dir=tmpdir)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(audio)
        with _voice_lock:
            r = subprocess.run([TX_PY, VOICE_HELPER, path, "--model", model, "--device", "auto" if model != "small.en" else "cpu", "--prompt", prompt],
                               capture_output=True, text=True, timeout=180, creationflags=NO_WINDOW_FLAG, cwd=APP_DIR)
        line = (r.stdout or "").strip().splitlines()[-1:] or [""]
        try:
            out = json.loads(line[0])
        except ValueError:
            raise todo_ai.AIError("the voice helper didn't answer")
        if out.get("error"):
            raise todo_ai.AIError("transcription failed: " + str(out["error"])[:160])
        return out
    except subprocess.TimeoutExpired:
        raise todo_ai.AIError("transcription took too long")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# ---------------------------------------------------------------- email (read-only Outlook + drag/paste)
def _email_cfg(data=None):
    return todo_email.email_settings(data if data is not None else load_data())


def email_inbox():
    data = load_data()
    cfg = _email_cfg(data)
    cache = todo_email.load_cache()
    st = todo_email.get_state(data)
    items = []
    for it in cache.get("items") or []:
        pub = todo_email.public_item(it)
        pub["state"] = "added" if it.get("id") in st["added"] else "dismissed" if it.get("id") in st["dismissed"] else "new"
        if pub["state"] == "added":
            pub["task_id"] = (st["added"].get(it["id"]) or {}).get("task_id", "")
        items.append(pub)
    status = cache.get("status") or {"state": "never"}
    if not cfg["auto"]:
        status = dict(status, state="off")
    return {"settings": cfg, "status": status, "me": cache.get("me") or {}, "items": items, "fetched": cache.get("fetched", ""),
            "running": _email_job["running"], "windows": os.name == "nt"}


def _email_worker(cfg):
    try:
        c = todo_email.fetch_inbox(cfg)
        todo_email.save_cache(c)
    except Exception:
        c = todo_email.load_cache()
        c["status"] = {"state": "error", "error": "couldn't read the inbox", "checked": time.strftime("%Y-%m-%dT%H:%M:%S")}
        try:
            todo_email.save_cache(c)
        except OSError:
            pass
    finally:
        _email_job["running"] = False
        _email_lock.release()


def email_refresh(force=False):
    cfg = _email_cfg()
    if not cfg["auto"] and not force:
        return {"started": False, "reason": "off"}
    if not _email_lock.acquire(blocking=False):
        return {"started": False, "reason": "running"}
    _email_job.update(running=True, started=time.time())
    threading.Thread(target=_email_worker, args=(cfg,), daemon=True).start()
    return {"started": True}


def email_task_for(rec, project_hint=None):
    data = load_data()
    hint = project_hint if any(p.get("id") == project_hint for p in data.get("projects", [])) else None
    return todo_email.suggest_task(rec, data, todo_ai, hint)


def email_suggest(eid, project_hint=None):
    it = next((x for x in todo_email.load_cache().get("items") or [] if x.get("id") == eid), None)
    if not it:
        raise todo_email.EmailError("that email isn't in the list any more. Refresh and try again.")
    return dict(email_task_for(it, project_hint), email=todo_email.public_item(it))


def email_import_file(handler):
    n = int(handler.headers.get("Content-Length", "0"))
    if n <= 0 or n > MAX_UPLOAD:
        raise todo_email.EmailError("that file is empty or bigger than 25 MB")
    raw = handler.rfile.read(n)
    name = urllib.request.unquote(handler.headers.get("X-Filename") or "email.eml")[:200]
    try:
        rec = todo_email.parse_file(name, raw)
    except todo_email.EmailError as e:
        if not name.lower().endswith(".msg"):
            raise
        rec = _msg_via_outlook(raw)  # backup: classic Outlook's OpenSharedItem, only if Outlook is set up
        if rec is None:
            raise todo_email.EmailError("couldn't read that .msg here. Open it, press Ctrl+A then Ctrl+C, and paste it into the box instead.")
    return dict(email_task_for(rec, handler.headers.get("X-Project-Hint")), email=todo_email.public_item(rec))


def _msg_via_outlook(raw):
    if todo_email.run_outlook("status", timeout=15).get("state") != "ready":
        return None
    tmpdir = os.path.join(DATA_DIR, "email-tmp")
    os.makedirs(tmpdir, exist_ok=True)
    fd, path = tempfile.mkstemp(suffix=".msg", dir=tmpdir)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(raw)
        r = todo_email.run_outlook("msg", WTD_MSGPATH=path)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    if r.get("state") != "ok":
        return None
    return todo_email._record(r.get("subject"), r.get("sender"), r.get("sender_email"), r.get("received"), r.get("to"), r.get("cc"), r.get("body"), "msg")


def email_add(body):
    """The confirm card's Add: validate one add_task with the same allowlist as the AI, save it, remember the email."""
    if todo_ai is None:
        raise todo_email.EmailError("the task module didn't load")
    act = body.get("action")
    if not isinstance(act, dict) or act.get("type") != "add_task":
        raise todo_email.EmailError("only adding a task is allowed here")
    eid = str(body.get("email_id") or "")
    if not re.fullmatch(r"[0-9a-f]{8,40}", eid):
        raise todo_email.EmailError("bad email id")
    with _data_lock:
        data = load_data()
        new, changes, undo = todo_ai.apply_actions(data, [act])
        tid = (undo.get("added_task_ids") or [""])[0]
        for t in new.get("tasks", []):
            if t.get("id") == tid:
                t["source"] = "email"
                t["email_id"] = eid
        todo_email.set_state(new, eid, "added", tid)
        save_data(new)
    return {"data": new, "changes": changes, "undo": undo, "task_id": tid}


def email_open(eid):
    it = next((x for x in todo_email.load_cache().get("items") or [] if x.get("id") == eid), None)
    if not it or not it.get("entry_id"):
        raise todo_email.EmailError("that email can't be opened from here")
    r = todo_email.run_outlook("open", WTD_ENTRYID=it["entry_id"])
    if r.get("state") != "ok":
        raise todo_email.EmailError("classic Outlook couldn't open it (" + str(r.get("state")) + ")")
    return {"ok": True}


def find_outlook():
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            for sub in (r"Microsoft Office\Root\Office16", r"Microsoft Office\Office16"):
                p = os.path.join(base, sub, "OUTLOOK.EXE")
                if os.path.exists(p):
                    return p
    return None


def email_setup_outlook():
    """Only launches classic Outlook so the user can sign in themselves. Creates nothing."""
    exe = find_outlook()
    if not exe:
        raise todo_email.EmailError("classic Outlook isn't installed on this PC")
    subprocess.Popen([exe], cwd=os.path.dirname(exe))
    return {"ok": True}



# ---------------------------------------------------------------- v7: Dashboard (Today at a glance)
MEET_DIR = os.environ.get("WORKTODO_MEET_DIR") or os.path.join(TX_DIR, "transcripts", "notes")
_brief_lock = threading.Lock()   # one AI brief at a time


def _meeting_when(stem, fallback_ts):
    import datetime as _dt
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2})(\d{2})", stem)
    if m:
        try:
            return _dt.datetime(*map(int, m.groups())).isoformat(timespec="seconds")
        except ValueError:
            pass
    return _dt.datetime.fromtimestamp(fallback_ts).isoformat(timespec="seconds")


def meetings_list(limit=6):
    """Latest Live Transcriber meeting notes (read-only: *-notes.json in transcriber\\transcripts\\notes)."""
    try:
        names = [n for n in os.listdir(MEET_DIR) if n.lower().endswith("-notes.json")]
    except OSError:
        return {"available": False, "items": []}
    paths = []
    for n in names:
        p = os.path.join(MEET_DIR, n)
        try:
            paths.append((os.path.getmtime(p), p))
        except OSError:
            pass
    items = []
    for mt, p in sorted(paths, reverse=True)[:limit]:
        try:
            with open(p, "r", encoding="utf-8") as f:
                j = json.load(f)
            n = j.get("notes") if isinstance(j, dict) else None
            if not isinstance(n, dict):
                continue
        except (OSError, ValueError):
            continue
        stem = os.path.basename(p)[:-len("-notes.json")]
        strs = lambda k, lim=6, ln=300: [str(x)[:ln] for x in (n.get(k) or []) if isinstance(x, (str, int, float)) and str(x).strip()][:lim]
        acts = []
        for a in n.get("action_items") or []:
            if isinstance(a, dict) and str(a.get("what") or "").strip():
                acts.append({"who": str(a.get("who") or "")[:80], "what": str(a.get("what"))[:300], "due": str(a.get("due") or "")[:40]})
        items.append({"id": stem, "title": str(n.get("title") or stem)[:160], "summary": str(n.get("summary") or "")[:1200],
                      "key_points": strs("key_points", 8), "decisions": strs("decisions"), "open_questions": strs("open_questions"),
                      "action_items": acts[:10], "source": str(j.get("source") or ""), "when": _meeting_when(stem, mt),
                      "updated": _meeting_when("", float(j.get("updated_ts") or mt)), "words": j.get("words") or 0})
    return {"available": True, "items": items}


def _first_line(s, n=90):
    s = (str(s or "").strip().splitlines() or [""])[0].strip()
    return s if len(s) <= n else s[:n - 1].rstrip() + "\u2026"


def _hm12(hm):
    try:
        h, m = map(int, hm.split(":"))
        return f"{(h % 12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}"
    except (ValueError, AttributeError):
        return ""


def brief_context(data, now=None):
    """Everything the daily brief talks about, straight from the app's own data (tasks, inbox cache, meeting notes)."""
    import datetime as _dt
    now = now or _dt.datetime.now()
    today = now.date().isoformat()
    week = (now.date() + _dt.timedelta(days=7)).isoformat()
    hm = now.strftime("%H:%M")
    pname = {p.get("id"): p.get("name", "") for p in data.get("projects", [])}
    rank = {"urgent": 0, "high": 1, "normal": 2}

    def tinfo(t):
        return {"title": _first_line(t.get("title"), 110), "project": pname.get(t.get("project_id"), ""),
                "priority": t.get("priority", "normal"), "due": t.get("due", ""), "due_time": t.get("due_time", "")}

    open_ = [t for t in data.get("tasks", []) if not t.get("done")]
    overdue = [t for t in open_ if t.get("due") and (t["due"] < today or (t["due"] == today and t.get("due_time") and t["due_time"] <= hm))]
    od_ids = {id(t) for t in overdue}
    due_today = [t for t in open_ if t.get("due") == today and id(t) not in od_ids]
    key = lambda t: (rank.get(t.get("priority"), 2), t.get("due") or "9999", t.get("due_time") or "99")
    upcoming = sorted([t for t in open_ if t.get("due") and today < t["due"] <= week], key=lambda t: (t["due"], t.get("due_time") or "99"))
    prio = sorted([t for t in open_ if t.get("priority") in ("urgent", "high")], key=key)
    done_today = [t for t in data.get("tasks", []) if t.get("done") and (t.get("completed") or "")[:10] == today]
    try:
        inbox = [i for i in (email_inbox().get("items") or []) if i.get("state") == "new"] if todo_email else []
    except Exception:
        inbox = []
    meets = meetings_list(3).get("items", [])
    return {"now": now.isoformat(timespec="minutes"), "weekday": now.strftime("%A"),
            "counts": {"open": len(open_), "overdue": len(overdue), "due_today": len(due_today), "done_today": len(done_today),
                       "upcoming_7d": len(upcoming), "urgent_high": len(prio), "new_emails": len(inbox)},
            "overdue": [tinfo(t) for t in sorted(overdue, key=key)[:8]],
            "due_today": [tinfo(t) for t in sorted(due_today, key=lambda t: (t.get("due_time") or "99", rank.get(t.get("priority"), 2)))[:8]],
            "urgent_high": [tinfo(t) for t in prio[:8]],
            "upcoming_7d": [tinfo(t) for t in upcoming[:8]],
            "new_emails": [{"from": i.get("sender") or i.get("sender_email") or "", "subject": _first_line(i.get("subject"), 120),
                            "received": i.get("received", ""), "high_importance": i.get("importance") == 2} for i in inbox[:8]],
            "recent_meetings": [{"title": m["title"], "when": m["when"], "summary": m["summary"][:300],
                                 "action_items": [a["what"][:140] for a in m["action_items"][:4]]} for m in meets]}


def local_brief(ctx):
    import datetime as _dt
    c = ctx["counts"]
    def lbl(t):
        return "\u201c" + t["title"] + "\u201d" + (f" ({t['project']})" if t.get("project") else "")
    def day(s):
        try:
            return _dt.date.fromisoformat(s).strftime("%a %b %d").replace(" 0", " ")
        except ValueError:
            return s
    head = []
    if c["overdue"]:
        head.append(f"{c['overdue']} overdue")
    if c["due_today"]:
        head.append(f"{c['due_today']} due today")
    if c["new_emails"]:
        head.append(f"{c['new_emails']} new email{'s' if c['new_emails'] != 1 else ''} to triage")
    headline = (", ".join(head) + ".") if head else "Nothing overdue or due today. A clear runway for priority work."
    b = []
    if ctx["overdue"]:
        b.append("Overdue: " + "; ".join(lbl(t) for t in ctx["overdue"][:3]) + ("" if c["overdue"] <= 3 else f" and {c['overdue'] - 3} more") + ".")
    timed = [t for t in ctx["due_today"] if t.get("due_time")]
    if ctx["due_today"]:
        b.append("Due today: " + "; ".join(lbl(t) + (f" at {_hm12(t['due_time'])}" if t.get("due_time") else "") for t in ctx["due_today"][:3]) + ".")
    pr = [t for t in ctx["urgent_high"] if t not in ctx["overdue"] and t not in ctx["due_today"]]
    if pr:
        b.append(f"{c['urgent_high']} urgent/high-priority open. Top: " + "; ".join(f"{lbl(t)} [{t['priority']}]" for t in pr[:2]) + ".")
    if ctx["upcoming_7d"]:
        t = ctx["upcoming_7d"][0]
        b.append(f"Coming up this week: {c['upcoming_7d']} task{'s' if c['upcoming_7d'] != 1 else ''}. Next: {lbl(t)} on {day(t['due'])}" + (f" at {_hm12(t['due_time'])}" if t.get("due_time") else "") + ".")
    if ctx["new_emails"]:
        e = ctx["new_emails"][0]
        b.append(f"Inbox: {c['new_emails']} email{'s' if c['new_emails'] != 1 else ''} for you not yet turned into tasks. Latest from {e['from']}: \u201c{e['subject']}\u201d.")
    if ctx["recent_meetings"]:
        m = ctx["recent_meetings"][0]
        b.append(f"Latest meeting notes: \u201c{_first_line(m['title'], 80)}\u201d ({day(m['when'][:10])})" + (f", {len(m['action_items'])} action item(s)." if m["action_items"] else "."))
    if c["done_today"]:
        b.append(f"Already done today: {c['done_today']}. Nice.")
    first = (ctx["overdue"] or ctx["due_today"] or ctx["urgent_high"] or [None])[0]
    focus = ("Start with " + lbl(first) + ".") if first else ("Triage the inbox first." if ctx["new_emails"] else "Pick a project and make progress.")
    return {"headline": headline, "bullets": b[:7], "focus": focus}


SYSTEM_BRIEF = ("You are a sharp, friendly executive assistant. From the JSON snapshot of the user's work to-do app, write a short "
                "brief for their workday. Use ONLY facts in the JSON; never invent tasks, people, times or emails. Reply with JSON only: "
                '{"headline": "<one sentence, max 22 words>", "bullets": ["<3 to 6 bullets, max 32 words each>"], '
                '"focus": "<the single thing to do first, max 25 words>"}')


def make_brief(data, use_ai):
    """Local brief always; the AI brief only when asked AND AI is turned on with a model. Any AI failure falls back to local."""
    import datetime as _dt
    ctx = brief_context(data)
    out = dict(local_brief(ctx), source="local", generated=_dt.datetime.now().isoformat(timespec="seconds"), note="", counts=ctx["counts"])
    if not use_ai:
        return out
    cfg = todo_ai.ai_settings(data) if todo_ai else {"enabled": False}
    if not cfg.get("enabled"):
        out["note"] = "AI is off (Reminders & settings \u2192 AI suggestions). Showing the built-in brief."
        return out
    if cfg.get("provider") != "openai_compatible":
        out["note"] = "AI provider is the built-in rules, so this is the built-in brief."
        return out
    if not _brief_lock.acquire(blocking=False):
        out["note"] = "An AI brief is already being written. Showing the built-in brief."
        return out
    try:
        prov = todo_ai.get_provider(cfg)
        reply = todo_ai._extract_json(prov._chat(SYSTEM_BRIEF, json.dumps(ctx), timeout=90))
        if not isinstance(reply, dict):
            raise todo_ai.AIError("the model didn't return a brief")
        bullets = [str(x).strip()[:300] for x in (reply.get("bullets") or []) if str(x).strip()][:7]
        head = str(reply.get("headline") or "").strip()[:240]
        if not head and not bullets:
            raise todo_ai.AIError("the model returned an empty brief")
        out.update(headline=head or out["headline"], bullets=bullets or out["bullets"],
                   focus=str(reply.get("focus") or out["focus"]).strip()[:240], source="ai", model=cfg.get("model", ""))
    except Exception as e:
        msg = str(e) if todo_ai and isinstance(e, todo_ai.AIError) else type(e).__name__
        out["note"] = "AI didn't answer (" + msg[:140] + "). Showing the built-in brief."
    finally:
        _brief_lock.release()
    return out


SYSTEM_REVIEW = ("You are a supportive, practical work coach. From the JSON snapshot of one person's week in their to-do app, write a short "
                 "weekly review. Reply with ONLY a JSON object: {\"headline\": one sentence, \"wins\": [2-4 short bullets], "
                 "\"slipped\": [0-3 short bullets about what slipped and why it might matter], \"next\": [2-4 short bullets for next week]}. "
                 "Use only facts from the snapshot. No markdown.")


def review_context(data, start=None, end=None):
    import datetime as _dt
    today = _dt.date.today()
    try:
        s0 = _dt.date.fromisoformat(str(start))
    except (TypeError, ValueError):
        s0 = today - _dt.timedelta(days=today.weekday())
    try:
        e0 = _dt.date.fromisoformat(str(end))
    except (TypeError, ValueError):
        e0 = s0 + _dt.timedelta(days=6)
    ns, ne = e0 + _dt.timedelta(days=1), e0 + _dt.timedelta(days=7)
    pname = {p.get("id"): p.get("name", "") for p in data.get("projects", [])}
    tasks = data.get("tasks", [])
    S, E = s0.isoformat(), e0.isoformat()
    done = [t for t in tasks if t.get("done") and S <= str(t.get("completed") or "")[:10] <= E]
    slipped = [t for t in tasks if not t.get("done") and t.get("due") and t["due"] <= min(E, today.isoformat()) and t["due"] < today.isoformat()]
    nxt = [t for t in tasks if not t.get("done") and t.get("due") and ns.isoformat() <= t["due"] <= ne.isoformat()]
    mins = {}
    for x in data.get("focus_log", []) or []:
        if isinstance(x, dict) and S <= str(x.get("start") or "")[:10] <= E:
            k = pname.get(x.get("project_id"), "No project")
            mins[k] = mins.get(k, 0) + int(x.get("minutes") or 0)
    by_proj = {}
    for t in done:
        k = pname.get(t.get("project_id"), "No project")
        by_proj[k] = by_proj.get(k, 0) + 1
    def row(t):
        return {"title": _first_line(t.get("title", "")), "project": pname.get(t.get("project_id"), ""), "priority": t.get("priority", "normal"), "due": t.get("due", "")}
    return {"week": [S, E], "today": today.isoformat(), "done_count": len(done), "done": [row(t) for t in done][:40],
            "done_by_project": by_proj, "slipped": [row(t) for t in slipped][:25], "next_week": [row(t) for t in nxt][:25],
            "focus_minutes_by_project": mins, "focus_total": sum(mins.values()), "open_count": sum(1 for t in tasks if not t.get("done"))}


def local_review(ctx):
    n, sl, nx, ft = ctx["done_count"], len(ctx["slipped"]), len(ctx["next_week"]), ctx["focus_total"]
    top = sorted(ctx["done_by_project"].items(), key=lambda kv: -kv[1])
    head = (f"You finished {n} task{'s' if n != 1 else ''} this week" + (f", mostly in {top[0][0]}" if top else "") + "."
            if n else "No tasks were marked done this week yet.")
    wins = [f"{k}: {v} done" for k, v in top[:3]]
    if ft:
        wins.append((f"{ft // 60}h {ft % 60}m" if ft >= 60 else f"{ft}m") + " of focused time logged")
    slipped = [f"{sl} task{'s' if sl != 1 else ''} slipped past the due date" + (f", e.g. \u201c{ctx['slipped'][0]['title']}\u201d" if sl else "")] if sl else []
    nxt = []
    if nx:
        nxt.append(f"{nx} task{'s' if nx != 1 else ''} due next week")
        urgent = [t for t in ctx["next_week"] if t["priority"] != "normal"]
        if urgent:
            nxt.append("Start with \u201c" + urgent[0]["title"] + "\u201d")
    if sl:
        nxt.append("Re-date or finish the slipped tasks first")
    if not nxt:
        nxt.append("Plan next week: give your top tasks a due date")
    return {"headline": head, "wins": wins or ["A quiet week. Small steps count."], "slipped": slipped, "next": nxt}


def make_review(data, start=None, end=None, use_ai=False):
    import datetime as _dt
    ctx = review_context(data, start, end)
    out = dict(local_review(ctx), source="local", note="", generated=_dt.datetime.now().isoformat(timespec="seconds"), week=ctx["week"])
    if not use_ai:
        return out
    cfg = todo_ai.ai_settings(data) if todo_ai else {"enabled": False}
    if not cfg.get("enabled") or cfg.get("provider") != "openai_compatible":
        out["note"] = "AI isn't set up (Reminders & settings \u2192 AI suggestions), so this is the built-in recap."
        return out
    if not _brief_lock.acquire(blocking=False):
        out["note"] = "The AI is busy with another request. Showing the built-in recap."
        return out
    try:
        reply = todo_ai._extract_json(todo_ai.get_provider(cfg)._chat(SYSTEM_REVIEW, json.dumps(ctx), timeout=90))
        if not isinstance(reply, dict):
            raise todo_ai.AIError("the model didn't return a recap")
        clean = lambda xs: [str(x).strip()[:300] for x in (xs or []) if str(x).strip()][:5]
        if not (reply.get("headline") or reply.get("wins")):
            raise todo_ai.AIError("the model returned an empty recap")
        out.update(headline=str(reply.get("headline") or out["headline"])[:240], wins=clean(reply.get("wins")) or out["wins"],
                   slipped=clean(reply.get("slipped")), next=clean(reply.get("next")) or out["next"], source="ai", model=cfg.get("model", ""))
    except Exception as e:
        msg = str(e) if todo_ai and isinstance(e, todo_ai.AIError) else type(e).__name__
        out["note"] = "AI didn't answer (" + msg[:140] + "). Showing the built-in recap."
    finally:
        _brief_lock.release()
    return out


def find_edge():
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.exists(p):
                return p
    return None


def open_window():
    if NO_WINDOW:
        return
    edge = find_edge()
    if edge:
        profile = os.path.join(DATA_DIR, "window-profile")
        subprocess.Popen([edge, f"--app={APP_URL}", f"--user-data-dir={profile}",
                          "--window-size=1200,780", "--no-first-run", "--no-default-browser-check"])
    else:
        webbrowser.open(APP_URL)


def watchdog(server):
    while True:
        time.sleep(10)
        if time.time() - _last_ping[0] > IDLE_LIMIT:
            server.shutdown()
            return


def main():
    after_update = "--after-update" in sys.argv
    server = None
    for i in range(40 if after_update else 1):   # after an update, wait for the old server to let go of the port
        try:
            server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
            break
        except OSError:
            if after_update:
                time.sleep(0.5)
    if server is None:
        open_window()  # already running: just bring up a window
        return
    _server[0] = server
    try:
        migrate_data()
    except Exception:
        pass
    threading.Thread(target=watchdog, args=(server,), daemon=True).start()
    threading.Thread(target=auto_update_check, daemon=True).start()
    if not after_update:   # after an update the open window reloads itself
        open_window()
    server.serve_forever()


HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>Work To-Do</title>
<link rel="icon" href="/favicon.ico">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{
  --bg:#0d1017; --panel:#141822; --panel2:#1a1f2c; --line:#262c3b; --text:#e8ebf2; --muted:#8a93a8;
  --accent:#7c6cff; --accent2:#4fc3ff; --urgent:#ff5d73; --high:#ffa94d; --normal:#6ea8fe; --ok:#3ddc97;
  --r:14px;
  --input:#0d1017; --hover:#1b2030; --text2:#c9cfdc; --line2:#34405a; --sw:#2a3142; --pop:#1d2333; --code:#9fe0ff; --okt:#9ff0cb;
  --ink:#ffffff; --strong:#ffffff; --scheme:dark; --side:linear-gradient(180deg,#121621,#0f121b); --modal:linear-gradient(180deg,#171c28,#131722);
}
html[data-theme="light"]{
  --bg:#f3f4f8; --panel:#ffffff; --panel2:#f8f9fc; --line:#e0e3eb; --text:#1b2030; --muted:#636c80; --accent2:#0b8fd6;
  --input:#f6f7fa; --hover:#eceef4; --text2:#394155; --line2:#c9cfdb; --sw:#cdd2dd; --pop:#ffffff; --code:#0a6b9e; --okt:#0d7a53;
  --ink:#000000; --strong:#11141c; --scheme:light; --side:linear-gradient(180deg,#fbfbfd,#f0f2f7); --modal:#ffffff;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:var(--bg);color:var(--text);
  font-family:"Segoe UI Variable Text","Segoe UI",system-ui,sans-serif;font-size:14px}
button,input,select{font:inherit;color:inherit}
.app{display:grid;grid-template-columns:270px 1fr;height:100vh}
/* sidebar */
.side{background:var(--side);border-right:1px solid var(--line);display:flex;flex-direction:column;padding:20px 14px}
.brand{display:flex;align-items:center;gap:10px;padding:4px 8px 22px}
.logo{width:34px;height:34px;border-radius:10px;background:linear-gradient(135deg,var(--accent),var(--accent2));
  display:grid;place-items:center;box-shadow:0 6px 18px color-mix(in srgb,var(--accent) 33%,transparent)}
.logo svg{width:20px;height:20px}
.brand b{font-size:17px;letter-spacing:.2px}
.brand small{display:block;color:var(--muted);font-size:11.5px;margin-top:1px}
.sect{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:1.2px;padding:14px 10px 8px}
.plist{flex:1;overflow:auto;margin:0 -4px;padding:0 4px}
.pitem{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:10px;cursor:pointer;color:var(--text2);position:relative;user-select:none}
.pitem:hover{background:var(--hover)}
.pitem.on{background:linear-gradient(90deg,color-mix(in srgb,var(--accent) 15%,transparent),color-mix(in srgb,var(--accent) 5%,transparent));color:#fff;box-shadow:inset 3px 0 0 var(--accent)}
.dot{width:10px;height:10px;border-radius:50%;flex:none}
.pname{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.count{font-size:11.5px;color:var(--muted);background:color-mix(in srgb,var(--ink) 5%,transparent);padding:1px 8px;border-radius:20px}
.pitem .tools{display:none;gap:2px}
.pitem:hover .tools{display:flex}
.pitem:hover .count{display:none}
.ibtn{background:none;border:0;padding:4px;border-radius:6px;cursor:pointer;color:var(--muted);display:grid;place-items:center}
.ibtn:hover{background:color-mix(in srgb,var(--ink) 8%,transparent);color:#fff}
.ibtn svg{width:15px;height:15px}
.pedit{flex:1;background:var(--input);border:1px solid var(--accent);border-radius:7px;padding:3px 7px;outline:none;min-width:0}
.newp{display:flex;gap:6px;margin-top:10px;padding-top:12px;border-top:1px solid var(--line)}
.newp input{flex:1;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:9px 11px;outline:none;min-width:0}
.newp input:focus{border-color:var(--accent)}
.btn{background:linear-gradient(135deg,var(--accent),color-mix(in srgb,var(--accent) 72%,#ffffff));border:0;border-radius:10px;padding:9px 14px;cursor:pointer;font-weight:600;color:#fff;box-shadow:0 4px 14px color-mix(in srgb,var(--accent) 27%,transparent);white-space:nowrap}
.btn:hover{filter:brightness(1.1)}
.btn.sq{padding:9px 12px}
/* main */
.main{overflow:auto;padding:28px 36px 40px}
.head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:22px}
.head h1{margin:0;font-size:28px;font-weight:700;letter-spacing:-.3px;display:flex;align-items:center;gap:12px}
.head h1 .dot{width:14px;height:14px}
.head .sub{color:var(--muted);margin-top:4px}
.search{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:9px 12px;display:flex;align-items:center;gap:8px;width:260px}
.search input{background:none;border:0;outline:none;flex:1;min-width:0}
.search svg{width:16px;height:16px;color:var(--muted)}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
.stat{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:14px 16px;position:relative;overflow:hidden}
.stat .n{font-size:26px;font-weight:700}
.stat .l{color:var(--muted);font-size:12.5px;margin-top:2px}
.stat::after{content:"";position:absolute;right:-20px;top:-20px;width:80px;height:80px;border-radius:50%;background:var(--c);opacity:.13}
.add{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:18px;box-shadow:0 10px 30px #00000033}
.add .title{flex:1 1 100%;resize:none;overflow:hidden;line-height:1.45;min-height:44px;max-height:240px;font-family:inherit;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:11px 13px;outline:none;font-size:14.5px}
.add .title:focus{border-color:var(--accent)}
.seg{display:flex;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:3px}
.seg button{background:none;border:0;padding:7px 11px;border-radius:7px;cursor:pointer;color:var(--muted);font-size:13px;display:flex;align-items:center;gap:6px}
.seg button .dot{width:8px;height:8px}
.seg button.on{background:color-mix(in srgb,var(--ink) 7%,transparent);color:#fff}
.field{background:var(--input);border:1px solid var(--line);border-radius:10px;padding:8px 10px;outline:none;color-scheme:var(--scheme)}
.toolbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:4px 0 12px;flex-wrap:wrap}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.chip{background:var(--panel);border:1px solid var(--line);border-radius:20px;padding:6px 13px;cursor:pointer;color:var(--text2);font-size:13px;display:flex;gap:7px;align-items:center}
.chip.on{border-color:var(--accent);background:color-mix(in srgb,var(--accent) 12%,transparent);color:#fff}
.chip .dot{width:8px;height:8px}
.toggle{display:flex;align-items:center;gap:9px;color:var(--muted);cursor:pointer;user-select:none;font-size:13px}
.sw{width:36px;height:20px;border-radius:20px;background:var(--sw);position:relative;transition:.2s}
.sw::after{content:"";position:absolute;width:16px;height:16px;border-radius:50%;background:#fff;top:2px;left:2px;transition:.2s}
.toggle.on .sw{background:var(--accent)} .toggle.on .sw::after{left:18px}
.list{display:flex;flex-direction:column;gap:8px}
.group{color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:1.2px;margin:14px 2px 2px;display:flex;align-items:center;gap:8px}
.group::after{content:"";flex:1;height:1px;background:var(--line)}
.task{display:flex;align-items:center;gap:14px;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:12px 14px;position:relative;transition:transform .12s, border-color .12s, background .12s;animation:pop .18s ease}
@keyframes pop{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
.task:hover{border-color:var(--line2);background:var(--panel2)}
.task::before{content:"";position:absolute;left:0;top:10px;bottom:10px;width:3px;border-radius:3px;background:var(--pc)}
.check{width:22px;height:22px;border-radius:50%;border:2px solid var(--pc);background:none;cursor:pointer;flex:none;display:grid;place-items:center;transition:.15s}
.check:hover{background:color-mix(in srgb,var(--pc) 25%,transparent)}
.check svg{width:13px;height:13px;opacity:0}
.task.done .check{background:var(--ok);border-color:var(--ok)} .task.done .check svg{opacity:1}
.task.done .ttl{text-decoration:line-through;color:var(--muted)}
.task.done{opacity:.7}
.body{flex:1;min-width:0}
.ttl{white-space:pre-wrap;line-height:1.5;font-size:14.5px;outline:none;border-radius:5px;padding:1px 3px;margin:-1px -3px;word-break:break-word}
.ttl:focus{background:var(--input);box-shadow:0 0 0 1px var(--accent)}
.meta{display:flex;gap:8px;margin-top:6px;flex-wrap:wrap;align-items:center}
.badge{font-size:11.5px;font-weight:600;padding:3px 9px;border-radius:20px;cursor:pointer;border:0;letter-spacing:.3px;text-transform:uppercase}
.b-urgent{background:#ff5d7322;color:var(--urgent)} .b-high{background:#ffa94d22;color:var(--high)} .b-normal{background:#6ea8fe1f;color:var(--normal)}
.tag{font-size:12px;color:var(--muted);display:flex;align-items:center;gap:6px;background:color-mix(in srgb,var(--ink) 3%,transparent);padding:3px 9px;border-radius:20px}
.tag .dot{width:7px;height:7px}
.tag.over{color:var(--urgent);background:#ff5d7318} .tag.today{color:var(--high);background:#ffa94d18}
.due{position:relative;cursor:pointer}
.due input{position:absolute;inset:0;opacity:0;cursor:pointer;color-scheme:var(--scheme)}
.task .acts{display:flex;gap:2px;opacity:0;transition:.15s}
.task:hover .acts{opacity:1}
.empty{text-align:center;color:var(--muted);padding:46px 20px 40px;border:1px dashed var(--line);border-radius:var(--r);
  background:radial-gradient(ellipse at 50% 0%,color-mix(in srgb,var(--accent) 8%,transparent),transparent 65%);animation:pop .25s ease}
.empty .big{font-size:40px;margin-bottom:8px}
.empty .art{width:74px;height:74px;margin:0 auto 16px;border-radius:22px;display:grid;place-items:center;
  background:linear-gradient(135deg,var(--ea,var(--accent)),var(--eb,var(--accent2)));box-shadow:0 14px 34px color-mix(in srgb,var(--accent) 25%,transparent);animation:float 3.2s ease-in-out infinite}
.empty .art svg{width:36px;height:36px}
.empty h3{margin:0 0 6px;color:var(--text);font-size:17px;font-weight:650}
.empty .eacts{display:flex;gap:10px;justify-content:center;align-items:center;margin-top:18px;flex-wrap:wrap}
@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-5px)}}
.toast{position:fixed;bottom:18px;right:22px;background:var(--pop);border:1px solid var(--line);padding:9px 14px;border-radius:10px;color:var(--muted);font-size:12.5px;opacity:0;transition:.3s;pointer-events:none}
.toast.show{opacity:1}
/* notes */
.nicon{width:18px;height:18px;flex:none;display:grid;place-items:center;color:var(--accent2)}
.nicon svg{width:16px;height:16px}
.nlink{display:inline-flex;align-items:center;gap:6px;margin-left:10px;font-size:12px;color:var(--accent2);background:#4fc3ff14;border:0;padding:3px 10px;border-radius:20px;cursor:pointer}
.nlink:hover{background:#4fc3ff26}
.nlink svg{width:12px;height:12px}
.neditor{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:14px;display:flex;flex-direction:column;gap:10px;margin-bottom:18px;box-shadow:0 10px 30px #00000033}
.neditor.editing{border-color:var(--accent);box-shadow:0 0 0 1px color-mix(in srgb,var(--accent) 33%,transparent),0 10px 30px #00000033}
.neditor .lbl{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;display:flex;align-items:center;gap:8px}
.neditor .lbl b{color:var(--accent);font-weight:600}
.ninput{background:var(--input);border:1px solid var(--line);border-radius:10px;padding:10px 13px;outline:none;font-size:15px;font-weight:600;width:100%}
.nbody{background:var(--input);border:1px solid var(--line);border-radius:10px;padding:11px 13px;outline:none;width:100%;resize:none;min-height:96px;max-height:420px;line-height:1.5;font-family:inherit;font-size:14px;color:inherit}
.ninput:focus,.nbody:focus{border-color:var(--accent)}
.nrow{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.nrow .hint{color:var(--muted);font-size:12px}
.btn.ghost{background:color-mix(in srgb,var(--ink) 5%,transparent);box-shadow:none;color:var(--text2);font-weight:500}
.btn.ghost:hover{background:color-mix(in srgb,var(--ink) 9%,transparent)}
.ngrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:14px;align-items:start}
.note{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:16px 18px 14px;position:relative;animation:pop .18s ease;transition:border-color .12s, background .12s}
.note:hover{border-color:var(--line2);background:var(--panel2)}
.note::before{content:"";position:absolute;left:18px;right:18px;top:0;height:2px;border-radius:0 0 3px 3px;background:var(--nc)}
.note h3{margin:0 70px 8px 0;font-size:15.5px;font-weight:650;line-height:1.35;word-break:break-word}
.note .nb{white-space:pre-wrap;line-height:1.55;color:var(--text2);word-break:break-word;font-size:13.5px;max-height:340px;overflow:auto}
.note code,.nb code{font-family:"Cascadia Code","Cascadia Mono",Consolas,monospace;font-size:12.5px;background:var(--input);border:1px solid var(--line);border-radius:6px;padding:1px 6px;color:var(--code)}
.note .nb .cmd{font-family:"Cascadia Code","Cascadia Mono",Consolas,monospace;font-size:12.5px;color:var(--code);background:var(--input);border:1px solid var(--line);border-radius:7px;padding:2px 8px;display:inline-block;margin:1px 0}
.note .nmeta{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;align-items:center;font-size:12px;color:var(--muted)}
.note .acts{position:absolute;top:10px;right:10px;display:flex;gap:2px;opacity:0;transition:.15s}
.note:hover .acts{opacity:1}
/* v4: smart views, banner, modals, shortcuts, undo */
.pitem{transition:background .15s,color .15s}
.sicon{width:18px;height:18px;flex:none;display:grid;place-items:center}
.sicon svg{width:16px;height:16px}
.count.red{background:#ff5d7326;color:var(--urgent);font-weight:600}
.count.amber{background:#ffa94d22;color:var(--high);font-weight:600}
.sfoot{margin-top:10px;padding-top:10px;border-top:1px solid var(--line);display:flex;flex-direction:column;gap:2px}
.sbtn{display:flex;align-items:center;gap:10px;width:100%;background:none;border:0;padding:8px 10px;border-radius:10px;color:var(--text2);cursor:pointer;text-align:left}
.sbtn:hover{background:var(--hover);color:#fff}
.sbtn svg{width:16px;height:16px;color:var(--muted)}
.sbtn span{flex:1}
kbd{font-family:"Cascadia Code","Cascadia Mono",Consolas,monospace;font-size:11px;background:var(--input);border:1px solid var(--line2);border-bottom-width:2px;border-radius:6px;padding:1px 6px;color:var(--text2);white-space:nowrap}
.search kbd{opacity:.8}
.banner{display:flex;align-items:center;gap:12px;background:linear-gradient(90deg,#ffa94d1c,#ff5d7312);border:1px solid #ffa94d44;border-radius:var(--r);padding:11px 12px 11px 14px;margin-bottom:16px;animation:slidein .3s ease}
.banner[hidden]{display:none}
.banner .bi{width:30px;height:30px;border-radius:9px;background:#ffa94d26;color:var(--high);display:grid;place-items:center;flex:none;animation:ring 2.4s ease-in-out 1}
.banner .bi svg{width:17px;height:17px}
.banner .bt{flex:1;min-width:0} .banner .bt b{color:#fff} .banner .bt small{display:block;color:var(--muted);font-size:12px;margin-top:1px}
.btn.sm{padding:7px 12px;font-size:13px}
@keyframes slidein{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}
@keyframes ring{0%,100%{transform:rotate(0)}10%,30%{transform:rotate(-12deg)}20%,40%{transform:rotate(12deg)}50%{transform:rotate(0)}}
.view-in{animation:viewin .22s ease}
@keyframes viewin{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.task.sel{border-color:var(--accent);box-shadow:0 0 0 1px color-mix(in srgb,var(--accent) 47%,transparent),0 8px 22px #00000040}
.task.leaving{animation:leave .3s ease forwards;pointer-events:none}
.task.leaving .check{background:var(--ok);border-color:var(--ok);transform:scale(1.15)} .task.leaving .check svg{opacity:1}
@keyframes leave{0%{opacity:1}35%{opacity:1;transform:none}100%{opacity:0;transform:translateX(26px)}}
.tag.ghost{opacity:0;transition:opacity .15s} .task:hover .tag.ghost,.task.sel .tag.ghost{opacity:1}
.tag.time input{position:absolute;inset:0;opacity:0;cursor:pointer;color-scheme:var(--scheme)}
.add .field.time{width:118px}
.modal-back{position:fixed;inset:0;background:#05070bb8;backdrop-filter:blur(5px);display:grid;place-items:center;z-index:50;animation:fadein .15s ease}
.modal-back[hidden]{display:none}
.modal{background:var(--modal);border:1px solid var(--line);border-radius:18px;width:min(780px,92vw);max-height:86vh;overflow:auto;padding:22px 26px 24px;box-shadow:0 30px 90px #000c;animation:zoom .2s cubic-bezier(.2,.9,.3,1.2)}
.modal.narrow{width:min(560px,92vw)}
.modal .mh{display:flex;align-items:center;gap:12px;margin-bottom:16px}
.modal .mh h2{margin:0;font-size:19px;flex:1}
.modal .mh .logo{width:30px;height:30px;border-radius:9px}
@keyframes fadein{from{opacity:0}to{opacity:1}}
@keyframes zoom{from{opacity:0;transform:scale(.96) translateY(8px)}to{opacity:1;transform:none}}
.kgrid{display:grid;grid-template-columns:1fr 1fr;gap:4px 34px}
.ksec{color:var(--accent2);font-size:11px;text-transform:uppercase;letter-spacing:1.2px;margin:14px 0 4px;grid-column:span 1}
.kcol{display:flex;flex-direction:column}
.krow{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:6px 0;border-bottom:1px solid color-mix(in srgb,var(--ink) 3%,transparent);font-size:13px;color:var(--text2)}
.krow .keys{display:flex;gap:4px;align-items:center;flex:none;color:var(--muted);font-size:11px}
.mfoot{margin-top:16px;color:var(--muted);font-size:12.5px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.srow{display:flex;align-items:center;gap:14px;padding:14px 0;border-bottom:1px solid color-mix(in srgb,var(--ink) 4%,transparent)}
.srow .st{flex:1} .srow .st b{display:block;font-weight:600} .srow .st small{color:var(--muted);font-size:12.5px;line-height:1.45;display:block;margin-top:2px}
.pill{font-size:12px;padding:3px 10px;border-radius:20px;background:color-mix(in srgb,var(--ink) 5%,transparent);color:var(--muted)}
.pill.ok{background:#3ddc9722;color:var(--ok)} .pill.bad{background:#ff5d7322;color:var(--urgent)}
.snack{position:fixed;left:50%;bottom:22px;transform:translate(-50%,24px);opacity:0;pointer-events:none;transition:.25s cubic-bezier(.2,.9,.3,1.1);
  background:var(--pop);border:1px solid var(--line2);border-radius:12px;padding:8px 8px 8px 16px;display:flex;align-items:center;gap:14px;box-shadow:0 14px 40px #0009;z-index:60;font-size:13.5px}
.snack.show{opacity:1;transform:translate(-50%,0);pointer-events:auto}
.snack button{background:color-mix(in srgb,var(--accent) 15%,transparent);color:var(--accent);border:0;border-radius:8px;padding:6px 12px;cursor:pointer;font-weight:600}
.snack button:hover{background:color-mix(in srgb,var(--accent) 27%,transparent)}
mark{background:#ffd43b40;color:inherit;border-radius:3px;padding:0 1px}
/* v5: AI suggestions */
.aiask{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:12px 14px;margin-bottom:18px;box-shadow:0 10px 30px #00000033;position:relative;overflow:hidden}
.aiask::before{content:"";position:absolute;inset:0 0 auto 0;height:2px;background:linear-gradient(90deg,var(--accent),var(--accent2))}
.aiask.off{opacity:.62} .aiask.off::before{background:var(--sw)}
.aiask .lbl{display:flex;align-items:center;gap:8px;font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;margin-bottom:9px}
.aiask .lbl svg{width:15px;height:15px;color:var(--accent)} .aiask .hint{text-transform:none;letter-spacing:0}
.aiaskrow{display:flex;gap:10px;align-items:flex-end}
.aiaskrow textarea{flex:1;resize:none;min-height:44px;max-height:160px;line-height:1.45;font-family:inherit;font-size:14.5px;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:11px 13px;outline:none;color:inherit}
.aiaskrow textarea:focus{border-color:var(--accent)} .aiaskrow textarea:disabled{cursor:not-allowed}
.micbtn{width:44px;height:44px;flex:none;border-radius:12px;border:1px solid var(--line);background:var(--input);color:var(--accent2);cursor:pointer;display:grid;place-items:center;transition:.15s}
.micbtn svg{width:19px;height:19px} .micbtn:hover:not(:disabled){border-color:var(--accent2);background:#4fc3ff14}
.micbtn.rec{color:#fff;background:var(--urgent);border-color:var(--urgent);animation:pulse 1.3s ease-in-out infinite}
.micbtn:disabled,.btn:disabled{opacity:.45;cursor:not-allowed;filter:none}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 #ff5d7366}50%{box-shadow:0 0 0 8px #ff5d7300}}
.aistat{color:var(--muted);font-size:12.5px;margin-top:8px;min-height:16px}
.aicard{display:flex;gap:14px;align-items:flex-start;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:13px 14px 12px 18px;position:relative;animation:pop .18s ease}
.aicard:hover{border-color:var(--line2);background:var(--panel2)}
.aicard::before{content:"";position:absolute;left:0;top:10px;bottom:10px;width:3px;border-radius:3px;background:var(--pc)}
.aicard .ttl{font-weight:600}
.aireason{color:var(--text2);font-size:13px;line-height:1.5;margin-top:3px}
.aichg{margin:8px 0 0;padding:0;list-style:none;display:flex;flex-direction:column;gap:4px}
.aichg li{font-size:12.5px;color:var(--muted);padding-left:16px;position:relative;line-height:1.45;word-break:break-word}
.aichg li::before{content:"→";position:absolute;left:0;color:var(--accent2)}
.aichg.big{margin:6px 0 4px;max-height:44vh;overflow:auto;background:var(--input);border:1px solid var(--line);border-radius:12px;padding:10px 12px;gap:7px}
.aichg.big li{font-size:13.5px;color:var(--text2)}
.aiacts{display:flex;gap:4px;align-items:center;flex:none}
.ibtxt{display:inline-flex;align-items:center;gap:6px} .ibtxt svg{width:14px;height:14px}
.toolbar .hint,.chips .hint{color:var(--muted);font-size:12.5px;align-self:center}
.aidorm .aiwhat{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:20px auto 0;max-width:760px;text-align:left}
.aidorm .aiwhat div{background:color-mix(in srgb,var(--ink) 2%,transparent);border:1px solid var(--line);border-radius:12px;padding:12px 14px;font-size:13px;line-height:1.5}
.aidorm .aiwhat b{color:var(--text);display:block;margin-bottom:3px}
.cfsub{color:var(--muted);font-size:13px;margin:-6px 0 12px;word-break:break-word}
.cflbl{display:block;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:1.1px;margin:4px 0 6px}
.cfgrid{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px;margin-top:10px} .cfgrid label{display:flex;flex-direction:column} .cfgrid .field{width:100%}
.ssec{display:flex;align-items:center;gap:9px;margin:22px 0 2px;padding-top:16px;border-top:1px solid var(--line);font-weight:650;font-size:15px}
.ssec svg{width:17px;height:17px;color:var(--accent)} .ssec .pill{margin-left:auto}
.sprov{display:flex;flex-direction:column;gap:10px;padding:12px 0 14px;border-bottom:1px solid color-mix(in srgb,var(--ink) 4%,transparent)}
.sprov[hidden]{display:none} .sprov label{display:flex;flex-direction:column} .sprov .field{width:100%}
.spre{display:flex;gap:8px;align-items:center;flex-wrap:wrap} .spre .hint,.sprov .hint{color:var(--muted);font-size:12.5px}
/* v6: Email */
.emimport{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:12px 14px;margin-bottom:16px;box-shadow:0 10px 30px #00000033;position:relative;overflow:hidden}
.emimport::before{content:"";position:absolute;inset:0 0 auto 0;height:2px;background:linear-gradient(90deg,var(--accent2),var(--ok))}
.emimport .lbl{display:flex;align-items:center;gap:8px;font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;margin-bottom:9px}
.emimport .lbl svg{width:15px;height:15px;color:var(--accent2)} .emimport .hint{text-transform:none;letter-spacing:0}
.emrow{display:grid;grid-template-columns:minmax(220px,.9fr) 1.6fr;gap:12px;align-items:stretch}
.drop{border:1.5px dashed var(--line2);border-radius:12px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;padding:12px;text-align:center;color:var(--muted);font-size:12.5px;cursor:pointer;transition:.15s;background:var(--input);min-height:92px}
.drop svg{width:24px;height:24px;color:var(--accent2);margin-bottom:2px}
.drop b{color:var(--text);font-size:13.5px;font-weight:600}
.drop:hover,.drop.over{border-color:var(--accent2);background:#4fc3ff12;color:var(--text)}
.drop.over{box-shadow:0 0 0 4px #4fc3ff1f}
.empaste{display:flex;flex-direction:column;gap:8px}
.empaste textarea{flex:1;resize:none;min-height:92px;max-height:220px;line-height:1.45;font-family:inherit;font-size:13.5px;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:10px 12px;outline:none;color:inherit}
.empaste textarea:focus{border-color:var(--accent2)}
.empaste .nrow{justify-content:flex-end}
.emstat{color:var(--muted);font-size:12.5px;margin-top:8px;min-height:16px}
.emsetup{display:flex;gap:14px;align-items:flex-start;background:linear-gradient(90deg,#4fc3ff14,color-mix(in srgb,var(--accent) 5%,transparent));border:1px solid #4fc3ff40;border-radius:var(--r);padding:14px 16px;margin-bottom:14px;animation:slidein .3s ease}
.emsetup.warn{background:linear-gradient(90deg,#ffa94d1c,#ff5d7310);border-color:#ffa94d44}
.emsetup .bi{width:34px;height:34px;border-radius:10px;background:#4fc3ff22;color:var(--accent2);display:grid;place-items:center;flex:none}
.emsetup.warn .bi{background:#ffa94d26;color:var(--high)}
.emsetup .bi svg{width:18px;height:18px}
.emsetup .bt{flex:1;min-width:0;line-height:1.5} .emsetup .bt b{color:#fff;font-size:14.5px}
.emsetup .bt small{display:block;color:var(--text2);font-size:12.5px;margin-top:3px}
.emsetup ol{margin:8px 0 0;padding-left:18px;color:var(--text2);font-size:12.8px;line-height:1.6}
.emsetup .eacts{display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap}
.mail{display:flex;gap:13px;align-items:flex-start;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:12px 14px;position:relative;animation:pop .18s ease;transition:border-color .12s,background .12s}
.mail:hover{border-color:var(--line2);background:var(--panel2)}
.mail.sel{border-color:var(--accent2);box-shadow:0 0 0 1px #4fc3ff66,0 8px 22px #00000040}
.mail.unread::before{content:"";position:absolute;left:0;top:12px;bottom:12px;width:3px;border-radius:3px;background:var(--accent2)}
.mail.added{opacity:.78} .mail.dismissed{opacity:.6}
.mail.leaving{animation:leave .3s ease forwards;pointer-events:none}
.av{width:36px;height:36px;border-radius:11px;flex:none;display:grid;place-items:center;font-weight:700;font-size:13px;color:#fff;letter-spacing:.3px}
.mail .body{flex:1;min-width:0}
.mtop{display:flex;align-items:baseline;gap:10px}
.mfrom{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;min-width:0}
.mfrom small{color:var(--muted);font-weight:400;margin-left:6px;font-size:12px}
.mwhen{color:var(--muted);font-size:12px;flex:none}
.msubj{margin-top:2px;font-size:14px;color:var(--text);word-break:break-word}
.mail.unread .msubj{font-weight:600}
.mprev{margin-top:3px;color:var(--muted);font-size:12.8px;line-height:1.45;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word}
.mail .meta{margin-top:7px}
.tag.why{color:var(--code);background:#4fc3ff14} .tag.imp{color:var(--high);background:#ffa94d18} .tag.ok{color:var(--ok);background:#3ddc9718}
.macts{display:flex;gap:4px;align-items:center;flex:none;align-self:center}
.mlink{background:none;border:0;color:var(--accent2);font-size:12.5px;cursor:pointer;padding:5px 8px;border-radius:8px;display:inline-flex;gap:5px;align-items:center}
.mlink:hover{background:#4fc3ff14} .mlink svg{width:13px;height:13px}
.emview.dragging .emimport .drop{border-color:var(--accent2);background:#4fc3ff12}
/* v7: dashboard */
.dash .hero{display:flex;align-items:center;gap:16px;background:radial-gradient(ellipse at 0% 0%,color-mix(in srgb,var(--accent) 18%,transparent),transparent 60%),radial-gradient(ellipse at 100% 0%,#4fc3ff1f,transparent 55%),var(--panel);border:1px solid var(--line);border-radius:18px;padding:18px 20px;margin-bottom:16px}
.dash .brief{flex:1;min-width:0}
.dash .brief .bh{display:flex;align-items:center;gap:8px;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:1.1px;margin-bottom:8px}
.dash .brief .bh svg{width:15px;height:15px;color:var(--accent)}
.dash .brief .bh .pill{text-transform:none;letter-spacing:0}
.dash .brief .hl{font-size:17px;font-weight:650;line-height:1.4;margin-bottom:8px}
.dash .brief ul{margin:0;padding-left:18px;color:var(--text2);line-height:1.6;font-size:13.5px}
.dash .brief li{margin:2px 0}
.dash .brief .focus{margin-top:10px;display:inline-flex;gap:8px;align-items:center;background:#3ddc9714;color:var(--okt);border:1px solid #3ddc9733;padding:6px 11px;border-radius:10px;font-size:13px}
.dash .brief .bnote{color:var(--muted);font-size:12px;margin-top:8px}
.dash .bacts{display:flex;flex-direction:column;gap:8px;align-items:stretch;flex:none}
.dash .bacts .btn{display:inline-flex;gap:8px;align-items:center;justify-content:center}
.dash .bacts .btn svg{width:15px;height:15px}
.dash .shimmer{height:12px;border-radius:6px;margin:8px 0;background:linear-gradient(90deg,color-mix(in srgb,var(--ink) 3%,transparent),color-mix(in srgb,var(--ink) 9%,transparent),color-mix(in srgb,var(--ink) 3%,transparent));background-size:200% 100%;animation:shim 1.2s linear infinite}
@keyframes shim{from{background-position:200% 0}to{background-position:-200% 0}}
.dash .stats{grid-template-columns:repeat(6,1fr)}
.dash .stat{cursor:pointer;transition:transform .15s,border-color .15s}
.dash .stat:hover{transform:translateY(-2px);border-color:color-mix(in srgb,var(--ink) 13%,transparent)}
.dgrid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.dcard{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:14px 16px 12px;display:flex;flex-direction:column;min-width:0;animation:pop .25s ease}
.dcard.wide{grid-column:1/-1}
.dcard .dh{display:flex;align-items:center;gap:9px;margin-bottom:10px}
.dcard .dh .sicon svg{width:17px;height:17px}
.dcard .dh b{font-size:14.5px;font-weight:650}
.dcard .dh .count{margin-left:2px}
.dcard .dh .mlink{margin-left:auto}
.dcard .dempty{color:var(--muted);font-size:13px;padding:14px 4px 10px;display:flex;align-items:center;gap:8px}
.dcard .dempty svg{width:16px;height:16px;color:var(--ok)}
.drow{display:flex;align-items:flex-start;gap:10px;padding:8px 8px;border-radius:10px;cursor:pointer;position:relative}
.drow:hover{background:color-mix(in srgb,var(--ink) 3%,transparent)}
.drow.sel{background:color-mix(in srgb,var(--accent) 11%,transparent);box-shadow:inset 0 0 0 1px color-mix(in srgb,var(--accent) 40%,transparent)}
.drow .dck{width:18px;height:18px;border-radius:6px;border:2px solid var(--line2);flex:none;margin-top:1px;display:grid;place-items:center;background:none;cursor:pointer;padding:0}
.drow .dck:hover{border-color:var(--ok)} .drow .dck svg{width:11px;height:11px;opacity:0} .drow .dck:hover svg{opacity:.6}
.drow .dck.done{background:var(--ok);border-color:var(--ok)} .drow .dck.done svg{opacity:1}
.drow .dt{flex:1;min-width:0}
.drow .dtt{font-size:13.5px;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.drow .dm{display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;align-items:center}
.drow .dm .tag{font-size:11.5px;padding:2px 8px}
.drow .dm .badge{font-size:10.5px;padding:2px 7px;cursor:default}
.drow .dr{flex:none;color:var(--muted);font-size:12px;margin-top:1px;white-space:nowrap}
.drow .dr.over{color:var(--urgent)} .drow .dr.now{color:var(--high);font-weight:600}
.dsub{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:1.1px;margin:10px 8px 4px}
.dfoot{display:flex;gap:8px;flex-wrap:wrap;align-items:center;border-top:1px solid var(--line);margin-top:auto;padding-top:10px;color:var(--muted);font-size:12px}
.dfoot .pill{font-size:11.5px}
.dproj{padding:6px 8px 4px}
.dproj .ph{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600;cursor:pointer}
.dproj .ph .pname{flex:1}
.dproj .bar{height:5px;border-radius:5px;background:color-mix(in srgb,var(--ink) 4%,transparent);margin:6px 0 2px;overflow:hidden;display:flex}
.dproj .bar span{display:block;height:100%}
.dmail{display:flex;gap:10px;align-items:flex-start;padding:8px;border-radius:10px}
.dmail:hover{background:color-mix(in srgb,var(--ink) 3%,transparent)}
.dmail .av{width:30px;height:30px;font-size:11.5px;flex:none}
.dmail .mt{flex:1;min-width:0}
.dmail .mf{font-size:12.5px;color:var(--muted);display:flex;gap:8px} .dmail .mf span:first-child{color:var(--text2);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dmail .ms{font-size:13.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px}
.dmail .ma{display:flex;gap:4px;align-items:center;flex:none;opacity:.85}
.dmail:hover .ma{opacity:1}
.dnote{padding:8px;border-radius:10px;cursor:pointer;border-left:3px solid var(--nc,var(--accent));margin:2px 0 6px;background:color-mix(in srgb,var(--ink) 2%,transparent)}
.dnote:hover{background:color-mix(in srgb,var(--ink) 4%,transparent)}
.dnote b{font-size:13.5px;display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dnote .nx{color:var(--muted);font-size:12.5px;margin-top:3px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;white-space:pre-line}
.dnote .nw{color:var(--muted);font-size:11.5px;margin-top:4px;display:flex;gap:8px;align-items:center}
.dmeet{padding:9px 8px;border-radius:10px;cursor:pointer}
.dmeet:hover{background:color-mix(in srgb,var(--ink) 3%,transparent)}
.dmeet .mh2{display:flex;gap:8px;align-items:center} .dmeet .mh2 b{flex:1;font-size:13.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dmeet .sx{color:var(--muted);font-size:12.5px;margin-top:3px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.dmeet .nw{color:var(--muted);font-size:11.5px;margin-top:4px}
.mdl h4{margin:14px 0 6px;font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:1px}
.mdl ul{margin:0;padding-left:18px;line-height:1.55;color:var(--text2)}
.mdl .msum{line-height:1.55;color:var(--text2)}
.mdl .nb{white-space:pre-wrap;line-height:1.55;max-height:55vh;overflow:auto;background:var(--input);border:1px solid var(--line);border-radius:10px;padding:12px}
.mdl .mrow{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}
@media (max-width:1100px){.dgrid{grid-template-columns:1fr}.dash .stats{grid-template-columns:repeat(3,1fr)}}
::-webkit-scrollbar{width:10px} ::-webkit-scrollbar-thumb{background:var(--line2);border-radius:10px;border:2px solid var(--bg)}
/* v8: themes, setup wizard, get-set-up checklist, focus timer, weekly review, tour, updates */
html[data-theme="light"] .pitem.on,html[data-theme="light"] .ibtn:hover,html[data-theme="light"] .seg button.on,html[data-theme="light"] .chip.on,
html[data-theme="light"] .sbtn:hover,html[data-theme="light"] .banner .bt b,html[data-theme="light"] .emsetup .bt b{color:var(--strong)}
html[data-theme="light"] .pitem{color:var(--text2)} html[data-theme="light"] .chip{color:var(--text2)}
html[data-theme="light"] .add,html[data-theme="light"] .neditor,html[data-theme="light"] .aiask,html[data-theme="light"] .emimport{box-shadow:0 6px 20px #1b203014}
html[data-theme="light"] .modal{box-shadow:0 30px 90px #1b203040} html[data-theme="light"] .modal-back{background:#1b203066}
html[data-theme="light"] .snack{box-shadow:0 14px 40px #1b203033} html[data-theme="light"] mark{background:#ffd43b80}
html[data-theme="light"] .task.sel{box-shadow:0 0 0 1px color-mix(in srgb,var(--accent) 45%,transparent),0 8px 22px #1b203018}
html[data-theme="light"] .dash .stat:hover{border-color:var(--line2)}
html[data-theme="light"] .b-normal{color:#2f6fd6} html[data-theme="light"] .b-high,html[data-theme="light"] .tag.today{color:#c46a00} html[data-theme="light"] .b-urgent,html[data-theme="light"] .tag.over{color:#d62845}
html[data-theme="light"] ::-webkit-scrollbar-thumb{background:#cdd2dd}
.wiz{position:fixed;inset:0;z-index:80;display:grid;place-items:center;background:radial-gradient(ellipse at 15% 0%,color-mix(in srgb,var(--accent) 22%,transparent),transparent 55%),radial-gradient(ellipse at 100% 100%,color-mix(in srgb,var(--accent2) 16%,transparent),transparent 55%),var(--bg);animation:fadein .2s ease}
.wiz[hidden]{display:none}
.wbox{width:min(720px,94vw);max-height:92vh;overflow:auto;background:var(--modal);border:1px solid var(--line);border-radius:22px;padding:26px 30px 22px;box-shadow:0 30px 90px #0007}
.wtop{display:flex;align-items:center;gap:12px;margin-bottom:18px}
.wtop .logo{width:36px;height:36px} .wtop b{font-size:16px} .wtop small{display:block;color:var(--muted);font-size:12px}
.wdots{margin-left:auto;display:flex;gap:6px} .wdots i{width:8px;height:8px;border-radius:8px;background:var(--line2);transition:.2s} .wdots i.on{width:22px;background:var(--accent)} .wdots i.done{background:color-mix(in srgb,var(--accent) 55%,transparent)}
.wstep{animation:viewin .25s ease;min-height:270px}
.wstep h2{margin:6px 0 6px;font-size:24px;letter-spacing:-.3px} .wstep p{color:var(--text2);line-height:1.6;margin:0 0 14px}
.wstep .big{font-size:15px;padding:12px 14px;width:100%;background:var(--input);border:1px solid var(--line);border-radius:12px;outline:none;color:var(--text)}
.wstep .big:focus{border-color:var(--accent)}
.wfeat{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0}
.wfeat div{background:color-mix(in srgb,var(--ink) 3%,transparent);border:1px solid var(--line);border-radius:12px;padding:11px 12px;font-size:12.8px;color:var(--text2);line-height:1.45} .wfeat b{display:block;color:var(--text);margin-bottom:2px}
.wnav{display:flex;align-items:center;gap:10px;margin-top:18px;padding-top:14px;border-top:1px solid var(--line)}
.wnav .hint{color:var(--muted);font-size:12px} .wnav .sp{flex:1}
.themes{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:8px 0 16px}
.thc{border:2px solid var(--line);border-radius:14px;padding:10px;cursor:pointer;background:none;text-align:left;color:var(--text)} .thc.on{border-color:var(--accent);box-shadow:0 0 0 3px color-mix(in srgb,var(--accent) 22%,transparent)}
.thc .pv{height:84px;border-radius:9px;display:grid;grid-template-columns:30% 1fr;overflow:hidden;margin-bottom:8px;border:1px solid #0002}
.thc .pv i{display:block} .thc .pv .pm{padding:10px;display:flex;flex-direction:column;gap:6px} .thc .pv .pm span{height:9px;border-radius:5px;display:block}
.swatches{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.swatch{width:34px;height:34px;border-radius:50%;border:3px solid transparent;cursor:pointer;position:relative;box-shadow:inset 0 0 0 2px #0002}
.swatch.on{border-color:var(--text)} .swatch small{position:absolute;bottom:-17px;left:50%;transform:translateX(-50%);font-size:10px;color:var(--muted)}
.wopt{display:flex;gap:12px;align-items:flex-start;border:1px solid var(--line);border-radius:14px;padding:13px 14px;margin:8px 0;background:color-mix(in srgb,var(--ink) 2%,transparent)}
.wopt .bi{width:34px;height:34px;border-radius:10px;display:grid;place-items:center;flex:none;background:color-mix(in srgb,var(--accent2) 16%,transparent);color:var(--accent2)} .wopt .bi svg{width:18px;height:18px}
.wopt .bt{flex:1;min-width:0;line-height:1.5;font-size:13px;color:var(--text2)} .wopt .bt b{color:var(--text);font-size:14px}
.wopt code{font-family:"Cascadia Code",Consolas,monospace;font-size:12px;background:var(--input);border:1px solid var(--line);border-radius:6px;padding:1px 6px;color:var(--code)}
.wrow{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:8px} .wrow .field{flex:1;min-width:180px}
.wmsg{font-size:12.5px;color:var(--muted);margin-top:8px;min-height:16px} .wmsg.ok{color:var(--ok)} .wmsg.bad{color:var(--urgent)}
.setup{background:linear-gradient(120deg,color-mix(in srgb,var(--accent) 14%,transparent),color-mix(in srgb,var(--accent2) 8%,transparent)),var(--panel);border:1px solid color-mix(in srgb,var(--accent) 35%,var(--line));border-radius:18px;padding:16px 18px 12px;margin-bottom:16px;animation:pop .25s ease}
.setup .sh{display:flex;align-items:center;gap:10px} .setup .sh b{font-size:15.5px} .setup .sh .pill{margin-left:6px}
.setup .sh .ibtn{margin-left:auto}
.setup .pbar{height:6px;border-radius:6px;background:color-mix(in srgb,var(--ink) 8%,transparent);margin:10px 0 8px;overflow:hidden} .setup .pbar span{display:block;height:100%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:6px;transition:width .4s}
.sgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:4px 14px}
.sitem{display:flex;align-items:center;gap:10px;padding:7px 6px;border-radius:10px;font-size:13.3px}
.sitem .ck{width:20px;height:20px;border-radius:50%;border:2px solid var(--line2);display:grid;place-items:center;flex:none}
.sitem.done .ck{background:var(--ok);border-color:var(--ok)} .sitem .ck svg{width:12px;height:12px;opacity:0} .sitem.done .ck svg{opacity:1}
.sitem.done .tx{color:var(--muted);text-decoration:line-through} .sitem .tx{flex:1}
.sitem .mlink{padding:3px 8px}
.tag.rep{color:var(--accent);background:color-mix(in srgb,var(--accent) 12%,transparent);cursor:pointer} .tag.fm{color:var(--ok);background:color-mix(in srgb,var(--ok) 12%,transparent)}
.add .field.rep{width:auto}
.drow .dfocus{opacity:0;transition:.15s;flex:none;margin-top:-2px} .drow:hover .dfocus,.drow.sel .dfocus{opacity:1}
.fpill{position:fixed;left:290px;bottom:18px;z-index:55;display:flex;align-items:center;gap:12px;background:var(--pop);border:1px solid var(--line2);border-radius:16px;padding:9px 10px 9px 14px;box-shadow:0 14px 40px #0006;animation:slidein .25s ease;max-width:min(520px,60vw)}
.fpill[hidden]{display:none}
.fpill .ring{width:38px;height:38px;flex:none} .fpill .ring circle{fill:none;stroke-width:4}
.fpill .ft{min-width:0} .fpill .ft b{display:block;font-size:20px;font-variant-numeric:tabular-nums;letter-spacing:.5px} .fpill .ft small{display:block;color:var(--muted);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:300px}
.fpill.brk b{color:var(--ok)} .fpill.paused b{opacity:.55}
.rv.dash .stats{grid-template-columns:repeat(4,1fr)}
.rv .rgrid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.rv .wk{display:flex;align-items:center;gap:8px}
.rv .hbar{display:flex;align-items:center;gap:10px;padding:6px 8px;font-size:13px} .rv .hbar .nm{width:150px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:flex;gap:8px;align-items:center}
.rv .hbar .tr{flex:1;height:8px;border-radius:8px;background:color-mix(in srgb,var(--ink) 7%,transparent);overflow:hidden} .rv .hbar .tr span{display:block;height:100%;border-radius:8px}
.rv .hbar .v{width:64px;text-align:right;color:var(--muted);font-variant-numeric:tabular-nums}
.recap h3{margin:0 0 8px;font-size:17px;line-height:1.4} .recap h4{margin:12px 0 4px;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:1.1px}
.recap ul{margin:0;padding-left:18px;line-height:1.6;color:var(--text2);font-size:13.5px}
.tourbox{width:min(560px,92vw)} .tourart{width:64px;height:64px;border-radius:18px;display:grid;place-items:center;background:linear-gradient(135deg,var(--accent),var(--accent2));margin-bottom:12px;box-shadow:0 14px 34px color-mix(in srgb,var(--accent) 30%,transparent)} .tourart svg{width:30px;height:30px}
.tourbox h2{margin:0 0 6px} .tourbox p{color:var(--text2);line-height:1.6;margin:0 0 10px}
.ubanner{display:flex;gap:14px;align-items:flex-start;background:linear-gradient(90deg,color-mix(in srgb,var(--ok) 14%,transparent),color-mix(in srgb,var(--accent2) 8%,transparent));border:1px solid color-mix(in srgb,var(--ok) 40%,transparent);border-radius:var(--r);padding:13px 16px;margin-bottom:14px;animation:slidein .3s ease}
.ubanner .bi{width:34px;height:34px;border-radius:10px;background:color-mix(in srgb,var(--ok) 20%,transparent);color:var(--ok);display:grid;place-items:center;flex:none} .ubanner .bi svg{width:18px;height:18px}
.ubanner .bt{flex:1;min-width:0;line-height:1.5} .ubanner .bt b{font-size:14.5px}
.ubanner pre{white-space:pre-wrap;font:12.5px/1.5 "Segoe UI",system-ui,sans-serif;color:var(--text2);max-height:160px;overflow:auto;margin:6px 0 0;background:color-mix(in srgb,var(--ink) 3%,transparent);border-radius:10px;padding:8px 10px}
.ubanner .eacts{display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap}
.busyover{position:fixed;inset:0;z-index:90;background:#05070bcc;display:grid;place-items:center;color:#fff;font-size:16px}
.busyover[hidden]{display:none}
</style></head>
<body>
<div class="app">
  <aside class="side">
    <div class="brand">
      <div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.2 4.2L19 7"/></svg></div>
      <div><b>Work To-Do</b><small>Projects &amp; priorities · v__WTD_VERSION__</small></div>
    </div>
    <div class="sect">Overview</div>
    <div id="dashItem"></div>
    <div id="allItem"></div>
    <div id="smartItems"></div>
    <div id="notesItem"></div>
    <div id="reviewItem"></div>
    <div id="aiItem"></div>
    <div id="emailItem"></div>
    <div class="sect">Projects</div>
    <div class="plist" id="plist"></div>
    <form class="newp" id="newp">
      <input id="newpName" placeholder="New project…" autocomplete="off">
      <button class="btn sq" title="Add project (Shift+P)">+</button>
    </form>
    <div class="sfoot">
      <button class="sbtn" id="kbBtn" type="button" title="Keyboard shortcuts (?)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="2.5" y="6" width="19" height="12" rx="2.5"/><path d="M6.5 10h.01M10 10h.01M13.5 10h.01M17 10h.01M7.5 14h9"/></svg><span>Keyboard shortcuts</span><kbd>?</kbd></button>
      <button class="sbtn" id="bkBtn" type="button" title="Back up everything to one file (Ctrl+Shift+E)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg><span>Back up</span><kbd>Ctrl+Shift+E</kbd></button>
      <button class="sbtn" id="setBtn" type="button" title="Settings (Ctrl+,)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/></svg><span>Reminders &amp; settings</span><span class="pill" id="remPill" style="flex:none">On</span></button>
    </div>
  </aside>
  <main class="main">
    <div class="head">
      <div><h1 id="h1"></h1><div class="sub" id="sub"></div></div>
      <label class="search"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg><input id="q" placeholder="Search tasks"><kbd>/</kbd></label>
    </div>
    <div id="taskView">
    <div class="banner" id="banner" hidden></div>
    <div class="stats" id="stats"></div>
    <form class="add" id="addForm">
      <textarea class="title" id="tTitle" rows="1" placeholder="What needs to get done?   Shift+Enter adds a new line"></textarea>
      <div class="seg" id="tPrio"></div>
      <input type="date" class="field" id="tDue" title="Due date (optional)">
      <input type="time" class="field time" id="tTime" title="Due time (optional). You get a reminder before it.">
      <select class="field" id="tProj" title="Project"></select>
      <select class="field rep" id="tRep" title="Repeat (optional)"><option value="">Doesn't repeat</option><option value="daily">Every day</option><option value="weekdays">Weekdays</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option></select>
      <select class="field rep" id="tRepDay" title="Which day" hidden></select>
      <button class="btn" style="margin-left:auto">Add task</button>
    </form>
    <div class="toolbar">
      <div class="chips" id="filters"></div>
      <div class="toggle" id="doneToggle"><span>Show completed</span><div class="sw"></div></div>
    </div>
    <div class="list" id="list"></div>
    </div>
    <div id="noteView" hidden>
      <form class="neditor" id="noteForm">
        <div class="lbl" id="noteLbl">New note</div>
        <input class="ninput" id="nTitle" placeholder="Title" autocomplete="off">
        <textarea class="nbody" id="nBody" placeholder="Write anything you want to keep handy. Line breaks are kept. Wrap commands in `backticks` to show them as code."></textarea>
        <div class="nrow">
          <select class="field" id="nProj" title="Link to a project (optional)"></select>
          <span class="hint"><kbd>Ctrl+Enter</kbd> save · <kbd>Esc</kbd> cancel</span>
          <span style="margin-left:auto;display:flex;gap:8px">
            <button type="button" class="btn ghost" id="nCancel" hidden>Cancel</button>
            <button class="btn" id="nSave">Add note</button>
          </span>
        </div>
      </form>
      <div class="toolbar"><div class="chips" id="nfilters"></div></div>
      <div class="ngrid" id="nlist"></div>
    </div>
    <div id="dashView" class="dash" hidden></div>
    <div id="aiView" hidden><div id="aiAsk"></div><div id="aiBody"></div></div>
    <div id="emailView" class="emview" hidden></div>
    <div id="reviewView" hidden></div>
  </main>
</div>
<div class="toast" id="toast">Saved</div>
<div class="snack" id="snack"><span id="snackMsg"></span><button type="button" id="snackUndo">Undo <kbd style="margin-left:4px">Ctrl+Z</kbd></button></div>
<div class="modal-back" id="modalBack" hidden><div class="modal" id="modal" role="dialog" aria-modal="true"></div></div>
<script>
const PRI=["urgent","high","normal"], RANK={urgent:0,high:1,normal:2};
const PCOL={urgent:"var(--urgent)",high:"var(--high)",normal:"var(--normal)"};
const PALETTE=["#7c6cff","#4fc3ff","#3ddc97","#ffa94d","#ff6fb5","#ffd43b","#20c997","#a78bfa","#f97316","#60a5fa"];
const ICON={
  edit:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 20h4L19 9l-4-4L4 16z"/></svg>',
  trash:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/></svg>',
  check:'<svg viewBox="0 0 24 24" fill="none" stroke="#0d1017" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.2 4.2L19 7"/></svg>',
  note:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/></svg>',
  copy:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a1 1 0 0 1 1-1h10"/></svg>',
  cal:'<svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4"/></svg>',
  clock:'<svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></svg>',
  sun:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M4.6 4.6l1.4 1.4M18 18l1.4 1.4M2.5 12h2M19.5 12h2M4.6 19.4L6 18M18 6l1.4-1.4"/></svg>',
  up:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="5" width="17" height="15" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4M10 15l2-2 2 2"/></svg>',
  alarm:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="13" r="7.5"/><path d="M12 9.5V13l2.2 1.6M4 4.5 6.5 2.5M20 4.5l-2.5-2M12 2.5v1"/></svg>',
  bell:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/></svg>',
  x:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  folder:'<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2.5h8a2 2 0 0 1 2 2V18a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M12 11v5M9.5 13.5h5"/></svg>',
  done:'<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.2 4.2L19 7"/></svg>',
  search:'<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round"><circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4-4"/></svg>',
  pen:'<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/></svg>'
};
const SMART={today:{name:"Today",icon:ICON.sun,color:"var(--high)",desc:"Due today, plus anything overdue"},
  upcoming:{name:"Upcoming",icon:ICON.up,color:"var(--accent2)",desc:"Everything with a future due date"},
  overdue:{name:"Overdue",icon:ICON.alarm,color:"var(--urgent)",desc:"Past its due date or time"}};
const isSmart=c=>Object.prototype.hasOwnProperty.call(SMART,c);
const pad2=n=>String(n).padStart(2,"0");
const nowHM=()=>{const d=new Date();return pad2(d.getHours())+":"+pad2(d.getMinutes());};
const nowIso=()=>{const d=new Date();return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate())+"T"+pad2(d.getHours())+":"+pad2(d.getMinutes())+":"+pad2(d.getSeconds());};
const addDays=(n)=>{const d=new Date();d.setDate(d.getDate()+n);return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate());};
const fmtTime=hm=>{if(!hm)return "";const [h,m]=hm.split(":").map(Number);const d=new Date();d.setHours(h,m,0,0);return d.toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"});};
function isOverdue(t){if(t.done||!t.due)return false;const td=todayStr();return t.due<td||(t.due===td&&!!t.due_time&&t.due_time<=nowHM());}
const isDueToday=t=>!t.done&&t.due===todayStr();
const DEF_SETTINGS={reminders:true,morning:"09:00",lead:15};
const settings=()=>Object.assign({},DEF_SETTINGS,data.settings||{});
let selTask=null, lastUndo=null, gPending=0, lastViewKey="", bannerHidden=sessionStorage.getItem("bannerHidden")||"";
let data={projects:[],tasks:[],notes:[]}, view=location.hash==="#notes"?"notes":location.hash==="#ai"?"ai":location.hash==="#email"?"email":(location.hash===""||location.hash==="#"||location.hash==="#dash")?"dash":"tasks", noteFilter="all", noteQuery="", editingNote=null, current="all", filter="all", showDone=false, newPrio="normal", query="", editingProject=null;
const $=s=>document.querySelector(s);
const el=(tag,cls,html)=>{const e=document.createElement(tag); if(cls)e.className=cls; if(html!=null)e.innerHTML=html; return e;};
const uid=()=>([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g,c=>(c^crypto.getRandomValues(new Uint8Array(1))[0]&15>>c/4).toString(16)).replace(/-/g,"");
const todayStr=()=>{const d=new Date();return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");};
const fmtDue=s=>{const [y,m,d]=s.split("-").map(Number);const dt=new Date(y,m-1,d);const t=new Date(todayStr()+"T00:00");
  const diff=Math.round((dt-t)/864e5); if(diff===0)return "Today"; if(diff===1)return "Tomorrow"; if(diff===-1)return "Yesterday";
  return dt.toLocaleDateString(undefined,{month:"short",day:"numeric",year:dt.getFullYear()!==t.getFullYear()?"numeric":undefined});};
const proj=id=>data.projects.find(p=>p.id===id);
const sortedProjects=()=>[...data.projects].sort((a,b)=>a.name.localeCompare(b.name,undefined,{sensitivity:"base"}));

let saveTimer=null;
function save(){clearTimeout(saveTimer); savePending=true; saveTimer=setTimeout(async()=>{
  try{const r=await fetch("/api/data",{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});
    if(r.ok&&r.headers.get("X-Rev"))lastRev=r.headers.get("X-Rev"); toast(r.ok?"Saved":"Couldn't save");}catch(e){toast("Couldn't save. Is the app still running?");} savePending=false;},150);}
let toastTimer=null;
function toast(msg){const t=$("#toast");t.textContent=msg;t.classList.add("show");clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.remove("show"),1200);}
function commit(){save();render();}

function ensureColors(){let i=0;for(const p of data.projects){if(!p.color){p.color=PALETTE[i%PALETTE.length];}i++;}}

function renderSidebar(){
  const openAll=data.tasks.filter(t=>!t.done).length;
  const all=$("#allItem"); all.innerHTML="";
  const a=el("div","pitem"+(view==="tasks"&&current==="all"?" on":""));
  a.append(el("span","dot",""),el("span","pname","All projects"),el("span","count",String(openAll)));
  a.querySelector(".dot").style.background="linear-gradient(135deg,var(--accent),var(--accent2))";
  a.onclick=()=>go("all"); all.append(a);
  const sm=$("#smartItems"); sm.innerHTML="";
  const od=data.tasks.filter(isOverdue).length, td=data.tasks.filter(t=>!t.done&&t.due&&t.due<=todayStr()).length, up=data.tasks.filter(t=>!t.done&&t.due&&t.due>todayStr()).length;
  for(const [k,n,cls] of [["today",td,od?"amber":""],["upcoming",up,""],["overdue",od,od?"red":""]]){
    const s=el("div","pitem"+(view==="tasks"&&current===k?" on":"")); const ic=el("span","sicon",SMART[k].icon); ic.style.color=SMART[k].color;
    s.append(ic,el("span","pname",SMART[k].name),el("span","count"+(cls&&n?" "+cls:""),String(n))); s.onclick=()=>go(k); sm.append(s);}
  const rp=$("#remPill"), on=settings().reminders; rp.textContent=on?"On":"Off"; rp.className="pill"+(on?" ok":"");
  document.title=(od+td? "("+td+") ":"")+"Work To-Do";
  const ni=$("#notesItem"); ni.innerHTML="";
  const nn=el("div","pitem"+(view==="notes"?" on":""));
  nn.append(el("span","nicon",ICON.note),el("span","pname","Notes"),el("span","count",String(data.notes.length)));
  nn.onclick=()=>go("notes"); ni.append(nn);
  const ri=$("#reviewItem"); ri.innerHTML=""; const rr=el("div","pitem"+(view==="review"?" on":"")); const ric=el("span","sicon",ICON.review); ric.style.color="var(--accent)";
  const [wS,wE]=weekRange(0); rr.append(ric,el("span","pname","Weekly review"),el("span","count",String(data.tasks.filter(t=>t.done&&(t.completed||"").slice(0,10)>=wS&&(t.completed||"").slice(0,10)<=wE).length)));
  rr.title="Weekly review (G then W)"; rr.onclick=()=>go("review"); ri.append(rr); renderAIItem(); renderEmailItem(); renderDashItem();
  const list=$("#plist"); list.innerHTML="";
  if(!data.projects.length){list.append(el("div","sect","No projects yet. Add one below."));}
  for(const p of sortedProjects()){
    const n=data.tasks.filter(t=>t.project_id===p.id&&!t.done).length;
    const it=el("div","pitem"+(view==="tasks"&&current===p.id?" on":""));
    const dot=el("span","dot"); dot.style.background=p.color;
    it.append(dot);
    if(editingProject===p.id){
      const inp=el("input","pedit"); inp.value=p.name; it.append(inp);
      const done=(ok)=>{if(editingProject!==p.id)return; editingProject=null; const v=inp.value.trim();
        if(ok&&v&&!data.projects.some(q=>q.id!==p.id&&q.name.toLowerCase()===v.toLowerCase())){p.name=v;commit();}else render();};
      inp.onkeydown=e=>{if(e.key==="Enter")done(true); if(e.key==="Escape")done(false);};
      inp.onblur=()=>done(true); setTimeout(()=>{inp.focus();inp.select();});
    } else {
      const nm=el("span","pname"); nm.textContent=p.name; it.append(nm, el("span","count",String(n)));
      const tools=el("span","tools");
      const eb=el("button","ibtn",ICON.edit); eb.title="Rename"; eb.onclick=e=>{e.stopPropagation();editingProject=p.id;render();};
      const db=el("button","ibtn",ICON.trash); db.title="Delete project"; db.onclick=e=>{e.stopPropagation();
        const c=data.tasks.filter(t=>t.project_id===p.id).length, nc=data.notes.filter(x=>x.project_id===p.id).length;
        if(confirm(`Delete "${p.name}" and its ${c} task${c===1?"":"s"}? This can't be undone.`+(nc?`\n\nIts ${nc} note${nc===1?"":"s"} will be kept and just unlinked from the project.`:""))){
          data.tasks=data.tasks.filter(t=>t.project_id!==p.id); data.projects=data.projects.filter(q=>q.id!==p.id);
          for(const x of data.notes){if(x.project_id===p.id)x.project_id=null;} if(noteFilter===p.id)noteFilter="all";
          if(current===p.id)current="all"; commit();}};
      tools.append(eb,db); it.append(tools);
      it.onclick=()=>go(p.id); it.ondblclick=()=>{editingProject=p.id;render();};
    }
    list.append(it);
  }
}

function scopeTasks(){const td=todayStr();
  if(current==="all")return data.tasks;
  if(current==="today")return data.tasks.filter(t=>t.due&&(t.done?t.due===td:t.due<=td));
  if(current==="overdue")return data.tasks.filter(isOverdue);
  if(current==="upcoming")return data.tasks.filter(t=>t.due&&t.due>td);
  return data.tasks.filter(t=>t.project_id===current);}
function go(k){if(k==="notes"||k==="ai"||k==="email"||k==="dash"||k==="review"){view=k;}else{view="tasks";current=k;} selTask=null; render(); $(".main").scrollTo({top:0});
  if(k==="ai"&&aiOn()&&aiCfg().provider==="rules"&&ai.recs)loadRecs(true);
  if(k==="email"){if(emCfg().auto&&Date.now()-em.lastAuto>600000)refreshEmail(false); else loadInbox();}
  if(k==="dash")dashOpen();
  if(k==="review")loadRecap(false);}

function renderHeader(){
  const h=$("#h1"); h.innerHTML="";
  if(current==="all"){h.textContent="All projects";}
  else if(isSmart(current)){const ic=el("span","sicon",SMART[current].icon); ic.style.cssText="width:28px;height:28px;color:"+SMART[current].color; ic.querySelector("svg").style.cssText="width:26px;height:26px"; h.append(ic,document.createTextNode(SMART[current].name));}
  else{const p=proj(current); const d=el("span","dot"); d.style.background=p.color; h.append(d,document.createTextNode(p.name));}
  const sub=$("#sub"); sub.textContent=new Date().toLocaleDateString(undefined,{weekday:"long",month:"long",day:"numeric"})+(isSmart(current)?" · "+SMART[current].desc:"");
  const nc=isSmart(current)?0:current==="all"?data.notes.length:data.notes.filter(x=>x.project_id===current).length;
  if(nc){const b=el("button","nlink",ICON.note); b.type="button"; b.append(document.createTextNode(nc+" note"+(nc===1?"":"s"))); b.title="Open notes";
    b.onclick=()=>{view="notes"; noteFilter=current; render();}; sub.append(b);}
  const ts=scopeTasks(), today=todayStr();
  const open=ts.filter(t=>!t.done);
  const weekAgo=new Date(Date.now()-7*864e5).toISOString().slice(0,10);
  const cards=[["Open tasks",open.length,"var(--accent)"],["Urgent",open.filter(t=>t.priority==="urgent").length,"var(--urgent)"],
    ["Overdue or due today",open.filter(t=>t.due&&t.due<=today).length,"var(--high)"],
    ["Done this week",ts.filter(t=>t.done&&(t.completed||"").slice(0,10)>=weekAgo).length,"var(--ok)"]];
  const s=$("#stats"); s.innerHTML="";
  for(const [l,n,c] of cards){const d=el("div","stat"); d.style.setProperty("--c",c); d.append(el("div","n",String(n)),el("div","l",l)); d.querySelector(".n").style.color=c; s.append(d);}
}

function renderAddForm(){
  const seg=$("#tPrio"); seg.innerHTML="";
  for(const p of ["normal","high","urgent"]){const b=el("button",newPrio===p?"on":""); b.type="button";
    const d=el("span","dot"); d.style.background=PCOL[p]; b.append(d,document.createTextNode(p[0].toUpperCase()+p.slice(1)));
    b.onclick=()=>{newPrio=p;renderAddForm();}; seg.append(b);}
  const sel=$("#tProj"); const prev=sel.value; sel.innerHTML="";
  for(const p of sortedProjects()){const o=el("option"); o.value=p.id; o.textContent=p.name; sel.append(o);}
  if(!data.projects.length){const o=el("option"); o.value=""; o.textContent="Add a project first"; sel.append(o);}
  const pick=current!=="all"&&!isSmart(current);
  sel.value=pick?current:(proj(prev)?prev:(sel.options[0]||{}).value||"");
  sel.style.display=pick?"none":"";
  $("#tDue").placeholder=""; $("#tTitle").placeholder=current==="today"?"What's due today?   Shift+Enter adds a new line":current==="upcoming"?"Add something for tomorrow or later   (Shift+Enter = new line)":"What needs to get done?   Shift+Enter adds a new line";
}

function renderFilters(){
  const f=$("#filters"); f.innerHTML="";
  for(const [k,label] of [["all","All"],["urgent","Urgent"],["high","High"],["normal","Normal"]]){
    const c=el("button","chip"+(filter===k?" on":"")); c.type="button";
    if(k!=="all"){const d=el("span","dot"); d.style.background=PCOL[k]; c.append(d);}
    c.append(document.createTextNode(label)); c.onclick=()=>{filter=k;render();}; f.append(c);}
  $("#doneToggle").classList.toggle("on",showDone);
}

function taskCard(t){
  const p=proj(t.project_id)||{name:"(no project)",color:"#555"};
  const c=el("div","task"+(t.done?" done":"")+(selTask===t.id?" sel":"")); c.dataset.id=t.id; c.style.setProperty("--pc",PCOL[t.priority]||PCOL.normal);
  c.addEventListener("mousedown",e=>{if(!e.target.closest("button,input,label")){selTask=t.id; document.querySelectorAll("#list .task.sel").forEach(x=>x.classList.remove("sel")); c.classList.add("sel");}});
  const ck=el("button","check",ICON.check); ck.title=t.done?"Mark not done":"Mark done";
  ck.onclick=()=>toggleDone(t,c);
  const body=el("div","body");
  const ttl=el("div","ttl"); ttl.textContent=t.title; ttl.title="Click to edit";
  ttl.contentEditable="plaintext-only"; ttl.spellcheck=true;
  ttl.onkeydown=e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();ttl.blur();} if(e.key==="Escape"){ttl.textContent=t.title;ttl.blur();}};
  ttl.onblur=()=>{const v=ttl.innerText.replace(/\r/g,"").replace(/\n{3,}/g,"\n\n").trim(); if(v&&v!==t.title){t.title=v;commit();} else ttl.textContent=t.title;};
  const meta=el("div","meta");
  const b=el("button","badge b-"+t.priority); b.textContent=t.priority; b.title="Click to change priority";
  b.onclick=()=>{t.priority=PRI[(PRI.indexOf(t.priority)+2)%3]; commit();};
  meta.append(b);
  if(current==="all"||isSmart(current)){const tg=el("span","tag"); const d=el("span","dot"); d.style.background=p.color; tg.append(d,document.createTextNode(p.name)); meta.append(tg);}
  const today=todayStr();
  const od=isOverdue(t);
  const due=el("label","tag due"+(od?" over":(!t.done&&t.due===today?" today":""))+(t.due?"":" ghost"));
  due.innerHTML=ICON.cal; due.append(document.createTextNode(t.due?(od?"Overdue · ":"")+fmtDue(t.due):"Add due date"));
  const di=el("input"); di.type="date"; di.value=t.due||""; di.title="Change due date";
  di.onchange=()=>{t.due=di.value||""; if(!t.due)t.due_time=""; commit();}; due.append(di); meta.append(due);
  if(t.due){const tm=el("label","tag due time"+(t.due_time?(od?" over":(!t.done&&t.due===today?" today":"")):" ghost"));
    tm.innerHTML=ICON.clock; tm.append(document.createTextNode(t.due_time?fmtTime(t.due_time):"Add time"));
    tm.title=t.due_time?"Change due time (you get a reminder before it)":"Add a due time for a reminder";
    const ti=el("input"); ti.type="time"; ti.value=t.due_time||""; ti.onchange=()=>{t.due_time=ti.value||""; commit();}; tm.append(ti); meta.append(tm);
    const clr=el("button","ibtn",ICON.x); clr.title="Clear due date and time"; clr.style.padding="2px"; clr.onclick=()=>{t.due="";t.due_time="";commit();}; meta.append(clr);}
  if(t.repeat){const rp=el("span","tag rep",ICON.repeat.replace("<svg ",'<svg width="12" height="12" ')); rp.append(document.createTextNode(repLabel(t.repeat))); rp.title="Repeats (Y to change)"; rp.onclick=()=>openRepeat(t); meta.append(rp);}
  if(+t.focus_min){const fm=el("span","tag fm",ICON.timer.replace("<svg ",'<svg width="12" height="12" ')); fm.append(document.createTextNode(fmtMin(+t.focus_min))); fm.title="Focused time on this task"; meta.append(fm);}
  body.append(ttl,meta);
  const acts=el("div","acts");
  if(!t.done){const fb=el("button","ibtn",ICON.timer); fb.title="Start a focus timer (F)"; fb.onclick=()=>startFocus(t); acts.append(fb);
    const rb=el("button","ibtn",ICON.repeat); rb.title="Repeat (Y)"; rb.onclick=()=>openRepeat(t); acts.append(rb);}
  const del=el("button","ibtn",ICON.trash); del.title="Delete task";
  del.onclick=()=>deleteTask(t);
  acts.append(del);
  c.append(ck,body,acts); return c;
}

function renderList(){
  const list=$("#list"); list.innerHTML="";
  const q=query.trim().toLowerCase();
  let ts=scopeTasks().filter(t=>(showDone||!t.done)&&(filter==="all"||t.priority===filter)&&(!q||t.title.toLowerCase().includes(q)||(proj(t.project_id)||{name:""}).name.toLowerCase().includes(q)));
  const dk=t=>(t.due||"9999-99-99")+" "+(t.due_time||"99:99");
  const byDate=current==="today"||current==="upcoming";
  ts.sort(byDate?(a,b)=>(a.done-b.done)||(dk(a)<dk(b)?-1:dk(a)>dk(b)?1:0)||(RANK[a.priority]-RANK[b.priority])||(a.created<b.created?-1:1)
    :(a,b)=>(a.done-b.done)||(RANK[a.priority]-RANK[b.priority])||(dk(a)<dk(b)?-1:dk(a)>dk(b)?1:0)||(a.created<b.created?-1:1));
  if(!data.projects.length){list.append(emptyState("folder","Start with a project","Add one on the left, like a client, a course, or an assignment you've been given.",[["New project",()=>$("#newpName").focus(),"Shift+P"]]));return;}
  if(!ts.length){
    if(q||filter!=="all"){list.append(emptyState("search","Nothing matches","Try a different word or priority.",[["Clear search & filter",()=>{query="";filter="all";render();},"Esc"]],"#4fc3ff","#7c6cff"));return;}
    const E={today:["done","Nothing due today","Enjoy it, or pull something forward from Upcoming."],overdue:["done","Nothing overdue","You're on top of everything."],
      upcoming:["up","Nothing scheduled yet","Give a task a due date and it shows up here, with a reminder."]}[current]||["done","All clear","Add a task above whenever something new comes in."];
    list.append(emptyState(E[0],E[1],E[2],[["Add a task",()=>$("#tTitle").focus(),"Ctrl+N"]],"#3ddc97","#4fc3ff"));return;}
  let last=null; const td=todayStr();
  const groupOf=t=>{if(t.done)return "Completed";
    if(current==="today")return isOverdue(t)?"Overdue":"Due today";
    if(current==="upcoming"){const diff=Math.round((new Date(t.due+"T00:00")-new Date(td+"T00:00"))/864e5);
      return diff===1?"Tomorrow":diff<7?new Date(t.due+"T00:00").toLocaleDateString(undefined,{weekday:"long",month:"short",day:"numeric"}):fmtDue(t.due);}
    return t.priority[0].toUpperCase()+t.priority.slice(1);};
  for(const t of ts){const g=groupOf(t);
    if(g!==last){list.append(el("div","group",g)); last=g;} list.append(taskCard(t));}
}

function emptyState(art,title,text,actions,ca,cb){const e=el("div","empty");
  const a=el("div","art",ICON[art]||ICON.done); if(ca){a.style.setProperty("--ea",ca);a.style.setProperty("--eb",cb);} e.append(a);
  const h=el("h3"); h.textContent=title; const p=el("div"); p.textContent=text; e.append(h,p);
  if(actions&&actions.length){const r=el("div","eacts"); for(const [label,fn,key] of actions){const b=el("button","btn sm"); b.type="button"; b.textContent=label; b.onclick=fn; r.append(b); if(key)r.append(keysHtml(key));} e.append(r);}
  return e;}
function renderBanner(){const b=$("#banner");
  const od=data.tasks.filter(isOverdue), dt=data.tasks.filter(t=>isDueToday(t)&&!isOverdue(t));
  const key=todayStr()+"/"+od.length+"/"+dt.length;
  if(!(od.length||dt.length)||current==="today"||current==="overdue"||bannerHidden===key){b.hidden=true;return;}
  b.hidden=false; b.innerHTML="";
  const parts=[]; if(od.length)parts.push(`<b>${od.length} overdue</b>`); if(dt.length)parts.push(`<b>${dt.length} due today</b>`);
  const s=settings(), next=dt.filter(t=>t.due_time).sort((a,c)=>a.due_time<c.due_time?-1:1)[0];
  const sm=s.reminders?(next?`Next: “${esc(next.title.split("\n")[0].slice(0,50))}” at ${fmtTime(next.due_time)} · desktop reminders are on`:`Desktop reminders are on · morning summary at ${fmtTime(s.morning)}`):"Desktop reminders are off. Turn them on in Reminders & settings.";
  b.append(el("div","bi",ICON.alarm),el("div","bt",`You have ${parts.join(" and ")}.<small>${sm}</small>`));
  const v=el("button","btn sm"); v.type="button"; v.textContent=od.length&&!dt.length?"Show overdue":"Show Today"; v.onclick=()=>go(od.length&&!dt.length?"overdue":"today");
  const x=el("button","ibtn",ICON.x); x.title="Hide for now"; x.onclick=()=>{bannerHidden=key; sessionStorage.setItem("bannerHidden",key); renderBanner();};
  b.append(v,x);}
function render(){ensureColors();renderSidebar();
  $("#taskView").hidden=view!=="tasks"; $("#noteView").hidden=view!=="notes"; $("#aiView").hidden=view!=="ai"; $("#emailView").hidden=view!=="email"; $("#dashView").hidden=view!=="dash"; $("#reviewView").hidden=view!=="review"; $(".search").style.visibility=view==="ai"||view==="dash"||view==="review"?"hidden":"";
  const vk=view+":"+(view==="tasks"?current:""); if(vk!==lastViewKey){lastViewKey=vk; const w=$(view==="notes"?"#noteView":view==="ai"?"#aiView":view==="email"?"#emailView":view==="dash"?"#dashView":view==="review"?"#reviewView":"#taskView"); w.classList.remove("view-in"); void w.offsetWidth; w.classList.add("view-in");}
  if(view==="ai"){renderAI();return;}
  if(view==="dash"){renderDash();return;}
  if(view==="review"){renderReview();return;}
  const q=$("#q"); q.placeholder=view==="notes"?"Search notes":view==="email"?"Search email":"Search tasks"; const want=view==="notes"?noteQuery:view==="email"?em.query:query; if(q.value!==want)q.value=want;
  if(view==="email"){renderEmail();return;}
  if(view==="notes"){renderNotes();return;}
  renderHeader();renderBanner();renderAddForm();renderFilters();renderList();}

/* ---------- notes ---------- */
const esc=s=>String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const reEsc=s=>s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
function hl(html,q){if(!q)return html; const re=new RegExp("("+reEsc(esc(q))+")","gi");
  return html.split(/(<[^>]+>)/).map(part=>part.startsWith("<")?part:part.replace(re,"<mark>$1</mark>")).join("");}
function noteHtml(body,q){
  return body.split("\n").map(line=>{
    const m=/^( {2,}|\t+)([^\s(].*)$/.exec(line);
    if(m&&!line.includes("`")&&m[2].length<=120)return m[1]+'<span class="cmd">'+hl(esc(m[2]),q)+"</span>";
    let h=esc(line).replace(/`([^`]+)`/g,"<code>$1</code>");
    return hl(h,q);}).join("\n");}
function fmtWhen(s){if(!s)return ""; const d=new Date(s); if(isNaN(d))return "";
  const today=new Date(); const same=d.toDateString()===today.toDateString();
  return same?"Today "+d.toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"}):d.toLocaleDateString(undefined,{month:"short",day:"numeric",year:d.getFullYear()!==today.getFullYear()?"numeric":undefined});}
function scopeNotes(){return data.notes.filter(n=>noteFilter==="all"||(noteFilter==="none"?!n.project_id||!proj(n.project_id):n.project_id===noteFilter));}
function renderNotes(){
  if(noteFilter!=="all"&&noteFilter!=="none"&&!proj(noteFilter))noteFilter="all";
  const h=$("#h1"); h.innerHTML=""; const ic=el("span","nicon",ICON.note); ic.style.width=ic.style.height="26px"; ic.querySelector("svg").style.cssText="width:24px;height:24px";
  h.append(ic,document.createTextNode("Notes"));
  const n=data.notes.length; $("#sub").textContent=n?`${n} note${n===1?"":"s"} · add, edit or delete anything you want to keep handy`:"Keep commands, how-tos and reference info here";
  // editor
  const f=$("#noteForm"), ed=editingNote&&data.notes.find(x=>x.id===editingNote);
  if(editingNote&&!ed)editingNote=null;
  f.classList.toggle("editing",!!ed);
  $("#noteLbl").innerHTML=ed?"Editing <b>"+esc(ed.title)+"</b>":"New note";
  $("#nSave").textContent=ed?"Save changes":"Add note"; $("#nCancel").hidden=!ed;
  const sel=$("#nProj"), prev=sel.value; sel.innerHTML="";
  const o0=el("option"); o0.value=""; o0.textContent="No project"; sel.append(o0);
  for(const p of sortedProjects()){const o=el("option"); o.value=p.id; o.textContent="Project: "+p.name; sel.append(o);}
  if(f.dataset.loaded!==(editingNote||"")){f.dataset.loaded=editingNote||"";
    $("#nTitle").value=ed?ed.title:""; $("#nBody").value=ed?ed.body:"";
    sel.value=ed?(proj(ed.project_id)?ed.project_id:""):(noteFilter!=="all"&&noteFilter!=="none"?noteFilter:"");}
  else sel.value=proj(prev)?prev:"";
  nautosize();
  // filters
  const fl=$("#nfilters"); fl.innerHTML="";
  const chips=[["all","All",null,data.notes.length],["none","No project",null,data.notes.filter(x=>!x.project_id||!proj(x.project_id)).length]];
  for(const p of sortedProjects()){const c=data.notes.filter(x=>x.project_id===p.id).length; if(c||noteFilter===p.id)chips.push([p.id,p.name,p.color,c]);}
  for(const [k,label,color,c] of chips){const b=el("button","chip"+(noteFilter===k?" on":"")); b.type="button";
    if(color){const d=el("span","dot"); d.style.background=color; b.append(d);}
    b.append(document.createTextNode(label)); const cn=el("span","count",String(c)); b.append(cn);
    b.onclick=()=>{noteFilter=k;render();}; fl.append(b);}
  // list
  const list=$("#nlist"); list.innerHTML="";
  const q=noteQuery.trim().toLowerCase();
  const ns=scopeNotes().filter(x=>!q||x.title.toLowerCase().includes(q)||x.body.toLowerCase().includes(q)||((proj(x.project_id)||{}).name||"").toLowerCase().includes(q))
    .sort((a,b)=>(b.updated||b.created||"").localeCompare(a.updated||a.created||""));
  if(!ns.length){const e=q||noteFilter!=="all"?emptyState("search","No notes match that","Try another word, or show all notes.",[["Show all notes",()=>{noteQuery="";noteFilter="all";render();},"Esc"]],"#4fc3ff","#7c6cff")
      :emptyState("pen","No notes yet","Write one above. Commands, links, steps, anything you want to remember.",[["Write a note",()=>$("#nTitle").focus(),"Ctrl+Shift+N"]]);
    e.style.gridColumn="1/-1"; list.append(e); return;}
  for(const x of ns)list.append(noteCard(x,noteQuery.trim()));
}
function noteCard(x,q){
  const p=proj(x.project_id);
  const c=el("div","note"); c.style.setProperty("--nc",p?p.color:"linear-gradient(90deg,var(--accent),var(--accent2))");
  const h=el("h3",null,hl(esc(x.title),q));
  const b=el("div","nb",noteHtml(x.body||"",q));
  const meta=el("div","nmeta");
  if(p){const tg=el("span","tag"); const d=el("span","dot"); d.style.background=p.color; tg.append(d,document.createTextNode(p.name)); meta.append(tg);}
  const when=el("span"); when.textContent=(x.updated&&x.updated!==x.created?"Edited ":"Added ")+fmtWhen(x.updated||x.created); meta.append(when);
  const acts=el("div","acts");
  const cp=el("button","ibtn",ICON.copy); cp.title="Copy note text"; cp.onclick=()=>{navigator.clipboard.writeText(x.title+"\n\n"+(x.body||"")).then(()=>toast("Copied"),()=>toast("Couldn't copy"));};
  const eb=el("button","ibtn",ICON.edit); eb.title="Edit note"; eb.onclick=()=>startEditNote(x.id);
  const db=el("button","ibtn",ICON.trash); db.title="Delete note"; db.onclick=()=>{const i=data.notes.indexOf(x);
    data.notes.splice(i,1); if(editingNote===x.id)editingNote=null; commit();
    snack("Note deleted",()=>{data.notes.splice(Math.min(i,data.notes.length),0,x);commit();});};
  acts.append(cp,eb,db);
  c.ondblclick=e=>{if(!e.target.closest(".acts"))startEditNote(x.id);};
  c.append(acts,h,b,meta); return c;
}
function startEditNote(id){editingNote=id; render(); $(".main").scrollTo({top:0,behavior:"smooth"}); const t=$("#nBody"); t.focus(); t.setSelectionRange(t.value.length,t.value.length);}
function cancelNote(){editingNote=null; $("#noteForm").dataset.loaded="\u0000"; render();}
function nautosize(){const t=$("#nBody"); t.style.height="auto"; t.style.height=Math.min(Math.max(t.scrollHeight+2,96),420)+"px"; t.style.overflowY=t.scrollHeight>418?"auto":"hidden";}

$("#newp").onsubmit=e=>{e.preventDefault(); const v=$("#newpName").value.trim(); if(!v)return;
  if(data.projects.some(p=>p.name.toLowerCase()===v.toLowerCase())){toast("That project already exists");return;}
  const p={id:uid(),name:v,color:PALETTE[data.projects.length%PALETTE.length]}; data.projects.push(p); current=p.id; $("#newpName").value=""; commit(); $("#tTitle").focus();};
$("#addForm").onsubmit=e=>{e.preventDefault(); const v=$("#tTitle").value.trim(); if(!v)return;
  const pid=current!=="all"?current:$("#tProj").value; if(!pid||!proj(pid)){toast("Add a project first");$("#newpName").focus();return;}
  const tm=$("#tTime").value||""; let due=$("#tDue").value||"";
  if(!due&&(tm||current==="today"||current==="overdue"))due=todayStr(); if(!due&&current==="upcoming")due=addDays(1);
  const rf=$("#tRep").value, rd=$("#tRepDay").value; const repeat=rf?(rf==="weekly"?{freq:rf,dow:+rd}:rf==="monthly"?{freq:rf,dom:+rd}:{freq:rf}):null;
  if(repeat&&!due)due=firstDue(repeat);
  const nt={id:uid(),project_id:pid,title:v,priority:newPrio,due,due_time:due?tm:"",done:false,created:nowIso(),completed:""}; if(repeat)nt.repeat=repeat; data.tasks.push(nt);
  if(repeat)toast("Repeats: "+repLabel(repeat)); $("#tRep").value=""; $("#tRepDay").hidden=true;
  if(current==="overdue"&&!(tm&&tm<=nowHM()))toast("Added. Find it in Today");
  $("#tTitle").value=""; autosize(); $("#tDue").value=""; $("#tTime").value=""; newPrio="normal"; commit(); $("#tTitle").focus();};
function autosize(){const t=$("#tTitle");t.style.height="auto";t.style.height=Math.min(t.scrollHeight+2,240)+"px";t.style.overflowY=t.scrollHeight>238?"auto":"hidden";}
$("#tTitle").addEventListener("input",autosize);
$("#tRep").onchange=()=>{const f=$("#tRep").value, d=$("#tRepDay"), base=$("#tDue").value?parseYmd($("#tDue").value):new Date();
  if(f==="weekly"){d.innerHTML=DOW.map((n,i)=>`<option value="${i}">on ${n}</option>`).join(""); d.value=String(base.getDay()); d.hidden=false;}
  else if(f==="monthly"){d.innerHTML=Array.from({length:31},(_,i)=>`<option value="${i+1}">on day ${i+1}</option>`).join(""); d.value=String(base.getDate()); d.hidden=false;}
  else d.hidden=true;};
$("#bkBtn").onclick=()=>exportBackup();
$("#tTitle").addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();$("#addForm").requestSubmit();}});
$("#noteForm").onsubmit=e=>{e.preventDefault();
  const title=$("#nTitle").value.trim(), body=$("#nBody").value.replace(/\r/g,"").replace(/\s+$/,"").replace(/^\n+/,""), pid=$("#nProj").value||null;
  if(!title&&!body){$("#nTitle").focus();return;}
  const t=title||body.split("\n")[0].slice(0,60), now=new Date().toISOString().slice(0,19);
  const ed=editingNote&&data.notes.find(x=>x.id===editingNote);
  if(ed){ed.title=t; ed.body=body; ed.project_id=pid; ed.updated=now; editingNote=null; toast("Note saved");}
  else data.notes.push({id:uid(),title:t,body,project_id:pid,created:now,updated:now});
  $("#noteForm").dataset.loaded="\u0000"; commit(); if(!ed)$("#nTitle").focus();};
$("#nCancel").onclick=cancelNote;
$("#nBody").addEventListener("input",nautosize);
$("#noteForm").addEventListener("keydown",e=>{if(e.key==="Enter"&&e.ctrlKey){e.preventDefault();$("#noteForm").requestSubmit();}
  else if(e.key==="Enter"&&e.target.id==="nTitle"){e.preventDefault();$("#nBody").focus();}
  else if(e.key==="Escape"&&editingNote){e.preventDefault();cancelNote();}});
$("#doneToggle").onclick=()=>{showDone=!showDone;render();};
$("#q").oninput=e=>{if(view==="notes"){noteQuery=e.target.value;renderNotes();}else if(view==="email"){em.query=e.target.value;renderEmail();}else{query=e.target.value;renderList();}};
/* ---------- v4: undo, reminders settings, shortcuts ---------- */
const isTyping=t=>!!t&&(t.isContentEditable||/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName));
function keysHtml(spec){const w=el("span","keys");
  spec.split(" ").forEach((chunk,i)=>{if(chunk==="then"||chunk==="or"){w.append(document.createTextNode(" "+chunk+" "));return;}
    const k=el("kbd"); k.textContent=chunk; w.append(k);}); return w;}
let snackTimer=null;
function snack(msg,undo){lastUndo=undo; $("#snackMsg").textContent=msg; const s=$("#snack"); s.classList.add("show");
  clearTimeout(snackTimer); snackTimer=setTimeout(()=>{s.classList.remove("show"); lastUndo=null;},6000);}
function doUndo(){if(!lastUndo)return false; const f=lastUndo; lastUndo=null; $("#snack").classList.remove("show"); f(); toast("Undone"); return true;}
$("#snackUndo").onclick=doUndo;
function toggleDone(t,card){
  const apply=()=>{const undo=setDone(t,!t.done); commit();
    if(t.done)snack("Marked done: "+t.title.split("\n")[0].slice(0,40)+(t.repeat_next?" · next one added":""),undo);};
  if(!t.done&&!showDone&&card){card.classList.add("leaving"); setTimeout(apply,300);} else apply();}
function deleteTask(t){const i=data.tasks.indexOf(t); if(i<0)return; data.tasks.splice(i,1); if(selTask===t.id)selTask=null; commit();
  snack("Task deleted",()=>{data.tasks.splice(Math.min(i,data.tasks.length),0,t);commit();});}

/* modals */
let modalKeys=null;
function openModal(html,cls){modalKeys=null; const m=$("#modal"); m.className="modal"+(cls?" "+cls:""); m.innerHTML=html; $("#modalBack").hidden=false; return m;}
function closeModal(){modalKeys=null; $("#modalBack").hidden=true; $("#modal").innerHTML=""; render();}
$("#modalBack").addEventListener("mousedown",e=>{if(e.target.id==="modalBack")closeModal();});
const SHORTCUTS=[
  ["Anywhere",[["Ctrl+N or N","New task"],["Ctrl+Shift+N or Shift+N","New note"],["Ctrl+F or /","Search"],["?","Show these shortcuts"],["Ctrl+,","Reminders & settings"],["Ctrl+Z","Undo delete, complete or an AI change"],["Ctrl+Shift+A","Ask AI to add a task"],["Esc","Close, cancel or clear"]]],
  ["Dashboard",[["G then D or Ctrl+D","Dashboard: today at a glance (opens first)"],["B","Brief me with AI (on the dashboard)"],["R","Refresh inbox, meetings and brief"],["J or K","Select a task on the dashboard"],["X","Mark the selected task done"],["Enter","Open the selected task in its project"]]],
  ["New in v8",[["Ctrl+Alt+N","Quick add a task from anywhere in Windows"],["F","Focus timer on the selected task"],["Shift+F","Pause / resume the focus timer"],["Ctrl+Shift+F","Stop the timer and log the time"],["Y","Repeat the selected task"],["G then W","Weekly review ([ and ] change week, A = AI recap)"],["Ctrl+Shift+E","Back up everything to a file"],["Ctrl+Shift+O","Restore from a backup"],["Ctrl+Shift+L","Light / dark theme"],["Ctrl+Shift+K","Next accent colour"],["C","Hide / show the Get set up checklist (dashboard)"],["Ctrl+Shift+U","Install an available update"],["T","Take the tour (in this list)"]]],
  ["Go to",[["G then A","All projects"],["G then T","Today"],["G then U","Upcoming"],["G then O","Overdue"],["G then N","Notes"],["G then I","AI suggestions"],["G then E","Email inbox"],["Ctrl+1 … Ctrl+9","Project 1 to 9 (sidebar order)"],["Shift+P","New project"]]],
  ["Selected task",[["J or ↓","Next task"],["K or ↑","Previous task"],["X","Mark done / not done"],["E or Enter","Edit the text"],["P","Change priority"],["D","Pick a due date"],["T","Due today"],["Delete","Delete (you can undo)"],["H","Show or hide completed"]]],
  ["Typing & Windows",[["Enter","Add the task"],["Shift+Enter","New line in a task"],["Ctrl+Enter","Save a note, or Add in the AI confirm card"],["Ctrl+Shift+Space","Start or stop voice in Ask AI"],["Win+H","Windows voice typing in any box"],["Ctrl+Alt+T","Open Work To-Do from anywhere in Windows"]]],
  ["Email",[["R","Refresh the inbox"],["J or K","Next / previous email"],["A or Enter","Add as task (opens the confirm card)"],["X or Delete","Dismiss (you can undo)"],["O","Open in classic Outlook"],["Ctrl+V","Paste an email to make a task"],["Ctrl+Enter","Make a task from the paste box"]]]];
function openShortcuts(){const m=openModal('<div class="mh"><div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"><rect x="2.5" y="6" width="19" height="12" rx="2.5"/><path d="M6.5 10h.01M10 10h.01M13.5 10h.01M17 10h.01M7.5 14h9"/></svg></div><h2>Keyboard shortcuts</h2></div><div class="kgrid" id="kgrid"></div><div class="mfoot">Press <kbd>?</kbd> any time to see this list · <kbd>Esc</kbd> closes it<span style="margin-left:auto"><button type="button" class="btn sm" id="kTour">Take the tour</button> <kbd>T</kbd></span></div>');
  m.querySelector("#kTour").onclick=()=>startTour(); modalKeys=(e,k,typing)=>{if(!typing&&(k==="t"||k==="T")){e.preventDefault(); startTour(); return true;}};
  const x=el("button","ibtn",ICON.x); x.title="Close (Esc)"; x.onclick=closeModal; m.querySelector(".mh").append(x);
  const g=m.querySelector("#kgrid"); const cols=[el("div","kcol"),el("div","kcol")];
  SHORTCUTS.forEach(([sec,rows],i)=>{const c=cols[i%2]; c.append(el("div","ksec",esc(sec)));
    for(const [k,d] of rows){const r=el("div","krow"); const s=el("span"); s.textContent=d; r.append(s,keysHtml(k)); c.append(r);}});
  g.append(...cols);}
async function openSettings(section){const s=settings();
  const m=openModal(`<div class="mh"><div class="logo">${ICON.bell.replace('stroke="currentColor"','stroke="#fff"')}</div><h2>Reminders &amp; settings</h2></div>
  <div class="srow"><div class="st"><b>Desktop reminders</b><small>Windows notifications for tasks due today, a heads-up before a task's due time, and a once-a-day nudge for overdue tasks. They work even when this window is closed.</small></div><div class="toggle${s.reminders?" on":""}" id="sRem"><div class="sw"></div></div></div>
  <div class="srow"><div class="st"><b>Morning summary</b><small>What's due today and what's overdue. If your PC is off then, it comes the first time it's on after.</small></div><input type="time" class="field" id="sMorning" value="${esc(s.morning)}"></div>
  <div class="srow"><div class="st"><b>Heads-up before a due time</b><small>For tasks with a time. You also get one at the due time.</small></div><select class="field" id="sLead">${[0,5,10,15,30,60].map(n=>`<option value="${n}"${+s.lead===n?" selected":""}>${n?n+" minutes before":"Only at the due time"}</option>`).join("")}</select></div>
  <div class="srow"><div class="st"><b>Background reminder helper</b><small>Starts with Windows. To stop it starting, press <kbd>Win+R</kbd>, type <kbd>shell:startup</kbd> and delete “Work To-Do Reminders”.</small></div><span class="pill" id="sStatus">Checking…</span><button type="button" class="btn ghost sm" id="sStart" hidden>Start it</button></div>
  <div class="mfoot" style="margin-top:12px"><button type="button" class="btn sm" id="sTest">Send a test notification</button><span id="sMsg"></span></div>
  ${v8SettingsHtml()}
  ${emailSettingsHtml()}
  ${aiSettingsHtml()}
  <div class="mfoot"><span style="margin-left:auto"><kbd>Esc</kbd> closes</span></div>`,"narrow");
  const x=el("button","ibtn",ICON.x); x.title="Close (Esc)"; x.onclick=closeModal; m.querySelector(".mh").append(x);
  const put=patch=>{data.settings=Object.assign(settings(),patch); save(); renderSidebar();};
  m.querySelector("#sRem").onclick=e=>{const on=!settings().reminders; e.currentTarget.classList.toggle("on",on); put({reminders:on}); toast(on?"Reminders on":"Reminders off");};
  m.querySelector("#sMorning").onchange=e=>{if(e.target.value)put({morning:e.target.value});};
  m.querySelector("#sLead").onchange=e=>put({lead:+e.target.value});
  const post=p=>fetch(p,{method:"POST",headers:{"X-WorkTodo":"1"}});
  m.querySelector("#sTest").onclick=async()=>{try{const r=await post("/api/reminder-test"); $("#sMsg").textContent=r.ok?"Sent. Look at the bottom-right of your screen.":"Couldn't send it";}catch(e){$("#sMsg").textContent="Couldn't reach the app";}};
  const refresh=async()=>{let st={running:false}; try{st=await (await fetch("/api/reminder-status")).json();}catch(e){}
    const p=$("#sStatus"); if(!p)return; p.textContent=st.running?"Running":"Not running"; p.className="pill "+(st.running?"ok":"bad"); $("#sStart").hidden=!!st.running;};
  m.querySelector("#sStart").onclick=async()=>{await post("/api/reminder-start").catch(()=>{}); setTimeout(refresh,1500);};
  refresh(); wireAISettings(m); wireEmailSettings(m); wireV8Settings(m);
  const secId={ai:"#sAISec",email:"#sEmSec",profile:"#sProfSec",focus:"#sFocusSec",backup:"#sBackupSec",updates:"#sUpdSec"}[section];
  if(secId)setTimeout(()=>{const x=m.querySelector(secId); if(x)x.scrollIntoView({block:"start",behavior:"smooth"}); if(section==="profile")m.querySelector("#sName").focus();},60);}
$("#kbBtn").onclick=openShortcuts; $("#setBtn").onclick=()=>openSettings();
function applyHash(){const h=location.hash.slice(1); if(isSmart(h)){view="tasks";current=h;} else if(h==="notes"||h==="ai"||h==="email"||h==="dash"||h==="review")view=h;
  if(h==="setup")setTimeout(openWizard); if(h==="tour")setTimeout(startTour);
  if(h==="shortcuts")setTimeout(openShortcuts); if(h==="settings")setTimeout(openSettings);}

/* ---------- v5: AI suggestions + "Ask AI to add" (dormant until turned on in Settings) ---------- */
const DEF_AI={enabled:false,provider:"rules",base_url:"http://127.0.0.1:1234/v1",model:"",api_key_env:"",voice_model:"small.en"};
const AI_PRESETS={lmstudio:["LM Studio","http://127.0.0.1:1234/v1",""],ollama:["Ollama","http://127.0.0.1:11434/v1",""],openrouter:["OpenRouter","https://openrouter.ai/api/v1","OPENROUTER_API_KEY"]};
const aiCfg=()=>Object.assign({},DEF_AI,(data.settings&&data.settings.ai)||{});
const aiOn=()=>aiCfg().enabled===true;
ICON.spark='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 4.9L19 9.7l-5.2 1.8L12 16.5l-1.8-5L5 9.7l5.2-1.8z"/><path d="M19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z"/></svg>';
ICON.mic='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="2.5" width="6" height="12" rx="3"/><path d="M5.5 11a6.5 6.5 0 0 0 13 0M12 17.5V21M8.5 21h7"/></svg>';
ICON.stop='<svg viewBox="0 0 24 24" fill="currentColor"><rect x="6" y="6" width="12" height="12" rx="2.5"/></svg>';
ICON.refresh='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7"/></svg>';
ICON.power='<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round"><path d="M12 3v8"/><path d="M6.3 6.8a8 8 0 1 0 11.4 0"/></svg>';
let ai={recs:null,loading:false,err:"",note:"",provider:"",generated:"",dismissed:new Set(),busy:false,
  mr:null,stream:null,chunks:[],recStart:0,recTimer:null,voice:"idle",voiceMsg:""};
async function aiPost(path,body,blob){
  if(!aiOn())throw new Error("AI suggestions are off");
  const r=await fetch("/api/ai/"+path,{method:"POST",headers:blob?{"X-WorkTodo":"1","X-Audio-Type":blob.type||"audio/webm"}:{"X-WorkTodo":"1","Content-Type":"application/json"},body:blob||JSON.stringify(body||{})});
  let j={}; try{j=await r.json();}catch(e){}
  if(!r.ok)throw new Error(j.error||("HTTP "+r.status)); return j;}
async function saveNow(){clearTimeout(saveTimer);
  const r=await fetch("/api/data",{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)}); savePending=false; if(!r.ok)throw new Error("Couldn't save first"); if(r.headers.get("X-Rev"))lastRev=r.headers.get("X-Rev");}
function renderAIItem(){const w=$("#aiItem"); if(!w)return; w.innerHTML="";
  const it=el("div","pitem"+(view==="ai"?" on":"")); const ic=el("span","sicon",ICON.spark); ic.style.color="var(--accent)";
  const on=aiOn(), n=on&&ai.recs?ai.recs.filter(r=>!ai.dismissed.has(r.id)).length:0;
  const c=el("span","count"+(on?(n?" amber":""):""),on?String(n||"0"):"Off"); if(!on)c.title="AI suggestions are off";
  it.append(ic,el("span","pname","AI suggestions"),c); it.onclick=()=>go("ai"); it.title="AI suggestions (Ctrl+Shift+A)"; w.append(it);}
async function loadRecs(force){if(!aiOn()||ai.loading)return; if(!force&&ai.recs&&aiCfg().provider!=="rules")return;
  ai.loading=true; ai.err=""; if(view==="ai")renderAI();
  try{await saveNow(); const j=await aiPost("recommend",{}); ai.recs=j.recommendations||[]; ai.note=j.note||""; ai.provider=j.provider; ai.generated=j.generated||"";}
  catch(e){ai.err=e.message; ai.recs=ai.recs||[];}
  ai.loading=false; renderSidebar(); if(view==="ai")renderAI();}
const SEV={high:"var(--urgent)",medium:"var(--high)",info:"var(--accent2)"};
function aiAskCard(on){const c=el("div","aiask"+(on?"":" off"));
  const cfg=aiCfg();
  c.innerHTML=`<div class="lbl">${ICON.spark}<span>Ask AI to add a task</span><span class="hint" style="margin-left:auto">${on?(cfg.provider==="rules"?"Built-in parser · on this PC":"Your model · rules as backup"):"Off"}</span></div>
  <div class="aiaskrow"><textarea id="aiText" rows="1" placeholder="${on?"e.g. remind me to send the Oracle keys to Mike tomorrow at 3, high priority":"Off. Turn on AI suggestions in Settings to use this."}" ${on?"":"disabled"}></textarea>
  <button type="button" class="micbtn${ai.voice==="rec"?" rec":""}" id="aiMic" ${on?"":"disabled"} title="${on?"Speak your task (Ctrl+Shift+Space). Recorded and transcribed on this PC by the Live Transcriber's local Whisper model; nothing is sent online. No mic set up? Click the box and press Win+H for Windows voice typing.":"Off. Turn on AI suggestions in Settings. (Windows voice typing, Win+H, always works in any text box.)"}">${ai.voice==="rec"?ICON.stop:ICON.mic}</button>
  <button type="button" class="btn" id="aiGo" ${on?"":"disabled"}>Add with AI</button></div>
  <div class="aistat" id="aiStat">${esc(ai.voiceMsg||(on?"Enter parses it · Shift+Enter new line · you check every field before it's added":"Win+H (Windows voice typing) works in any text box, even while this is off."))}</div>`;
  if(on){const t=c.querySelector("#aiText");
    t.addEventListener("input",()=>{t.style.height="auto";t.style.height=Math.min(t.scrollHeight+2,160)+"px";});
    t.addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();askAI();}});
    c.querySelector("#aiGo").onclick=askAI; c.querySelector("#aiMic").onclick=toggleMic;}
  return c;}
function renderAI(){
  const h=$("#h1"); h.innerHTML=""; const ic=el("span","sicon",ICON.spark); ic.style.cssText="width:28px;height:28px;color:var(--accent)"; ic.querySelector("svg").style.cssText="width:26px;height:26px";
  h.append(ic,document.createTextNode("AI suggestions")); const on=aiOn(), cfg=aiCfg();
  $("#sub").textContent=on?("On · "+(cfg.provider==="rules"?"built-in rules, everything stays on this PC":"OpenAI-compatible: "+(cfg.model||"no model set")+" at "+cfg.base_url)+" · nothing changes until you confirm"):"Off · nothing runs and nothing connects anywhere until you turn it on";
  const ask=$("#aiAsk"), keep=ask.querySelector("#aiText"), kept=keep?keep.value:""; ask.innerHTML=""; ask.append(aiAskCard(on)); if(kept&&on)ask.querySelector("#aiText").value=kept;
  const b=$("#aiBody"); b.innerHTML="";
  if(!on){const d=el("div","empty aidorm");
    d.innerHTML=`<div class="art" style="--ea:#3a4157;--eb:#262c3b;box-shadow:none">${ICON.power}</div><h3>Off, turn on in Settings</h3>
    <div>AI suggestions are built in but switched off. While off, nothing runs: no suggestions, no voice, no network calls.</div>
    <div class="aiwhat"><div><b>Suggestions</b>Raise priority on tasks due within a day, reschedule overdue or overloaded days, date old tasks, split long tasks, add follow-ups.</div>
    <div><b>Ask AI to add</b>Type or say a sentence and it becomes a task with project, priority and due time.</div>
    <div><b>You stay in control</b>Every change shows a confirm card first, and Ctrl+Z undoes it.</div></div>`;
    const r=el("div","eacts"); const bt=el("button","btn sm","Open settings"); bt.type="button"; bt.onclick=()=>openSettings("ai"); r.append(bt,keysHtml("Ctrl+,")); d.append(r); b.append(d); return;}
  const vis=(ai.recs||[]).filter(r=>!ai.dismissed.has(r.id));
  const bar=el("div","toolbar");
  const left=el("div","chips"); left.innerHTML=`<span class="pill ${ai.err?"bad":"ok"}">${ai.loading?"Thinking…":ai.err?"Problem":vis.length+" suggestion"+(vis.length===1?"":"s")}</span>`+(ai.generated?`<span class="hint">Checked ${esc(fmtWhen(ai.generated))}${ai.provider==="rules"?" by the built-in rules":" by your model"}</span>`:"");
  const right=el("div","chips"); const rf=el("button","btn ghost sm",ICON.refresh+"<span>Refresh</span>"); rf.type="button"; rf.className+=" ibtxt"; rf.onclick=()=>loadRecs(true);
  const all=el("button","btn sm","Apply all ("+vis.length+")"); all.type="button"; all.disabled=!vis.length||ai.loading; all.onclick=()=>confirmApply(vis.flatMap(r=>r.actions),"Apply all "+vis.length+" suggestions",vis.map(r=>r.id));
  right.append(rf,all); bar.append(left,right); b.append(bar);
  if(ai.err){const e=el("div","banner"); e.innerHTML=`<div class="bi">${ICON.alarm}</div><div class="bt"><b>Couldn't get suggestions</b><small>${esc(ai.err)}</small></div>`; b.append(e);}
  if(ai.note){const e=el("div","aistat"); e.textContent=ai.note; e.style.margin="0 0 10px"; b.append(e);}
  if(!ai.recs&&!ai.loading){loadRecs(); return;}
  if(!vis.length&&!ai.loading){b.append(emptyState("done","Nothing to suggest right now","Your list looks in good shape. Check back after adding or finishing tasks.",[["Check again",()=>loadRecs(true),""]],"#3ddc97","#4fc3ff")); return;}
  const list=el("div","list");
  for(const r of vis){const c=el("div","aicard"); c.style.setProperty("--pc",SEV[r.severity]||SEV.info);
    const bd=el("div","body"); const t=el("div","ttl"); t.textContent=r.title; const rs=el("div","aireason"); rs.textContent=r.reason;
    const ch=el("ul","aichg"); for(const x of (r.changes||[])){const li=el("li"); li.textContent=x; ch.append(li);}
    bd.append(t,rs,ch);
    const acts=el("div","aiacts"); const ap=el("button","btn sm","Apply"); ap.type="button"; ap.onclick=()=>confirmApply(r.actions,r.title,[r.id]);
    const dm=el("button","ibtn",ICON.x); dm.title="Dismiss for now"; dm.onclick=()=>{ai.dismissed.add(r.id); renderSidebar(); renderAI();};
    acts.append(ap,dm); c.append(bd,acts); list.append(c);}
  b.append(list);}
/* confirm then apply: every change goes through this */
async function confirmApply(actions,label,recIds){if(!actions||!actions.length||ai.busy)return;
  let pv; try{pv=await aiPost("preview",{actions});}catch(e){toast(e.message);return;}
  const m=openModal(`<div class="mh"><div class="logo">${ICON.spark.replace('stroke="currentColor"','stroke="#fff"')}</div><h2>Do you really want to do this action?</h2></div>
    <div class="cfsub">${esc(label)}</div><div class="cflbl">Exactly what will change (${pv.changes.length})</div><ul class="aichg big" id="cfList"></ul>
    <div class="mfoot"><span>Nothing is deleted. You can undo right after with <kbd>Ctrl+Z</kbd>.</span><span style="margin-left:auto;display:flex;gap:8px"><button type="button" class="btn ghost" id="cfNo">Cancel</button><button type="button" class="btn" id="cfYes">Confirm</button></span></div>`,"narrow");
  const ul=m.querySelector("#cfList"); for(const x of pv.changes){const li=el("li"); li.textContent=x; ul.append(li);}
  m.querySelector("#cfNo").onclick=closeModal; m.querySelector("#cfNo").focus();
  m.querySelector("#cfYes").onclick=async()=>{const y=m.querySelector("#cfYes"); y.disabled=true; y.textContent="Applying…";
    const ok=await applyActions(pv.actions,pv.changes.length===1?pv.changes[0]:"Applied "+pv.changes.length+" changes"); if(ok){(recIds||[]).forEach(id=>ai.dismissed.add(id)); closeModal(); if(aiCfg().provider==="rules")loadRecs(true);} else {y.disabled=false;y.textContent="Confirm";}};}
async function applyActions(actions,msg){ai.busy=true;
  try{await saveNow(); const j=await aiPost("apply",{actions});
    const u=j.undo||{}; const mod=u.modified_tasks||[], addT=new Set(u.added_task_ids||[]), addN=new Set(u.added_note_ids||[]);
    data=j.data; data.notes=Array.isArray(data.notes)?data.notes:[]; render();
    snack(msg.length>70?msg.slice(0,69)+"…":msg,()=>{data.tasks=data.tasks.filter(t=>!addT.has(t.id)); data.notes=data.notes.filter(n=>!addN.has(n.id));
      for(const o of mod){const i=data.tasks.findIndex(t=>t.id===o.id); if(i>=0)data.tasks[i]=o; else data.tasks.push(o);} commit(); if(view==="ai"&&aiCfg().provider==="rules")loadRecs(true);});
    return true;}
  catch(e){toast(e.message); return false;} finally{ai.busy=false;}}
/* Ask AI to add */
function aiProjectHint(){return current!=="all"&&!isSmart(current)&&proj(current)?current:null;}
async function askAI(){const t=$("#aiText"); if(!t||!aiOn())return; const text=t.value.trim(); if(!text){t.focus();return;}
  setAIStat("Reading it…"); $("#aiGo").disabled=true;
  try{const j=await aiPost("parse",{text,project_hint:aiProjectHint()}); setAIStat(""); openTaskConfirm(j.task,j.note,text);}
  catch(e){setAIStat("Couldn't read that: "+e.message);} finally{const g=$("#aiGo"); if(g)g.disabled=false;}}
function setAIStat(s){ai.voiceMsg=s; const x=$("#aiStat"); if(x)x.textContent=s||"Enter parses it · Shift+Enter new line · you check every field before it's added";}
function openTaskConfirm(task,note,said,opts){opts=opts||{};
  const m=openModal(`<div class="mh"><div class="logo">${(opts.add?ICON.mail:ICON.spark).replace('stroke="currentColor"','stroke="#fff"')}</div><h2>${esc(opts.heading||"Add this task?")}</h2></div>
    <div class="cfsub">${opts.sub?esc(opts.sub):`From: “${esc(said.slice(0,160))}”${task.parsed_by==="model"?" · read by your model":" · read by the built-in parser"}`}</div>
    ${note?`<div class="aistat" style="margin:0 0 10px">${esc(note)}</div>`:""}
    <div class="cflbl">Title (extra lines become notes on the task)</div><textarea class="nbody" id="cfTitle" style="min-height:70px"></textarea>
    <div class="cfgrid"><label><span class="cflbl">Project${task.project_matched?"":" · not mentioned, check it"}</span><select class="field" id="cfProj"></select></label>
    <label><span class="cflbl">Priority</span><div class="seg" id="cfPrio"></div></label>
    <label><span class="cflbl">Due date</span><input type="date" class="field" id="cfDue"></label>
    <label><span class="cflbl">Due time</span><input type="time" class="field" id="cfTime"></label></div>
    <div class="mfoot"><span>Check and edit anything above. <kbd>Ctrl+Enter</kbd> adds · <kbd>Esc</kbd> cancels</span><span style="margin-left:auto;display:flex;gap:8px"><button type="button" class="btn ghost" id="cfNo">Cancel</button><button type="button" class="btn" id="cfYes">Add</button></span></div>`,"narrow");
  const ti=m.querySelector("#cfTitle"); ti.value=task.title||""; const sel=m.querySelector("#cfProj");
  for(const p of sortedProjects()){const o=el("option"); o.value=p.id; o.textContent=p.name; sel.append(o);}
  sel.value=proj(task.project_id)?task.project_id:(aiProjectHint()||(sortedProjects()[0]||{}).id||"");
  if(!data.projects.length){const o=el("option"); o.value=""; o.textContent="Add a project first"; sel.append(o);}
  let pr=task.priority||"normal"; const seg=m.querySelector("#cfPrio");
  const drawSeg=()=>{seg.innerHTML=""; for(const p of ["normal","high","urgent"]){const b=el("button",pr===p?"on":""); b.type="button"; const d=el("span","dot"); d.style.background=PCOL[p]; b.append(d,document.createTextNode(p[0].toUpperCase()+p.slice(1))); b.onclick=()=>{pr=p;drawSeg();}; seg.append(b);}}; drawSeg();
  m.querySelector("#cfDue").value=task.due||""; m.querySelector("#cfTime").value=task.due_time||"";
  const add=async()=>{const title=ti.value.replace(/\r/g,"").replace(/\n{3,}/g,"\n\n").trim(); if(!title){ti.focus();return;}
    let due=m.querySelector("#cfDue").value||"", tm=m.querySelector("#cfTime").value||""; if(tm&&!due)due=todayStr();
    const y=m.querySelector("#cfYes"); y.disabled=true;
    const action={type:"add_task",title,project_id:sel.value,priority:pr,due,due_time:due?tm:""};
    if(opts.add){const ok=await opts.add(action,title); if(ok)closeModal(); else y.disabled=false; return;}
    const ok=await applyActions([action],"Added: "+title.split("\n")[0]);
    if(ok){closeModal(); const a=$("#aiText"); if(a)a.value=""; setAIStat("Added. It's in "+((proj(sel.value)||{}).name||"your list")+".");} else y.disabled=false;};
  m.querySelector("#cfYes").onclick=add; m.querySelector("#cfNo").onclick=closeModal;
  m.addEventListener("keydown",e=>{if(e.key==="Enter"&&e.ctrlKey){e.preventDefault();add();}});
  setTimeout(()=>{ti.focus(); if(opts.add){ti.style.height=Math.min(ti.scrollHeight+4,230)+"px"; ti.setSelectionRange(0,0); ti.scrollTop=0;} else ti.setSelectionRange(ti.value.length,ti.value.length);});}
/* voice: record in the window, transcribe on this PC with the Live Transcriber's Whisper (never the cloud Web Speech API) */
async function toggleMic(){if(!aiOn())return;
  if(ai.voice==="rec"){stopMic();return;} if(ai.voice==="busy")return;
  if(!navigator.mediaDevices||!window.MediaRecorder){setAIStat("Recording isn't available here. Press Win+H in the box for Windows voice typing.");return;}
  try{ai.stream=await navigator.mediaDevices.getUserMedia({audio:{channelCount:1,echoCancellation:true,noiseSuppression:true}});}
  catch(e){setAIStat("Couldn't use the microphone ("+e.name+"). Allow it when Edge asks, or press Win+H for Windows voice typing."); return;}
  const type=["audio/webm;codecs=opus","audio/webm","audio/ogg;codecs=opus"].find(t=>MediaRecorder.isTypeSupported(t))||"";
  ai.chunks=[]; ai.mr=new MediaRecorder(ai.stream,type?{mimeType:type}:undefined);
  ai.mr.ondataavailable=e=>{if(e.data&&e.data.size)ai.chunks.push(e.data);};
  ai.mr.onstop=()=>{ai.stream.getTracks().forEach(t=>t.stop()); clearInterval(ai.recTimer); sendClip(new Blob(ai.chunks,{type:ai.mr.mimeType||"audio/webm"}));};
  ai.mr.start(250); ai.voice="rec"; ai.recStart=Date.now(); renderAI();
  const tick=()=>{const s=Math.round((Date.now()-ai.recStart)/1000); setAIStat("Listening… 0:"+pad2(s)+" · click the mic or press Ctrl+Shift+Space to stop"); if(s>=60)stopMic();};
  tick(); ai.recTimer=setInterval(tick,500);}
function stopMic(){if(ai.mr&&ai.mr.state!=="inactive"){ai.voice="busy"; ai.mr.stop(); renderAI();}}
async function sendClip(blob){if(Date.now()-ai.recStart<600||blob.size<1500){ai.voice="idle"; setAIStat("That was too short. Try again."); renderAI(); return;}
  setAIStat("Transcribing on this PC…"); renderAI();
  try{const j=await aiPost("transcribe",null,blob); ai.voice="idle"; const text=(j.text||"").trim();
    if(!text){setAIStat("Didn't catch anything. Try again a little closer to the mic."); renderAI(); return;}
    setAIStat("Heard it ("+j.model+" on "+(j.device==="cuda"?"GPU":"CPU")+", "+j.seconds+" s)."); renderAI(); const t=$("#aiText"); if(t)t.value=text; askAI();}
  catch(e){ai.voice="idle"; setAIStat("Voice didn't work: "+e.message+". Win+H (Windows voice typing) is a backup."); renderAI();}}
function aiSettingsHtml(){const c=aiCfg(), oa=c.provider==="openai_compatible";
  return `<div class="ssec" id="sAISec">${ICON.spark}<span>AI suggestions</span><span class="pill${c.enabled?" ok":""}" id="sAIPill">${c.enabled?"On":"Off"}</span></div>
  <div class="srow"><div class="st"><b>Turn on AI suggestions</b><small>Off by default. When on, it reads your to-do list and suggests changes, and "Ask AI to add" turns a sentence or voice clip into a task. Nothing changes until you confirm. While off, nothing runs and nothing connects anywhere.</small></div><div class="toggle${c.enabled?" on":""}" id="sAI"><div class="sw"></div></div></div>
  <div class="srow"><div class="st"><b>Provider</b><small>Built-in rules run on this PC with no network. OpenAI-compatible works with LM Studio, Ollama or OpenRouter. If it fails, the built-in rules are used.</small></div>
    <select class="field" id="sProv"><option value="rules"${oa?"":" selected"}>Built-in rules (local)</option><option value="openai_compatible"${oa?" selected":""}>OpenAI-compatible</option></select></div>
  <div class="sprov" id="sProvF"${oa?"":" hidden"}>
    <div class="spre"><span class="hint">Presets:</span>${Object.entries(AI_PRESETS).map(([k,v])=>`<button type="button" class="chip" data-pre="${k}">${v[0]}</button>`).join("")}</div>
    <label><span class="cflbl">Base URL</span><input class="field" id="sBase" value="${esc(c.base_url)}" placeholder="http://127.0.0.1:1234/v1" spellcheck="false"></label>
    <label><span class="cflbl">Model</span><input class="field" id="sModel" value="${esc(c.model)}" placeholder="e.g. qwen2.5-7b-instruct or llama3.1" spellcheck="false"></label>
    <label><span class="cflbl">API key environment variable (the NAME, not the key)</span><input class="field" id="sKeyEnv" value="${esc(c.api_key_env)}" placeholder="Leave empty for LM Studio / Ollama · e.g. OPENROUTER_API_KEY" spellcheck="false" autocomplete="off"></label>
    <div class="nrow"><span class="hint" id="sKeyMsg">The key itself is never saved in this app. Set it in Windows as an environment variable, then restart the app.</span><button type="button" class="btn ghost sm" id="sAITest" style="margin-left:auto">Test connection</button></div></div>
  <div class="srow"><div class="st"><b>Voice model</b><small>Speech in "Ask AI to add" is transcribed on this PC with the Live Transcriber's Whisper. Nothing is sent online. Win+H (Windows voice typing) also works in any box.</small></div>
    <span class="pill" id="sVoice">…</span><select class="field" id="sVModel"><option value="small.en"${c.voice_model==="small.en"?" selected":""}>small.en (fast, CPU)</option><option value="large-v3-turbo"${c.voice_model==="large-v3-turbo"?" selected":""}>large-v3-turbo (most accurate, GPU)</option></select></div>`;}
function wireAISettings(m){
  const putAI=patch=>{data.settings=Object.assign(settings(),{ai:Object.assign(aiCfg(),patch)}); save(); renderSidebar();};
  const pill=()=>{const p=m.querySelector("#sAIPill"), on=aiOn(); p.textContent=on?"On":"Off"; p.className="pill"+(on?" ok":"");};
  m.querySelector("#sAI").onclick=e=>{const on=!aiOn(); e.currentTarget.classList.toggle("on",on); if(!on&&ai.voice==="rec")stopMic(); putAI({enabled:on}); ai.recs=null; ai.err=""; pill(); toast(on?"AI suggestions on":"AI suggestions off");};
  m.querySelector("#sProv").onchange=e=>{putAI({provider:e.target.value}); m.querySelector("#sProvF").hidden=e.target.value!=="openai_compatible"; ai.recs=null;};
  const envOk=v=>!v||/^[A-Za-z_][A-Za-z0-9_]{0,63}$/.test(v);
  m.querySelector("#sBase").onchange=e=>{const v=e.target.value.trim(); if(v&&!/^https?:\/\/\S+$/.test(v)){toast("Use a URL that starts with http:// or https://");return;} putAI({base_url:v}); ai.recs=null;};
  m.querySelector("#sModel").onchange=e=>{putAI({model:e.target.value.trim()}); ai.recs=null;};
  m.querySelector("#sKeyEnv").onchange=e=>{const v=e.target.value.trim();
    if(!envOk(v)){e.target.value=aiCfg().api_key_env; m.querySelector("#sKeyMsg").textContent="That doesn't look like a variable name. Type a NAME like OPENROUTER_API_KEY, never the key itself."; return;}
    putAI({api_key_env:v}); ai.recs=null; refreshAIStatus();};
  m.querySelectorAll("[data-pre]").forEach(b=>b.onclick=()=>{const [,u,k]=AI_PRESETS[b.dataset.pre]; m.querySelector("#sBase").value=u; m.querySelector("#sKeyEnv").value=k; putAI({base_url:u,api_key_env:k}); ai.recs=null; refreshAIStatus(); toast(b.textContent+" preset");});
  m.querySelector("#sVModel").onchange=e=>putAI({voice_model:e.target.value});
  m.querySelector("#sAITest").onclick=async()=>{const msg=m.querySelector("#sKeyMsg");
    if(!aiOn()){msg.textContent="Turn AI suggestions on first. Nothing connects while it's off."; return;}
    msg.textContent="Testing…"; try{await saveNow(); const j=await aiPost("test",{}); msg.textContent=j.rules?"Built-in rules need no connection.":"Connected. "+j.models+" model"+(j.models===1?"":"s")+" listed"+(j.model_found===false?", but your model name isn't one of them.":j.model_found?", including yours.":".");}
    catch(e){msg.textContent="Didn't connect: "+e.message;}};
  const refreshAIStatus=async()=>{let st={}; try{st=await (await fetch("/api/ai/status")).json();}catch(e){}
    const v=m.querySelector("#sVoice"); if(!v)return; v.textContent=st.voice_available?"Whisper ready":"Whisper not found"; v.className="pill "+(st.voice_available?"ok":"bad");
    const k=aiCfg().api_key_env, km=m.querySelector("#sKeyMsg"); if(k&&km)km.textContent=st.api_key_env_set?`${k} is set on this PC. The key itself is never saved here.`:`${k} isn't set on this PC yet. Set it in Windows, then restart the app.`;};
  refreshAIStatus();}

/* ---------- v6: Email inbox (classic Outlook, read-only) + drag / paste an email ---------- */
ICON.mail='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M3.5 7l8.5 6 8.5-6"/></svg>';
ICON.inbox='<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 13l2.5-7.5A2 2 0 0 1 7.4 4h9.2a2 2 0 0 1 1.9 1.5L21 13v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M3 13h5l1.5 2.5h5L16 13h5"/></svg>';
ICON.drop='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>';
ICON.ext='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/></svg>';
ICON.undo='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 14L4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/></svg>';
const DEF_EMAIL={auto:true,name:"",address:"",aliases:"",days:7,max:50};
const emCfg=()=>Object.assign({},DEF_EMAIL,(data.settings&&data.settings.email)||{});
let em={inbox:null,filter:"new",query:"",sel:null,pollT:null,polls:0,busy:false,stat:"",lastAuto:0};
const emHint="Drop or paste an email · you check every field before it's added";
async function emPost(path,body,raw){
  const h={"X-WorkTodo":"1"}; if(!raw)h["Content-Type"]="application/json"; else Object.assign(h,raw.headers);
  const r=await fetch("/api/email/"+path,{method:"POST",headers:h,body:raw?raw.body:JSON.stringify(body||{})});
  let j={}; try{j=await r.json();}catch(e){} if(!r.ok)throw new Error(j.error||("HTTP "+r.status)); return j;}
async function loadInbox(){
  try{em.inbox=await (await fetch("/api/email/inbox")).json();}catch(e){return;}
  renderEmailItem(); if(view==="email"&&$("#modalBack").hidden)renderEmail(); if(view==="dash"&&$("#modalBack").hidden){renderDash(); if(!em.inbox.running)loadBrief(false);}
  clearTimeout(em.pollT); if(em.inbox.running&&em.polls<40){em.polls++; em.pollT=setTimeout(loadInbox,1500);} else em.polls=0;}
async function refreshEmail(force){
  try{await emPost("refresh",{force:!!force}); em.lastAuto=Date.now();}catch(e){}
  em.polls=0; if(em.inbox)em.inbox.running=true; if(view==="email")renderEmail(); setTimeout(loadInbox,300);}
const emItems=()=>(em.inbox&&em.inbox.items)||[];
const emCount=s=>emItems().filter(x=>x.state===s).length;
function renderEmailItem(){const w=$("#emailItem"); if(!w)return; w.innerHTML="";
  const it=el("div","pitem"+(view==="email"?" on":"")); const ic=el("span","sicon",ICON.mail); ic.style.color="var(--accent2)";
  const st=em.inbox&&em.inbox.status&&em.inbox.status.state, n=emCount("new");
  const c=el("span","count"+(n?" amber":""),st==="no_profile"?"Set up":st==="off"?"Off":String(n));
  it.append(ic,el("span","pname","Email"),c); it.onclick=()=>go("email"); it.title="Email inbox (G then E)"; w.append(it);}
function emWhen(s){if(!s)return ""; const d=new Date(s); if(isNaN(d))return "";
  const t=new Date(), y=new Date(); y.setDate(t.getDate()-1);
  if(d.toDateString()===t.toDateString())return d.toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"});
  if(d.toDateString()===y.toDateString())return "Yesterday "+d.toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"});
  return d.toLocaleDateString(undefined,{weekday:"short",month:"short",day:"numeric"});}
function avatar(name){const n=(name||"?").replace(/[<>"]/g,"").trim(); const parts=n.includes(",")?n.split(",").reverse():n.split(/\s+/);
  const ini=((parts[0]||"?")[0]+((parts[1]||"")[0]||"")).toUpperCase(); let h=0; for(const ch of n)h=(h*31+ch.charCodeAt(0))>>>0;
  const a=el("div","av"); a.textContent=ini; a.style.background=`linear-gradient(135deg,${PALETTE[h%PALETTE.length]},${PALETTE[(h>>>4)%PALETTE.length]})`; return a;}
function emStatusCard(){
  const ib=em.inbox||{}, st=(ib.status||{}).state||"never", running=ib.running;
  const card=(cls,icon,title,small,extra)=>{const c=el("div","emsetup"+(cls?" "+cls:"")); c.innerHTML=`<div class="bi">${icon}</div><div class="bt"><b>${title}</b><small>${small}</small>${extra||""}</div>`; return c;};
  if(st==="no_profile"){const c=card("",ICON.mail,"Connect your work inbox (one-time)","Classic Outlook is installed but not signed in yet. This app reads your Inbox through it, read-only. You can keep using the new Outlook every day; classic Outlook just needs your sign-in once.",
      `<ol><li>Click <b>Set up classic Outlook</b>. It opens classic Outlook (not the new one).</li><li>Type your work email address, click <b>Connect</b>, and sign in the way you normally do (password and Authenticator).</li><li>When your Inbox shows, you're done. You can close classic Outlook.</li><li>Come back here and press <kbd>R</kbd> or <b>Refresh</b>.</li></ol>`);
    const r=el("div","eacts"); const b=el("button","btn sm","Set up classic Outlook"); b.type="button"; b.onclick=setupOutlook;
    const rf=el("button","btn ghost sm","I've signed in, refresh"); rf.type="button"; rf.onclick=()=>refreshEmail(true);
    r.append(b,rf,el("span","hint","Until then, drag or paste emails above. Nothing is ever sent, moved, deleted or marked read.")); r.querySelector(".hint").style.cssText="color:var(--muted);font-size:12px";
    c.querySelector(".bt").append(r); return c;}
  if(st==="not_installed")return card("warn",ICON.mail,"Classic Outlook isn't installed","The automatic inbox needs classic Outlook. Drag or paste emails above instead; that works with the new Outlook.");
  if(st==="not_windows")return card("warn",ICON.mail,"Automatic inbox needs Windows","Drag or paste emails above instead.");
  if(st==="busy"||st==="timeout"){const c=card("warn",ICON.alarm,st==="timeout"?"Outlook took too long to answer":"Outlook is busy",
      "It may be starting up, syncing, or waiting on a window. If classic Outlook is showing a sign-in or a “A program is trying to access…” prompt, answer it (choose Allow), then refresh. Your last list stays below.");
    const r=el("div","eacts"); const rf=el("button","btn sm","Try again"); rf.type="button"; rf.onclick=()=>refreshEmail(true); r.append(rf); c.querySelector(".bt").append(r); return c;}
  if(st==="error"){const c=card("warn",ICON.alarm,"Couldn't read the inbox",esc((ib.status||{}).error||"Something went wrong talking to Outlook.")+" Your last list stays below.");
    const r=el("div","eacts"); const rf=el("button","btn sm","Try again"); rf.type="button"; rf.onclick=()=>refreshEmail(true); r.append(rf); c.querySelector(".bt").append(r); return c;}
  if(st==="off"){const c=card("",ICON.power,"Automatic inbox is off","Turn it on to see recent emails sent to you here. Drag and paste still work.");
    const r=el("div","eacts"); const b=el("button","btn sm","Turn on"); b.type="button"; b.onclick=()=>{data.settings=Object.assign(settings(),{email:Object.assign(emCfg(),{auto:true})}); saveNow().catch(()=>{}).then(()=>refreshEmail(true));}; r.append(b); c.querySelector(".bt").append(r); return c;}
  if(st==="ok"&&(ib.status||{}).identity_known===false)return card("",ICON.mail,"Showing your whole Inbox","Outlook didn't say who you are, so every recent email is listed. Add your name and address in Reminders &amp; settings → Email to see only the ones for you.");
  if((st==="never")&&running)return card("",ICON.refresh,"Checking Outlook…","Reading recent Inbox headers. This can take a few seconds the first time.");
  return null;}
function emImportCard(){const c=el("div","emimport");
  c.innerHTML=`<div class="lbl">${ICON.drop}<span>Add a task from any email</span><span class="hint" style="margin-left:auto">Works with the new Outlook</span></div>
  <div class="emrow"><label class="drop" id="emDrop" title="Drop .eml or .msg files, or drag an email from Outlook"><input type="file" id="emFile" accept=".eml,.msg,.txt,message/rfc822" hidden>${ICON.drop}<b>Drop an email here</b><span>.eml or .msg file, or drag it from Outlook · click to browse</span></label>
  <div class="empaste"><textarea id="emPaste" placeholder="Or paste an email: open it in Outlook, press Ctrl+A then Ctrl+C, click here and Ctrl+V. From/Sent/Subject lines are picked up."></textarea>
  <div class="nrow"><span class="hint" style="margin-right:auto"><kbd>Ctrl+Enter</kbd> make the task · <kbd>Ctrl+V</kbd> anywhere on this page</span><button type="button" class="btn sm" id="emMake">Make a task</button></div></div></div>
  <div class="emstat" id="emStat">${esc(em.stat||emHint)}</div>`;
  const f=c.querySelector("#emFile"); f.onchange=()=>{if(f.files[0])importFile(f.files[0]); f.value="";};
  const t=c.querySelector("#emPaste"); t.addEventListener("keydown",e=>{if(e.key==="Enter"&&e.ctrlKey){e.preventDefault();importText(t.value);}});
  c.querySelector("#emMake").onclick=()=>importText(t.value); return c;}
function setEmStat(s){em.stat=s; const x=$("#emStat"); if(x)x.textContent=s||emHint;}
function mailCard(it){
  const c=el("div","mail"+(it.unread&&it.state==="new"?" unread":"")+(it.state!=="new"?" "+it.state:"")+(em.sel===it.id?" sel":"")); c.dataset.id=it.id;
  c.addEventListener("mousedown",e=>{if(!e.target.closest("button")){em.sel=it.id; document.querySelectorAll(".mail.sel").forEach(x=>x.classList.remove("sel")); c.classList.add("sel");}});
  c.ondblclick=e=>{if(!e.target.closest("button"))addFromEmail(it);};
  const b=el("div","body"); const top=el("div","mtop");
  const fr=el("div","mfrom"); fr.textContent=it.sender||it.sender_email||"(unknown sender)"; if(it.sender_email&&it.sender&&!it.sender.includes("@")){const sm=el("small"); sm.textContent=it.sender_email; fr.append(sm);}
  const wh=el("div","mwhen"); wh.textContent=emWhen(it.received); wh.title=it.received?new Date(it.received).toLocaleString():"";
  top.append(fr,wh); const sj=el("div","msubj"); sj.textContent=it.subject||"(no subject)"; const pv=el("div","mprev"); pv.textContent=it.preview||"";
  const meta=el("div","meta"); const W={to:"To you",cc:"CC'd you",mention:"Mentions you",inbox:"Inbox"};
  if(W[it.why]){const g=el("span","tag why"); g.textContent=W[it.why]; meta.append(g);}
  if(it.importance===2){const g=el("span","tag imp"); g.textContent="High importance"; meta.append(g);}
  if(it.state==="added"){const g=el("span","tag ok"); g.innerHTML=ICON.check.replace('stroke="#0d1017"','stroke="currentColor"').replace('<svg ','<svg width="11" height="11" '); g.append(document.createTextNode(" Added as a task")); meta.append(g);
    const t=data.tasks.find(x=>x.id===it.task_id); if(t){const l=el("button","mlink","Show task"); l.type="button"; l.onclick=()=>{go(t.project_id); setTimeout(()=>{selTask=t.id; render(); const cd=selCard(); if(cd)cd.scrollIntoView({block:"center"});},50);}; meta.append(l);}}
  if(it.state==="dismissed"){const g=el("span","tag"); g.textContent="Dismissed"; meta.append(g);}
  b.append(top,sj,pv,meta);
  const a=el("div","macts");
  if(it.can_open){const o=el("button","mlink",ICON.ext+"<span>Open in Outlook</span>"); o.type="button"; o.title="Open in classic Outlook (O)"; o.onclick=()=>openInOutlook(it); a.append(o);}
  if(it.state==="dismissed"){const r=el("button","btn ghost sm ibtxt",ICON.undo+"<span>Restore</span>"); r.type="button"; r.onclick=()=>emState(it,"undismiss"); a.append(r);}
  else{const ad=el("button","btn sm",it.state==="added"?"Add again":"Add as task"); ad.type="button"; if(it.state==="added")ad.className="btn ghost sm"; ad.title="Add as task (A)"; ad.onclick=()=>addFromEmail(it); a.append(ad);
    if(it.state==="new"){const x=el("button","ibtn",ICON.x); x.title="Dismiss (X)"; x.onclick=()=>dismissEmail(it,c); a.append(x);}}
  c.append(avatar(it.sender||it.sender_email),b,a); return c;}
function emVisible(){const q=em.query.trim().toLowerCase();
  return emItems().filter(x=>(em.filter==="all"||x.state===em.filter)&&(!q||[x.subject,x.sender,x.sender_email,x.preview].some(v=>(v||"").toLowerCase().includes(q))));}
function renderEmail(){
  const h=$("#h1"); h.innerHTML=""; const ic=el("span","sicon",ICON.mail); ic.style.cssText="width:28px;height:28px;color:var(--accent2)"; ic.querySelector("svg").style.cssText="width:26px;height:26px";
  h.append(ic,document.createTextNode("Email")); const cfg=emCfg(), ib=em.inbox||{}, st=(ib.status||{}).state||"never";
  $("#sub").textContent="Inbox for you · last "+cfg.days+" day"+(cfg.days===1?"":"s")+" · read-only from classic Outlook · nothing is added until you click Add";
  const w=$("#emailView"), keep=w.querySelector("#emPaste"), kept=keep?keep.value:""; w.innerHTML="";
  w.append(emImportCard()); if(kept)w.querySelector("#emPaste").value=kept;
  const sc=emStatusCard(); if(sc)w.append(sc);
  const bar=el("div","toolbar"); const chips=el("div","chips");
  for(const [k,l] of [["new","For you"],["added","Added"],["dismissed","Dismissed"],["all","All"]]){
    const n=k==="all"?emItems().length:emCount(k); const b=el("button","chip"+(em.filter===k?" on":"")); b.type="button"; b.append(document.createTextNode(l),el("span","count",String(n))); b.onclick=()=>{em.filter=k; em.sel=null; renderEmail();}; chips.append(b);}
  const right=el("div","chips"); const s=ib.status||{};
  const info=el("span","hint"); info.textContent=ib.running?"Checking Outlook…":s.checked?("Checked "+emWhen(s.checked)+(st==="ok"?" · "+(s.scanned||0)+" scanned":"")):(st==="no_profile"?"Waiting for classic Outlook sign-in":"");
  const rf=el("button","btn ghost sm ibtxt",ICON.refresh+"<span>"+(ib.running?"Refreshing…":"Refresh")+"</span>"); rf.type="button"; rf.disabled=!!ib.running; rf.title="Refresh (R) · also every 10 minutes while this window is open"; rf.onclick=()=>refreshEmail(true);
  right.append(info,rf); bar.append(chips,right); w.append(bar);
  const list=el("div","list"); const vis=emVisible();
  if(!vis.length){
    if(em.query)list.append(emptyState("search","No emails match","Try another word.",[["Clear search",()=>{em.query="";$("#q").value="";renderEmail();},"Esc"]],"#4fc3ff","#7c6cff"));
    else if(st==="ok"||emItems().length)list.append(emptyState("done",em.filter==="new"?"You're all caught up":"Nothing here","New emails sent to you show up here. Add the ones that need doing, dismiss the rest.",[["Refresh",()=>refreshEmail(true),"R"]],"#4fc3ff","#3ddc97"));
  } else for(const it of vis)list.append(mailCard(it));
  w.append(list);}
/* actions */
async function addFromEmail(it){if(em.busy)return; em.busy=true;
  try{const j=await emPost("suggest",{id:it.id,project_hint:aiProjectHint()}); openEmailConfirm(j);}catch(e){toast(e.message);} finally{em.busy=false;}}
function openEmailConfirm(j){const e=j.email||{};
  openTaskConfirm(j.task,j.note,"",{heading:"Add this email as a task?",sub:"Email from "+(e.sender||e.sender_email||"unknown")+(e.received?" · "+emWhen(e.received):"")+" · “"+(e.subject||"(no subject)").slice(0,120)+"”"+(j.task.parsed_by==="model"?" · read by your model":" · dates read by the built-in parser"),
    add:(action,title)=>emailAdd(e.id,action,title)});}
async function emailAdd(eid,action,title){
  try{await saveNow(); const j=await emPost("add",{email_id:eid,action}); data=j.data; data.notes=Array.isArray(data.notes)?data.notes:[];
    const it=emItems().find(x=>x.id===eid); if(it){it.state="added"; it.task_id=j.task_id;}
    render(); setEmStat("Added: "+title.split("\n")[0].slice(0,60)+" → "+((proj(action.project_id)||{}).name||"your list"));
    snack("Added: "+title.split("\n")[0].slice(0,50),()=>{data.tasks=data.tasks.filter(t=>t.id!==j.task_id); commit(); emPost("state",{id:eid,op:"unadd"}).then(loadInbox,()=>{});});
    return true;}
  catch(e){toast(e.message); return false;}}
async function dismissEmail(it,card){const go2=async()=>{try{await emPost("state",{id:it.id,op:"dismiss"}); it.state="dismissed"; renderEmailItem(); renderEmail();
    snack("Dismissed: "+(it.subject||"email").slice(0,50),()=>emState(it,"undismiss"));}catch(e){toast(e.message);}};
  if(card&&em.filter==="new"){card.classList.add("leaving"); setTimeout(go2,280);} else go2();}
async function emState(it,op){try{await emPost("state",{id:it.id,op}); it.state="new"; renderEmailItem(); if(view==="email")renderEmail(); toast("Restored");}catch(e){toast(e.message);}}
async function openInOutlook(it){toast("Opening in Outlook…"); try{await emPost("open",{id:it.id});}catch(e){toast(e.message);}}
async function setupOutlook(){try{await emPost("setup-outlook",{}); setEmStat("Classic Outlook is opening. Sign in with your work account, wait for your Inbox, then press R here.");}catch(e){toast(e.message);}}
async function importFile(file){if(!file)return; const ext=(file.name.split(".").pop()||"").toLowerCase();
  if(!["eml","msg","txt"].includes(ext)&&file.type!=="message/rfc822"){setEmStat("That's not an email file. Drop a .eml or .msg, or paste the email text.");return;}
  setEmStat("Reading "+file.name+"…");
  try{const h={"X-Filename":encodeURIComponent(file.name),"Content-Type":"application/octet-stream"}; const ph=aiProjectHint(); if(ph)h["X-Project-Hint"]=ph;
    const j=await emPost("import",null,{headers:h,body:await file.arrayBuffer()}); setEmStat(""); openEmailConfirm(j);}
  catch(e){setEmStat("Couldn't read "+file.name+": "+e.message);}}
async function importText(text){text=(text||"").trim(); if(!text){setEmStat("Paste an email first (Ctrl+A, Ctrl+C in Outlook, then Ctrl+V here)."); const t=$("#emPaste"); if(t)t.focus(); return;}
  setEmStat("Reading it…");
  try{const j=await emPost("import-text",{text,project_hint:aiProjectHint()}); setEmStat(""); openEmailConfirm(j); const t=$("#emPaste"); if(t)t.value="";}
  catch(e){setEmStat("Couldn't read that: "+e.message);}}
function emHandleTransfer(dt){
  const files=[...(dt.files||[])]; if(files.length){importFile(files[0]); if(files.length>1)toast("One email at a time. Took the first."); return true;}
  const txt=dt.getData("text/plain")||dt.getData("text/html")||""; if(txt.trim()){importText(txt); return true;}
  return false;}
function wireEmailDrop(){const w=$("#emailView"); let depth=0;
  w.addEventListener("dragenter",e=>{e.preventDefault(); depth++; w.classList.add("dragging"); const d=$("#emDrop"); if(d)d.classList.add("over");});
  w.addEventListener("dragover",e=>{e.preventDefault(); e.dataTransfer.dropEffect="copy";});
  w.addEventListener("dragleave",()=>{if(--depth<=0){depth=0; w.classList.remove("dragging"); const d=$("#emDrop"); if(d)d.classList.remove("over");}});
  w.addEventListener("drop",e=>{e.preventDefault(); depth=0; w.classList.remove("dragging"); const d=$("#emDrop"); if(d)d.classList.remove("over");
    if(e.target.id==="emPaste"&&!e.dataTransfer.files.length){const t=e.dataTransfer.getData("text/plain"); if(t){e.target.value=t;} return;}
    if(!emHandleTransfer(e.dataTransfer))setEmStat("Couldn't read that drop. In the new Outlook, open the email, press Ctrl+A then Ctrl+C, and paste it here. Or drag it to your Desktop first and drop the file.");});
  document.addEventListener("paste",e=>{if(view!=="email"||!$("#modalBack").hidden||isTyping(e.target))return; if(emHandleTransfer(e.clipboardData))e.preventDefault();});}
function emMoveSel(d){const cards=[...document.querySelectorAll("#emailView .mail")]; if(!cards.length)return;
  let i=cards.findIndex(c=>c.dataset.id===em.sel); i=i<0?(d>0?0:cards.length-1):Math.max(0,Math.min(cards.length-1,i+d));
  cards.forEach(c=>c.classList.remove("sel")); cards[i].classList.add("sel"); em.sel=cards[i].dataset.id; cards[i].scrollIntoView({block:"nearest",behavior:"smooth"});}
function emailKey(e,k){const it=emItems().find(x=>x.id===em.sel), card=em.sel&&document.querySelector(`.mail[data-id="${em.sel}"]`);
  if(k==="j"||k==="ArrowDown"){e.preventDefault();emMoveSel(1);return true;}
  if(k==="k"||k==="ArrowUp"){e.preventDefault();emMoveSel(-1);return true;}
  if(k==="r"){e.preventDefault();refreshEmail(true);return true;}
  if(!it||!card)return false;
  if(k==="a"||k==="Enter"){e.preventDefault(); if(it.state!=="dismissed")addFromEmail(it); return true;}
  if(k==="x"||k==="Delete"||k==="Backspace"){e.preventDefault(); if(it.state==="new"){const cards=[...document.querySelectorAll("#emailView .mail")], i=cards.indexOf(card), nx=cards[i+1]||cards[i-1]; em.sel=nx?nx.dataset.id:null; dismissEmail(it,card);} return true;}
  if(k==="o"){e.preventDefault(); if(it.can_open)openInOutlook(it); return true;}
  return false;}
function emailSettingsHtml(){const c=emCfg(), me=(em.inbox&&em.inbox.me)||{}, st=((em.inbox||{}).status||{}).state||"never";
  const pill={ok:["Connected","ok"],no_profile:["Not signed in","bad"],not_installed:["Not installed","bad"],busy:["Busy",""],timeout:["Busy",""],error:["Problem","bad"],off:["Off",""],never:["Not checked",""]}[st]||["…",""];
  return `<div class="ssec" id="sEmSec">${ICON.mail}<span>Email</span><span class="pill ${pill[1]}">${pill[0]}</span></div>
  <div class="srow"><div class="st"><b>Automatic inbox</b><small>Reads recent Inbox emails sent to you through classic Outlook, read-only, every 10 minutes while this window is open. It never sends, moves, deletes, flags or marks anything read. Drag and paste work either way.</small></div><div class="toggle${c.auto?" on":""}" id="sEmAuto"><div class="sw"></div></div></div>
  <div class="sprov"><label><span class="cflbl">Your name${me.name?" · Outlook says “"+esc(me.name)+"”":""}</span><input class="field" id="sEmName" value="${esc(c.name)}" placeholder="${esc(me.name||"Auto-detected from Outlook once it's signed in")}" spellcheck="false"></label>
    <label><span class="cflbl">Your work email address${(me.smtp||[]).length?" · Outlook says "+esc(me.smtp.join(", ")):""}</span><input class="field" id="sEmAddr" value="${esc(c.address)}" placeholder="${esc((me.smtp||[])[0]||"Auto-detected from Outlook once it's signed in")}" spellcheck="false"></label>
    <label><span class="cflbl">Other names people use for you (comma separated)</span><input class="field" id="sEmAlias" value="${esc(c.aliases)}" placeholder="e.g. Sam, S. Lee" spellcheck="false"></label>
    <div class="cfgrid" style="margin-top:0"><label><span class="cflbl">Look back</span><select class="field" id="sEmDays">${[1,3,7,14,30].map(n=>`<option value="${n}"${+c.days===n?" selected":""}>${n} day${n===1?"":"s"}</option>`).join("")}</select></label>
    <label><span class="cflbl">Show up to</span><select class="field" id="sEmMax">${[25,50,100].map(n=>`<option value="${n}"${+c.max===n?" selected":""}>${n} emails</option>`).join("")}</select></label></div>
    <div class="nrow"><span class="hint">Leave name and address empty to use what Outlook reports. Emails where you're on To or CC, or that mention your name, are shown.</span><button type="button" class="btn ghost sm" id="sEmSetup" style="margin-left:auto">Set up classic Outlook</button></div></div>`;}
function wireEmailSettings(m){
  const putEm=(patch,refresh)=>{data.settings=Object.assign(settings(),{email:Object.assign(emCfg(),patch)}); renderSidebar(); saveNow().then(()=>{toast("Saved"); if(refresh)refreshEmail(true);},()=>toast("Couldn't save"));};
  m.querySelector("#sEmAuto").onclick=e=>{const on=!emCfg().auto; e.currentTarget.classList.toggle("on",on); putEm({auto:on},on); setTimeout(loadInbox,600); toast(on?"Automatic inbox on":"Automatic inbox off");};
  m.querySelector("#sEmName").onchange=e=>putEm({name:e.target.value.trim()},true);
  m.querySelector("#sEmAddr").onchange=e=>{const v=e.target.value.trim(); if(v&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)){toast("That doesn't look like an email address");return;} putEm({address:v},true);};
  m.querySelector("#sEmAlias").onchange=e=>putEm({aliases:e.target.value.trim()},true);
  m.querySelector("#sEmDays").onchange=e=>putEm({days:+e.target.value},true);
  m.querySelector("#sEmMax").onchange=e=>putEm({max:+e.target.value},true);
  m.querySelector("#sEmSetup").onclick=setupOutlook;}


/* ---------- v7: Dashboard (Today at a glance) ---------- */
ICON.grid='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="3.5" width="7" height="8" rx="2"/><rect x="13.5" y="3.5" width="7" height="5" rx="2"/><rect x="13.5" y="11.5" width="7" height="9" rx="2"/><rect x="3.5" y="14.5" width="7" height="6" rx="2"/></svg>';
ICON.flag='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 21V4M5 4h11l-2 4 2 4H5"/></svg>';
ICON.mic2='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h16M4 12h10M4 18h7"/><circle cx="18" cy="16" r="3"/></svg>';
ICON.target='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r=".8" fill="currentColor"/></svg>';
let dash={brief:null,briefKey:"",briefBusy:false,aiBrief:null,meet:null,meetAt:0,rem:null,remAt:0,sel:null};
try{const s=JSON.parse(sessionStorage.getItem("aiBrief")||"null"); if(s&&s.day===todayStr())dash.aiBrief=s.brief;}catch(e){}
function myName(){const pn=(prof().name||"").trim(); if(pn)return pn.split(/\s+/)[0]; const n=(emCfg().name||"").trim(); if(n)return (n.includes(",")?n.split(",")[1]:n).trim().split(/\s+/)[0]; return "";}
function greet(){const h=new Date().getHours(); return h<12?"Good morning":h<17?"Good afternoon":"Good evening";}
function renderDashItem(){const w=$("#dashItem"); if(!w)return; w.innerHTML="";
  const it=el("div","pitem"+(view==="dash"?" on":"")); const ic=el("span","sicon",ICON.grid); ic.style.color="var(--accent)";
  const n=data.tasks.filter(t=>isOverdue(t)||isDueToday(t)).length;
  it.append(ic,el("span","pname","Dashboard"),el("span","count"+(n?" amber":""),String(n))); it.onclick=()=>go("dash"); it.title="Dashboard: today at a glance (G then D or Ctrl+D)"; w.append(it);}
async function loadMeetings(force){if(!force&&dash.meet&&Date.now()-dash.meetAt<300000)return;
  dash.meetAt=Date.now(); try{dash.meet=await (await fetch("/api/meetings")).json();}catch(e){dash.meet={available:false,items:[]};} if(view==="dash")renderDash();}
async function loadRem(force){if(!force&&dash.rem&&Date.now()-dash.remAt<120000)return;
  dash.remAt=Date.now(); try{dash.rem=await (await fetch("/api/reminder-status")).json();}catch(e){dash.rem={running:false};} if(view==="dash")renderDash();}
async function loadBrief(ai){
  if(dash.briefBusy)return; const key=JSON.stringify([data.tasks,emItems().map(x=>x.id+x.state),todayStr()]);
  if(!ai&&dash.brief&&dash.briefKey===key)return;
  dash.briefBusy=ai?"ai":"local"; if(view==="dash")renderDash();
  try{await saveNow();
    const r=await fetch("/api/brief",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:JSON.stringify({ai:!!ai})});
    const j=await r.json(); if(!r.ok)throw new Error(j.error||("HTTP "+r.status));
    if(ai){dash.aiBrief=j; sessionStorage.setItem("aiBrief",JSON.stringify({day:todayStr(),brief:j})); toast(j.source==="ai"?"Brief ready":"Built-in brief (AI didn't answer)");}
    else{dash.brief=j; dash.briefKey=key;}}
  catch(e){toast("Couldn't make the brief: "+e.message);}
  dash.briefBusy=false; if(view==="dash")renderDash();}
function dCard(icon,color,title,count,link,cls){const c=el("div","dcard"+(cls?" "+cls:"")); const h=el("div","dh"); const ic=el("span","sicon",ICON[icon]); ic.style.color=color;
  const b=el("b"); b.textContent=title; h.append(ic,b); if(count!=null)h.append(el("span","count",String(count)));
  if(link){const l=el("button","mlink"); l.type="button"; l.textContent=link[0]+" \u2192"; l.title=link[2]||""; l.onclick=link[1]; h.append(l);} c.append(h); return c;}
function dEmpty(text){const d=el("div","dempty",ICON.check.replace('stroke="#0d1017"','stroke="currentColor"')); d.append(document.createTextNode(text)); return d;}
function openTask(t){go(t.project_id&&proj(t.project_id)?t.project_id:"all"); setTimeout(()=>{selTask=t.id; render(); const cd=selCard(); if(cd)cd.scrollIntoView({block:"center"});},50);}
function dTaskRow(t,right,rcls){const r=el("div","drow"+(dash.sel===t.id?" sel":"")); r.dataset.id=t.id;
  const ck=el("button","dck"+(t.done?" done":""),ICON.check); ck.type="button"; ck.title=t.done?"Mark not done":"Mark done (X)";
  ck.onclick=e=>{e.stopPropagation(); dToggle(t,r);};
  const tx=el("div","dt"); const tt=el("div","dtt"); tt.textContent=t.title.split("\n")[0]; tt.title=t.title; tx.append(tt);
  const m=el("div","dm"); const p=proj(t.project_id);
  if(t.priority!=="normal")m.append(el("span","badge b-"+t.priority,t.priority));
  if(p){const g=el("span","tag"); const d=el("span","dot"); d.style.background=p.color; g.append(d,document.createTextNode(p.name)); m.append(g);}
  if(t.source==="email"){const g=el("span","tag",ICON.mail.replace("<svg ",'<svg width="11" height="11" ')); g.append(document.createTextNode("Email")); m.append(g);}
  if(t.repeat){const g=el("span","tag rep",ICON.repeat.replace("<svg ",'<svg width="11" height="11" ')); g.append(document.createTextNode(repLabel(t.repeat))); m.append(g);}
  tx.append(m); r.append(ck,tx);
  if(!t.done){const fb=el("button","ibtn dfocus",focus&&focus.task_id===t.id?ICON.pause:ICON.timer); fb.type="button"; fb.title="Focus timer on this task (F)"; fb.onclick=e=>{e.stopPropagation(); startFocus(t);}; r.append(fb);}
  if(right){const rr=el("div","dr"+(rcls?" "+rcls:"")); rr.textContent=right; r.append(rr);}
  r.onclick=()=>{dash.sel=t.id; document.querySelectorAll(".drow.sel").forEach(x=>x.classList.remove("sel")); r.classList.add("sel");};
  r.ondblclick=()=>openTask(t); return r;}
function dToggle(t,row){const undo=setDone(t,!t.done);
  if(row&&t.done){row.style.transition=".3s"; row.style.opacity=".35";}
  setTimeout(()=>{commit(); if(t.done)snack("Marked done: "+t.title.split("\n")[0].slice(0,40)+(t.repeat_next?" · next one added":""),undo);},t.done?250:0);}
function relDay(s){const [y,m,d]=s.split("-").map(Number); const dt=new Date(y,m-1,d), t=new Date(todayStr()+"T00:00"); const diff=Math.round((dt-t)/864e5);
  if(diff===0)return "Today"; if(diff===1)return "Tomorrow"; if(diff===-1)return "Yesterday"; if(diff<0)return (-diff)+" days ago";
  if(diff<7)return dt.toLocaleDateString(undefined,{weekday:"long"}); return dt.toLocaleDateString(undefined,{month:"short",day:"numeric"});}
function briefCard(){const hero=el("div","hero"); const b=el("div","brief"); const showAI=dash.aiBrief&&dash.aiBrief.source==="ai";
  const br=showAI?dash.aiBrief:dash.brief; const busy=dash.briefBusy;
  const bh=el("div","bh",ICON.spark); bh.append(document.createTextNode(showAI?"AI brief for today":"Your day in brief"));
  if(showAI){const p=el("span","pill ok"); p.textContent=(dash.aiBrief.model||"AI")+" \u00b7 "+fmtWhen(dash.aiBrief.generated); bh.append(p);}
  else{const p=el("span","pill"); p.textContent="Built-in \u00b7 on this PC"; bh.append(p);}
  b.append(bh);
  if(!br){b.append(el("div","shimmer"),el("div","shimmer"),el("div","shimmer")); b.querySelectorAll(".shimmer")[1].style.width="80%"; b.querySelectorAll(".shimmer")[2].style.width="60%";}
  else{const hl=el("div","hl"); hl.textContent=br.headline; b.append(hl); const ul=el("ul"); for(const x of br.bullets||[]){const li=el("li"); li.textContent=x; ul.append(li);} if((br.bullets||[]).length)b.append(ul);
    if(br.focus){const f=el("div","focus",ICON.target.replace("<svg ",'<svg width="15" height="15" ')); const s=el("span"); s.textContent=br.focus; f.append(s); b.append(f);}}
  const note=(dash.aiBrief&&dash.aiBrief.source!=="ai"&&dash.aiBrief.note)||""; if(busy==="ai")b.append(el("div","bnote","Asking your AI model ("+esc(aiCfg().model||"model")+")\u2026 free models can take up to a minute. The built-in brief stays here meanwhile."));
  else if(note){const n=el("div","bnote"); n.textContent=note; b.append(n);}
  const acts=el("div","bacts"); const on=aiOn()&&aiCfg().provider==="openai_compatible";
  const ab=el("button","btn sm",ICON.spark.replace('stroke="currentColor"','stroke="#fff"')+"<span>"+(busy==="ai"?"Writing\u2026":showAI?"Refresh AI brief":"Brief me with AI")+"</span>"); ab.type="button"; ab.disabled=!!busy||!on;
  ab.title=on?"Ask your AI model for a brief (B)":"Turn on AI suggestions with an OpenAI-compatible model in Reminders & settings to use this"; ab.onclick=()=>loadBrief(true);
  acts.append(ab); acts.append(keysHtml("B")); acts.lastChild.style.cssText="align-self:center;font-size:11px";
  if(showAI){const lb=el("button","btn ghost sm","Show built-in"); lb.type="button"; lb.onclick=()=>{dash.aiBrief=null; sessionStorage.removeItem("aiBrief"); renderDash();}; acts.append(lb);}
  hero.append(b,acts); return hero;}
function renderDash(){
  const h=$("#h1"); h.innerHTML=""; const nm=myName(); h.textContent=greet()+(nm?", "+nm:"");
  const sub=$("#sub"); sub.textContent=new Date().toLocaleDateString(undefined,{weekday:"long",month:"long",day:"numeric",year:"numeric"})+" \u00b7 "+new Date().toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"})+" \u00b7 everything for today at a glance";
  const w=$("#dashView"), st=w.scrollTop; w.innerHTML="";
  const td=todayStr(), open=data.tasks.filter(t=>!t.done), od=open.filter(isOverdue), dt=open.filter(t=>isDueToday(t)&&!isOverdue(t));
  const doneT=data.tasks.filter(t=>t.done&&(t.completed||"").slice(0,10)===td), newMail=emItems().filter(x=>x.state==="new");
  const ub=updateBanner(); if(ub)w.append(ub);
  const su=setupCard(); if(su)w.append(su);
  w.append(briefCard());
  const s=el("div","stats");
  for(const [l,n,c,fn] of [["Open tasks",open.length,"var(--accent)",()=>go("all")],["Due today",dt.length,"var(--high)",()=>go("today")],["Overdue",od.length,"var(--urgent)",()=>go("overdue")],
      ["Done today",doneT.length,"var(--ok)",()=>{showDone=true;go("all");}],["New emails",newMail.length,"var(--accent2)",()=>go("email")],["Focus today",fmtMin(focusMinutes(td,td)),"var(--accent)",()=>go("review")]]){
    const d=el("div","stat"); d.style.setProperty("--c",c); d.append(el("div","n",String(n)),el("div","l",l)); d.querySelector(".n").style.color=c; d.onclick=fn; s.append(d);}
  w.append(s);
  const g=el("div","dgrid"); w.append(g);
  // 1) needs attention
  const rank=t=>RANK[t.priority]??2;
  const att=dCard("alarm","var(--urgent)","Needs attention",od.length+dt.length,["Today",()=>go("today"),"G then T"]);
  if(!od.length&&!dt.length)att.append(dEmpty("Nothing overdue or due today."));
  if(od.length){att.append(el("div","dsub","Overdue")); for(const t of [...od].sort((a,b)=>(a.due+(a.due_time||"")).localeCompare(b.due+(b.due_time||""))||rank(a)-rank(b)).slice(0,8))att.append(dTaskRow(t,relDay(t.due)+(t.due_time?" "+fmtTime(t.due_time):""),"over"));}
  if(dt.length){att.append(el("div","dsub","Due today")); for(const t of [...dt].sort((a,b)=>(a.due_time||"99").localeCompare(b.due_time||"99")||rank(a)-rank(b)))att.append(dTaskRow(t,t.due_time?fmtTime(t.due_time):"Today"));}
  if(doneT.length){att.append(el("div","dsub","Done today")); for(const t of doneT.slice(0,4))att.append(dTaskRow(t,""));}
  g.append(att);
  // 2) schedule & reminders
  const sc=dCard("clock","var(--high)","Schedule & reminders",null,["Upcoming",()=>go("upcoming"),"G then U"]);
  const timed=data.tasks.filter(t=>!t.done&&t.due===td&&t.due_time).sort((a,b)=>a.due_time.localeCompare(b.due_time));
  const nextT=timed.find(t=>t.due_time>nowHM());
  sc.append(el("div","dsub","Today"));
  if(!timed.length)sc.append(el("div","dempty","No timed tasks today."));
  for(const t of timed)sc.append(dTaskRow(t,(t===nextT?"Next \u00b7 ":"")+fmtTime(t.due_time),t.due_time<=nowHM()?"over":t===nextT?"now":""));
  const wk=addDays(7), up=data.tasks.filter(t=>!t.done&&t.due&&t.due>td&&t.due<=wk).sort((a,b)=>(a.due+(a.due_time||"99")).localeCompare(b.due+(b.due_time||"99")));
  sc.append(el("div","dsub","Next 7 days"));
  if(!up.length)sc.append(el("div","dempty","Nothing due in the next 7 days."));
  for(const t of up.slice(0,6))sc.append(dTaskRow(t,relDay(t.due)+(t.due_time?" \u00b7 "+fmtTime(t.due_time):"")));
  const ss=settings(), rem=dash.rem, f=el("div","dfoot");
  const rp=el("span","pill"+(ss.reminders&&rem&&rem.running?" ok":ss.reminders?"":" bad")); rp.textContent=!ss.reminders?"Desktop reminders off":rem?(rem.running?"Reminders on":"Reminder helper not running"):"Checking reminders\u2026"; f.append(rp);
  if(ss.reminders){const x=el("span"); x.textContent="Morning summary "+fmtTime(ss.morning)+" \u00b7 heads-up "+(+ss.lead?ss.lead+" min before":"at the due time"); f.append(x);}
  const sb=el("button","mlink","Settings"); sb.type="button"; sb.style.marginLeft="auto"; sb.onclick=()=>openSettings(); f.append(sb); sc.append(f);
  g.append(sc);
  // 3) priorities by project
  const pc=dCard("flag","var(--accent)","Priorities by project",open.filter(t=>t.priority!=="normal").length,["All projects",()=>go("all"),"Ctrl+0"]);
  const groups=sortedProjects().map(p=>({p,ts:open.filter(t=>t.project_id===p.id)})).filter(x=>x.ts.length).sort((a,b)=>b.ts.filter(t=>t.priority!=="normal").length-a.ts.filter(t=>t.priority!=="normal").length||b.ts.length-a.ts.length);
  const orphan=open.filter(t=>!proj(t.project_id)); if(orphan.length)groups.push({p:{id:"all",name:"No project",color:"#8a93a8"},ts:orphan});
  if(!groups.length)pc.append(dEmpty("No open tasks. Enjoy it."));
  const maxN=Math.max(1,...groups.map(x=>x.ts.length));
  for(const {p,ts} of groups){const d=el("div","dproj"); const ph=el("div","ph"); const dot=el("span","dot"); dot.style.background=p.color; const nmE=el("span","pname"); nmE.textContent=p.name;
    const u=ts.filter(t=>t.priority==="urgent").length, hi=ts.filter(t=>t.priority==="high").length;
    ph.append(dot,nmE,el("span","count",ts.length+" open")); ph.onclick=()=>go(p.id); ph.title="Open "+p.name; d.append(ph);
    const bar=el("div","bar"); bar.style.width=Math.max(12,ts.length/maxN*100)+"%";
    for(const [n,c] of [[u,"var(--urgent)"],[hi,"var(--high)"],[ts.length-u-hi,"var(--normal)"]]){if(!n)continue; const sp=el("span"); sp.style.cssText=`width:${n/ts.length*100}%;background:${c};opacity:.85`; bar.append(sp);} d.append(bar);
    for(const t of ts.filter(t=>t.priority!=="normal").sort((a,b)=>rank(a)-rank(b)).slice(0,4))d.append(dTaskRow(t,t.due?relDay(t.due):"",isOverdue(t)?"over":""));
    pc.append(d);}
  g.append(pc);
  // 4) inbox: new emails not yet tasks
  const ib=em.inbox||{}, est=(ib.status||{}).state||"never";
  const mc=dCard("mail","var(--accent2)","Inbox \u00b7 not yet tasks",newMail.length,["Email",()=>go("email"),"G then E"]);
  if(est==="no_profile"||est==="not_installed"||est==="off"){const d=el("div","dempty"); d.textContent=est==="off"?"The automatic inbox is off. Open Email to turn it on.":"Connect classic Outlook once to see your inbox here (Email \u2192 Set up)."; mc.append(d);}
  else if(!newMail.length)mc.append(dEmpty(ib.running?"Checking Outlook\u2026":"You're all caught up. No new emails for you."));
  for(const it of [...newMail].sort((a,b)=>(b.received||"").localeCompare(a.received||"")).slice(0,6)){
    const r=el("div","dmail"); const t=el("div","mt"); const fr=el("div","mf"); const a1=el("span"); a1.textContent=it.sender||it.sender_email||"(unknown)"; const a2=el("span"); a2.textContent=emWhen(it.received); fr.append(a1,a2);
    if(it.importance===2){const im=el("span"); im.textContent="High importance"; im.style.color="var(--urgent)"; fr.append(im);}
    const sj=el("div","ms"); sj.textContent=it.subject||"(no subject)"; sj.title=it.preview||""; t.append(fr,sj);
    const ma=el("div","ma"); const ad=el("button","btn sm","Add as task"); ad.type="button"; ad.title="Review it in the confirm card, then Add. Nothing is added automatically."; ad.onclick=()=>addFromEmail(it);
    const x=el("button","ibtn",ICON.x); x.title="Dismiss (you can undo)"; x.onclick=async()=>{r.style.opacity=".35"; try{await emPost("state",{id:it.id,op:"dismiss"}); it.state="dismissed"; render(); snack("Dismissed: "+(it.subject||"email").slice(0,50),()=>emState(it,"undismiss"));}catch(e){toast(e.message); r.style.opacity="";}};
    ma.append(ad,x); r.append(avatar(it.sender||it.sender_email),t,ma); r.ondblclick=e=>{if(!e.target.closest("button"))addFromEmail(it);}; mc.append(r);}
  if(newMail.length>6){const m2=el("div","dsub"); m2.textContent="+ "+(newMail.length-6)+" more in Email"; mc.append(m2);}
  const ef=el("div","dfoot"); const s2=ib.status||{}; const info=el("span"); info.textContent=ib.running?"Checking Outlook\u2026":s2.checked?"Checked "+emWhen(s2.checked)+" \u00b7 read-only, nothing is added until you click Add":"Read-only \u00b7 nothing is added until you click Add"; ef.append(info);
  const rf=el("button","mlink",ICON.refresh+"<span>Refresh</span>"); rf.type="button"; rf.style.marginLeft="auto"; rf.title="Refresh inbox and meeting notes (R)"; rf.onclick=()=>dashRefresh(); ef.append(rf); mc.append(ef);
  g.append(mc);
  // 5) recent notes
  const ns=[...data.notes].sort((a,b)=>(b.updated||b.created||"").localeCompare(a.updated||a.created||"")).slice(0,4);
  const nc=dCard("note","#c9cfdc","Recent notes",data.notes.length,["Notes",()=>go("notes"),"G then N"]);
  if(!ns.length)nc.append(el("div","dempty","No notes yet. Ctrl+Shift+N writes one."));
  for(const x of ns){const p=proj(x.project_id); const d=el("div","dnote"); if(p)d.style.setProperty("--nc",p.color);
    const b=el("b"); b.textContent=x.title; const bx=el("div","nx"); bx.textContent=(x.body||"").trim().slice(0,240);
    const nw=el("div","nw"); if(p){const tg=el("span","tag"); const dd=el("span","dot"); dd.style.background=p.color; tg.append(dd,document.createTextNode(p.name)); nw.append(tg);} const wh=el("span"); wh.textContent=fmtWhen(x.updated||x.created); nw.append(wh);
    d.append(b,bx,nw); d.onclick=()=>openNoteModal(x); nc.append(d);}
  g.append(nc);
  // 6) meeting notes (Live Transcriber, read-only)
  const mt=dash.meet; if(mt&&mt.available){const its=mt.items||[];
    const mm=dCard("mic2","var(--ok)","Recent meeting notes",its.length||null,null);
    if(!its.length)mm.append(el("div","dempty","No meeting notes yet from the Live Transcriber."));
    for(const m of its.slice(0,3)){const d=el("div","dmeet"); const hh=el("div","mh2"); const b=el("b"); b.textContent=m.title; const pl=el("span","pill"+(m.source==="ai"?" ok":"")); pl.textContent=m.source==="ai"?"AI notes":"Outline"; hh.append(b,pl);
      const sx=el("div","sx"); sx.textContent=m.summary; const nw=el("div","nw"); const wd=new Date(m.when);
      nw.textContent=(isNaN(wd)?"":wd.toLocaleDateString(undefined,{weekday:"short",month:"short",day:"numeric"})+" "+wd.toLocaleTimeString(undefined,{hour:"numeric",minute:"2-digit"}))+(m.action_items.length?" \u00b7 "+m.action_items.length+" action item"+(m.action_items.length===1?"":"s"):"")+(m.key_points.length?" \u00b7 "+m.key_points.length+" key points":"");
      d.append(hh,sx,nw); d.onclick=()=>openMeetModal(m); mm.append(d);}
    const ff=el("div","dfoot"); ff.textContent="From Live Transcriber \u00b7 read-only"; mm.append(ff);
    g.append(mm);}
  w.scrollTop=st;}
function openNoteModal(x){const p=proj(x.project_id);
  const m=openModal(`<div class="mh"><div class="logo">${ICON.pen}</div><h2></h2></div><div class="mdl"><div class="cfsub"></div><div class="nb"></div><div class="mrow"></div></div>`);
  m.querySelector("h2").textContent=x.title; m.querySelector(".cfsub").textContent=(p?p.name+" \u00b7 ":"")+"Updated "+fmtWhen(x.updated||x.created); m.querySelector(".nb").innerHTML=noteHtml(x.body||"","");
  const cl=el("button","ibtn",ICON.x); cl.title="Close (Esc)"; cl.onclick=closeModal; m.querySelector(".mh").append(cl);
  const r=m.querySelector(".mrow"); const cp=el("button","btn ghost sm","Copy"); cp.type="button"; cp.onclick=()=>navigator.clipboard.writeText(x.title+"\n\n"+(x.body||"")).then(()=>toast("Copied"),()=>toast("Couldn't copy"));
  const ed=el("button","btn sm","Edit in Notes"); ed.type="button"; ed.onclick=()=>{$("#modalBack").hidden=true; $("#modal").innerHTML=""; view="notes"; startEditNote(x.id);}; r.append(cp,ed);}
function openMeetModal(mt){
  const m=openModal(`<div class="mh"><div class="logo">${ICON.mic2.replace('stroke="currentColor"','stroke="#fff"')}</div><h2></h2></div><div class="mdl"><div class="cfsub"></div><div class="msum"></div><div id="mdlBody"></div><div class="mrow"></div></div>`);
  m.querySelector("h2").textContent=mt.title; const wd=new Date(mt.when);
  m.querySelector(".cfsub").textContent=(isNaN(wd)?"":wd.toLocaleString(undefined,{weekday:"long",month:"long",day:"numeric",hour:"numeric",minute:"2-digit"}))+(mt.words?" \u00b7 "+mt.words+" words":"")+" \u00b7 "+(mt.source==="ai"?"AI notes":"local outline")+" \u00b7 Live Transcriber (read-only)";
  m.querySelector(".msum").textContent=mt.summary||"";
  const body=m.querySelector("#mdlBody");
  const sec=(t,arr,fmt)=>{if(!arr||!arr.length)return; const h=el("h4"); h.textContent=t; const ul=el("ul"); for(const a of arr){const li=el("li"); li.textContent=fmt?fmt(a):a; ul.append(li);} body.append(h,ul);};
  sec("Action items",mt.action_items,a=>(a.who&&a.who!=="Unassigned"?a.who+": ":"")+a.what+(a.due?" (due "+a.due+")":"")); sec("Decisions",mt.decisions); sec("Key points",mt.key_points); sec("Open questions",mt.open_questions);
  const cl=el("button","ibtn",ICON.x); cl.title="Close (Esc)"; cl.onclick=closeModal; m.querySelector(".mh").append(cl);
  const cp=el("button","btn ghost sm","Copy notes"); cp.type="button"; cp.onclick=()=>navigator.clipboard.writeText(mt.title+"\n\n"+mt.summary+"\n\n"+[...body.querySelectorAll("h4,li")].map(e=>e.tagName==="H4"?"\n"+e.textContent:"- "+e.textContent).join("\n")).then(()=>toast("Copied"),()=>toast("Couldn't copy"));
  m.querySelector(".mrow").append(cp);}
function dashRefresh(){refreshEmail(true); loadMeetings(true); loadRem(true); dash.briefKey=""; setTimeout(()=>loadBrief(false),1200); toast("Refreshing\u2026");}
function dashOpen(){loadMeetings(false); loadRem(false); loadBrief(false);}
function dashKey(e,k){
  if(k==="b"||k==="B"){e.preventDefault(); if(aiOn()&&aiCfg().provider==="openai_compatible")loadBrief(true); else toast("Turn on AI with a model in Reminders & settings to use the AI brief"); return;}
  if(k==="r"||k==="R"){e.preventDefault(); dashRefresh(); return;}
  const rows=[...document.querySelectorAll("#dashView .drow")]; if(!rows.length)return;
  if(k==="j"||k==="ArrowDown"||k==="k"||k==="ArrowUp"){e.preventDefault(); const d=(k==="j"||k==="ArrowDown")?1:-1;
    let i=rows.findIndex(r=>r.classList.contains("sel")); i=i<0?(d>0?0:rows.length-1):Math.max(0,Math.min(rows.length-1,i+d));
    rows.forEach(r=>r.classList.remove("sel")); rows[i].classList.add("sel"); dash.sel=rows[i].dataset.id; rows[i].scrollIntoView({block:"nearest",behavior:"smooth"}); return;}
  const row=document.querySelector("#dashView .drow.sel"), t=row&&data.tasks.find(x=>x.id===row.dataset.id); if(!t)return;
  if(k==="x"||k===" "){e.preventDefault(); dToggle(t,row); return;}
  if(k==="Enter"||k==="o"){e.preventDefault(); openTask(t); return;}}

/* ---------- v8: profile, themes, setup wizard, checklist, recurring, focus timer, weekly review, backup, tour, updates ---------- */
const APP_VERSION="__WTD_VERSION__";
const ACCENTS={violet:["#7c6cff","Violet"],blue:["#3b82f6","Blue"],teal:["#14b8a6","Teal"],green:["#22c55e","Green"],orange:["#f97316","Orange"],pink:["#ec4899","Pink"]};
const DOW=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
ICON.repeat='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 2l3 3-3 3"/><path d="M4 11V9a4 4 0 0 1 4-4h12M7 22l-3-3 3-3"/><path d="M20 13v2a4 4 0 0 1-4 4H4"/></svg>';
ICON.timer='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2M9.5 2.5h5M12 2.5V5"/></svg>';
ICON.play='<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5.5v13l11-6.5z"/></svg>';
ICON.pause='<svg viewBox="0 0 24 24" fill="currentColor"><rect x="6.5" y="5" width="4" height="14" rx="1.2"/><rect x="13.5" y="5" width="4" height="14" rx="1.2"/></svg>';
ICON.review='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19V5M4 19h16"/><path d="M8 15l3-4 3 2 5-6"/></svg>';
ICON.save='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>';
ICON.rocket='<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 15c-1.5 1.3-2 5-2 5s3.7-.5 5-2"/><path d="M9 12a20 20 0 0 1 9-9c1.5 0 3 1.5 3 3a20 20 0 0 1-9 9z"/><path d="M9 12l3 3M12.5 7.5h.01M14 5.5l4.5 4.5"/></svg>';
ICON.gift='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="8" width="18" height="4" rx="1"/><path d="M12 8v13M19 12v9H5v-9M7.5 8a2.5 2.5 0 0 1 0-5C10 3 12 8 12 8s2-5 4.5-5a2.5 2.5 0 0 1 0 5"/></svg>';
const prof=()=>Object.assign({name:"",setup_done:false},(data.settings&&data.settings.profile)||{});
const onb=()=>Object.assign({dismissed:false,tour_seen:false,outlook:"",ai:"",backup_at:""},(data.settings&&data.settings.onboarding)||{});
function putSettings(patch){data.settings=Object.assign({},data.settings||{},patch); save();}
function putOnb(patch){putSettings({onboarding:Object.assign(onb(),patch)});}
function putProf(patch){putSettings({profile:Object.assign(prof(),patch)});}
/* themes */
function applyTheme(theme,accent){const h=document.documentElement; h.dataset.theme=theme==="light"?"light":"dark";
  const a=(ACCENTS[accent]||ACCENTS.violet)[0]; h.style.setProperty("--accent",a);
  try{localStorage.setItem("wtdTheme",JSON.stringify({theme:h.dataset.theme,accent:accent||"violet"}));}catch(e){}}
function themeNow(){const s=data.settings||{}; return {theme:s.theme==="light"?"light":"dark",accent:ACCENTS[s.accent]?s.accent:"violet"};}
function setTheme(theme,accent){const t=themeNow(); theme=theme||t.theme; accent=accent||t.accent; putSettings({theme,accent}); applyTheme(theme,accent);}
function toggleTheme(){const t=themeNow(); setTheme(t.theme==="light"?"dark":"light"); toast(themeNow().theme==="light"?"Light theme":"Dark theme");}
function cycleAccent(){const k=Object.keys(ACCENTS), t=themeNow(); const n=k[(k.indexOf(t.accent)+1)%k.length]; setTheme(null,n); toast("Accent: "+ACCENTS[n][1]);}
function themeCards(onPick){const w=el("div","themes"); const t=themeNow();
  for(const [k,lbl,bg,side,panel,line] of [["dark","Dark","#0d1017","#121621","#1a1f2c","#2a3142"],["light","Light","#f3f4f8","#fbfbfd","#ffffff","#e0e3eb"]]){
    const b=el("button","thc"+(t.theme===k?" on":"")); b.type="button";
    b.innerHTML=`<div class="pv" style="background:${bg}"><i style="background:${side}"></i><div class="pm"><span style="width:60%;background:var(--accent)"></span><span style="background:${panel};border:1px solid ${line}"></span><span style="width:80%;background:${panel};border:1px solid ${line}"></span></div></div><b>${lbl}</b> <kbd>${k[0].toUpperCase()}</kbd>`;
    b.onclick=()=>{setTheme(k); onPick&&onPick();}; w.append(b);}
  return w;}
function accentSwatches(onPick){const w=el("div","swatches"); const t=themeNow(); let i=1;
  for(const [k,[c,lbl]] of Object.entries(ACCENTS)){const b=el("button","swatch"+(t.accent===k?" on":"")); b.type="button"; b.style.background=c; b.title=lbl+" ("+i+")"; b.innerHTML="<small>"+i+"</small>"; i++;
    b.onclick=()=>{setTheme(null,k); onPick&&onPick();}; w.append(b);}
  return w;}
/* recurring tasks */
const ymd=d=>d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate());
const parseYmd=s=>{const [y,m,d]=s.split("-").map(Number); return new Date(y,m-1,d);};
function nextDue(rep,base){const d=new Date(base.getFullYear(),base.getMonth(),base.getDate()); if(!rep)return null;
  if(rep.freq==="daily"){d.setDate(d.getDate()+1); return d;}
  if(rep.freq==="weekdays"){do{d.setDate(d.getDate()+1);}while(d.getDay()===0||d.getDay()===6); return d;}
  if(rep.freq==="weekly"){const dow=rep.dow??base.getDay(); do{d.setDate(d.getDate()+1);}while(d.getDay()!==+dow); return d;}
  if(rep.freq==="monthly"){const dom=Math.max(1,Math.min(31,+(rep.dom||base.getDate()))); const y=d.getFullYear(), m=d.getMonth()+1; const last=new Date(y,m+1,0).getDate(); return new Date(y,m,Math.min(dom,last));}
  return null;}
function firstDue(rep){const t=parseYmd(todayStr());
  if(rep.freq==="weekly"){const d=new Date(t); while(d.getDay()!==+rep.dow)d.setDate(d.getDate()+1); return ymd(d);}
  if(rep.freq==="monthly"){const dom=+rep.dom; const last=new Date(t.getFullYear(),t.getMonth()+1,0).getDate(); const d=new Date(t.getFullYear(),t.getMonth(),Math.min(dom,last)); return d>=t?ymd(d):ymd(nextDue(rep,d));}
  if(rep.freq==="weekdays"){const d=new Date(t); while(d.getDay()===0||d.getDay()===6)d.setDate(d.getDate()+1); return ymd(d);}
  return ymd(t);}
function repLabel(r){if(!r)return ""; if(r.freq==="daily")return "Daily"; if(r.freq==="weekdays")return "Weekdays";
  if(r.freq==="weekly")return "Weekly \u00b7 "+DOW[+r.dow].slice(0,3); if(r.freq==="monthly"){const n=+r.dom; const sfx=n%10===1&&n!==11?"st":n%10===2&&n!==12?"nd":n%10===3&&n!==13?"rd":"th"; return "Monthly \u00b7 "+n+sfx;} return "";}
/* completes (or un-completes) a task; a repeating task gets its next occurrence. Returns an undo function. */
function setDone(t,done){const was={done:t.done,completed:t.completed||"",repeat_next:t.repeat_next||""}; let spawned=null;
  t.done=done; t.completed=done?nowIso():"";
  if(done&&t.repeat&&!t.repeat_next){const today=parseYmd(todayStr()); const base=t.due&&t.due>todayStr()?parseYmd(t.due):today; const nd=nextDue(t.repeat,base);
    if(nd){spawned={id:uid(),project_id:t.project_id,title:t.title,priority:t.priority,due:ymd(nd),due_time:t.due_time||"",done:false,created:nowIso(),completed:"",repeat:Object.assign({},t.repeat)};
      data.tasks.push(spawned); t.repeat_next=spawned.id;}}
  return ()=>{t.done=was.done; t.completed=was.completed; if(spawned){data.tasks=data.tasks.filter(x=>x.id!==spawned.id); if(was.repeat_next)t.repeat_next=was.repeat_next; else delete t.repeat_next;} commit();};}
function repeatFields(rep,prefix){const f=rep?rep.freq:""; const opts=[["","Doesn't repeat"],["daily","Every day"],["weekdays","Every weekday (Mon\u2013Fri)"],["weekly","Every week on\u2026"],["monthly","Every month on day\u2026"]];
  return `<select class="field" id="${prefix}Freq">${opts.map(([v,l])=>`<option value="${v}"${v===f?" selected":""}>${l}</option>`).join("")}</select>
    <select class="field" id="${prefix}Dow"${f==="weekly"?"":" hidden"}>${DOW.map((d,i)=>`<option value="${i}"${rep&&+rep.dow===i?" selected":""}>${d}</option>`).join("")}</select>
    <select class="field" id="${prefix}Dom"${f==="monthly"?"":" hidden"}>${Array.from({length:31},(_,i)=>`<option value="${i+1}"${rep&&+rep.dom===i+1?" selected":""}>Day ${i+1}</option>`).join("")}</select>`;}
function wireRepeat(root,prefix,base){const fq=root.querySelector("#"+prefix+"Freq"), dw=root.querySelector("#"+prefix+"Dow"), dm=root.querySelector("#"+prefix+"Dom");
  const b=base?parseYmd(base):new Date(); if(!fq.value){dw.value=String(b.getDay()); dm.value=String(b.getDate());}
  fq.onchange=()=>{dw.hidden=fq.value!=="weekly"; dm.hidden=fq.value!=="monthly";};
  return ()=>fq.value?(fq.value==="weekly"?{freq:"weekly",dow:+dw.value}:fq.value==="monthly"?{freq:"monthly",dom:+dm.value}:{freq:fq.value}):null;}
function openRepeat(t){const m=openModal(`<div class="mh"><div class="logo">${ICON.repeat.replace('stroke="currentColor"','stroke="#fff"')}</div><h2>Repeat</h2></div>
  <div class="cfsub">${esc(t.title.split("\n")[0].slice(0,90))}</div><div class="nrow" id="rpBox">${repeatFields(t.repeat,"rp")}</div>
  <div class="cfsub" style="margin-top:10px">When you mark it done, the next one is created automatically with the next due date.</div>
  <div class="mrow mdl"><div class="mrow"><button type="button" class="btn ghost" id="rpOff">Stop repeating</button><button type="button" class="btn" id="rpSave">Save <kbd style="margin-left:4px">Enter</kbd></button></div></div>`,"narrow");
  const get=wireRepeat(m,"rp",t.due); const saveR=()=>{const r=get(); if(r){t.repeat=r; if(!t.due)t.due=firstDue(r);} else delete t.repeat; delete t.repeat_next; closeModal(); commit(); toast(r?"Repeats: "+repLabel(r):"Not repeating");};
  m.querySelector("#rpSave").onclick=saveR; m.querySelector("#rpOff").onclick=()=>{delete t.repeat; closeModal(); commit(); toast("Not repeating");};
  modalKeys=(e,k)=>{if(k==="Enter"){e.preventDefault(); saveR(); return true;}}; setTimeout(()=>m.querySelector("#rpFreq").focus(),30);}
/* focus timer (Pomodoro) */
const focusCfg=()=>Object.assign({work:25,brk:5,sound:true},(data.settings&&data.settings.focus)||{});
let focus=null; try{focus=JSON.parse(localStorage.getItem("wtdFocus")||"null");}catch(e){focus=null;}
const fSave=()=>{try{focus?localStorage.setItem("wtdFocus",JSON.stringify(focus)):localStorage.removeItem("wtdFocus");}catch(e){}};
const fLeft=()=>!focus?0:focus.paused!=null?focus.paused:Math.max(0,focus.ends-Date.now());
function chime(){if(!focusCfg().sound)return; try{const C=new (window.AudioContext||window.webkitAudioContext)(); [0,.18,.36].forEach((t,i)=>{const o=C.createOscillator(), g=C.createGain(); o.frequency.value=[660,880,990][i]; g.gain.setValueAtTime(.0001,C.currentTime+t); g.gain.exponentialRampToValueAtTime(.18,C.currentTime+t+.02); g.gain.exponentialRampToValueAtTime(.0001,C.currentTime+t+.35); o.connect(g).connect(C.destination); o.start(C.currentTime+t); o.stop(C.currentTime+t+.4);});}catch(e){}}
function startFocus(t){if(!t)return; if(focus&&focus.task_id===t.id&&focus.phase==="work"){pauseFocus(); return;}
  if(focus&&focus.phase==="work")stopFocus(true);
  const c=focusCfg(); focus={task_id:t.id,title:t.title.split("\n")[0].slice(0,80),phase:"work",len:c.work*60000,ends:Date.now()+c.work*60000,paused:null,start:nowIso()}; fSave(); renderFocus(); toast("Focus: "+c.work+" min on \u201c"+focus.title.slice(0,40)+"\u201d");}
function pauseFocus(){if(!focus||focus.phase==="done")return; if(focus.paused!=null){focus.ends=Date.now()+focus.paused; focus.paused=null; toast("Resumed");} else {focus.paused=fLeft(); toast("Paused");} fSave(); renderFocus();}
function logFocus(minutes){if(!focus||minutes<1)return; const t=data.tasks.find(x=>x.id===focus.task_id);
  data.focus_log=Array.isArray(data.focus_log)?data.focus_log:[]; data.focus_log.push({id:uid(),task_id:focus.task_id,project_id:t?t.project_id:null,title:focus.title,start:focus.start,minutes:Math.round(minutes)});
  if(t)t.focus_min=(+t.focus_min||0)+Math.round(minutes); commit();}
function stopFocus(silent){if(!focus)return; if(focus.phase==="work"){const m=(focus.len-fLeft())/60000; if(m>=1){logFocus(m); if(!silent)toast("Logged "+Math.round(m)+" min");} else if(!silent)toast("Timer stopped (under a minute, not logged)");}
  else if(!silent)toast("Timer stopped"); focus=null; fSave(); renderFocus();}
function focusTick(){if(!focus)return; if(focus.paused==null&&focus.phase!=="done"&&Date.now()>=focus.ends){
    if(focus.phase==="work"){logFocus(focus.len/60000); const c=focusCfg(); focus.phase="break"; focus.len=c.brk*60000; focus.ends=Date.now()+focus.len; chime(); toast("Nice work! Take a "+c.brk+"-minute break.");}
    else {focus.phase="done"; chime(); toast("Break's over. Press F to start another round.");}
    fSave();}
  renderFocus();}
function renderFocus(){let p=$("#fpill"); if(!focus){if(p)p.hidden=true; document.title=document.title.replace(/^\u23f1 \S+ /,""); return;}
  if(!p){p=el("div","fpill"); p.id="fpill"; document.body.append(p);}
  p.hidden=false; const left=fLeft(), mm=Math.floor(left/60000), ss=Math.floor(left%60000/1000), frac=focus.phase==="done"?1:1-left/focus.len;
  p.className="fpill"+(focus.phase!=="work"?" brk":"")+(focus.paused!=null?" paused":"");
  const col=focus.phase==="work"?"var(--accent)":"var(--ok)", C=2*Math.PI*16;
  const label=focus.phase==="work"?"Focus":focus.phase==="break"?"Break":"Break over";
  p.innerHTML=`<svg class="ring" viewBox="0 0 38 38"><circle cx="19" cy="19" r="16" stroke="color-mix(in srgb,var(--ink) 10%,transparent)"/><circle cx="19" cy="19" r="16" stroke="${col}" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C*(1-frac)}" transform="rotate(-90 19 19)"/></svg>
   <div class="ft"><b>${focus.phase==="done"?"0:00":mm+":"+pad2(ss)}</b><small>${label} \u00b7 ${esc(focus.title)}</small></div>`;
  const pb=el("button","ibtn",focus.paused!=null||focus.phase==="done"?ICON.play:ICON.pause); pb.title=focus.phase==="done"?"Start another round (F)":"Pause / resume (Shift+F)";
  pb.onclick=()=>{if(focus.phase==="done"){const t=data.tasks.find(x=>x.id===focus.task_id); focus=null; startFocus(t);} else pauseFocus();};
  const sb=el("button","ibtn",ICON.x); sb.title="Stop and log (Ctrl+Shift+F)"; sb.onclick=()=>stopFocus(); p.append(pb,sb);
  if(focus.phase==="work"&&focus.paused==null){const base=document.title.replace(/^\u23f1 \S+ /,""); document.title="\u23f1 "+mm+":"+pad2(ss)+" "+base;}}
setInterval(focusTick,1000);
const focusMinutes=(from,to,pid)=>(data.focus_log||[]).filter(x=>{const d=(x.start||"").slice(0,10); return d>=from&&d<=to&&(pid===undefined||x.project_id===pid);}).reduce((a,x)=>a+(+x.minutes||0),0);
const fmtMin=m=>m>=60?Math.floor(m/60)+"h "+(m%60?m%60+"m":""):m+"m";
/* backup & restore */
async function exportBackup(){try{await saveNow(); const r=await fetch("/api/export"); if(!r.ok)throw new Error("HTTP "+r.status); const blob=await r.blob();
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="WorkToDo-backup-"+todayStr()+".json"; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),4000);
    let path=""; try{const b=await fetch("/api/backup-now",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:"{}"}); if(b.ok)path=(await b.json()).path;}catch(e){}
    putOnb({backup_at:nowIso()}); toast("Backup saved to Downloads"+(path?" (and a copy in the app's backups folder)":"")); if(view==="dash")renderDash();}
  catch(e){toast("Couldn't make a backup: "+e.message);}}
function pickRestore(){const i=document.createElement("input"); i.type="file"; i.accept=".json,application/json";
  i.onchange=async()=>{const f=i.files[0]; if(!f)return; let j; try{j=JSON.parse(await f.text());}catch(e){toast("That isn't a JSON backup file"); return;}
    const d=j&&j.data?j.data:j; if(!d||!Array.isArray(d.projects)||!Array.isArray(d.tasks)){toast("That file isn't a Work To-Do backup"); return;}
    const m=openModal(`<div class="mh"><div class="logo">${ICON.save.replace('stroke="currentColor"','stroke="#fff"')}</div><h2>Restore this backup?</h2></div>
      <div class="cfsub">${esc(f.name)}${j.exported?" \u00b7 made "+esc(fmtWhen(j.exported)):""}</div>
      <ul class="aichg big"><li>${d.projects.length} projects, ${d.tasks.length} tasks, ${(d.notes||[]).length} notes</li><li>This replaces what's in the app now (${data.projects.length} projects, ${data.tasks.length} tasks, ${data.notes.length} notes).</li><li>A copy of your current data is saved to the backups folder first, so you can undo it.</li></ul>
      <div class="mdl"><div class="mrow"><button type="button" class="btn ghost" id="rsNo">Cancel</button><button type="button" class="btn" id="rsYes">Restore <kbd style="margin-left:4px">Ctrl+Enter</kbd></button></div></div>`,"narrow");
    const go2=async()=>{try{const r=await fetch("/api/restore",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:JSON.stringify({backup:j})});
        const out=await r.json(); if(!r.ok)throw new Error(out.error||"HTTP "+r.status); data=normData(out.data); lastRev=r.headers.get("X-Rev")||lastRev; closeModal(); applyTheme(themeNow().theme,themeNow().accent); render(); toast("Restored. Your previous data is in the backups folder.");}
      catch(e){toast("Restore failed: "+e.message);}};
    m.querySelector("#rsNo").onclick=closeModal; m.querySelector("#rsYes").onclick=go2; modalKeys=(e,k)=>{if(k==="Enter"&&(e.ctrlKey||e.metaKey)){e.preventDefault(); go2(); return true;}};};
  i.click();}
function normData(d){d=d||{}; d.projects=d.projects||[]; d.tasks=d.tasks||[]; d.notes=Array.isArray(d.notes)?d.notes:[]; return d;}
/* get set up checklist */
function setupItems(){const o=onb(), st=((em.inbox||{}).status||{}).state||"", c=aiCfg();
  return [["name","Add your name",!!prof().name.trim(),()=>openSettings("profile")],
    ["project","Create your first project",data.projects.length>0,()=>{$("#newpName").focus(); toast("Type a project name on the left, then Enter");}],
    ["task","Add your first task",data.tasks.length>0,()=>newTask()],
    ["outlook","Connect Outlook (or skip it)",o.outlook==="skipped"||o.outlook==="connected"||st==="ok"||st==="ready",()=>go("email"),()=>{putOnb({outlook:"skipped"}); putSettings({email:Object.assign({},(data.settings||{}).email||{},{auto:false})}); renderDash();}],
    ["ai","Turn on AI (or skip it)",o.ai==="skipped"||(c.enabled&&c.provider==="openai_compatible"),()=>openSettings("ai"),()=>{putOnb({ai:"skipped"}); renderDash();}],
    ["tour","Take the shortcuts tour",!!o.tour_seen,()=>startTour()],
    ["backup","Make your first backup",!!o.backup_at,()=>exportBackup()]];}
function setupCard(){const o=onb(); if(o.dismissed)return null; const items=setupItems(), done=items.filter(x=>x[2]).length;
  const c=el("div","setup"); const h=el("div","sh"); h.innerHTML=`<span class="sicon" style="color:var(--accent)">${ICON.rocket.replace('stroke="#fff"','stroke="currentColor"')}</span><b>${done===items.length?"You're all set!":"Get set up"}</b><span class="pill${done===items.length?" ok":""}">${done} of ${items.length}</span>`;
  const x=el("button","ibtn",ICON.x); x.title="Dismiss the checklist (C). Bring it back in Settings."; x.onclick=()=>{putOnb({dismissed:true}); renderDash(); toast("Checklist hidden. Settings \u2192 Profile brings it back.");}; h.append(x); c.append(h);
  const bar=el("div","pbar"); const sp=el("span"); sp.style.width=(done/items.length*100)+"%"; bar.append(sp); c.append(bar);
  const g=el("div","sgrid");
  for(const [k,label,ok,act,skip] of items){const r=el("div","sitem"+(ok?" done":"")); r.dataset.k=k; r.innerHTML=`<span class="ck">${ICON.check}</span>`; const t=el("span","tx"); t.textContent=label; r.append(t);
    if(!ok){const b=el("button","mlink","Do it"); b.type="button"; b.onclick=act; r.append(b); if(skip){const s2=el("button","mlink","Skip"); s2.type="button"; s2.style.color="var(--muted)"; s2.onclick=skip; r.append(s2);}}
    g.append(r);}
  c.append(g); return c;}
/* weekly review */
let rv={offset:0,recap:null,recapKey:"",busy:false};
function weekRange(off){const t=parseYmd(todayStr()); const mon=new Date(t); mon.setDate(t.getDate()-((t.getDay()+6)%7)+off*7); const sun=new Date(mon); sun.setDate(mon.getDate()+6); return [ymd(mon),ymd(sun)];}
async function loadRecap(ai){const [S,E]=weekRange(rv.offset); const key=S+"|"+(ai?"ai":"local")+"|"+data.tasks.length+"|"+(data.focus_log||[]).length;
  if(!ai&&rv.recap&&rv.recapKey===key)return; if(rv.busy)return; rv.busy=ai?"ai":"local"; if(view==="review")renderReview();
  try{await saveNow(); const r=await fetch("/api/review",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:JSON.stringify({start:S,end:E,ai:!!ai})});
    const j=await r.json(); if(!r.ok)throw new Error(j.error||"HTTP "+r.status); rv.recap=j; rv.recapKey=key; if(ai)toast(j.source==="ai"?"AI recap ready":"Built-in recap (AI didn't answer)");}
  catch(e){toast("Couldn't make the recap: "+e.message);}
  rv.busy=false; if(view==="review")renderReview();}
function renderReview(){const [S,E]=weekRange(rv.offset); const fmt=s=>parseYmd(s).toLocaleDateString(undefined,{month:"short",day:"numeric"});
  const h=$("#h1"); h.innerHTML=""; const ic=el("span","sicon",ICON.review); ic.style.cssText="width:28px;height:28px;color:var(--accent)"; ic.querySelector("svg").style.cssText="width:26px;height:26px"; h.append(ic,document.createTextNode("Weekly review"));
  $("#sub").textContent=fmt(S)+" \u2013 "+fmt(E)+(rv.offset===0?" \u00b7 this week":rv.offset===-1?" \u00b7 last week":"");
  const w=$("#reviewView"); w.innerHTML=""; w.className="rv dash";
  const td=todayStr(), [NS,NE]=weekRange(rv.offset+1);
  const done=data.tasks.filter(t=>t.done&&(t.completed||"").slice(0,10)>=S&&(t.completed||"").slice(0,10)<=E);
  const slipped=data.tasks.filter(t=>!t.done&&t.due&&t.due<td&&t.due<=E);
  const next=data.tasks.filter(t=>!t.done&&t.due&&t.due>=NS&&t.due<=NE).sort((a,b)=>(a.due+(a.due_time||"")).localeCompare(b.due+(b.due_time||"")));
  const ft=focusMinutes(S,E);
  const s=el("div","stats"); for(const [l,n,c] of [["Done this week",done.length,"var(--ok)"],["Slipped / overdue",slipped.length,"var(--urgent)"],["Due next week",next.length,"var(--accent2)"],["Focused time",fmtMin(ft),"var(--accent)"]]){
    const d=el("div","stat"); d.style.setProperty("--c",c); d.append(el("div","n",String(n)),el("div","l",l)); d.querySelector(".n").style.color=c; s.append(d);} w.append(s);
  const nav=el("div","toolbar"); const wk=el("div","wk"); const pb=el("button","btn ghost sm","\u2190 Previous"); pb.type="button"; pb.onclick=()=>{rv.offset--; rv.recap=null; renderReview(); loadRecap(false);};
  const nb=el("button","btn ghost sm","Next \u2192"); nb.type="button"; nb.disabled=rv.offset>=0; nb.onclick=()=>{if(rv.offset<0){rv.offset++; rv.recap=null; renderReview(); loadRecap(false);}};
  wk.append(pb,keysHtml("["),nb,keysHtml("]")); nav.append(wk); w.append(nav);
  // recap
  const rc=el("div","hero recap"); rc.style.marginBottom="16px"; const b=el("div","brief"); const r=rv.recap;
  const bh=el("div","bh",ICON.spark); bh.append(document.createTextNode(r&&r.source==="ai"?"AI recap":"Your week in brief")); const pl=el("span","pill"+(r&&r.source==="ai"?" ok":"")); pl.textContent=r&&r.source==="ai"?(r.model||"AI"):"Built-in \u00b7 on this PC"; bh.append(pl); b.append(bh);
  if(!r){b.append(el("div","shimmer"),el("div","shimmer"));}
  else{const h3=el("h3"); h3.textContent=r.headline; b.append(h3);
    for(const [k,lbl] of [["wins","Wins"],["slipped","What slipped"],["next","For next week"]]){if(!(r[k]||[]).length)continue; b.append(el("h4","",lbl)); const ul=el("ul"); for(const x of r[k]){const li=el("li"); li.textContent=x; ul.append(li);} b.append(ul);}
    if(r.note){const n=el("div","bnote"); n.textContent=r.note; b.append(n);}}
  if(rv.busy==="ai")b.append(el("div","bnote","Asking your AI model\u2026 this can take up to a minute."));
  const acts=el("div","bacts"); const on=aiOn()&&aiCfg().provider==="openai_compatible";
  const ab=el("button","btn sm",ICON.spark.replace('stroke="currentColor"','stroke="#fff"')+"<span>"+(rv.busy==="ai"?"Writing\u2026":"AI recap")+"</span>"); ab.type="button"; ab.disabled=!!rv.busy||!on; ab.title=on?"Ask your AI model for a recap (A)":"Set up AI in Reminders & settings to use this"; ab.onclick=()=>loadRecap(true);
  acts.append(ab,keysHtml("A")); rc.append(b,acts); w.append(rc);
  const g=el("div","rgrid"); w.append(g);
  const list=(card,ts,right)=>{for(const t of ts.slice(0,12))card.append(dTaskRow(t,right(t))); if(ts.length>12){const m=el("div","dsub"); m.textContent="+ "+(ts.length-12)+" more"; card.append(m);}};
  const c1=dCard("checkc","var(--ok)","Done this week",done.length,null); if(!done.length)c1.append(el("div","dempty","Nothing marked done in this week yet.")); list(c1,[...done].sort((a,b)=>(b.completed||"").localeCompare(a.completed||"")),t=>relDay((t.completed||"").slice(0,10))); g.append(c1);
  const c2=dCard("alarm","var(--urgent)","Slipped / overdue",slipped.length,["Overdue",()=>go("overdue"),"G then O"]); if(!slipped.length)c2.append(dEmpty("Nothing slipped. Great job.")); list(c2,[...slipped].sort((a,b)=>a.due.localeCompare(b.due)),t=>relDay(t.due)); g.append(c2);
  const c3=dCard("up","var(--accent2)","Next week",next.length,["Upcoming",()=>go("upcoming"),"G then U"]); if(!next.length)c3.append(el("div","dempty","Nothing due next week yet. Give your top tasks a date.")); list(c3,next,t=>relDay(t.due)+(t.due_time?" \u00b7 "+fmtTime(t.due_time):"")); g.append(c3);
  const c4=dCard("timer","var(--accent)","Time per project",ft?fmtMin(ft):null,null);
  const per=[...data.projects.map(p=>({p,m:focusMinutes(S,E,p.id)})),{p:{name:"No project",color:"#8a93a8"},m:(data.focus_log||[]).filter(x=>{const d=(x.start||"").slice(0,10); return d>=S&&d<=E&&!proj(x.project_id);}).reduce((a,x)=>a+(+x.minutes||0),0)}].filter(x=>x.m>0).sort((a,b)=>b.m-a.m);
  if(!per.length)c4.append(el("div","dempty","No focus time logged this week. Select a task on the dashboard and press F."));
  const mx=Math.max(1,...per.map(x=>x.m));
  for(const {p,m} of per){const r2=el("div","hbar"); const nm=el("span","nm"); const d=el("span","dot"); d.style.background=p.color; const t=el("span"); t.textContent=p.name; nm.append(d,t);
    const tr=el("div","tr"); const sp=el("span"); sp.style.cssText=`width:${m/mx*100}%;background:${p.color}`; tr.append(sp); const v=el("span","v"); v.textContent=fmtMin(m); r2.append(nm,tr,v); c4.append(r2);}
  g.append(c4);}
function reviewKey(e,k){if(k==="["){e.preventDefault(); rv.offset--; rv.recap=null; renderReview(); loadRecap(false); return;}
  if(k==="]"){e.preventDefault(); if(rv.offset<0){rv.offset++; rv.recap=null; renderReview(); loadRecap(false);} return;}
  if(k==="a"||k==="A"){e.preventDefault(); if(aiOn()&&aiCfg().provider==="openai_compatible")loadRecap(true); else toast("Set up AI in Reminders & settings to use the AI recap"); return;}
  const rows=[...document.querySelectorAll("#reviewView .drow")]; if(!rows.length)return;
  if(k==="j"||k==="k"||k==="ArrowDown"||k==="ArrowUp"){e.preventDefault(); const d=(k==="j"||k==="ArrowDown")?1:-1; let i=rows.findIndex(r=>r.classList.contains("sel")); i=i<0?(d>0?0:rows.length-1):Math.max(0,Math.min(rows.length-1,i+d)); rows.forEach(r=>r.classList.remove("sel")); rows[i].classList.add("sel"); dash.sel=rows[i].dataset.id; rows[i].scrollIntoView({block:"nearest"}); return;}
  const row=document.querySelector("#reviewView .drow.sel"), t=row&&data.tasks.find(x=>x.id===row.dataset.id); if(!t)return;
  if(k==="Enter"||k==="o"){e.preventDefault(); openTask(t);}}
/* tour */
const TOUR=[["grid","Your dashboard","Everything for today in one place: what's overdue or due, your schedule, priorities by project, new emails and a brief for the day.",[["Ctrl+D or G then D","Dashboard"],["B","AI brief"]]],
  ["edit","Add tasks fast","Type a task and press Enter. Pick a priority, a due date and time for a reminder, and make it repeat (daily, weekdays, weekly or monthly).",[["N","New task"],["Y","Repeat the selected task"]]],
  ["spark","Quick add from anywhere","Press Ctrl+Alt+N anywhere in Windows, even with this window closed, to pop up a tiny box. Try \u201cCall Sam tomorrow 3pm high\u201d.",[["Ctrl+Alt+N","Quick add (anywhere)"],["Ctrl+Alt+T","Open Work To-Do"]]],
  ["timer","Focus timer","Select a task on the dashboard and press F for a 25-minute focus round and a 5-minute break. Time is logged per task and project.",[["F","Start focus on the selected task"],["Shift+F","Pause / resume"]]],
  ["review","Weekly review","See what got done, what slipped, what's next week and where your time went. Ask AI for a recap if you've set it up.",[["G then W","Weekly review"],["A","AI recap"]]],
  ["save","Back up and restore","Save everything to one file whenever you like, and restore it on any PC. Restoring always keeps a copy of what was there first.",[["Ctrl+Shift+E","Back up"],["Ctrl+Shift+O","Restore"]]],
  ["sun","Make it yours","Switch between light and dark and pick an accent colour. Press ? any time for every shortcut, and replay this tour from there.",[["Ctrl+Shift+L","Light / dark"],["?","All shortcuts"]]]];
let tourI=0;
function startTour(){tourI=0; showTour();}
function showTour(){const [ic,title,text,keys]=TOUR[tourI]; const last=tourI===TOUR.length-1;
  const m=openModal(`<div class="tourbox"><div class="mh"><div class="wdots" style="margin-left:0">${TOUR.map((_,i)=>`<i class="${i===tourI?"on":i<tourI?"done":""}"></i>`).join("")}</div><span style="flex:1"></span></div>
   <div class="tourart">${(ICON[ic]||ICON.spark).replace(/stroke="(currentColor|#fff)"/,'stroke="#fff"')}</div><h2>${esc(title)}</h2><p>${esc(text)}</p><div class="kcol" id="tourKeys"></div>
   <div class="wnav"><span class="hint">${tourI+1} of ${TOUR.length} \u00b7 <kbd>\u2190</kbd> <kbd>\u2192</kbd> \u00b7 <kbd>Esc</kbd> closes</span><span class="sp"></span>${tourI?'<button type="button" class="btn ghost sm" id="tBack">Back</button>':""}<button type="button" class="btn sm" id="tNext">${last?"Done":"Next"}</button></div></div>`,"narrow");
  const kc=m.querySelector("#tourKeys"); for(const [k,d] of keys){const r=el("div","krow"); const s=el("span"); s.textContent=d; r.append(s,keysHtml(k)); kc.append(r);}
  const next=()=>{if(last){finishTour();} else {tourI++; showTour();}}, back=()=>{if(tourI){tourI--; showTour();}};
  m.querySelector("#tNext").onclick=next; const bb=m.querySelector("#tBack"); if(bb)bb.onclick=back;
  modalKeys=(e,k)=>{if(k==="ArrowRight"||k==="Enter"){e.preventDefault(); next(); return true;} if(k==="ArrowLeft"){e.preventDefault(); back(); return true;} if(k==="Escape"){e.preventDefault(); finishTour(); return true;}};}
function finishTour(){putOnb({tour_seen:true}); closeModal(); toast("Press ? any time to see shortcuts or replay the tour");}
/* setup wizard */
let wz={i:0,steps:["welcome","name","theme","outlook","ai","reminders","finish"],msg:""};
function openWizard(){wz.i=0; let w=$("#wiz"); if(!w){w=el("div","wiz"); w.id="wiz"; document.body.append(w);} w.hidden=false; renderWizard();}
function wizOpen(){const w=$("#wiz"); return !!w&&!w.hidden;}
function wizNext(){const st=wz.steps[wz.i];
  if(st==="name"){const v=($("#wzName").value||"").trim(); if(!v){wz.msg="Type your name (a first name is fine)."; renderWizard(); return;}
    const emc=Object.assign({},(data.settings||{}).email||{}); if(!emc.name)emc.name=v; putSettings({profile:Object.assign(prof(),{name:v}),email:emc});}
  if(st==="reminders"){const t=$("#wzMorning").value; putSettings({morning:t||"09:00",lead:+$("#wzLead").value});}
  if(st==="finish"){finishWizard(); return;}
  wz.msg=""; wz.i=Math.min(wz.steps.length-1,wz.i+1); renderWizard();}
function wizBack(){if(wz.i){wz.i--; wz.msg=""; renderWizard();}}
function finishWizard(){putSettings({profile:Object.assign(prof(),{setup_done:true,setup_at:nowIso()})}); $("#wiz").hidden=true; saveNow().catch(()=>{}); go("dash");
  if(emCfg().auto&&onb().outlook==="connected")refreshEmail(true);
  if(!onb().tour_seen)setTimeout(startTour,350);}
function renderWizard(){const w=$("#wiz"); const st=wz.steps[wz.i], s=settings(), o=onb(), c=aiCfg();
  const dots=wz.steps.map((_,i)=>`<i class="${i===wz.i?"on":i<wz.i?"done":""}"></i>`).join("");
  let body="";
  if(st==="welcome")body=`<h2>Welcome to Work To-Do \ud83d\udc4b</h2><p>A fast, private to-do app for your projects. Everything stays on this PC. Let's set it up for you. It takes about a minute, and you can change any of it later in Settings.</p>
    <div class="wfeat"><div><b>Dashboard</b>Today at a glance, with an optional AI brief.</div><div><b>Quick add</b>Ctrl+Alt+N from anywhere in Windows.</div><div><b>Focus & review</b>Pomodoro timer and a weekly review.</div><div><b>Recurring tasks</b>Daily, weekdays, weekly, monthly.</div><div><b>Reminders</b>Windows notifications before things are due.</div><div><b>Backup</b>One file, restore anywhere.</div></div>`;
  if(st==="name")body=`<h2>What should we call you?</h2><p>Used for your greeting and to spot emails meant for you. It stays on this PC.</p><input class="big" id="wzName" placeholder="Your name" autocomplete="off" value="${esc(prof().name)}">`;
  if(st==="theme")body=`<h2>Pick a look</h2><p>Light or dark, plus an accent colour. Press <kbd>L</kbd> or <kbd>D</kbd>, and <kbd>1</kbd>\u2013<kbd>6</kbd> for colours.</p><div id="wzThemes"></div><div style="padding-bottom:18px" id="wzAcc"></div>`;
  if(st==="outlook")body=`<h2>Connect Outlook? <span class="pill">optional</span></h2><p>Work To-Do can show emails sent to you from <b>classic Outlook</b> on this PC, so you can turn them into tasks with one click. It's read-only: nothing is ever sent, moved, deleted or marked read.</p>
    <div class="wopt"><div class="bi">${ICON.mail}</div><div class="bt"><b>Connect classic Outlook</b><br>Opens classic Outlook so you can sign in once (if it isn't set up yet). Then your inbox shows up on the dashboard and in Email. New Outlook and web Outlook don't support this, but you can still drag or paste emails in.</div></div>
    <div class="wrow"><button type="button" class="btn" id="wzOlYes">Connect Outlook</button><button type="button" class="btn ghost" id="wzOlNo">Skip for now <kbd style="margin-left:4px">S</kbd></button></div><div class="wmsg${o.outlook?" ok":""}" id="wzOlMsg">${o.outlook==="connected"?"\u2713 Outlook inbox turned on.":o.outlook==="skipped"?"Skipped. You can connect later from Email.":""}</div>`;
  if(st==="ai")body=`<h2>Turn on AI? <span class="pill">optional</span></h2><p>AI can write a daily brief and a weekly recap, suggest priorities and turn a sentence into a task. It uses <b>OpenRouter</b> (openrouter.ai), which has free and paid models. Without it, everything still works with built-in rules.</p>
    <div class="wopt"><div class="bi">${ICON.spark}</div><div class="bt"><b>Option 1: set it yourself</b><br>Create a key at openrouter.ai, then in Windows search \u201cEdit environment variables for your account\u201d and add <code>OPENROUTER_API_KEY</code>. Restart Work To-Do.<br><b>Option 2: paste it here</b><br>It's saved only in <b>your own Windows user environment variable</b> (<code>OPENROUTER_API_KEY</code>), never in the app's files.</div></div>
    <div class="wrow"><input type="password" class="field" id="wzKey" placeholder="Paste your OpenRouter key (optional)" autocomplete="off"><button type="button" class="btn ghost" id="wzKeySave">Save key</button></div>
    <div class="wrow"><label style="flex:1"><span class="cflbl">Model (from openrouter.ai/models; ids ending in :free cost nothing)</span><input class="field" id="wzModel" style="width:100%" value="${esc(c.provider==="openai_compatible"&&c.model?c.model:"openrouter/auto")}"></label></div>
    <div class="wrow"><button type="button" class="btn" id="wzAiOn">Turn on AI</button><button type="button" class="btn ghost" id="wzAiNo">Skip for now <kbd style="margin-left:4px">S</kbd></button></div><div class="wmsg" id="wzAiMsg">${o.ai==="skipped"?"Skipped. Turn it on any time in Settings.":c.enabled?"\u2713 AI is on.":""}</div>`;
  if(st==="reminders")body=`<h2>Reminders</h2><p>Windows notifications for what's due, even when this window is closed. A small helper runs in the background for this, and it also powers the <kbd>Ctrl+Alt+N</kbd> quick-add box.</p>
    <div class="srow"><div class="st"><b>Desktop reminders</b><small>Due-today summary, a heads-up before due times, and overdue nudges.</small></div><div class="toggle${s.reminders?" on":""}" id="wzRem"><div class="sw"></div></div></div>
    <div class="srow"><div class="st"><b>Morning summary</b><small>What's due today and overdue, once a day.</small></div><input type="time" class="field" id="wzMorning" value="${esc(s.morning)}"></div>
    <div class="srow"><div class="st"><b>Heads-up before a due time</b></div><select class="field" id="wzLead">${[0,5,10,15,30,60].map(n=>`<option value="${n}"${+s.lead===n?" selected":""}>${n?n+" minutes before":"Only at the due time"}</option>`).join("")}</select></div>
    <div class="wrow"><button type="button" class="btn ghost sm" id="wzRemStart">Start the helper now</button><span class="wmsg" id="wzRemMsg" style="margin:0"></span></div>`;
  if(st==="finish")body=`<h2>You're all set${prof().name?", "+esc(prof().name.split(" ")[0]):""}! \ud83c\udf89</h2><p>Your dashboard has a short <b>Get set up</b> checklist to help you add your first project and task. Next you'll get a 30-second tour of the shortcuts.</p>
    <div class="wfeat"><div><b>Ctrl+Alt+T</b>Open Work To-Do</div><div><b>Ctrl+Alt+N</b>Quick add from anywhere</div><div><b>?</b>All shortcuts and the tour</div></div>`;
  w.innerHTML=`<div class="wbox"><div class="wtop"><div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.2 4.2L19 7"/></svg></div><div><b>Work To-Do setup</b><small>Step ${wz.i+1} of ${wz.steps.length} \u00b7 v${APP_VERSION}</small></div><div class="wdots">${dots}</div></div>
    <div class="wstep">${body}<div class="wmsg bad" id="wzMsg">${esc(wz.msg)}</div></div>
    <div class="wnav"><span class="hint"><kbd>Enter</kbd> next${wz.i?" \u00b7 <kbd>Alt+\u2190</kbd> back":""}</span><span class="sp"></span>${wz.i?'<button type="button" class="btn ghost" id="wzBack">Back</button>':""}<button type="button" class="btn" id="wzNext">${st==="welcome"?"Let's go":st==="finish"?"Open my dashboard":"Next"}</button></div></div>`;
  w.querySelector("#wzNext").onclick=wizNext; const bk=w.querySelector("#wzBack"); if(bk)bk.onclick=wizBack;
  if(st==="name"){const n=w.querySelector("#wzName"); setTimeout(()=>{n.focus(); n.select();},40);}
  if(st==="theme"){w.querySelector("#wzThemes").append(themeCards(renderWizard)); w.querySelector("#wzAcc").append(accentSwatches(renderWizard));}
  if(st==="outlook"){w.querySelector("#wzOlYes").onclick=async()=>{putOnb({outlook:"connected"}); putSettings({email:Object.assign({},(data.settings||{}).email||{},{auto:true,name:((data.settings||{}).email||{}).name||prof().name})});
      const msg=$("#wzOlMsg"); msg.className="wmsg ok"; msg.textContent="\u2713 Outlook inbox turned on.";
      try{await saveNow(); await emPost("setup-outlook",{}); msg.textContent="\u2713 Classic Outlook is opening. Sign in with your work account, then come back here.";}catch(e){msg.textContent="\u2713 Inbox turned on. ("+e.message+". Set up classic Outlook later from Email.)";}};
    w.querySelector("#wzOlNo").onclick=()=>{wzSkip();};}
  if(st==="ai"){const msg=w.querySelector("#wzAiMsg");
    fetch("/api/ai/status").then(r=>r.json()).then(j=>{if(j.openrouter_env_set&&!msg.textContent){msg.className="wmsg ok"; msg.textContent="\u2713 OPENROUTER_API_KEY is already set on this PC.";}}).catch(()=>{});
    w.querySelector("#wzKeySave").onclick=async()=>{const k=w.querySelector("#wzKey"); if(!k.value.trim()){k.focus(); return;}
      try{const r=await fetch("/api/setup/ai-key",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:JSON.stringify({key:k.value.trim()})}); const j=await r.json(); k.value="";
        if(!r.ok)throw new Error(j.error||"HTTP "+r.status); msg.className="wmsg ok"; msg.textContent="\u2713 Saved to "+j.saved+". Now press Turn on AI.";}catch(e){msg.className="wmsg bad"; msg.textContent="Couldn't save the key: "+e.message;}};
    w.querySelector("#wzAiOn").onclick=()=>{const model=(w.querySelector("#wzModel").value||"").trim()||"openrouter/auto";
      putSettings({ai:Object.assign({},aiCfg(),{enabled:true,provider:"openai_compatible",base_url:"https://openrouter.ai/api/v1",model,api_key_env:"OPENROUTER_API_KEY"})}); putOnb({ai:"set"});
      msg.className="wmsg ok"; msg.textContent="\u2713 AI is on with "+model+". Change it any time in Settings."; setTimeout(wizNext,700);};
    w.querySelector("#wzAiNo").onclick=()=>wzSkip();}
  if(st==="reminders"){w.querySelector("#wzRem").onclick=e=>{const on=!settings().reminders; e.currentTarget.classList.toggle("on",on); putSettings({reminders:on});};
    w.querySelector("#wzRemStart").onclick=async()=>{const m=$("#wzRemMsg"); try{await fetch("/api/reminder-start",{method:"POST",headers:{"X-WorkTodo":"1"}}); m.className="wmsg ok"; m.textContent="\u2713 Started.";}catch(e){m.textContent="Couldn't start it";}};}}
function wzSkip(){const st=wz.steps[wz.i];
  if(st==="outlook"){putOnb({outlook:"skipped"}); putSettings({email:Object.assign({},(data.settings||{}).email||{},{auto:false})}); wz.i++; renderWizard();}
  else if(st==="ai"){putOnb({ai:"skipped"}); putSettings({ai:Object.assign({},aiCfg(),{enabled:false})}); wz.i++; renderWizard();}}
function wizardKey(e){const k=e.key, typing=isTyping(e.target), st=wz.steps[wz.i];
  if(k==="Enter"&&!(e.target&&e.target.tagName==="BUTTON")){e.preventDefault(); if(st==="ai"&&e.target.id==="wzKey"){$("#wzKeySave").click(); return;} wizNext(); return;}
  if(k==="ArrowLeft"&&e.altKey){e.preventDefault(); wizBack(); return;}
  if(typing)return;
  if((k==="s"||k==="S")&&(st==="outlook"||st==="ai")){e.preventDefault(); wzSkip(); return;}
  if(st==="theme"){if(k==="l"||k==="L"){setTheme("light"); renderWizard();} else if(k==="d"||k==="D"){setTheme("dark"); renderWizard();} else if(/^[1-6]$/.test(k)){setTheme(null,Object.keys(ACCENTS)[+k-1]); renderWizard();}}}
document.addEventListener("keydown",e=>{if(wizOpen()){wizardKey(e); e.stopImmediatePropagation();}},true);
/* updates (GitHub Releases) */
let upd=null, updHidden=sessionStorage.getItem("updHidden")||"";
async function loadUpdate(){try{upd=await (await fetch("/api/update/status")).json();}catch(e){upd=null;} if(view==="dash")renderDash();}
async function checkUpdates(msgEl){if(msgEl)msgEl.textContent="Checking\u2026"; try{const r=await fetch("/api/update/check",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:"{}"}); const j=await r.json(); if(!r.ok)throw new Error(j.error||"HTTP "+r.status); upd=j;
    const l=j.last||{}; const t=j.available?"Version "+l.tag+" is available.":l.state==="ok"?"You're up to date (v"+APP_VERSION+").":l.state==="none"?"No releases published yet. You're on v"+APP_VERSION+".":l.state==="offline"?"Couldn't reach GitHub. Try again later.":"Couldn't check ("+(l.detail||"error")+").";
    if(msgEl)msgEl.textContent=t; else toast(t); updHidden=""; if(view==="dash")renderDash();}
  catch(e){if(msgEl)msgEl.textContent="Couldn't check: "+e.message; else toast("Couldn't check: "+e.message);}}
function updateBanner(){if(!upd||!upd.available||!upd.last)return null; const l=upd.last; if(updHidden===l.tag)return null;
  const b=el("div","ubanner"); b.append(el("div","bi",ICON.gift)); const t=el("div","bt"); const h=el("b"); h.textContent="Update available: "+(l.name||l.tag); t.append(h);
  const sm=el("small"); sm.style.cssText="display:block;color:var(--muted)"; sm.textContent="You have v"+APP_VERSION+(l.published?" \u00b7 released "+fmtWhen(l.published):"")+(l.sha256||l.sha_asset_url?" \u00b7 checksum verified on download":""); t.append(sm);
  if(l.notes){const pre=el("pre"); pre.textContent=l.notes.replace(/^\s*sha-?256\s*[:=]?.*$/gim,"").trim().slice(0,3000); if(pre.textContent)t.append(pre);}
  const a=el("div","eacts"); const go2=el("button","btn sm",upd.can_apply?"Update now":"Open release page"); go2.type="button"; go2.onclick=()=>upd.can_apply?applyUpdate():window.open(l.url,"_blank");
  a.append(go2); if(upd.can_apply)a.append(keysHtml("Ctrl+Shift+U")); const later=el("button","btn ghost sm","Later"); later.type="button"; later.onclick=()=>{updHidden=l.tag; sessionStorage.setItem("updHidden",l.tag); renderDash();}; a.append(later);
  if(!upd.can_apply){const n=el("span"); n.style.cssText="color:var(--muted);font-size:12px"; n.textContent="This copy wasn't installed from a release, so download and run the new installer."; a.append(n);}
  t.append(a); b.append(t); return b;}
async function applyUpdate(){if(!upd||!upd.available||!upd.can_apply)return; const l=upd.last;
  if(!confirm("Update Work To-Do to "+l.tag+"?\n\nYour tasks, notes and settings are kept. The current version is backed up first, and the app restarts when it's done."))return;
  let ov=$("#busyover"); if(!ov){ov=el("div","busyover"); ov.id="busyover"; document.body.append(ov);} ov.hidden=false; ov.textContent="Downloading and installing "+l.tag+"\u2026";
  try{await saveNow(); const r=await fetch("/api/update/apply",{method:"POST",headers:{"X-WorkTodo":"1","Content-Type":"application/json"},body:"{}"}); const j=await r.json(); if(!r.ok)throw new Error(j.error||"HTTP "+r.status);
    ov.textContent="Installed "+l.tag+". Restarting\u2026"; const t0=Date.now();
    const poll=async()=>{try{const s2=await (await fetch("/api/update/status",{cache:"no-store"})).json(); if(s2.version!==APP_VERSION){location.reload(); return;}}catch(e){} if(Date.now()-t0<45000)setTimeout(poll,1000); else ov.textContent="Updated. Close this window and open Work To-Do again (Ctrl+Alt+T).";};
    setTimeout(poll,2500);}
  catch(e){ov.hidden=true; alert("The update didn't install: "+e.message+"\n\nNothing was changed. You can try again later or download the new installer.");}}
/* settings sections (v8) */
function v8SettingsHtml(){const p=prof(), fc=focusCfg(), s=data.settings||{}, uOn=upd?upd.enabled:(s.updates===true);
  return `<div class="ssec" id="sProfSec">${ICON.sun}<span>Profile & appearance</span></div>
  <div class="srow"><div class="st"><b>Your name</b><small>Used for the greeting and to spot emails meant for you.</small></div><input class="field" id="sName" value="${esc(p.name)}" placeholder="Your name" style="width:200px"></div>
  <div class="srow" style="display:block"><div class="st" style="margin-bottom:8px"><b>Theme</b> <small style="display:inline">Ctrl+Shift+L switches light/dark \u00b7 Ctrl+Shift+K cycles accents</small></div><div id="sThemes"></div><div id="sAcc" style="padding-bottom:14px"></div></div>
  <div class="srow"><div class="st"><b>Get set up checklist</b><small>The new-user checklist on the dashboard.</small></div><button type="button" class="btn ghost sm" id="sShowSetup">${onb().dismissed?"Show it again":"Hide it"}</button></div>
  <div class="srow"><div class="st"><b>Setup wizard & tour</b><small>Walk through first-run setup again, or replay the shortcuts tour.</small></div><button type="button" class="btn ghost sm" id="sWizard">Run setup again</button><button type="button" class="btn ghost sm" id="sTour">Take the tour</button></div>
  <div class="ssec" id="sFocusSec">${ICON.timer}<span>Focus timer & quick add</span></div>
  <div class="srow"><div class="st"><b>Focus / break length</b><small>Minutes. Select a task and press F to start.</small></div><input type="number" min="5" max="120" class="field" id="sFWork" value="${+fc.work}" style="width:80px"><input type="number" min="1" max="60" class="field" id="sFBrk" value="${+fc.brk}" style="width:80px"></div>
  <div class="srow"><div class="st"><b>Chime when a round ends</b></div><div class="toggle${fc.sound?" on":""}" id="sFSound"><div class="sw"></div></div></div>
  <div class="srow"><div class="st"><b>Quick add hotkey (Ctrl+Alt+N)</b><small>Pops up a tiny add-task box from anywhere in Windows, even with this window closed. Needs the background reminder helper running.</small></div><span class="pill" id="sQaPill">Checking\u2026</span><div class="toggle${s.quick_add!==false?" on":""}" id="sQa"><div class="sw"></div></div></div>
  <div class="ssec" id="sBackupSec">${ICON.save}<span>Backup & restore</span></div>
  <div class="srow"><div class="st"><b>Back up everything to one file</b><small>Projects, tasks, notes, settings and focus logs. ${onb().backup_at?"Last backup "+esc(fmtWhen(onb().backup_at))+".":"No backup yet."} Restoring keeps a copy of your current data first.</small></div><button type="button" class="btn sm" id="sExport">Back up</button><button type="button" class="btn ghost sm" id="sRestore">Restore\u2026</button></div>
  <div class="ssec" id="sUpdSec">${ICON.gift}<span>Updates</span><span class="pill">v${APP_VERSION}</span></div>
  <div class="srow"><div class="st"><b>Check for updates automatically</b><small>Once a day, from GitHub Releases (${esc((upd&&upd.repo)||"")}). Your data is never touched by an update.${upd&&!upd.installed?" This copy wasn't installed from a release, so it can't update itself; a check still tells you if there's a newer version.":""}</small></div><div class="toggle${uOn?" on":""}" id="sUpd"><div class="sw"></div></div></div>
  <div class="mfoot" style="margin-top:8px"><button type="button" class="btn ghost sm" id="sUpdCheck">Check for updates</button><span id="sUpdMsg">${upd&&upd.last&&upd.last.checked?"Last checked "+esc(fmtWhen(upd.last.checked)):""}</span></div>`;}
function wireV8Settings(m){
  m.querySelector("#sName").onchange=e=>{const v=e.target.value.trim(); if(v){putProf({name:v}); toast("Name saved");}};
  const th=()=>{const a=m.querySelector("#sThemes"), b=m.querySelector("#sAcc"); a.innerHTML=""; b.innerHTML=""; a.append(themeCards(th)); b.append(accentSwatches(th));}; th();
  m.querySelector("#sShowSetup").onclick=e=>{const d=!onb().dismissed; putOnb({dismissed:d}); e.target.textContent=d?"Show it again":"Hide it"; toast(d?"Checklist hidden":"Checklist is back on the dashboard");};
  m.querySelector("#sWizard").onclick=()=>{closeModal(); openWizard();}; m.querySelector("#sTour").onclick=()=>startTour();
  const fput=patch=>putSettings({focus:Object.assign(focusCfg(),patch)});
  m.querySelector("#sFWork").onchange=e=>fput({work:Math.max(5,Math.min(120,+e.target.value||25))}); m.querySelector("#sFBrk").onchange=e=>fput({brk:Math.max(1,Math.min(60,+e.target.value||5))});
  m.querySelector("#sFSound").onclick=e=>{const on=!focusCfg().sound; e.currentTarget.classList.toggle("on",on); fput({sound:on});};
  m.querySelector("#sQa").onclick=e=>{const on=(data.settings||{}).quick_add===false; e.currentTarget.classList.toggle("on",on); putSettings({quick_add:on}); toast(on?"Quick add hotkey on":"Quick add hotkey off");};
  fetch("/api/reminder-status").then(r=>r.json()).then(j=>{const p=m.querySelector("#sQaPill"); if(!p)return; p.textContent=!j.running?"Helper not running":j.hotkey_ok?"Ready":(j.hotkey_error||"Not available"); p.className="pill "+(j.running&&j.hotkey_ok?"ok":"bad");}).catch(()=>{});
  m.querySelector("#sExport").onclick=()=>exportBackup(); m.querySelector("#sRestore").onclick=()=>{closeModal(); pickRestore();};
  m.querySelector("#sUpd").onclick=e=>{const on=!(upd?upd.enabled:false); e.currentTarget.classList.toggle("on",on); putSettings({updates:on}); if(upd)upd.enabled=on; toast(on?"Update checks on":"Update checks off");};
  m.querySelector("#sUpdCheck").onclick=()=>checkUpdates(m.querySelector("#sUpdMsg"));}
/* outside changes (quick add, restore): reload when the file changed and nothing is pending here */
let lastRev="", savePending=false;
async function pollRev(){try{const j=await (await fetch("/api/rev")).json(); if(!lastRev){lastRev=j.rev; return;}
    if(j.rev!==lastRev&&!savePending&&!isTyping(document.activeElement)&&$("#modalBack").hidden&&!wizOpen()&&!document.querySelector(".task.leaving")){
      const d=await (await fetch("/api/data")).json(); lastRev=j.rev; const before=data.tasks.length; data=normData(d); render(); if(data.tasks.length>before)toast("Added from quick add");}}catch(e){}}
ICON.checkc='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.2 4.2L19 7"/></svg>';
try{const t0=JSON.parse(localStorage.getItem("wtdTheme")||"null"); if(t0)applyTheme(t0.theme,t0.accent);}catch(e){}

/* keyboard */
function newTask(){if(view!=="tasks"){view="tasks";render();} const t=$("#tTitle"); $(".main").scrollTo({top:0,behavior:"smooth"}); t.focus();}
function newNote(){if(view!=="notes"){view="notes";render();} $(".main").scrollTo({top:0,behavior:"smooth"}); $("#nTitle").focus();}
function focusSearch(){if(view==="dash")go("all"); const q=$("#q"); q.focus(); q.select();}
function selCard(){return selTask&&document.querySelector(`#list .task[data-id="${selTask}"]`);}
function moveSel(d){const cards=[...document.querySelectorAll("#list .task")]; if(!cards.length)return;
  let i=cards.findIndex(c=>c.dataset.id===selTask); i=i<0?(d>0?0:cards.length-1):Math.max(0,Math.min(cards.length-1,i+d));
  cards.forEach(c=>c.classList.remove("sel")); cards[i].classList.add("sel"); selTask=cards[i].dataset.id; cards[i].scrollIntoView({block:"nearest",behavior:"smooth"});}
const selT=()=>data.tasks.find(t=>t.id===selTask);
document.addEventListener("keydown",e=>{
  if(e.defaultPrevented)return;
  const k=e.key, typing=isTyping(e.target), ctrl=e.ctrlKey||e.metaKey;
  if(!$("#modalBack").hidden){if(modalKeys&&modalKeys(e,k,typing))return; if(k==="Escape"){e.preventDefault();closeModal();} else if(k==="?"&&!typing){e.preventDefault();closeModal();} return;}
  if(k==="Escape"){
    if(typing){if(e.target.id==="q"&&e.target.value){e.target.value="";if(view==="notes"){noteQuery="";renderNotes();}else if(view==="email"){em.query="";renderEmail();}else{query="";renderList();}} e.target.blur(); return;}
    if(selTask){selTask=null;document.querySelectorAll(".task.sel").forEach(c=>c.classList.remove("sel"));return;}
    if(view==="email"&&em.sel){em.sel=null;document.querySelectorAll(".mail.sel").forEach(c=>c.classList.remove("sel"));return;}
    if(query||filter!=="all"||noteQuery||em.query){query="";noteQuery="";em.query="";filter="all";render();} return;}
  if(ctrl&&!e.altKey){const lk=k.toLowerCase();
    if(e.shiftKey&&lk==="a"){e.preventDefault(); if(view!=="ai")go("ai"); setTimeout(()=>{const t=$("#aiText"); if(t&&!t.disabled)t.focus();}); return;}
    if(e.shiftKey&&(k===" "||e.code==="Space")){if(view==="ai"&&aiOn()){e.preventDefault(); toggleMic();} return;}
    if(lk==="n"){e.preventDefault(); e.shiftKey?newNote():newTask(); return;}
    if(lk==="f"&&!e.shiftKey){e.preventDefault(); focusSearch(); return;}
    if(k===","){e.preventDefault(); openSettings(); return;}
    if(lk==="z"&&!typing){if(doUndo())e.preventDefault(); return;}
    if(/^[1-9]$/.test(k)){const p=sortedProjects()[+k-1]; e.preventDefault(); if(p)go(p.id); return;}
    if(k==="0"){e.preventDefault(); go("all"); return;}
    if(lk==="d"&&!e.shiftKey){e.preventDefault(); go("dash"); return;}
    if(e.shiftKey&&lk==="e"){e.preventDefault(); exportBackup(); return;}
    if(e.shiftKey&&lk==="o"){e.preventDefault(); pickRestore(); return;}
    if(e.shiftKey&&lk==="l"){e.preventDefault(); toggleTheme(); return;}
    if(e.shiftKey&&lk==="k"){e.preventDefault(); cycleAccent(); return;}
    if(e.shiftKey&&lk==="f"){e.preventDefault(); stopFocus(); return;}
    if(e.shiftKey&&lk==="u"){e.preventDefault(); if(upd&&upd.available&&upd.can_apply)applyUpdate(); else checkUpdates(); return;}
    return;}
  if(typing||e.altKey||ctrl)return;
  if(gPending&&Date.now()-gPending<1500){gPending=0; const dest={a:"all",t:"today",u:"upcoming",o:"overdue",n:"notes",i:"ai",e:"email",d:"dash",w:"review"}[k.toLowerCase()]; if(dest){e.preventDefault();go(dest);} return;}
  gPending=0;
  switch(k){
    case "g": case "G": gPending=Date.now(); return;
    case "/": e.preventDefault(); focusSearch(); return;
    case "?": e.preventDefault(); openShortcuts(); return;
    case "n": e.preventDefault(); newTask(); return;
    case "N": e.preventDefault(); newNote(); return;
    case "P": e.preventDefault(); $("#newpName").focus(); return;
    case "h": showDone=!showDone; render(); toast(showDone?"Showing completed":"Hiding completed"); return;
    case "F": e.preventDefault(); if(focus)pauseFocus(); else toast("No focus timer running. Select a task and press F."); return;
  }
  if(k==="f"){const sel=view==="dash"||view==="review"?document.querySelector("#"+(view==="dash"?"dashView":"reviewView")+" .drow.sel"):null;
    const t=sel?data.tasks.find(x=>x.id===sel.dataset.id):view==="tasks"?selT():null;
    if(t){e.preventDefault(); startFocus(t); if(view==="dash")renderDash(); return;}
    if(focus&&focus.phase==="done"){e.preventDefault(); const t2=data.tasks.find(x=>x.id===focus.task_id); focus=null; startFocus(t2); return;}
    if(view==="dash"||view==="tasks"){e.preventDefault(); toast("Select a task first (J / K), then press F"); return;}}
  if(view==="dash"&&(k==="c"||k==="C")){e.preventDefault(); putOnb({dismissed:!onb().dismissed}); renderDash(); return;}
  if(view==="review"){reviewKey(e,k);return;}
  if(view==="email"){emailKey(e,k);return;}
  if(view==="dash"){dashKey(e,k);return;}
  if(view!=="tasks")return;
  if(k==="j"||k==="ArrowDown"){e.preventDefault();moveSel(1);return;}
  if(k==="k"||k==="ArrowUp"){e.preventDefault();moveSel(-1);return;}
  const t=selT(), card=selCard(); if(!t||!card)return;
  if(k==="x"||k===" "){e.preventDefault(); toggleDone(t,card); return;}
  if(k==="e"||k==="Enter"){e.preventDefault(); const tt=card.querySelector(".ttl"); tt.focus(); const r=document.createRange(); r.selectNodeContents(tt); r.collapse(false); const s=getSelection(); s.removeAllRanges(); s.addRange(r); return;}
  if(k==="p"){t.priority=PRI[(PRI.indexOf(t.priority)+2)%3]; commit(); toast("Priority: "+t.priority); return;}
  if(k==="y"){e.preventDefault(); openRepeat(t); return;}
  if(k==="t"){t.due=todayStr(); commit(); toast("Due today"); return;}
  if(k==="d"){const di=card.querySelector('.due:not(.time) input'); try{di.showPicker();}catch(_){di.focus();} return;}
  if(k==="Delete"||k==="Backspace"){e.preventDefault(); const cards=[...document.querySelectorAll("#list .task")]; const i=cards.indexOf(card);
    const next=cards[i+1]||cards[i-1]; deleteTask(t); if(next&&next.dataset.id!==t.id){selTask=next.dataset.id; render();} return;}
});


(async()=>{try{const r=await fetch("/api/data");data=await r.json();}catch(e){toast("Couldn't load your tasks");}
  data.projects=data.projects||[];data.tasks=data.tasks||[];data.notes=Array.isArray(data.notes)?data.notes:[]; const before=JSON.stringify(data);
  applyTheme(themeNow().theme,themeNow().accent);
  applyHash(); render(); if(JSON.stringify(data)!==before)save(); if(view==="tasks"&&$("#modalBack").hidden)$("#tTitle").focus();
  const ready=prof().setup_done;
  wireEmailDrop(); if(ready&&emCfg().auto)refreshEmail(false); else loadInbox(); if(view==="dash")dashOpen(); loadUpdate(); renderFocus(); pollRev();
  if(!ready)openWizard(); else if(!onb().tour_seen&&$("#modalBack").hidden)setTimeout(startTour,500);})();
setInterval(()=>{if(emCfg().auto&&prof().setup_done)refreshEmail(false);},600000);
setInterval(()=>{if(!isTyping(document.activeElement)&&$("#modalBack").hidden&&!document.querySelector(".task.leaving"))render();},60000);
setInterval(()=>{fetch("/api/ping").catch(()=>{}); pollRev();},5000);
</script>
</body></html>"""

if __name__ == "__main__":
    main()
