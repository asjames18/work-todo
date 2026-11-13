"""Work To-Do: AI suggestions layer (dormant by default).

Structure
  build_context(data)       -> JSON snapshot of the to-do state (tasks, overdue, due today, workload per day ...)
  RulesProvider             -> built-in, deterministic, local. No network.
  OpenAICompatibleProvider  -> any OpenAI-style /v1/chat/completions endpoint (LM Studio, Ollama, OpenRouter ...).
                               The API key is read from an environment variable whose NAME is stored in settings;
                               the key itself is never stored, logged or returned.
  validate_actions()        -> the fixed allowlist. Everything a provider suggests goes through it.
  apply_actions()           -> applies validated actions to a copy of the data and reports what changed (for undo).
  parse_task()              -> "Ask AI to add": free text -> one structured task (rules parser, or the model with
                               a rules fallback).
Nothing here runs unless settings.ai.enabled is true; the server checks that before calling anything.
"""
import json, os, re, uuid, datetime as dt, urllib.request, urllib.error

PRIORITIES = ("normal", "high", "urgent")
RANK = {"urgent": 0, "high": 1, "normal": 2}
ACTION_TYPES = ("set_priority", "set_due", "mark_done", "add_task", "move_task", "add_note")
DEFAULT_AI = {"enabled": False, "provider": "rules", "base_url": "http://127.0.0.1:1234/v1",
              "model": "", "api_key_env": "", "voice_model": "small.en"}
PRESETS = {"lmstudio": "http://127.0.0.1:1234/v1", "ollama": "http://127.0.0.1:11434/v1",
           "openrouter": "https://openrouter.ai/api/v1"}
OVERLOAD = 5          # open tasks on one day that count as "overloaded"
STALE_DAYS = 14
MAX_ACTIONS = 60
MAX_RECS = 15
ENV_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,63}$")


class AIError(Exception):
    pass


# ---------------------------------------------------------------- settings
def ai_settings(data):
    s = dict(DEFAULT_AI)
    raw = (data.get("settings") or {}).get("ai") if isinstance(data.get("settings"), dict) else None
    if isinstance(raw, dict):
        for k in DEFAULT_AI:
            if k in raw:
                s[k] = raw[k]
    s["enabled"] = s.get("enabled") is True
    if s["provider"] not in ("rules", "openai_compatible"):
        s["provider"] = "rules"
    if s.get("voice_model") not in ("small.en", "large-v3-turbo"):
        s["voice_model"] = "small.en"
    for k in ("base_url", "model", "api_key_env"):
        s[k] = str(s.get(k) or "").strip()
    return s


def is_enabled(data):
    return ai_settings(data)["enabled"]


# ---------------------------------------------------------------- time helpers
def _now():
    return dt.datetime.now().replace(microsecond=0)


def _iso(d):
    return d.strftime("%Y-%m-%dT%H:%M:%S")


def _date(s):
    try:
        return dt.date.fromisoformat(s) if isinstance(s, str) and len(s) == 10 else None
    except ValueError:
        return None


def _valid_time(s):
    return isinstance(s, str) and re.fullmatch(r"([01]\d|2[0-3]):[0-5]\d", s) is not None


def _parse_created(s):
    try:
        return dt.datetime.fromisoformat(str(s)[:19])
    except ValueError:
        return None


def is_overdue(t, now):
    if t.get("done") or not t.get("due"):
        return False
    today = now.date().isoformat()
    return t["due"] < today or (t["due"] == today and bool(t.get("due_time")) and t["due_time"] <= now.strftime("%H:%M"))


def _first(title, n=60):
    line = (title or "").split("\n")[0].strip()
    return line if len(line) <= n else line[: n - 1] + "…"


def _fmt_date(s, now):
    d = _date(s)
    if not d:
        return "no date"
    diff = (d - now.date()).days
    label = {0: "Today", 1: "Tomorrow", -1: "Yesterday"}.get(diff)
    base = f"{d:%a %b} {d.day}"
    return f"{label} ({base})" if label else base


def _fmt_time(hm):
    if not hm:
        return ""
    h, m = map(int, hm.split(":"))
    return f"{(h % 12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}"


# ---------------------------------------------------------------- context snapshot
def build_context(data, now=None):
    now = now or _now()
    today = now.date()
    projects = {p.get("id"): p for p in data.get("projects", [])}
    tasks, workload = [], {}
    for i in range(14):
        workload[(today + dt.timedelta(days=i)).isoformat()] = 0
    for t in data.get("tasks", []):
        created = _parse_created(t.get("created"))
        row = {"id": t.get("id"), "title": (t.get("title") or "")[:400],
               "project": (projects.get(t.get("project_id")) or {}).get("name", ""), "project_id": t.get("project_id"),
               "priority": t.get("priority", "normal"), "due": t.get("due") or "", "due_time": t.get("due_time") or "",
               "done": bool(t.get("done")), "age_days": (now - created).days if created else None,
               "overdue": is_overdue(t, now), "due_today": (not t.get("done")) and t.get("due") == today.isoformat(),
               "lines": len([l for l in (t.get("title") or "").split("\n") if l.strip()])}
        if t.get("done"):
            row["completed"] = t.get("completed") or ""
        tasks.append(row)
        if not t.get("done") and t.get("due") in workload:
            workload[t["due"]] += 1
    open_tasks = [r for r in tasks if not r["done"]]
    return {
        "now": _iso(now), "today": today.isoformat(), "weekday": today.strftime("%A"),
        "projects": [{"id": p.get("id"), "name": p.get("name"),
                      "open_tasks": sum(1 for r in open_tasks if r["project_id"] == p.get("id"))}
                     for p in data.get("projects", [])],
        "tasks": tasks,
        "overdue": [r["id"] for r in open_tasks if r["overdue"]],
        "due_today": [r["id"] for r in open_tasks if r["due_today"]],
        "workload_by_day": workload,
        "open_count": len(open_tasks),
        "notes": [{"id": n.get("id"), "title": n.get("title", ""), "body": (n.get("body") or "")[:600],
                   "project": (projects.get(n.get("project_id")) or {}).get("name", "")} for n in data.get("notes", [])],
    }


# ---------------------------------------------------------------- allowlist validation
def _task(data, tid):
    for t in data.get("tasks", []):
        if t.get("id") == tid:
            return t
    raise AIError("that task no longer exists")


def _project_ok(data, pid):
    return any(p.get("id") == pid for p in data.get("projects", []))


def _clean_due(a):
    due = a.get("due") or ""
    tm = a.get("due_time") or ""
    if due and not _date(due):
        raise AIError("bad due date")
    if tm and not _valid_time(tm):
        raise AIError("bad due time")
    if tm and not due:
        raise AIError("a due time needs a due date")
    return due, tm


def validate_action(a, data):
    """Return a normalized copy of one action, or raise AIError. Unknown keys are dropped."""
    if not isinstance(a, dict):
        raise AIError("action must be an object")
    typ = a.get("type")
    if typ not in ACTION_TYPES:
        raise AIError(f"action type not allowed: {str(typ)[:40]}")
    if typ in ("set_priority", "set_due", "mark_done", "move_task"):
        t = _task(data, a.get("task_id"))
        out = {"type": typ, "task_id": t["id"]}
        if typ == "set_priority":
            if a.get("priority") not in PRIORITIES:
                raise AIError("bad priority")
            out["priority"] = a["priority"]
        elif typ == "set_due":
            out["due"], out["due_time"] = _clean_due(a)
        elif typ == "mark_done":
            if t.get("done"):
                raise AIError("task is already done")
        elif typ == "move_task":
            if not _project_ok(data, a.get("project_id")):
                raise AIError("unknown project")
            out["project_id"] = a["project_id"]
        return out
    if typ == "add_task":
        title = str(a.get("title") or "").replace("\r", "").strip()
        if not title or len(title) > 4000:
            raise AIError("task title is empty or too long")
        if not _project_ok(data, a.get("project_id")):
            raise AIError("pick a project for the new task")
        pr = a.get("priority") or "normal"
        if pr not in PRIORITIES:
            raise AIError("bad priority")
        due, tm = _clean_due(a)
        return {"type": typ, "title": title, "project_id": a["project_id"], "priority": pr, "due": due, "due_time": tm}
    # add_note
    title = str(a.get("title") or "").strip()
    body = str(a.get("body") or "").replace("\r", "")
    if not title or len(title) > 200 or len(body) > 20000:
        raise AIError("note title is empty or too long")
    pid = a.get("project_id") or None
    if pid is not None and not _project_ok(data, pid):
        raise AIError("unknown project")
    return {"type": typ, "title": title, "body": body, "project_id": pid}


def validate_actions(actions, data):
    """Validate a list in order against a working copy (so 'add then move' style chains are checked properly)."""
    if not isinstance(actions, list) or not actions:
        raise AIError("no actions")
    if len(actions) > MAX_ACTIONS:
        raise AIError("too many actions at once")
    work = json.loads(json.dumps(data))
    out = []
    for a in actions:
        v = validate_action(a, work)
        _apply_one(work, v, _now(), {"modified": set(), "added_tasks": [], "added_notes": []})
        out.append(v)
    return out


def describe(action, data, now=None):
    now = now or _now()
    projects = {p.get("id"): p.get("name") for p in data.get("projects", [])}
    typ = action["type"]
    if typ == "add_task":
        due = (" · due " + _fmt_date(action["due"], now) + (" " + _fmt_time(action["due_time"]) if action["due_time"] else "")) if action["due"] else ""
        return f"Add task “{_first(action['title'], 80)}” to {projects.get(action['project_id'], '?')} · {action['priority'].title()}{due}"
    if typ == "add_note":
        return f"Add note “{_first(action['title'], 80)}”" + (f" to {projects[action['project_id']]}" if action.get("project_id") in projects else "")
    t = next((x for x in data.get("tasks", []) if x.get("id") == action["task_id"]), {})
    name = f"“{_first(t.get('title'), 70)}”"
    if typ == "set_priority":
        return f"Task {name}: priority {str(t.get('priority', 'normal')).title()} → {action['priority'].title()}"
    if typ == "set_due":
        old = _fmt_date(t.get("due"), now) + (" " + _fmt_time(t.get("due_time")) if t.get("due_time") else "") if t.get("due") else "no due date"
        new = _fmt_date(action["due"], now) + (" " + _fmt_time(action["due_time"]) if action["due_time"] else "") if action["due"] else "no due date"
        return f"Task {name}: due {old} → {new}"
    if typ == "mark_done":
        return f"Task {name}: mark done"
    if typ == "move_task":
        return f"Task {name}: move from {projects.get(t.get('project_id'), '(no project)')} to {projects.get(action['project_id'], '?')}"
    return typ


def describe_all(actions, data):
    work = json.loads(json.dumps(data))
    lines, now = [], _now()
    for a in actions:
        lines.append(describe(a, work, now))
        _apply_one(work, a, now, {"modified": set(), "added_tasks": [], "added_notes": []})
    return lines


def _apply_one(data, a, now, rec):
    typ = a["type"]
    if typ == "add_task":
        tid = uuid.uuid4().hex
        data.setdefault("tasks", []).append({"id": tid, "project_id": a["project_id"], "title": a["title"], "priority": a["priority"],
                                             "due": a["due"], "due_time": a["due_time"], "done": False, "created": _iso(now), "completed": ""})
        rec["added_tasks"].append(tid)
        return
    if typ == "add_note":
        nid = uuid.uuid4().hex
        data.setdefault("notes", []).append({"id": nid, "title": a["title"], "body": a["body"], "project_id": a["project_id"],
                                             "created": _iso(now), "updated": _iso(now)})
        rec["added_notes"].append(nid)
        return
    t = _task(data, a["task_id"])
    if t["id"] not in rec["added_tasks"]:
        rec["modified"].add(t["id"])
    if typ == "set_priority":
        t["priority"] = a["priority"]
    elif typ == "set_due":
        t["due"], t["due_time"] = a["due"], a["due_time"]
    elif typ == "mark_done":
        t["done"], t["completed"] = True, _iso(now)
    elif typ == "move_task":
        t["project_id"] = a["project_id"]


def apply_actions(data, actions):
    """Validate, then apply to a copy. Returns (new_data, changes, undo_info)."""
    actions = validate_actions(actions, data)
    changes = describe_all(actions, data)
    new = json.loads(json.dumps(data))
    rec = {"modified": set(), "added_tasks": [], "added_notes": []}
    now = _now()
    for a in actions:
        _apply_one(new, a, now, rec)
    undo = {"modified_tasks": [t for t in data.get("tasks", []) if t.get("id") in rec["modified"]],
            "added_task_ids": rec["added_tasks"], "added_note_ids": rec["added_notes"]}
    return new, changes, undo


# ---------------------------------------------------------------- built-in rules provider
_SPLIT_RE = re.compile(r"^\s*(?:[-*•▪]|\d{1,2}[.)]|\[ ?\])\s*")
_FOLLOW_RE = re.compile(r"\b(send|sent|email|e-mail|submit|submitted|request|requested|ask|asked|reach out|ticket)\b", re.I)
_NOTE_TODO_RE = re.compile(r"^\s*(?:TODO[:\-]\s*|- \[ \]\s*|\[ \]\s*)(.+)$", re.I)


def _next_weekday(d, wd, min_ahead=1):
    days = (wd - d.weekday()) % 7
    if days < min_ahead:
        days += 7
    return d + dt.timedelta(days=days)


def _biz_days(d, n):
    while n > 0:
        d += dt.timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def _norm(s):
    return re.sub(r"\s+", " ", (s or "").lower()).strip()


class RulesProvider:
    name = "rules"
    label = "Built-in rules (local)"

    def recommend(self, ctx, data, now=None):
        now = now or dt.datetime.fromisoformat(ctx["now"])
        today = now.date()
        recs = []
        tasks = {t["id"]: t for t in data.get("tasks", [])}
        rows = [r for r in ctx["tasks"] if not r["done"]]
        titles = {_norm(t.get("title")) for t in data.get("tasks", [])}
        titles_first = {_norm((t.get("title") or "").split("\n")[0]) for t in data.get("tasks", [])}

        def add(kind, key, title, reason, actions, sev="info"):
            recs.append({"id": f"{kind}:{key}", "kind": kind, "title": title, "reason": reason, "severity": sev, "actions": actions})

        # 1. overdue
        for r in rows:
            if not r["overdue"]:
                continue
            keep = r["due_time"] if r["due_time"] and r["due_time"] > now.strftime("%H:%M") and r["due"] == today.isoformat() else ""
            acts = [{"type": "set_due", "task_id": r["id"], "due": today.isoformat(), "due_time": keep}]
            if r["priority"] == "normal":
                acts.append({"type": "set_priority", "task_id": r["id"], "priority": "high"})
            late = (today - _date(r["due"])).days
            when = "earlier today" if late == 0 else f"{late} day{'s' if late != 1 else ''} ago"
            add("overdue", r["id"], f"Overdue: {_first(r['title'])}",
                f"It was due {when}. Move it to today" + (" and raise it to High" if r["priority"] == "normal" else "") + ", or mark it done if it's finished.",
                acts, "high")
        # 2. due within 24h but still Normal
        for r in rows:
            if r["overdue"] or not r["due"] or r["priority"] != "normal":
                continue
            due_at = dt.datetime.fromisoformat(r["due"] + "T" + (r["due_time"] or "23:59"))
            if now < due_at <= now + dt.timedelta(hours=24):
                hrs = max(1, round((due_at - now).total_seconds() / 3600))
                add("due_soon", r["id"], f"Due soon: {_first(r['title'])}",
                    f"Due in about {hrs} hour{'s' if hrs != 1 else ''} and still marked Normal. Raise it to High so it sorts to the top.",
                    [{"type": "set_priority", "task_id": r["id"], "priority": "high"}], "medium")
        # 3. overloaded days
        load = dict(ctx["workload_by_day"])
        for day, n in sorted(ctx["workload_by_day"].items()):
            if n < OVERLOAD:
                continue
            movable = sorted([r for r in rows if r["due"] == day and not r["due_time"] and not r["overdue"]],
                             key=lambda r: (-RANK.get(r["priority"], 2), r["age_days"] or 0))
            movable = [r for r in movable if r["priority"] == "normal"]
            need = n - (OVERLOAD - 1)
            acts, moved = [], []
            for r in movable[:need]:
                cands = [d for d in sorted(load) if d > day and dt.date.fromisoformat(d).weekday() < 5 and load[d] < OVERLOAD - 1]
                if not cands:
                    break
                target = min(cands, key=lambda d: (load[d], d))
                load[target] += 1
                load[day] -= 1
                acts.append({"type": "set_due", "task_id": r["id"], "due": target, "due_time": ""})
                moved.append(_first(r["title"], 30))
            if acts:
                add("overload", day, f"{_fmt_date(day, now)} is overloaded ({n} tasks)",
                    f"Spread the lighter Normal tasks to quieter weekdays: {', '.join(moved)}.", acts, "medium")
        # 4. stale, no due date
        target = _next_weekday(today, 4, 2)  # a Friday at least 2 days out
        for r in rows:
            if r["due"] or (r["age_days"] or 0) <= STALE_DAYS:
                continue
            add("stale", r["id"], f"Waiting a while: {_first(r['title'])}",
                f"Added {r['age_days']} days ago with no due date. Give it a date so it gets a reminder, or mark it done if it's no longer needed.",
                [{"type": "set_due", "task_id": r["id"], "due": target.isoformat(), "due_time": ""}])
        # 5. very long multi-line tasks -> split
        for r in rows:
            t = tasks.get(r["id"]) or {}
            lines = [l.rstrip() for l in (t.get("title") or "").split("\n") if l.strip()]
            if len(lines) < 4:
                continue
            items = [_SPLIT_RE.sub("", l).strip() for l in lines]
            items = [i for i in items if 3 <= len(i) <= 400]
            if len(items) < 4 or len(items) > 25:
                continue
            acts = [{"type": "add_task", "title": i, "project_id": t.get("project_id"), "priority": t.get("priority", "normal"),
                     "due": t.get("due") or "", "due_time": t.get("due_time") or ""} for i in items if _norm(i) not in titles]
            if len(acts) < 3:
                continue
            acts.append({"type": "mark_done", "task_id": r["id"]})
            add("split", r["id"], f"Split into {len(acts) - 1} tasks: {_first(r['title'], 40)}",
                f"This task has {len(lines)} lines. Make each line its own task so you can check them off one at a time. "
                "The original is marked done and stays under Completed.", acts)
        # 6. conservative follow-ups: recently completed "send/request/submit" tasks, and TODO lines in notes
        cutoff = now - dt.timedelta(days=7)
        fcount = 0
        for t in data.get("tasks", []):
            if fcount >= 3 or not t.get("done"):
                continue
            c = _parse_created(t.get("completed"))
            first = (t.get("title") or "").split("\n")[0].strip()
            if not c or c < cutoff or not _FOLLOW_RE.search(first):
                continue
            title = "Follow up: " + first[:200]
            if _norm(title) in titles_first:
                continue
            due = _biz_days(c.date(), 3)
            if due < today:
                due = _biz_days(today, 1)
            add("follow_up", t["id"], f"Follow up on: {_first(first, 45)}",
                "You finished this recently and it sounds like it's waiting on someone. Add a follow-up so it doesn't slip.",
                [{"type": "add_task", "title": title, "project_id": t.get("project_id"), "priority": "normal",
                  "due": due.isoformat(), "due_time": ""}])
            fcount += 1
        default_pid = (data.get("projects") or [{}])[0].get("id")
        for n in data.get("notes", []):
            pid = n.get("project_id") if _project_ok(data, n.get("project_id")) else default_pid
            if not pid:
                break
            for line in (n.get("body") or "").split("\n"):
                m = _NOTE_TODO_RE.match(line)
                if not m or fcount >= 6:
                    continue
                item = m.group(1).strip()[:300]
                if len(item) < 3 or _norm(item) in titles or _norm(item) in titles_first:
                    continue
                add("note_todo", f"{n.get('id')}:{_norm(item)[:40]}", f"From your note: {_first(item, 50)}",
                    f"Your note “{_first(n.get('title'), 40)}” has an open to-do line that isn't a task yet.",
                    [{"type": "add_task", "title": item, "project_id": pid, "priority": "normal", "due": "", "due_time": ""}])
                titles.add(_norm(item))
                fcount += 1
        order = {"high": 0, "medium": 1, "info": 2}
        recs.sort(key=lambda r: order.get(r["severity"], 3))
        return recs[:MAX_RECS]

    def parse(self, text, ctx, project_hint=None, now=None):
        return parse_task_rules(text, ctx, project_hint, now)


# ---------------------------------------------------------------- "Ask AI to add" rules parser
WEEKDAYS = {"mon": 0, "monday": 0, "tue": 1, "tues": 1, "tuesday": 1, "wed": 2, "weds": 2, "wednesday": 2, "thu": 3, "thur": 3,
            "thurs": 3, "thursday": 3, "fri": 4, "friday": 4, "sat": 5, "saturday": 5, "sun": 6, "sunday": 6}
MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
MONTHS.update({"january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7, "august": 8, "sept": 9,
               "september": 9, "october": 10, "november": 11, "december": 12})
_WD = "|".join(sorted(WEEKDAYS, key=len, reverse=True))
_MO = "|".join(sorted(MONTHS, key=len, reverse=True))
_NUMWORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
             "ten": 10, "eleven": 11, "twelve": 12, "a": 1, "an": 1}


def _hm(h, m, ampm):
    h, m = int(h), int(m or 0)
    ampm = (ampm or "").lower().replace(".", "")
    if ampm.startswith("p") and h < 12:
        h += 12
    elif ampm.startswith("a") and h == 12:
        h = 0
    elif not ampm and 1 <= h <= 7:
        h += 12  # "at 3" during a workday means 3 PM
    if not (0 <= h <= 23 and 0 <= m <= 59):
        return None
    return f"{h:02d}:{m:02d}"


def parse_task_rules(text, ctx, project_hint=None, now=None):
    now = now or (dt.datetime.fromisoformat(ctx["now"]) if ctx and ctx.get("now") else _now())
    today = now.date()
    raw = re.sub(r"\s+", " ", str(text or "").replace("\r", " ")).strip()
    notes = ""
    m = re.search(r"(?:\s+-{2,}\s+|\s*\b(?:notes?|details?)\s*:\s*)(.+)$", raw, re.I)
    if m:
        notes, raw = re.sub(r"^(?:notes?|details?)\s*:\s*", "", m.group(1).strip(), flags=re.I), raw[: m.start()].strip()
    s = " " + raw + " "
    found = {}

    def take(pattern, fn, key):
        nonlocal s
        mm = re.search(pattern, s, re.I)
        if mm and key not in found:
            val = fn(mm)
            if val is not None:
                found[key] = val
                s = s[: mm.start()] + " " + s[mm.end():]
                return True
        return False

    # priority
    take(r"[\s,;]+(?:with\s+|as\s+|at\s+)?(?:an?\s+)?(urgent|critical|asap|top)(?:\s+priority)?[\s,.;!]", lambda mm: "urgent", "priority")
    take(r"[\s,;]+(?:with\s+|as\s+|at\s+|it'?s\s+)?(?:an?\s+)?(?:high|important)(?:\s+priority|\s+pri)?[\s,.;!]", lambda mm: "high", "priority")
    take(r"[\s,;]+priority\s+(urgent|high|normal|low)[\s,.;!]", lambda mm: {"low": "normal"}.get(mm.group(1).lower(), mm.group(1).lower()), "priority")
    take(r"[\s,;]+(?:with\s+|as\s+)?(?:a\s+)?(?:low|normal|medium)\s+priority[\s,.;!]", lambda mm: "normal", "priority")
    # project: "for/in/under project X", "#X", or "in|for|under <exact project name>"
    projects = sorted((ctx or {}).get("projects", []), key=lambda p: -len(p.get("name") or ""))
    for p in projects:
        nm = re.escape(p.get("name") or "")
        if not nm:
            continue
        if take(r"[\s,;]+(?:(?:in|for|under|to)\s+(?:the\s+)?)?(?:project\s+)" + nm + r"(?:\s+project)?(?=[\s,.;!])", lambda mm, p=p: p["id"], "project") \
                or take(r"[\s,;]+#" + nm.replace(r"\ ", r"[\s_-]?") + r"(?=[\s,.;!])", lambda mm, p=p: p["id"], "project") \
                or take(r"[\s,;]+(?:in|under|for)\s+(?:the\s+)?" + nm + r"(?:\s+project)?(?=[\s,.;!])", lambda mm, p=p: p["id"], "project"):
            break
    if "project" not in found:  # bare mention: keep the words in the title, just link the project
        low = " " + raw.lower() + " "
        for p in projects:
            nm = (p.get("name") or "").lower()
            if nm and re.search(r"(?<![\w])" + re.escape(nm) + r"(?![\w])", low):
                found["project"] = p["id"]
                break
        if "project" not in found:  # spacing/punctuation differences, e.g. speech gives "Acme 451" for Acme451
            squash = re.sub(r"[^a-z0-9]", "", low)
            for p in projects:
                nm = re.sub(r"[^a-z0-9]", "", (p.get("name") or "").lower())
                if len(nm) >= 6 and nm in squash:
                    found["project"] = p["id"]
                    break
        if "project" not in found:  # also try the first word of multi-word project names, e.g. "Payroll" for "Payroll System"
            for p in projects:
                head = re.split(r"[\s\-–:]+", (p.get("name") or "").strip())[0].lower()
                if len(head) >= 5 and re.search(r"(?<![\w])" + re.escape(head) + r"(?![\w])", low):
                    found["project"] = p["id"]
                    break
    # time words
    take(r"[\s,;]+(?:by\s+|at\s+)?(?:end of (?:the )?day|eod|close of business|cob)(?=[\s,.;!])", lambda mm: "17:00", "time")
    take(r"[\s,;]+(?:at\s+)?(?:noon|midday)(?=[\s,.;!])", lambda mm: "12:00", "time")
    take(r"[\s,;]+(?:by\s+|at\s+|@\s*)?(\d{1,2})(?::(\d{2}))?\s*(a\.?m\.?|p\.?m\.?|am|pm)(?=[\s,.;!])",
         lambda mm: _hm(mm.group(1), mm.group(2), mm.group(3)), "time")
    take(r"[\s,;]+(?:by\s+|at\s+|@\s*)(\d{1,2}):(\d{2})(?=[\s,.;!])", lambda mm: _hm(mm.group(1), mm.group(2), ""), "time")
    take(r"[\s,;]+(?:at|@)\s*(\d{1,2})(?:\s+o'?clock)?(?=[\s,.;!])", lambda mm: _hm(mm.group(1), 0, ""), "time")
    # dates
    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?(?:the\s+)?day after tomorrow(?=[\s,.;!])", lambda mm: today + dt.timedelta(days=2), "date")
    take(r"[\s,;]+(?:by\s+|due\s+)?(?:today|tonight|this (?:morning|afternoon|evening))(?=[\s,.;!])",
         lambda mm: today, "date")
    take(r"[\s,;]+(?:by\s+|due\s+)?tomorrow(?:\s+(?:morning|afternoon|evening|night))?(?=[\s,.;!])", lambda mm: today + dt.timedelta(days=1), "date")
    take(r"[\s,;]+(?:in|within)\s+(\d+|" + "|".join(_NUMWORDS) + r")\s+(day|days|week|weeks)(?=[\s,.;!])",
         lambda mm: today + dt.timedelta(days=(int(mm.group(1)) if mm.group(1).isdigit() else _NUMWORDS[mm.group(1).lower()]) * (7 if mm.group(2).lower().startswith("w") else 1)), "date")
    take(r"[\s,;]+(?:by\s+|due\s+)?next week(?=[\s,.;!])", lambda mm: _next_weekday(today, 0), "date")
    take(r"[\s,;]+(?:by\s+|due\s+)?(?:the\s+)?end of (?:the |this )?week(?=[\s,.;!])", lambda mm: _next_weekday(today, 4, 0), "date")
    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?next\s+(" + _WD + r")\.?(?=[\s,.;!])",
         lambda mm: _next_weekday(today, WEEKDAYS[mm.group(1).lower()]), "date")
    take(r"[\s,;]+(on\s+|by\s+|due\s+)?(this\s+)?(" + _WD + r")\.?(?=[\s,.;!])",  # "by/this Friday" on a Friday = today
         lambda mm: _next_weekday(today, WEEKDAYS[mm.group(3).lower()], 0 if (mm.group(2) or (mm.group(1) or "").strip().lower() in ("by", "due")) else 1), "date")

    def ymd(y, mo, d):
        try:
            x = dt.date(int(y), int(mo), int(d))
        except ValueError:
            return None
        return x

    def roll(mo, d, y=None):
        if y:
            y = int(y)
            if y < 100:
                y += 2000
            return ymd(y, mo, d)
        x = ymd(today.year, mo, d)
        if x and x < today:
            x = ymd(today.year + 1, mo, d)
        return x

    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?(\d{4})-(\d{2})-(\d{2})(?=[\s,.;!])", lambda mm: ymd(mm.group(1), mm.group(2), mm.group(3)), "date")
    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?(?=[\s,.;!])", lambda mm: roll(mm.group(1), mm.group(2), mm.group(3)), "date")
    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?(?:the\s+)?(" + _MO + r")\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?(?=[\s,.;!])",
         lambda mm: roll(MONTHS[mm.group(1).lower()], mm.group(2), mm.group(3)), "date")
    take(r"[\s,;]+(?:on\s+|by\s+|due\s+)?(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?(" + _MO + r")\.?(?:,?\s+(\d{4}))?(?=[\s,.;!])",
         lambda mm: roll(MONTHS[mm.group(2).lower()], mm.group(1), mm.group(3)), "date")
    # "tonight"/"this evening" without a time: leave the time empty (a date alone is fine)
    due = found.get("date")
    tm = found.get("time") or ""
    if tm and not due:
        due = today if tm > now.strftime("%H:%M") else today + dt.timedelta(days=1)
    # tidy the title
    title = s.strip()
    title = re.sub(r"^(?:hey\s+|ok(?:ay)?\s+|so\s+)?(?:can you\s+|could you\s+|please\s+)?(?:remind me to|remind me|remember to|don'?t forget to|"
                   r"i need to|i have to|i've got to|i gotta|need to|add (?:a )?(?:new )?task(?: to)?|new task|create (?:a )?task(?: to)?|todo:?|to-?do:?)\s+",
                   "", title, flags=re.I)
    title = re.sub(r"^(?:please|to)\s+", "", title, flags=re.I)
    title = re.sub(r"\s+(?:please|thanks|thank you)[.!]*$", "", title, flags=re.I)
    title = re.sub(r"\s+(?:by|on|at|due|for|and|with|,)\s*$", "", title, flags=re.I)
    title = re.sub(r"\s*[,;]\s*(?=[,;]|$)", "", title)
    title = re.sub(r"\s{2,}", " ", title).strip(" ,;-")
    title = re.sub(r"\s+([,.;!?])", r"\1", title)
    if title:
        title = title[0].upper() + title[1:]
    if notes:
        title = (title + "\n" + notes).strip()
    pid = found.get("project")
    if not pid and project_hint and any(p.get("id") == project_hint for p in (ctx or {}).get("projects", [])):
        pid = project_hint
    return {"title": title, "notes": notes, "project_id": pid or None, "project_matched": bool(found.get("project")),
            "priority": found.get("priority", "normal"), "due": due.isoformat() if due else "", "due_time": tm if due else "",
            "parsed_by": "rules"}


# ---------------------------------------------------------------- OpenAI-compatible provider (stub, off until enabled)
SYSTEM_RECOMMEND = """You help one person manage a work to-do list. You will get a JSON snapshot of the list.
Reply with ONLY a JSON object, no prose, no markdown:
{"recommendations":[{"title":"short title","reason":"one or two plain sentences","actions":[ACTION,...]}]}
ACTION must be one of these exact shapes (ids must come from the snapshot):
{"type":"set_priority","task_id":"...","priority":"normal|high|urgent"}
{"type":"set_due","task_id":"...","due":"YYYY-MM-DD or empty","due_time":"HH:MM or empty"}
{"type":"mark_done","task_id":"..."}
{"type":"add_task","title":"...","project_id":"...","priority":"normal|high|urgent","due":"YYYY-MM-DD or empty","due_time":"HH:MM or empty"}
{"type":"move_task","task_id":"...","project_id":"..."}
{"type":"add_note","title":"...","body":"...","project_id":"... or null"}
No other action types exist. Never delete anything. At most 8 recommendations. Be conservative."""

SYSTEM_PARSE = """Turn the user's request into ONE to-do task. Reply with ONLY a JSON object, no prose:
{"title":"short imperative title","notes":"extra detail or empty","project_id":"an id from the projects list or null",
 "priority":"normal|high|urgent","due":"YYYY-MM-DD or empty","due_time":"HH:MM 24h or empty"}
Use the given 'now' to resolve words like tomorrow or next Monday. Don't invent details."""


def _extract_json(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    try:
        return json.loads(text)
    except ValueError:
        pass
    a, b = text.find("{"), text.rfind("}")
    if a >= 0 and b > a:
        return json.loads(text[a: b + 1])
    raise AIError("the model didn't return JSON")


class OpenAICompatibleProvider:
    name = "openai_compatible"
    label = "OpenAI-compatible endpoint"

    def __init__(self, cfg, opener=None):
        self.base_url = cfg.get("base_url", "").rstrip("/")
        self.model = cfg.get("model", "")
        self.key_env = cfg.get("api_key_env", "")
        self._open = opener or urllib.request.urlopen
        if not re.match(r"^https?://[^\s/]+", self.base_url):
            raise AIError("set a base URL that starts with http:// or https://")
        if not self.model:
            raise AIError("set a model name")
        if self.key_env and not ENV_NAME_RE.match(self.key_env):
            raise AIError("the API key setting must be an environment variable NAME, like OPENROUTER_API_KEY")

    def _headers(self):
        h = {"Content-Type": "application/json", "User-Agent": "WorkToDo/1.0"}
        if self.key_env:
            key = os.environ.get(self.key_env, "")
            if not key:
                raise AIError(f"environment variable {self.key_env} isn't set on this PC")
            h["Authorization"] = "Bearer " + key
        if "openrouter.ai" in self.base_url:
            h["X-Title"] = "Work To-Do"
        return h

    def _chat(self, system, user, timeout=90):
        body = json.dumps({"model": self.model, "temperature": 0.2, "stream": False,
                           "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}).encode("utf-8")
        req = urllib.request.Request(self.base_url + "/chat/completions", data=body, headers=self._headers(), method="POST")
        try:
            with self._open(req, timeout=timeout) as r:
                out = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            raise AIError(f"the AI endpoint answered HTTP {e.code}")
        except (urllib.error.URLError, OSError, ValueError) as e:
            raise AIError(f"couldn't reach the AI endpoint ({type(e).__name__})")
        try:
            return out["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            raise AIError("unexpected reply from the AI endpoint")

    def test(self):
        req = urllib.request.Request(self.base_url + "/models", headers=self._headers())
        try:
            with self._open(req, timeout=10) as r:
                out = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            raise AIError(f"HTTP {e.code}")
        except (urllib.error.URLError, OSError, ValueError) as e:
            raise AIError(f"couldn't connect ({type(e).__name__})")
        ids = [m.get("id") for m in out.get("data", []) if isinstance(m, dict)]
        return {"models": len(ids), "model_found": self.model in ids if ids else None}

    def recommend(self, ctx, data, now=None):
        reply = _extract_json(self._chat(SYSTEM_RECOMMEND, json.dumps(ctx)))
        recs = []
        for i, r in enumerate((reply.get("recommendations") or [])[:MAX_RECS] if isinstance(reply, dict) else []):
            if not isinstance(r, dict):
                continue
            try:
                acts = validate_actions(r.get("actions") or [], data)
            except AIError:
                continue  # drop anything outside the allowlist
            recs.append({"id": f"ai:{i}", "kind": "ai", "title": str(r.get("title") or "Suggestion")[:120],
                         "reason": str(r.get("reason") or "")[:400], "severity": "info", "actions": acts})
        return recs

    def parse(self, text, ctx, project_hint=None, now=None):
        user = json.dumps({"now": ctx["now"], "weekday": ctx["weekday"], "current_project_id": project_hint,
                           "projects": [{"id": p["id"], "name": p["name"]} for p in ctx["projects"]], "request": str(text)[:2000]})
        r = _extract_json(self._chat(SYSTEM_PARSE, user, timeout=60))
        if not isinstance(r, dict):
            raise AIError("bad JSON shape")
        title = str(r.get("title") or "").strip()
        notes = str(r.get("notes") or "").strip()
        pid = r.get("project_id") if any(p["id"] == r.get("project_id") for p in ctx["projects"]) else None
        out = {"title": (title + ("\n" + notes if notes else "")).strip(), "notes": notes, "project_id": pid or project_hint,
               "project_matched": bool(pid), "priority": r.get("priority") if r.get("priority") in PRIORITIES else "normal",
               "due": r.get("due") if _date(r.get("due") or "") else "", "due_time": r.get("due_time") if _valid_time(r.get("due_time") or "") else "",
               "parsed_by": "model"}
        if not out["due"]:
            out["due_time"] = ""
        if not out["title"]:
            raise AIError("the model returned an empty title")
        return out


def get_provider(cfg, opener=None):
    if cfg.get("provider") == "openai_compatible":
        return OpenAICompatibleProvider(cfg, opener)
    return RulesProvider()


# ---------------------------------------------------------------- entry points the server calls (only when enabled)
def recommend(data, opener=None):
    cfg = ai_settings(data)
    if not cfg["enabled"]:
        raise AIError("AI suggestions are off")
    ctx = build_context(data)
    note = ""
    try:
        prov = get_provider(cfg, opener)
        recs = prov.recommend(ctx, data)
        used = prov.name
    except AIError as e:
        recs, used, note = RulesProvider().recommend(ctx, data), "rules", f"The AI provider failed ({e}), so these come from the built-in rules."
    for r in recs:
        r["changes"] = describe_all(r["actions"], data)
    return {"provider": used, "recommendations": recs, "note": note, "generated": ctx["now"]}


def parse_task(data, text, project_hint=None, opener=None):
    cfg = ai_settings(data)
    if not cfg["enabled"]:
        raise AIError("AI suggestions are off")
    text = str(text or "").strip()
    if not text:
        raise AIError("type or say what to add")
    ctx = build_context(data)
    note = ""
    if cfg["provider"] == "openai_compatible":
        try:
            return {"task": get_provider(cfg, opener).parse(text, ctx, project_hint), "note": ""}
        except AIError as e:
            note = f"The AI provider failed ({e}), so the built-in parser filled this in."
    return {"task": parse_task_rules(text, ctx, project_hint), "note": note}
