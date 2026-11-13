"""Work To-Do quick add: a tiny always-on-top box for typing one task from anywhere.

Opened by the global hotkey (Ctrl+Alt+N), which the background reminder helper (reminder.pyw) registers.
Type naturally, e.g. "Send the report to Sam friday 3pm high #Clients", then Enter. Esc closes.
If the app is running it adds the task through the app (so an open window picks it up right away);
otherwise it writes todo-data.json directly (atomic replace). Stdlib only (tkinter)."""
import json, os, sys, time, urllib.request, uuid
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wtd_paths as P
try:
    import todo_ai
except Exception:
    todo_ai = None

DATA_FILE = os.environ.get("TODO_DATA") or os.path.join(P.DATA_DIR, "todo-data.json")
PORT = int(os.environ.get("TODO_PORT", "48231"))


def load():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def theme(data):
    s = data.get("settings") or {}
    light = s.get("theme") == "light"
    accents = {"violet": "#7c6cff", "blue": "#3b82f6", "teal": "#14b8a6", "green": "#22c55e", "orange": "#f97316", "pink": "#ec4899"}
    acc = accents.get(s.get("accent"), "#7c6cff")
    if light:
        return dict(bg="#f5f6fa", panel="#ffffff", text="#1b2030", muted="#5d667a", line="#d9dde7", acc=acc)
    return dict(bg="#0d1017", panel="#141822", text="#e8ebf2", muted="#8a93a8", line="#262c3b", acc=acc)


def parse(text, data, pid):
    now = time.localtime()
    if todo_ai:
        try:
            ctx = {"now": time.strftime("%Y-%m-%dT%H:%M:%S", now), "projects": data.get("projects", [])}
            r = todo_ai.parse_task_rules(text, ctx, pid)
            return r
        except Exception:
            pass
    return {"title": text.strip(), "project_id": pid, "priority": "normal", "due": "", "due_time": ""}


def add_via_server(task):
    req = urllib.request.Request(f"http://127.0.0.1:{PORT}/api/quick-add", data=json.dumps(task).encode("utf-8"),
                                 headers={"Content-Type": "application/json", "X-WorkTodo": "1"}, method="POST")
    with urllib.request.urlopen(req, timeout=4) as r:
        return json.loads(r.read().decode("utf-8"))


def add_direct(task):
    d = load()
    d.setdefault("projects", [])
    d.setdefault("tasks", [])
    d.setdefault("notes", [])
    d["tasks"].append(task)
    P.ensure_data_dir()
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2)
    os.replace(tmp, DATA_FILE)
    return {"task": task}


def main():
    import tkinter as tk
    data = load()
    projects = sorted(data.get("projects", []), key=lambda p: (p.get("name") or "").lower())
    th = theme(data)
    root = tk.Tk()
    root.title("Quick add - Work To-Do")
    root.configure(bg=th["bg"])
    root.attributes("-topmost", True)
    root.resizable(False, False)
    try:
        root.iconbitmap(os.path.join(P.APP_DIR, "todo.ico"))
    except Exception:
        pass
    w, h = 560, 150
    root.geometry(f"{w}x{h}+{(root.winfo_screenwidth() - w) // 2}+{int(root.winfo_screenheight() * 0.22)}")
    font = ("Segoe UI", 12)
    tk.Label(root, text="Quick add a task", bg=th["bg"], fg=th["muted"], font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=16, pady=(12, 4))
    box = tk.Frame(root, bg=th["acc"], padx=1, pady=1)
    box.pack(fill="x", padx=16)
    entry = tk.Entry(box, font=font, bg=th["panel"], fg=th["text"], insertbackground=th["text"], relief="flat", bd=8)
    entry.pack(fill="x")
    row = tk.Frame(root, bg=th["bg"])
    row.pack(fill="x", padx=16, pady=(8, 0))
    last = (data.get("settings") or {}).get("last_project")
    names = [p.get("name") or "?" for p in projects] or ["(no projects yet)"]
    var = tk.StringVar(value=next((p["name"] for p in projects if p.get("id") == last), names[0]))
    om = tk.OptionMenu(row, var, *names)
    om.configure(bg=th["panel"], fg=th["text"], activebackground=th["line"], activeforeground=th["text"], relief="flat",
                 highlightthickness=0, font=("Segoe UI", 9))
    om["menu"].configure(bg=th["panel"], fg=th["text"])
    om.pack(side="left")
    status = tk.Label(row, text="Enter adds  ·  Esc closes  ·  try \"tomorrow 3pm high\"", bg=th["bg"], fg=th["muted"], font=("Segoe UI", 9))
    status.pack(side="left", padx=10)

    def submit(_e=None):
        text = entry.get().strip()
        if not text:
            return
        pid = next((p["id"] for p in projects if p.get("name") == var.get()), None)
        r = parse(text, data, pid)
        task = {"id": uuid.uuid4().hex, "project_id": r.get("project_id") or pid, "title": r.get("title") or text,
                "priority": r.get("priority") or "normal", "due": r.get("due") or "", "due_time": r.get("due_time") or "",
                "done": False, "created": time.strftime("%Y-%m-%dT%H:%M:%S"), "completed": "", "source": "quick-add"}
        if not task["project_id"]:
            status.configure(text="Add a project in the app first (Ctrl+Alt+T).", fg="#ff5d73")
            return
        try:
            add_via_server(task)
        except Exception:
            try:
                add_direct(task)
            except Exception as e:
                status.configure(text="Couldn't save: " + type(e).__name__, fg="#ff5d73")
                return
        pn = next((p["name"] for p in projects if p["id"] == task["project_id"]), "")
        status.configure(text="\u2713 Added to " + pn + (" \u00b7 due " + task["due"] if task["due"] else ""), fg="#3ddc97")
        entry.delete(0, "end")
        root.after(900, root.destroy)

    entry.bind("<Return>", submit)
    root.bind("<Escape>", lambda e: root.destroy())
    root.after(30000, root.destroy)  # never linger forever

    def grab_focus():
        root.lift()
        root.focus_force()
        entry.focus_set()
        if os.name == "nt":
            try:
                import ctypes
                ctypes.windll.user32.SetForegroundWindow(root.winfo_id())
                ctypes.windll.user32.SetForegroundWindow(ctypes.windll.user32.GetParent(root.winfo_id()))
            except Exception:
                pass
    root.after(60, grab_focus)
    root.mainloop()


if __name__ == "__main__":
    main()
