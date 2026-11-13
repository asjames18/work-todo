"""Work To-Do: a small local to-do app with projects and priorities.
Data is saved automatically to todo-data.json next to this file."""
import json
import os
import uuid
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(APP_DIR, "todo-data.json")
PRIORITIES = ["urgent", "high", "normal"]
RANK = {p: i for i, p in enumerate(PRIORITIES)}
ALL = "__all__"


def load_data():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        data = {}
    except (json.JSONDecodeError, OSError) as e:
        messagebox.showerror("Work To-Do", f"Couldn't read {DATA_FILE}:\n{e}\n\nStarting with an empty list; the old file is kept as .bak.")
        try:
            os.replace(DATA_FILE, DATA_FILE + ".bak")
        except OSError:
            pass
        data = {}
    data.setdefault("projects", [])
    data.setdefault("tasks", [])
    return data


def save_data(data):
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, DATA_FILE)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Work To-Do")
        self.geometry("1000x600")
        self.minsize(800, 450)
        self.data = load_data()
        self.current = ALL
        self._build()
        self.refresh_projects()
        self.refresh_tasks()

    # ---------- layout ----------
    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=24)

        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        left = ttk.Frame(paned, padding=4)
        ttk.Label(left, text="Projects", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.proj_list = tk.Listbox(left, exportselection=False, font=("Segoe UI", 10), activestyle="none")
        self.proj_list.pack(fill="both", expand=True, pady=4)
        self.proj_list.bind("<<ListboxSelect>>", self.on_project_select)
        self.proj_list.bind("<Double-Button-1>", lambda e: self.rename_project())
        btns = ttk.Frame(left)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add", command=self.add_project).pack(side="left", expand=True, fill="x")
        ttk.Button(btns, text="Rename", command=self.rename_project).pack(side="left", expand=True, fill="x")
        ttk.Button(btns, text="Delete", command=self.delete_project).pack(side="left", expand=True, fill="x")
        paned.add(left, weight=1)

        right = ttk.Frame(paned, padding=4)
        paned.add(right, weight=4)

        form = ttk.LabelFrame(right, text="New task", padding=6)
        form.pack(fill="x")
        ttk.Label(form, text="Task").grid(row=0, column=0, sticky="w")
        self.title_var = tk.StringVar()
        title_entry = ttk.Entry(form, textvariable=self.title_var)
        title_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        title_entry.bind("<Return>", lambda e: self.add_task())
        ttk.Label(form, text="Project").grid(row=0, column=1, sticky="w")
        self.task_proj_var = tk.StringVar()
        self.task_proj_box = ttk.Combobox(form, textvariable=self.task_proj_var, state="readonly", width=18)
        self.task_proj_box.grid(row=1, column=1, padx=(0, 6))
        ttk.Label(form, text="Priority").grid(row=0, column=2, sticky="w")
        self.prio_var = tk.StringVar(value="normal")
        ttk.Combobox(form, textvariable=self.prio_var, values=["normal", "high", "urgent"], state="readonly", width=9).grid(row=1, column=2, padx=(0, 6))
        ttk.Label(form, text="Due (YYYY-MM-DD, optional)").grid(row=0, column=3, sticky="w")
        self.due_var = tk.StringVar()
        due_entry = ttk.Entry(form, textvariable=self.due_var, width=14)
        due_entry.grid(row=1, column=3, sticky="w", padx=(0, 6))
        due_entry.bind("<Return>", lambda e: self.add_task())
        ttk.Button(form, text="Add task", command=self.add_task).grid(row=1, column=4)
        form.columnconfigure(0, weight=1)

        bar = ttk.Frame(right, padding=(0, 6))
        bar.pack(fill="x")
        self.heading = ttk.Label(bar, text="", font=("Segoe UI", 11, "bold"))
        self.heading.pack(side="left")
        self.show_done = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="Show completed", variable=self.show_done, command=self.refresh_tasks).pack(side="right")
        self.filter_var = tk.StringVar(value="All priorities")
        fbox = ttk.Combobox(bar, textvariable=self.filter_var, values=["All priorities", "urgent", "high", "normal"], state="readonly", width=14)
        fbox.pack(side="right", padx=6)
        fbox.bind("<<ComboboxSelected>>", lambda e: self.refresh_tasks())
        ttk.Label(bar, text="Filter:").pack(side="right")

        cols = ("priority", "title", "project", "due", "status")
        tree_frame = ttk.Frame(right)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        for c, w, label in [("priority", 80, "Priority"), ("title", 380, "Task"), ("project", 150, "Project"), ("due", 100, "Due"), ("status", 80, "Status")]:
            self.tree.heading(c, text=label)
            self.tree.column(c, width=w, anchor="w", stretch=(c == "title"))
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("urgent", foreground="#C62828", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("high", foreground="#E65100")
        self.tree.tag_configure("normal", foreground="#212121")
        self.tree.tag_configure("done", foreground="#9E9E9E")
        self.tree.tag_configure("overdue", background="#FFEBEE")
        self.tree.bind("<Double-Button-1>", lambda e: self.toggle_done())
        self.tree.bind("<Delete>", lambda e: self.delete_tasks())

        actions = ttk.Frame(right, padding=(0, 6, 0, 0))
        actions.pack(fill="x")
        ttk.Button(actions, text="Mark done / undo", command=self.toggle_done).pack(side="left")
        ttk.Button(actions, text="Change priority", command=self.change_priority).pack(side="left", padx=6)
        ttk.Button(actions, text="Edit task", command=self.edit_task).pack(side="left")
        ttk.Button(actions, text="Delete task", command=self.delete_tasks).pack(side="left", padx=6)
        self.status = ttk.Label(actions, text="")
        self.status.pack(side="right")

    # ---------- helpers ----------
    def project_name(self, pid):
        for p in self.data["projects"]:
            if p["id"] == pid:
                return p["name"]
        return "(none)"

    def sorted_projects(self):
        return sorted(self.data["projects"], key=lambda p: p["name"].lower())

    def save(self):
        try:
            save_data(self.data)
        except OSError as e:
            messagebox.showerror("Work To-Do", f"Couldn't save:\n{e}")

    def selected_task_ids(self):
        return list(self.tree.selection())

    def find_task(self, tid):
        return next((t for t in self.data["tasks"] if t["id"] == tid), None)

    # ---------- projects ----------
    def refresh_projects(self):
        self.proj_list.delete(0, "end")
        self._proj_ids = [ALL]
        open_total = sum(1 for t in self.data["tasks"] if not t["done"])
        self.proj_list.insert("end", f"All projects  ({open_total})")
        for p in self.sorted_projects():
            n = sum(1 for t in self.data["tasks"] if t["project_id"] == p["id"] and not t["done"])
            self.proj_list.insert("end", f"{p['name']}  ({n})")
            self._proj_ids.append(p["id"])
        idx = self._proj_ids.index(self.current) if self.current in self._proj_ids else 0
        self.current = self._proj_ids[idx]
        self.proj_list.selection_clear(0, "end")
        self.proj_list.selection_set(idx)
        names = [p["name"] for p in self.sorted_projects()]
        self.task_proj_box["values"] = names
        if self.current != ALL:
            self.task_proj_var.set(self.project_name(self.current))
        elif self.task_proj_var.get() not in names:
            self.task_proj_var.set(names[0] if names else "")

    def on_project_select(self, _e=None):
        sel = self.proj_list.curselection()
        if not sel:
            return
        self.current = self._proj_ids[sel[0]]
        if self.current != ALL:
            self.task_proj_var.set(self.project_name(self.current))
        self.refresh_tasks()

    def ask_name(self, title, initial=""):
        name = simpledialog.askstring(title, "Project name:", initialvalue=initial, parent=self)
        if name is None:
            return None
        name = name.strip()
        if not name:
            return None
        if any(p["name"].lower() == name.lower() and p["name"] != initial for p in self.data["projects"]):
            messagebox.showwarning("Work To-Do", f'A project named "{name}" already exists.')
            return None
        return name

    def add_project(self):
        name = self.ask_name("Add project")
        if not name:
            return
        pid = uuid.uuid4().hex
        self.data["projects"].append({"id": pid, "name": name})
        self.current = pid
        self.save()
        self.refresh_projects()
        self.refresh_tasks()

    def rename_project(self):
        if self.current == ALL:
            return
        p = next(p for p in self.data["projects"] if p["id"] == self.current)
        name = self.ask_name("Rename project", p["name"])
        if not name:
            return
        p["name"] = name
        self.save()
        self.refresh_projects()
        self.refresh_tasks()

    def delete_project(self):
        if self.current == ALL:
            messagebox.showinfo("Work To-Do", "Pick a project to delete first.")
            return
        name = self.project_name(self.current)
        n = sum(1 for t in self.data["tasks"] if t["project_id"] == self.current)
        if not messagebox.askyesno("Delete project", f'Delete "{name}" and its {n} task(s)? This can\'t be undone.'):
            return
        self.data["tasks"] = [t for t in self.data["tasks"] if t["project_id"] != self.current]
        self.data["projects"] = [p for p in self.data["projects"] if p["id"] != self.current]
        self.current = ALL
        self.save()
        self.refresh_projects()
        self.refresh_tasks()

    # ---------- tasks ----------
    def parse_due(self, text):
        text = text.strip()
        if not text:
            return ""
        try:
            return dt.date.fromisoformat(text).isoformat()
        except ValueError:
            messagebox.showwarning("Work To-Do", "Due date should look like 2026-10-31, or leave it blank.")
            return None

    def add_task(self):
        title = self.title_var.get().strip()
        if not title:
            return
        if not self.data["projects"]:
            messagebox.showinfo("Work To-Do", "Add a project first (left side), then add tasks to it.")
            return
        pname = self.task_proj_var.get()
        pid = next((p["id"] for p in self.data["projects"] if p["name"] == pname), None)
        if pid is None:
            messagebox.showinfo("Work To-Do", "Pick a project for this task.")
            return
        due = self.parse_due(self.due_var.get())
        if due is None:
            return
        self.data["tasks"].append({
            "id": uuid.uuid4().hex, "project_id": pid, "title": title,
            "priority": self.prio_var.get() or "normal", "due": due, "done": False,
            "created": dt.datetime.now().isoformat(timespec="seconds"), "completed": "",
        })
        self.title_var.set("")
        self.due_var.set("")
        self.prio_var.set("normal")
        self.save()
        self.refresh_projects()
        self.refresh_tasks()

    def refresh_tasks(self):
        self.tree.delete(*self.tree.get_children())
        flt = self.filter_var.get()
        today = dt.date.today().isoformat()
        tasks = [t for t in self.data["tasks"]
                 if (self.current == ALL or t["project_id"] == self.current)
                 and (self.show_done.get() or not t["done"])
                 and (flt == "All priorities" or t["priority"] == flt)]
        tasks.sort(key=lambda t: (t["done"], RANK.get(t["priority"], 9), t["due"] or "9999-99-99", t["created"]))
        for t in tasks:
            tags = ["done"] if t["done"] else [t["priority"]]
            if not t["done"] and t["due"] and t["due"] < today:
                tags.append("overdue")
            self.tree.insert("", "end", iid=t["id"], tags=tags, values=(
                t["priority"].capitalize(), t["title"], self.project_name(t["project_id"]),
                t["due"] or "", "Done" if t["done"] else ("Overdue" if "overdue" in tags else "Open")))
        self.heading.config(text="All projects" if self.current == ALL else self.project_name(self.current))
        open_n = sum(1 for t in tasks if not t["done"])
        urgent_n = sum(1 for t in tasks if not t["done"] and t["priority"] == "urgent")
        self.status.config(text=f"{open_n} open, {urgent_n} urgent")

    def toggle_done(self):
        ids = self.selected_task_ids()
        for tid in ids:
            t = self.find_task(tid)
            if t:
                t["done"] = not t["done"]
                t["completed"] = dt.datetime.now().isoformat(timespec="seconds") if t["done"] else ""
        if ids:
            self.save()
            self.refresh_projects()
            self.refresh_tasks()

    def change_priority(self):
        ids = self.selected_task_ids()
        if not ids:
            return
        p = simpledialog.askstring("Change priority", "New priority (normal, high, urgent):", parent=self)
        if p is None:
            return
        p = p.strip().lower()
        if p not in RANK:
            messagebox.showwarning("Work To-Do", "Priority must be normal, high, or urgent.")
            return
        for tid in ids:
            t = self.find_task(tid)
            if t:
                t["priority"] = p
        self.save()
        self.refresh_tasks()

    def edit_task(self):
        ids = self.selected_task_ids()
        if len(ids) != 1:
            messagebox.showinfo("Work To-Do", "Select one task to edit.")
            return
        t = self.find_task(ids[0])
        title = simpledialog.askstring("Edit task", "Task:", initialvalue=t["title"], parent=self)
        if title is None or not title.strip():
            return
        due_text = simpledialog.askstring("Edit task", "Due date (YYYY-MM-DD, blank for none):", initialvalue=t["due"], parent=self)
        if due_text is None:
            return
        due = self.parse_due(due_text)
        if due is None:
            return
        t["title"], t["due"] = title.strip(), due
        self.save()
        self.refresh_tasks()

    def delete_tasks(self):
        ids = self.selected_task_ids()
        if not ids:
            return
        if not messagebox.askyesno("Delete task", f"Delete {len(ids)} task(s)? This can't be undone."):
            return
        self.data["tasks"] = [t for t in self.data["tasks"] if t["id"] not in ids]
        self.save()
        self.refresh_projects()
        self.refresh_tasks()


if __name__ == "__main__":
    App().mainloop()
