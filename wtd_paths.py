"""Where Work To-Do keeps its code and its data, plus the app version.

Installed from the setup zip (Install-WorkToDo.cmd), the layout is:
    %LOCALAPPDATA%\\WorkToDo\\app\\    code only (replaced by updates)
    %LOCALAPPDATA%\\WorkToDo\\data\\   your tasks, settings, caches, window profile and backups (never touched by updates)
The installer drops app\\installed.json, which is how the app knows it is an installed copy.
Run from a plain folder (no installed.json), data stays next to the code, exactly like older versions."""
import json, os

VERSION = "8.0.1"
UPDATE_REPO = os.environ.get("WORKTODO_UPDATE_REPO") or "asjames18/work-todo"   # GitHub owner/repo whose Releases carry updates

APP_DIR = os.path.dirname(os.path.abspath(__file__))
INSTALL_FILE = os.path.join(APP_DIR, "installed.json")


def _installed():
    try:
        with open(INSTALL_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return None


INSTALLED = _installed()            # None = not an installed copy (e.g. a dev folder)
IS_INSTALLED = INSTALLED is not None
INSTALL_ROOT = os.path.dirname(APP_DIR) if IS_INSTALLED else APP_DIR
if os.environ.get("WORKTODO_DATA_DIR"):
    DATA_DIR = os.environ["WORKTODO_DATA_DIR"]
elif IS_INSTALLED:
    DATA_DIR = os.path.normpath(os.path.join(APP_DIR, INSTALLED.get("data_dir") or os.path.join("..", "data")))
else:
    DATA_DIR = APP_DIR
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
# Files and folders that are user data. Updates never write, move or delete these.
DATA_NAMES = {"todo-data.json", "email-cache.json", "reminder-state.json", "reminder.log", "reminder.log.1", "update-state.json",
              "installed.json", "window-profile", "backups", "voice-tmp", "email-tmp", "data", "settings.json"}


def ensure_data_dir():
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except OSError:
        pass


def version_tuple(v):
    """'v8.1.0' -> (8, 1, 0). Anything after a '-' or '+' is ignored; junk sorts as (0,)."""
    v = str(v or "").strip().lstrip("vV").split("-")[0].split("+")[0]
    out = []
    for part in v.split("."):
        num = "".join(ch for ch in part if ch.isdigit())
        if not num:
            break
        out.append(int(num))
    return tuple(out) or (0,)


def is_newer(tag, current=VERSION):
    a, b = list(version_tuple(tag)), list(version_tuple(current))
    n = max(len(a), len(b))
    return a + [0] * (n - len(a)) > b + [0] * (n - len(b))
