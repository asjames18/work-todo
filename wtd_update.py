"""Work To-Do updates from GitHub Releases (stdlib only).

check():  GET https://api.github.com/repos/<UPDATE_REPO>/releases/latest  (quiet on 404 / offline)
apply():  download the release's .zip asset to a temp folder, verify SHA-256 (a "sha256: <hex>" line in the release
          notes or a .sha256 asset, when there is one), back up the current app folder, copy in ONLY the code files
          from the zip's app\\ folder, and roll everything back if any step fails. User data (DATA_DIR) is never touched.
Only installed copies (app\\installed.json, see wtd_paths) can apply updates."""
import hashlib, json, os, re, shutil, tempfile, time, urllib.error, urllib.request, zipfile
import wtd_paths as P

API = "https://api.github.com/repos/{repo}/releases/latest"
STATE_FILE = os.path.join(P.DATA_DIR, "update-state.json")
CHECK_EVERY = 20 * 3600          # "at most once a day" with a little slack
MAX_ZIP = 60 * 1024 * 1024
UA = "WorkToDo-Updater/" + P.VERSION
SHA_RE = re.compile(r"sha-?256\s*[:=]?\s*`?([0-9a-fA-F]{64})`?", re.I)
REQUIRED = ("todo.pyw", "wtd_paths.py")


class UpdateError(Exception):
    pass


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            s = json.load(f)
        return s if isinstance(s, dict) else {}
    except (OSError, ValueError):
        return {}


def save_state(s):
    P.ensure_data_dir()
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(s, f, indent=2)
    os.replace(tmp, STATE_FILE)


def _get(url, timeout=10, accept="application/vnd.github+json", opener=None):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": accept})
    return (opener or urllib.request.urlopen)(req, timeout=timeout)


def _pick_assets(rel):
    assets = [a for a in rel.get("assets") or [] if isinstance(a, dict)]
    zips = [a for a in assets if str(a.get("name", "")).lower().endswith(".zip")]
    zips.sort(key=lambda a: (0 if str(a.get("name", "")).lower().startswith("worktodo") else 1, str(a.get("name"))))
    z = zips[0] if zips else None
    sha = None
    if z:
        sha = next((a for a in assets if str(a.get("name", "")).lower() in (z["name"].lower() + ".sha256", z["name"].lower()[:-4] + ".sha256")), None)
    sha = sha or next((a for a in assets if str(a.get("name", "")).lower().endswith(".sha256")), None)
    return z, sha


def check(repo=None, opener=None, now=None):
    """Returns a small dict. Never raises: offline / 404 / bad JSON become state 'offline' / 'none' / 'error'."""
    repo = repo or P.UPDATE_REPO
    now = now or time.time()
    out = {"checked": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(now)), "repo": repo, "current": P.VERSION, "available": False}
    try:
        with _get(API.format(repo=repo), opener=opener) as r:
            rel = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        out["state"] = "none" if e.code == 404 else "error"
        out["detail"] = "no published release yet" if e.code == 404 else f"GitHub answered HTTP {e.code}"
        return out
    except (urllib.error.URLError, OSError, ValueError) as e:
        out["state"] = "offline"
        out["detail"] = "couldn't reach GitHub"
        return out
    if not isinstance(rel, dict) or not rel.get("tag_name"):
        out.update(state="error", detail="unexpected reply from GitHub")
        return out
    z, sha = _pick_assets(rel)
    body = str(rel.get("body") or "")[:8000]
    m = SHA_RE.search(body)
    out.update(state="ok", tag=str(rel["tag_name"])[:40], name=str(rel.get("name") or rel["tag_name"])[:120], notes=body,
               url=str(rel.get("html_url") or "")[:300], published=str(rel.get("published_at") or "")[:40],
               asset=(z or {}).get("name"), asset_url=(z or {}).get("browser_download_url"), asset_size=(z or {}).get("size"),
               sha_asset_url=(sha or {}).get("browser_download_url"), sha256=(m.group(1).lower() if m else None))
    out["available"] = P.is_newer(out["tag"]) and bool(out["asset_url"])
    return out


def due(state, now=None):
    now = now or time.time()
    return now - float(state.get("last_check_ts") or 0) >= CHECK_EVERY


def check_and_store(repo=None, opener=None):
    res = check(repo, opener)
    s = load_state()
    s.update(last_check_ts=time.time(), last=res)
    try:
        save_state(s)
    except OSError:
        pass
    return res


def _expected_sha(info, opener=None):
    if info.get("sha256"):
        return info["sha256"]
    if info.get("sha_asset_url"):
        try:
            with _get(info["sha_asset_url"], accept="application/octet-stream", opener=opener) as r:
                m = re.search(r"\b([0-9a-fA-F]{64})\b", r.read(4096).decode("utf-8", "replace"))
            if m:
                return m.group(1).lower()
        except (urllib.error.URLError, OSError):
            raise UpdateError("couldn't download the checksum file")
        raise UpdateError("the checksum file doesn't contain a SHA-256")
    return None


def _download(url, dest, opener=None):
    h = hashlib.sha256()
    n = 0
    try:
        with _get(url, timeout=60, accept="application/octet-stream", opener=opener) as r, open(dest, "wb") as f:
            while True:
                chunk = r.read(65536)
                if not chunk:
                    break
                n += len(chunk)
                if n > MAX_ZIP:
                    raise UpdateError("the download is unexpectedly large")
                h.update(chunk)
                f.write(chunk)
    except (urllib.error.URLError, OSError) as e:
        raise UpdateError("the download failed (" + type(e).__name__ + ")")
    return h.hexdigest()


def _find_app_dir(root):
    for cur, dirs, files in os.walk(root):
        depth = os.path.relpath(cur, root).count(os.sep)
        if depth > 3:
            dirs[:] = []
            continue
        if all(f in files for f in REQUIRED):
            return cur
    return None


def _code_files(app_src):
    """Relative paths of the files an update may write: everything in the package's app folder except user-data names."""
    out = []
    for cur, dirs, files in os.walk(app_src):
        dirs[:] = [d for d in dirs if d not in P.DATA_NAMES and d != "__pycache__"]
        for f in files:
            if f in P.DATA_NAMES or f.endswith((".pyc", ".tmp")):
                continue
            out.append(os.path.relpath(os.path.join(cur, f), app_src))
    return out


def apply(info, app_dir=None, opener=None, log=None):
    """Install the release described by info (from check()). Returns {"version", "backup", "files"}. Rolls back on failure."""
    log = log or (lambda m: None)
    app_dir = app_dir or P.APP_DIR
    if not P.IS_INSTALLED and app_dir == P.APP_DIR:
        raise UpdateError("this copy wasn't installed from a release, so it can't update itself. Run the latest installer instead.")
    if not info or not info.get("available") or not info.get("asset_url"):
        raise UpdateError("there's no newer release to install")
    if not str(info["asset_url"]).startswith("https://"):
        raise UpdateError("refusing a non-https download")
    work = tempfile.mkdtemp(prefix="worktodo-update-")
    backup = None
    written = []
    try:
        zpath = os.path.join(work, "update.zip")
        expected = _expected_sha(info, opener)
        got = _download(info["asset_url"], zpath, opener)
        if expected and got != expected:
            raise UpdateError("the download's SHA-256 doesn't match the release, so nothing was changed")
        log(f"downloaded {info.get('asset')} sha256={got} verified={bool(expected)}")
        ex = os.path.join(work, "x")
        with zipfile.ZipFile(zpath) as z:
            for m in z.infolist():   # no absolute paths or ".." escapes
                target = os.path.normpath(os.path.join(ex, m.filename))
                if not target.startswith(os.path.normpath(ex) + os.sep) and target != os.path.normpath(ex):
                    raise UpdateError("the update zip has an unsafe path")
            z.extractall(ex)
        src = _find_app_dir(ex)
        if not src:
            raise UpdateError("the update zip doesn't contain the app")
        files = _code_files(src)
        # back up the current code (not data) first
        stamp = time.strftime("%Y%m%d-%H%M%S")
        P.ensure_data_dir()
        backup = os.path.join(P.BACKUP_DIR, f"app-{P.VERSION}-{stamp}")
        shutil.copytree(app_dir, backup, ignore=shutil.ignore_patterns("__pycache__", *P.DATA_NAMES))
        log(f"backed up code to {backup}")
        for rel in files:
            dst = os.path.join(app_dir, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            existed = os.path.exists(dst)
            shutil.copy2(os.path.join(src, rel), dst)
            written.append((rel, existed))
        log(f"installed {len(files)} files")
        return {"version": info.get("tag"), "backup": backup, "files": len(files), "verified": bool(expected), "sha256": got}
    except Exception as e:
        if backup and os.path.isdir(backup):
            for rel, existed in written:   # roll back: restore old files, remove new ones
                dst = os.path.join(app_dir, rel)
                old = os.path.join(backup, rel)
                try:
                    if existed and os.path.exists(old):
                        shutil.copy2(old, dst)
                    elif not existed and os.path.exists(dst):
                        os.remove(dst)
                except OSError:
                    pass
            log(f"rolled back after: {e!r}")
        if isinstance(e, UpdateError):
            raise
        raise UpdateError("the update failed (" + type(e).__name__ + "); your app was left as it was")
    finally:
        shutil.rmtree(work, ignore_errors=True)
