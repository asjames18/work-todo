"""Work To-Do email helpers (stdlib only).

- Classic Outlook inbox, READ-ONLY, through outlook_mail.ps1 in a PowerShell subprocess with a timeout.
  It never sends, moves, deletes, flags, marks read or saves mail.
- .eml parsing (email module), .msg parsing (a small built-in Compound File reader, Outlook COM as a backup),
  and pasted email text (e.g. copied from the new Outlook).
- Turning an email into a suggested task with the existing parser in todo_ai.py.
Only headers and a ~300 character preview are kept (email-cache.json). Nothing here writes logs.
"""
import datetime as dt, email, email.policy, email.utils, hashlib, html, html.parser, json, os, re, struct, subprocess, tempfile, threading, time

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.environ.get("TODO_EMAIL_CACHE") or os.path.join(APP_DIR, "email-cache.json")
PS_SCRIPT = os.path.join(APP_DIR, "outlook_mail.ps1")
NO_WINDOW_FLAG = getattr(subprocess, "CREATE_NO_WINDOW", 0)
PREVIEW = 300
COM_TIMEOUT = int(os.environ.get("TODO_EMAIL_TIMEOUT", "30"))
DEFAULT_EMAIL = {"auto": True, "name": "", "address": "", "aliases": "", "days": 7, "max": 50}
STATE_CAP = 3000


class EmailError(Exception):
    pass


def email_settings(data):
    s = dict(DEFAULT_EMAIL)
    raw = ((data or {}).get("settings") or {}).get("email")
    if isinstance(raw, dict):
        for k in DEFAULT_EMAIL:
            if k in raw:
                s[k] = raw[k]
    s["auto"] = s["auto"] is not False
    for k in ("name", "address", "aliases"):
        s[k] = str(s[k] or "")[:300]
    try:
        s["days"] = max(1, min(30, int(s["days"])))
    except (TypeError, ValueError):
        s["days"] = 7
    try:
        s["max"] = max(5, min(200, int(s["max"])))
    except (TypeError, ValueError):
        s["max"] = 50
    return s


def clean_settings(raw):
    """What the PUT handler keeps for settings.email (known keys only)."""
    if not isinstance(raw, dict):
        raise ValueError("bad email settings")
    return {k: v for k, v in raw.items() if k in DEFAULT_EMAIL}


def email_id(*parts):
    return hashlib.sha1("\x1f".join(str(p or "") for p in parts).encode("utf-8", "replace")).hexdigest()[:24]


# ------------------------------------------------------------------ text helpers
class _Strip(html.parser.HTMLParser):
    BLOCK = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "table", "blockquote"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head", "title"):
            self.skip += 1
        elif tag in self.BLOCK:
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head", "title"):
            self.skip = max(0, self.skip - 1)
        elif tag in self.BLOCK:
            self.out.append("\n")

    def handle_data(self, d):
        if not self.skip:
            self.out.append(d)


def html_to_text(s):
    p = _Strip()
    try:
        p.feed(s or "")
        p.close()
    except Exception:
        return re.sub(r"<[^>]+>", " ", s or "")
    return "".join(p.out)


_BANNER = re.compile(r"^\s*(?:\[?external\]?|caution\s*[:\-]|warning\s*[:\-]|this (?:e-?mail|message) (?:originated|came) from outside|"
                     r"you don'?t often get email from|some people who received this message don'?t often)", re.I)
_QUOTE_CUT = re.compile(r"^\s*(?:-{2,}\s*original message\s*-{2,}|_{8,}|from:\s.+|on .{6,120} wrote:|sent from my )", re.I)


def make_preview(body, n=PREVIEW):
    """Readable start of the message: no banners, no quoted reply chain, whitespace collapsed, about n characters."""
    lines = []
    for line in str(body or "").replace("\r", "").split("\n"):
        if _QUOTE_CUT.match(line) and lines:
            break
        if line.lstrip().startswith(">") or _BANNER.match(line):
            continue
        lines.append(line)
    text = re.sub(r"\s+", " ", " ".join(lines)).strip()
    text = re.sub(r"<(?:https?|mailto):[^>]{1,300}>", "", text).strip()
    if len(text) > n:
        cut = text[:n].rsplit(" ", 1)[0] if " " in text[:n] else text[:n]
        text = cut.rstrip(" ,;:.") + "…"
    return text


_SUBJ_PREFIX = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|wg|tr|sv|vs)\s*(?:\[\d+\])?\s*:\s*|\[(?:external|ext|caution)\]\s*:?\s*|external\s*:\s*)+", re.I)


def clean_subject(s):
    return re.sub(r"\s+", " ", _SUBJ_PREFIX.sub("", str(s or ""))).strip()


def _addr(s):
    name, a = email.utils.parseaddr(str(s or ""))
    return (name or "").strip().strip('"'), (a or "").strip()


def _norm_date(value):
    """-> ISO 'YYYY-MM-DDTHH:MM:SS' in local time, or ''."""
    if not value:
        return ""
    if isinstance(value, dt.datetime):
        d = value
    else:
        v = str(value).strip()
        d = None
        v2 = re.sub(r"^\w+day,\s*", "", v)  # e.g. "Friday, October 30, 2026 3:12 PM" from Outlook's copy/forward header
        for fmt in ("%B %d, %Y %I:%M %p", "%b %d, %Y %I:%M %p", "%m/%d/%Y %I:%M %p", "%B %d, %Y %H:%M", "%d %B %Y %H:%M", "%m/%d/%Y %H:%M"):
            try:
                d = dt.datetime.strptime(v2, fmt)
                break
            except ValueError:
                pass
        if d is None:
            try:
                d = dt.datetime.fromisoformat(v.replace("Z", "+00:00"))
            except ValueError:
                d = None
        if d is None:
            try:
                d = email.utils.parsedate_to_datetime(v)
            except (TypeError, ValueError, IndexError):
                d = None
        if d is None:
            return ""
    if d.tzinfo is not None:
        d = d.astimezone().replace(tzinfo=None)
    return d.replace(microsecond=0).isoformat()


def _record(subject, sender, sender_email, received, to, cc, body, source, importance=1, entry_id="", mid=""):
    rec = {"subject": str(subject or "").strip()[:300], "sender": str(sender or "").strip()[:160], "sender_email": str(sender_email or "").strip()[:300],
           "received": _norm_date(received), "to": str(to or "")[:600], "cc": str(cc or "")[:600], "preview": make_preview(body),
           "importance": importance, "source": source}
    rec["id"] = email_id("ol", entry_id) if entry_id else email_id("mid", mid) if mid else email_id("x", rec["subject"], rec["sender_email"] or rec["sender"], rec["received"], rec["preview"][:80])
    return rec


# ------------------------------------------------------------------ .eml
def parse_eml(raw):
    if not raw:
        raise EmailError("the file is empty")
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    if not (msg.get("subject") or msg.get("from")):
        raise EmailError("that doesn't look like an email file")
    body = ""
    try:
        part = msg.get_body(preferencelist=("plain", "html"))
        if part is not None:
            body = part.get_content()
            if part.get_content_type() == "text/html":
                body = html_to_text(body)
    except Exception:
        body = ""
    name, addr = _addr(msg.get("from"))
    imp = 2 if re.search(r"high|urgent|^1", str(msg.get("importance") or msg.get("x-priority") or ""), re.I) else 1
    return _record(msg.get("subject"), name or addr, addr, msg.get("date"), msg.get("to"), msg.get("cc"), body, "eml", imp, mid=str(msg.get("message-id") or ""))


# ------------------------------------------------------------------ .msg (OLE compound file, read-only, stdlib)
class _CFB:
    FREE, EOC = 0xFFFFFFFF, 0xFFFFFFFE

    def __init__(self, b):
        if b[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise EmailError("that isn't an Outlook .msg file")
        self.b = b
        self.ssz = 1 << struct.unpack_from("<H", b, 30)[0]
        self.mssz = 1 << struct.unpack_from("<H", b, 32)[0]
        nfat, dir1, _, self.cutoff, mfat1, nmfat, difat1, ndifat = struct.unpack_from("<IIIIIIII", b, 44)
        difat = list(struct.unpack_from("<109I", b, 76))
        s, guard = difat1, 0
        while s not in (self.FREE, self.EOC) and guard < ndifat + 1:
            ints = struct.unpack_from(f"<{self.ssz // 4}I", self._sec(s))
            difat += ints[:-1]
            s, guard = ints[-1], guard + 1
        self.fat = []
        for s in difat[:nfat]:
            if s in (self.FREE, self.EOC):
                continue
            self.fat += struct.unpack_from(f"<{self.ssz // 4}I", self._sec(s))
        d = self._chain(dir1)
        self.dirs = []
        for i in range(len(d) // 128):
            e = d[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
            typ = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start, size = struct.unpack_from("<IQ", e, 116)
            if self.ssz == 512:
                size &= 0xFFFFFFFF
            self.dirs.append({"name": name, "type": typ, "l": left, "r": right, "c": child, "start": start, "size": size})
        root = self.dirs[0]
        self.ministream = self._chain(root["start"])[:root["size"]] if root["start"] not in (self.FREE, self.EOC) else b""
        self.minifat = []
        if nmfat and mfat1 not in (self.FREE, self.EOC):
            mf = self._chain(mfat1)
            self.minifat = list(struct.unpack_from(f"<{len(mf) // 4}I", mf))

    def _sec(self, s):
        off = (s + 1) * self.ssz
        return self.b[off:off + self.ssz]

    def _chain(self, s):
        out, seen = [], set()
        while s < len(self.fat) and s not in seen:
            seen.add(s)
            out.append(self._sec(s))
            s = self.fat[s]
        return b"".join(out)

    def read(self, e):
        if e["size"] < self.cutoff:
            out, s, seen = [], e["start"], set()
            while s not in (self.FREE, self.EOC) and s < len(self.minifat) and s not in seen:
                seen.add(s)
                out.append(self.ministream[s * self.mssz:(s + 1) * self.mssz])
                s = self.minifat[s]
            return b"".join(out)[:e["size"]]
        return self._chain(e["start"])[:e["size"]]

    def children(self, idx):
        """Direct children of storage idx (walks the red-black tree)."""
        out, stack, seen = [], [self.dirs[idx]["c"]], set()
        while stack:
            i = stack.pop()
            if i == self.FREE or i >= len(self.dirs) or i in seen:
                continue
            seen.add(i)
            out.append(i)
            stack += [self.dirs[i]["l"], self.dirs[i]["r"]]
        return out


def parse_msg(raw):
    """Subject, sender, date and body from a .msg, reading only top-level message properties (attachments are skipped)."""
    try:
        cf = _CFB(raw)
    except EmailError:
        raise
    except Exception:
        raise EmailError("couldn't read that .msg file")
    props = {}
    for i in cf.children(0):
        e = cf.dirs[i]
        m = re.match(r"__substg1\.0_([0-9A-Fa-f]{4})(001F|001E|0102)$", e["name"])
        if e["type"] == 2 and m:
            pid, typ = m.group(1).upper(), m.group(2)
            if pid not in ("0037", "0C1A", "0C1F", "5D01", "0065", "0E04", "0E03", "1000", "1013", "007D", "0042"):
                continue
            data = cf.read(e)
            if typ == "001F":
                val = data.decode("utf-16-le", "replace").rstrip("\x00")
            elif typ == "001E":
                val = data.decode("cp1252", "replace").rstrip("\x00")
            else:
                val = data.decode("utf-8", "replace") if pid == "1013" else ""
            if val and pid not in props:
                props[pid] = val
    received = ""
    for i in cf.children(0):  # fixed-size props: PR_MESSAGE_DELIVERY_TIME (0E06) / PR_CLIENT_SUBMIT_TIME (0039); PR_IMPORTANCE (0017)
        if cf.dirs[i]["name"] == "__properties_version1.0":
            pb = cf.read(cf.dirs[i])
            imp = 1
            for off in range(32, len(pb) - 15, 16):
                ptype, pid = struct.unpack_from("<HH", pb, off)
                if ptype == 0x0040 and pid in (0x0E06, 0x0039) and not received:
                    ft = struct.unpack_from("<Q", pb, off + 8)[0]
                    if ft:
                        received = (dt.datetime(1601, 1, 1, tzinfo=dt.timezone.utc) + dt.timedelta(microseconds=ft // 10))
                if ptype == 0x0003 and pid == 0x0017:
                    imp = struct.unpack_from("<I", pb, off + 8)[0]
            props["_imp"] = imp
    hdr = props.get("007D", "")
    if not received and hdr:
        m = re.search(r"^Date:\s*(.+)$", hdr, re.M | re.I)
        received = m.group(1).strip() if m else ""
    body = props.get("1000") or html_to_text(props.get("1013", ""))
    sender_email = props.get("5D01") or props.get("0065") or props.get("0C1F", "")
    if sender_email.startswith("/"):
        sender_email = ""
    if not (props.get("0037") or props.get("0C1A") or body):
        raise EmailError("couldn't find an email inside that .msg file")
    mid = ""
    if hdr:
        m = re.search(r"^Message-ID:\s*(\S+)", hdr, re.M | re.I)
        mid = m.group(1) if m else ""
    return _record(props.get("0037"), props.get("0C1A") or props.get("0042") or sender_email, sender_email, received, props.get("0E04"),
                   props.get("0E03"), body, "msg", 2 if props.get("_imp") == 2 else 1, mid=mid)


# ------------------------------------------------------------------ pasted / dragged text
_HDR = re.compile(r"^\s*(from|sent|date|to|cc|subject|importance)\s*:\s*(.*)$", re.I)


def parse_text(text):
    """Pasted email text. Reads Outlook-style 'From: / Sent: / To: / Cc: / Subject:' lines when they're there."""
    t = str(text or "").replace("\r", "").strip()
    if not t:
        raise EmailError("paste an email first")
    if re.search(r"<(?:html|div|p|table|span)[\s>]", t[:2000], re.I):
        t = html_to_text(t).strip()
    lines = t.split("\n")
    hdr, body_start, seen_hdr = {}, 0, False
    for i, line in enumerate(lines[:25]):
        m = _HDR.match(line)
        if m and m.group(1).lower() not in hdr:
            hdr[m.group(1).lower()] = m.group(2).strip()
            seen_hdr, body_start = True, i + 1
        elif seen_hdr and line.strip() and not _HDR.match(line) and i - body_start > 1:
            break
        elif seen_hdr and not line.strip():
            body_start = i + 1
            if "subject" in hdr:
                break
    body = "\n".join(lines[body_start:]) if seen_hdr else t
    subject = hdr.get("subject", "")
    if not subject:
        first = next((l.strip() for l in body.split("\n") if l.strip()), "")
        subject = first[:120]
        if not seen_hdr:
            body = body[len(first):] if body.startswith(first) else body
    name, addr = _addr(hdr.get("from", ""))
    if not addr and "@" in hdr.get("from", ""):
        addr = re.search(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+", hdr["from"]).group(0)
    imp = 2 if re.search(r"high", hdr.get("importance", ""), re.I) else 1
    return _record(subject, name or addr or hdr.get("from", ""), addr, hdr.get("sent") or hdr.get("date"), hdr.get("to"), hdr.get("cc"), body, "paste", imp)


def parse_file(name, raw):
    ext = os.path.splitext(str(name or "").lower())[1]
    if ext == ".msg" or raw[:8] == bytes.fromhex("D0CF11E0A1B11AE1"):
        return parse_msg(raw)
    if ext in (".txt", ".text"):
        return parse_text(raw.decode("utf-8", "replace"))
    return parse_eml(raw)


# ------------------------------------------------------------------ who is "me"
def identity(settings, detected=None):
    det = detected or {}
    names = [settings.get("name") or det.get("name") or ""]
    names += [a.strip() for a in re.split(r"[,;]", settings.get("aliases") or "") if a.strip()]
    addrs = [settings.get("address") or ""] + list(det.get("smtp") or [])
    if det.get("x500"):
        addrs.append(det["x500"])
    names = [n for n in dict.fromkeys(n.strip() for n in names) if len(n) >= 2]
    addrs = [a.lower() for a in dict.fromkeys(a.strip() for a in addrs) if a]
    firsts = []
    for n in names:
        if "," in n:  # "Doe, Jane"
            n = n.split(",", 1)[1]
        f = n.strip().split(" ")[0]
        if len(f) >= 3 and f.lower() not in (x.lower() for x in firsts):
            firsts.append(f)
    return {"names": names, "addresses": addrs, "first_names": firsts, "known": bool(names or addrs)}


def _name_in(names, text):
    low = (text or "").lower()
    for n in names:
        variants = {n.lower()}
        if "," in n:
            last, first = [x.strip() for x in n.split(",", 1)]
            variants.add(f"{first} {last}".lower())
        for v in variants:
            if v and re.search(r"(?<![\w])" + re.escape(v) + r"(?![\w])", low):
                return True
    return False


def classify(item, ident):
    """-> 'to', 'cc', 'mention' or '' for one fetched Outlook item."""
    addrs, names = set(ident["addresses"]), ident["names"]
    for r in item.get("recips") or []:
        a = str(r.get("a") or "").lower()
        if (a and a in addrs) or (names and _name_in(names, r.get("n") or "")):
            return "to" if r.get("t") == 1 else "cc" if r.get("t") == 2 else "to"
    if names and _name_in(names, item.get("to")):
        return "to"
    if names and _name_in(names, item.get("cc")):
        return "cc"
    text = (item.get("subject") or "") + " " + (item.get("body") or "")[:1500]
    if ident["first_names"] and re.search(r"(?<![\w])@?(?:" + "|".join(re.escape(f) for f in ident["first_names"]) + r")(?![\w])", text, re.I):
        return "mention"
    for a in addrs:
        if "@" in a and a in text.lower():
            return "mention"
    return ""


# ------------------------------------------------------------------ classic Outlook (read-only)
def run_outlook(action, timeout=None, **env):
    """Run outlook_mail.ps1 for one action. Never raises for Outlook problems; returns a dict with 'state'."""
    if os.name != "nt" and not os.environ.get("TODO_EMAIL_FAKE_PS"):
        return {"state": "not_windows"}
    cmd = os.environ.get("TODO_EMAIL_FAKE_PS")  # tests only: a stand-in script
    args = [cmd] if cmd else ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", PS_SCRIPT]
    e = dict(os.environ, WTD_ACTION=action, **{k: str(v) for k, v in env.items()})
    try:
        r = subprocess.run(args, capture_output=True, timeout=timeout or COM_TIMEOUT, env=e, cwd=APP_DIR, creationflags=NO_WINDOW_FLAG)
    except subprocess.TimeoutExpired:
        return {"state": "timeout"}
    except OSError as ex:
        return {"state": "error", "error": "couldn't start PowerShell"}
    out = (r.stdout or b"").decode("utf-8", "replace").strip().splitlines()
    for line in reversed(out):
        line = line.strip().lstrip("\ufeff")
        if line.startswith("{"):
            try:
                return json.loads(line)
            except ValueError:
                break
    return {"state": "error", "error": "Outlook didn't answer (exit %s)" % r.returncode}


def load_cache():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            c = json.load(f)
        if isinstance(c, dict):
            return c
    except (OSError, ValueError):
        pass
    return {"items": [], "status": {"state": "never"}, "me": {}}


def save_cache(c):
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(c, f, indent=1)
    os.replace(tmp, CACHE_FILE)


def fetch_inbox(settings, runner=None):
    """Read the Inbox through classic Outlook and return a new cache dict (does not save)."""
    runner = runner or run_outlook
    started = time.time()
    res = runner("fetch", days=settings["days"], scan=max(150, settings["max"] * 6))
    now = dt.datetime.now().replace(microsecond=0).isoformat()
    old = load_cache()
    status = {"state": res.get("state", "error"), "checked": now, "seconds": round(time.time() - started, 1)}
    for k in ("error", "profiles", "running", "new_outlook_running", "exe"):
        if k in res:
            status[k] = res[k]
    if status["state"] != "ok":
        old["status"] = status  # keep the last good list on screen
        return old
    me = res.get("me") or {}
    ident = identity(settings, me)
    items = []
    for it in res.get("items") or []:
        why = classify(it, ident) if ident["known"] else "inbox"
        if not why:
            continue
        rec = _record(it.get("subject"), it.get("sender"), it.get("sender_email"), it.get("received"), it.get("to"), it.get("cc"),
                      it.get("body"), "outlook", it.get("importance") or 1, entry_id=it.get("entry_id"))
        rec.update(entry_id=str(it.get("entry_id") or ""), why=why, unread=bool(it.get("unread")))
        items.append(rec)
        if len(items) >= settings["max"]:
            break
    items.sort(key=lambda x: x.get("received") or "", reverse=True)
    status.update(scanned=res.get("scanned", 0), found=len(items), identity_known=ident["known"])
    return {"items": items, "status": status, "me": {"name": me.get("name", ""), "smtp": list(dict.fromkeys(me.get("smtp") or []))[:4]}, "fetched": now}


def public_item(it):
    """What the window gets for one item (no EntryID; it opens by our short id)."""
    return {k: it.get(k) for k in ("id", "subject", "sender", "sender_email", "received", "preview", "why", "unread", "importance", "source", "to", "cc")} | {"can_open": bool(it.get("entry_id"))}


# ------------------------------------------------------------------ remembered state (dismissed / added) lives in todo-data.json
def get_state(data):
    st = data.get("email_state")
    if not isinstance(st, dict):
        st = {}
    st.setdefault("dismissed", {})
    st.setdefault("added", {})
    return st


def set_state(data, eid, op, task_id=None):
    if not re.fullmatch(r"[0-9a-f]{8,40}", str(eid or "")):
        raise EmailError("bad email id")
    st = get_state(data)
    now = dt.datetime.now().replace(microsecond=0).isoformat()
    if op == "dismiss":
        st["dismissed"][eid] = now
    elif op == "undismiss":
        st["dismissed"].pop(eid, None)
    elif op == "added":
        st["added"][eid] = {"task_id": task_id or "", "at": now}
        st["dismissed"].pop(eid, None)
    elif op == "unadd":
        st["added"].pop(eid, None)
    else:
        raise EmailError("bad op")
    for k in ("dismissed", "added"):  # keep the file small
        if len(st[k]) > STATE_CAP:
            keep = sorted(st[k].items(), key=lambda kv: kv[1] if isinstance(kv[1], str) else kv[1].get("at", ""))[-STATE_CAP:]
            st[k] = dict(keep)
    data["email_state"] = st
    return st


# ------------------------------------------------------------------ email -> suggested task
def suggest_task(rec, data, todo_ai, project_hint=None, now=None):
    """Title from the subject, notes with sender/date/preview, due/priority/project from the existing parser."""
    subj = clean_subject(rec.get("subject")) or ("Email from " + (rec.get("sender") or "someone"))
    text = subj + ". " + (rec.get("preview") or "")
    text = re.sub(r"[?()\[\]\"“”]", " , ", text)  # the parser's word boundaries know , . ; ! and spaces, not ? or brackets
    note = ""
    parsed = None
    cfg = todo_ai.ai_settings(data) if todo_ai else {"enabled": False}
    if todo_ai and cfg.get("enabled") and cfg.get("provider") == "openai_compatible":
        try:
            parsed = todo_ai.parse_task(data, text, project_hint)["task"]  # only when the user turned the AI on
        except Exception as e:
            note = "Your AI provider didn't answer, so the built-in parser filled this in."
    if parsed is None and todo_ai:
        ctx = todo_ai.build_context(data)
        base = now
        if base is None and rec.get("received"):
            try:  # "tomorrow" in an email means the day after it was sent
                base = dt.datetime.fromisoformat(rec["received"])
            except ValueError:
                base = None
        parsed = todo_ai.parse_task_rules(text, ctx, project_hint, base)
    parsed = parsed or {}
    due, tm = parsed.get("due") or "", parsed.get("due_time") or ""
    today = dt.date.today().isoformat()
    if due and due < today:
        note = (note + " " if note else "") + f"The email mentioned {due}, which has passed. Check the due date."
    when = ""
    if rec.get("received"):
        try:
            d = dt.datetime.fromisoformat(rec["received"])
            when = d.strftime("%a %b %d, %I:%M %p").replace(" 0", " ")
        except ValueError:
            when = rec["received"]
    who = rec.get("sender") or ""
    if rec.get("sender_email") and rec["sender_email"].lower() not in who.lower():
        who = (who + " <" + rec["sender_email"] + ">").strip()
    lines = [subj, "From: " + (who or "unknown") + (" · " + when if when else "")]
    if rec.get("preview"):
        lines.append(rec["preview"])
    pr = parsed.get("priority") or "normal"
    if rec.get("importance") == 2 and pr == "normal":
        pr = "high"
    return {"task": {"title": "\n".join(lines), "project_id": parsed.get("project_id"), "project_matched": bool(parsed.get("project_matched")),
                     "priority": pr, "due": due, "due_time": tm if due else "", "parsed_by": parsed.get("parsed_by", "rules")}, "note": note}
