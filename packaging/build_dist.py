"""Build WorkToDo-Setup-<version>.zip (+ .sha256) from the clean source.
Usage: python packaging/build_dist.py <src_dir> <out_dir>   (e.g. python packaging/build_dist.py . dist)
Refuses to build if it finds keys or personal strings in the files it packs."""
import hashlib, os, re, sys, zipfile
SRC, OUT = sys.argv[1], sys.argv[2]
HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.join(HERE, "pkg") if os.path.isdir(os.path.join(HERE, "pkg")) else HERE
APP_FILES = ["todo.pyw", "reminder.pyw", "quickadd.pyw", "todo_ai.py", "todo_email.py", "voice_helper.py", "outlook_mail.ps1",
             "wtd_paths.py", "wtd_update.py", "todo.ico", "toast-icon.png"]
ver = re.search(r'VERSION = "([^"]+)"', open(os.path.join(SRC, "wtd_paths.py"), encoding="utf-8").read()).group(1)
root = f"WorkToDo-Setup-{ver}"
os.makedirs(OUT, exist_ok=True)
zpath = os.path.join(OUT, root + ".zip")
BAD = [r"sk-or-v1-[A-Za-z0-9]{10,}", r"\bsk-[A-Za-z0-9]{20,}", r"ghp_[A-Za-z0-9]{20,}", r"AKIA[0-9A-Z]{16}", r"jamesa", r"Antonio", r"Element\s?451", r"Tony, A",
       r"asjames18@", r"sfsc", r"southflorida", r"@[a-z0-9-]+\.edu\b"]

def crlf(b):
    return b.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")

entries = []
for f in APP_FILES:
    entries.append((f"{root}/app/{f}", open(os.path.join(SRC, f), "rb").read()))
for f in ("Install-WorkToDo.cmd", "Uninstall-WorkToDo.cmd", "README.md"):
    entries.append((f"{root}/{f}", open(os.path.join(PKG, f), "rb").read()))
for f in ("Install-WorkToDo.ps1", "Uninstall-WorkToDo.ps1"):
    entries.append((f"{root}/installer/{f}", open(os.path.join(PKG, "installer", f), "rb").read()))
problems = []
for name, b in entries:
    if name.endswith((".png", ".ico")):
        continue
    txt = b.decode("utf-8", "replace")
    for pat in BAD:
        for m in re.finditer(pat, txt):
            problems.append(f"{name}: {pat} -> {txt[max(0, m.start()-30):m.end()+30]!r}")
if problems:
    print("REFUSING TO BUILD, personal data or secrets found:\n" + "\n".join(problems))
    sys.exit(1)
with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
    for name, b in entries:
        if name.endswith((".cmd", ".ps1")):
            b = crlf(b)
        zi = zipfile.ZipInfo(name, date_time=(2026, 11, 13, 12, 0, 0))
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = 0o644 << 16
        z.writestr(zi, b)
sha = hashlib.sha256(open(zpath, "rb").read()).hexdigest()
open(zpath + ".sha256", "w").write(f"{sha}  {os.path.basename(zpath)}\n")
print(zpath, os.path.getsize(zpath), "bytes", "sha256", sha)
