"""Work To-Do voice helper: transcribe one short clip with the Live Transcriber's local Whisper models.
Run with the transcriber's venv python (it has faster-whisper). Reads the models from transcriber\\models,
never downloads anything (HF_HUB_OFFLINE=1) and never writes into the transcriber folder.
Usage: python voice_helper.py <audio file> [--model small.en|large-v3-turbo] [--device auto|cpu|cuda] [--prompt "names"]
Prints one JSON line: {"text": ..., "model": ..., "device": ..., "seconds": ...} or {"error": ...}."""
import os, sys, glob, json, time, argparse

TX_DIR = os.environ.get("WORKTODO_TX_DIR") or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "transcriber")
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")


def add_cuda_dll_dirs():
    for d in glob.glob(os.path.join(sys.prefix, "Lib", "site-packages", "nvidia", "*", "bin")):
        try:
            os.add_dll_directory(d)
        except (OSError, AttributeError):
            pass
        os.environ["PATH"] = d + os.pathsep + os.environ.get("PATH", "")


def load_audio(path):
    """Decode any clip (webm/opus from the window, wav, ...) to 16 kHz mono float32 with PyAV directly.
    (faster-whisper's own decode_audio passes an option newer PyAV versions no longer accept.)"""
    import av, numpy as np
    parts = []
    with av.open(path) as c:
        rs = av.AudioResampler(format="s16", layout="mono", rate=16000)
        for f in c.decode(audio=0):
            for g in rs.resample(f):
                parts.append(g.to_ndarray().reshape(-1))
        try:
            for g in rs.resample(None):
                parts.append(g.to_ndarray().reshape(-1))
        except Exception:
            pass
    if not parts:
        return np.zeros(0, np.float32)
    return np.concatenate(parts).astype(np.float32) / 32768.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("audio")
    ap.add_argument("--model", default="small.en")
    ap.add_argument("--device", default="cpu")
    ap.add_argument("--prompt", default="")
    a = ap.parse_args()
    t0 = time.time()
    add_cuda_dll_dirs()
    from faster_whisper import WhisperModel
    models = os.path.join(TX_DIR, "models")
    audio = load_audio(a.audio)
    if audio.size < 1600:
        print(json.dumps({"text": "", "model": a.model, "device": "-", "seconds": round(time.time() - t0, 2), "audio_seconds": 0}))
        return 0
    tries = []
    if a.device in ("auto", "cuda"):
        tries.append(("cuda", "float16"))
    if a.device in ("auto", "cpu"):
        tries.append(("cpu", "int8"))
    err = None
    for device, compute in tries:
        name = a.model if device == "cuda" or a.model == "small.en" else "small.en"
        try:
            kw = dict(device=device, compute_type=compute, download_root=models, local_files_only=True)
            if device == "cpu":
                kw["cpu_threads"] = max(4, (os.cpu_count() or 8) // 2)
            m = WhisperModel(name, **kw)
            segs, info = m.transcribe(audio, language="en", beam_size=5, vad_filter=True,
                                      initial_prompt=(a.prompt[:600] or None), condition_on_previous_text=False)
            text = " ".join(s.text.strip() for s in segs).strip()
            print(json.dumps({"text": text, "model": name, "device": device, "seconds": round(time.time() - t0, 2),
                              "audio_seconds": round(getattr(info, "duration", 0) or 0, 2)}))
            return 0
        except Exception as e:  # try the next device
            err = f"{type(e).__name__}: {str(e)[:200]}"
    print(json.dumps({"error": err or "no device to try"}))
    return 1


if __name__ == "__main__":
    sys.exit(main())
