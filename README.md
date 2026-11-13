# Work To-Do

A local to-do and planning app for Windows. Your projects, tasks and notes stay on this PC.

[![Latest release](https://img.shields.io/github/v/release/asjames18/work-todo)](https://github.com/asjames18/work-todo/releases/latest)
![Platform](https://img.shields.io/badge/Windows-10%20%2F%2011-0078D6?logo=windows&logoColor=white)
![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/license-not%20specified-lightgrey)

## Overview

Work To-Do is a desktop planner with a Today dashboard, projects, recurring tasks, desktop reminders, a focus timer and a weekly review. The core app uses the Python standard library only. It runs a local server on `127.0.0.1` (default port `48231`) and opens in a Microsoft Edge window, or the system browser if Edge is not available.

The ready-made installer can set up Python if it is missing, and it can optionally add on-device voice input. Optional Outlook and AI features stay off until you turn them on.

## Key features

- **Today dashboard** — A daily brief, overdue and due-today tasks, priorities by project, today's schedule and the next seven days, and a first-run checklist.
- **Projects, tasks and notes** — Priorities, due dates, smart views and undo.
- **Recurring tasks** — Daily, weekdays, weekly on a chosen day, or monthly on day N. Completing one creates the next.
- **Focus timer** — A Pomodoro-style timer (25/5 by default, adjustable) that logs time per task and project.
- **Weekly review** — Completed work, slipped or overdue items, the week ahead, time per project, and an optional recap.
- **Desktop reminders** — Morning summary, heads-up and due alerts, including when the main window is closed (via `reminder.pyw`).
- **Backup and restore** — Everything in one JSON file. Restore makes a backup first.
- **Themes** — Light and dark, with six accent colours.
- **Keyboard shortcuts and a first-run tour** — Press `?` in the app for the full list. A global quick-add shortcut (`Ctrl+Alt+N`) works even when the window is closed.
- **Optional voice input** — Speak a task and it is transcribed on this PC with Whisper. If [Live Transcriber](https://github.com/asjames18/live-transcriber) is already installed, Work To-Do can reuse that speech engine.
- **Optional classic Outlook inbox** — Read-only. You can turn an email into a task; nothing is sent, moved, deleted or marked read.
- **Optional AI suggestions** — Off until you enable them. Built-in rules run on this PC with no network. An OpenAI-compatible provider such as OpenRouter can be used if you set it up. Suggested changes are shown for you to confirm.
- **In-app updates** — Installed copies can check GitHub Releases, verify the SHA-256, back up the current app folder and roll back if something fails.

![Work To-Do Today dashboard](docs/screenshot-dashboard.png)

| Setup wizard | Weekly review | Light theme |
| --- | --- | --- |
| ![Setup wizard](docs/screenshot-wizard.png) | ![Weekly review](docs/screenshot-review.png) | ![Light theme](docs/screenshot-light.png) |

## Download

The installer is the supported way to install Work To-Do:

- [Latest GitHub release](https://github.com/asjames18/work-todo/releases/latest) — `WorkToDo-Setup-8.0.2.zip`
- [melanatedintech.com/apps](https://melanatedintech.com/apps)

Current release: **8.0.2**. The zip is 189 KB. Tested on Windows 11.

## Install

1. Download `WorkToDo-Setup-8.0.2.zip` and open it.
2. Double-click `Install-WorkToDo.cmd`. If you run it from inside the zip, it unpacks itself first. Extracting the zip first also works.
3. If Windows SmartScreen says "Windows protected your PC", click **More info**, then **Run anyway**. The installer is not code-signed.
4. If Python 3.10 or newer is missing, the installer can install Python 3.13.17 for your account from python.org (the download's SHA-256 and signature are checked; no administrator account is required).
5. It will offer optional voice input. That step downloads a private faster-whisper environment and the `small.en` Whisper model (about 550 MB, roughly 1 GB on disk, usually 3–10 minutes). If Live Transcriber is already installed, Work To-Do can reuse that model and skip the download. Voice input needs Python 3.12 or newer.
6. The installer copies the app to `%LOCALAPPDATA%\WorkToDo\app`, keeps your data in `%LOCALAPPDATA%\WorkToDo\data`, adds Desktop and Start menu shortcuts (`Ctrl+Alt+T`), and can start the reminder helper at sign-in. The first launch opens a short setup wizard.

Nothing is downloaded without asking. If you are offline or a download fails, the installer stops safely; running it again resumes and skips what is already done. No administrator rights are needed for a normal install. A longer end-user guide ships in the zip as `README.md` (the same text as [`packaging/README.md`](packaging/README.md) in this repository).

To uninstall, double-click `Uninstall-WorkToDo.cmd` or use **Start menu > Uninstall Work To-Do**. Your tasks are kept unless you choose to delete them. Python stays installed.

## Verify the SHA-256

The SHA-256 of `WorkToDo-Setup-8.0.2.zip` is:

```
690429c47931433382bb09df75c61dd14bf0198a88a252c983f3297b4164d273
```

On Windows, open Command Prompt in the folder that contains the zip and run:

```bat
certutil -hashfile WorkToDo-Setup-8.0.2.zip SHA256
```

The hash should match the value above and the `WorkToDo-Setup-8.0.2.zip.sha256` file attached to the [release](https://github.com/asjames18/work-todo/releases/latest). If it does not match, do not run the installer.

## System requirements

- Windows 10 or 11 (64-bit)
- Python 3.10 or newer for the core app (the installer can install 3.13.17 if needed). Voice input needs Python 3.12 or newer.
- Microsoft Edge, or another system browser as a fallback
- Internet access only for the first-time Python or Whisper downloads, optional AI, and optional in-app updates
- Classic Outlook (optional), for the read-only work-email inbox
- [Live Transcriber](https://github.com/asjames18/live-transcriber) (optional), if you want Work To-Do to reuse an existing Whisper installation

The core app does not need pip packages. Voice input uses the packages in `requirements-voice.txt`.

## How to use it

1. Start Work To-Do from the Desktop or Start menu shortcut. The Today dashboard opens first.
2. Add projects from the sidebar, then add tasks and notes. Priorities and due dates can be set from the task card or with keyboard shortcuts (`P`, `D`, `T`).
3. Turn on desktop reminders in **Reminders & settings** (`Ctrl+,`) and keep the reminder helper running if you want notifications when the window is closed.
4. Optionally connect classic Outlook on the Email view, or drag or paste an email and confirm it as a task.
5. Optionally enable AI suggestions and review each proposed change before it is applied.
6. Press `?` at any time for keyboard shortcuts. Press `T` after that to replay the tour.

## Privacy

Work To-Do is local-first.

- The server binds to `127.0.0.1` only and rejects requests whose `Host` header is not `127.0.0.1` or `localhost` on the app port.
- Projects, tasks, notes and settings are stored in `%LOCALAPPDATA%\WorkToDo\data\todo-data.json` (or next to the app when you run from source).
- The reminder helper never changes your tasks. It only records which notifications it has already shown, in `reminder-state.json`.
- The Outlook helper is read-only. The local email cache (`email-cache.json`) keeps headers and a short preview, not the full body.
- AI is off until you enable it. Built-in rules make no network calls. When an OpenAI-compatible provider is selected, the app sends a snapshot of to-do state (and, for notes or parse requests, the relevant text) to the endpoint you configure. If you paste an OpenRouter key in the setup wizard, it is written only to the Windows user environment variable `OPENROUTER_API_KEY`, never to the app's files. The key is read at call time and is not stored in `todo-data.json` or logs.
- Voice clips used by “Ask AI to add” are transcribed on this PC. That audio is not sent online.
- Meeting notes from Live Transcriber, if you use them, are read only.

## Build from source

From a clone of this repository, with Python 3.10 or newer:

```bat
pythonw todo.pyw
```

The reminder helper is a separate no-window process:

```bat
pythonw reminder.pyw
```

`todo-classic.pyw` is an older Tkinter version kept for reference.

To build the installer zip:

```bat
python packaging\build_dist.py . dist
```

This writes `dist\WorkToDo-Setup-8.0.2.zip` and `dist\WorkToDo-Setup-8.0.2.zip.sha256`. The packager refuses to build if packed files contain API keys or personal strings. Scripts are converted to CRLF. The zip layout is `app\`, `installer\`, `Install-WorkToDo.cmd`, `Uninstall-WorkToDo.cmd` and `README.md`.

Useful environment variables when running from source:

| Variable | Purpose |
| --- | --- |
| `WORKTODO_DATA_DIR` | Folder for data files (default: `..\data` when installed, otherwise next to the app) |
| `TODO_DATA` | Path to the data file (overrides the folder above) |
| `TODO_PORT` | Main app port (default `48231`) |
| `TODO_REMINDER_PORT` | Reminder helper port (default `48232`) |
| `WORKTODO_TX_DIR` | Live Transcriber install folder |
| `WORKTODO_UPDATE_REPO` | `owner/repo` to check for releases (default `asjames18/work-todo`) |

Copy `settings.example.json` if you want a reference for the settings shape. Live settings are stored inside `todo-data.json`, which is created on first run and is not committed.

## Troubleshooting

**SmartScreen blocks the installer.** Click **More info**, then **Run anyway**. The installer is not code-signed.

**The installer says Python is missing.** Let it install Python 3.13.17 for your account, or install Python 3.10+ (3.12+ for voice) from [python.org](https://www.python.org/) yourself and run the installer again.

**Voice input is skipped or fails.** It needs a network connection for the first download (about 550 MB) unless Live Transcriber is already installed. You can run the installer again later and accept the voice-input step.

**The window does not open.** Confirm Microsoft Edge or another browser is available. If port `48231` is already in use, set `TODO_PORT` to a free port.

**Reminders do not appear when the window is closed.** Start `reminder.pyw` (the installer can add it at sign-in). In the app, open **Reminders & settings** to send a test notification.

**An in-app update stops.** Installed copies verify the SHA-256 before replacing files and restore the previous app folder if a step fails. Check your connection, or download the zip from the [latest release](https://github.com/asjames18/work-todo/releases/latest) and run the installer again.

## License and third-party notices

This repository does not include a `LICENSE` file, so a license is not specified here.

The core app uses the Python standard library. Optional voice input uses packages listed in `requirements-voice.txt`, including faster-whisper, CTranslate2, PyAV, NumPy, Hugging Face Hub, tokenizers and ONNX Runtime. Those projects are licensed by their authors; see each project for its terms. Optional AI features may call an OpenAI-compatible endpoint that you configure (OpenRouter by default). Python itself is licensed under the [PSF License](https://docs.python.org/3/license.html).

## Made by

[Antonio James](https://antoniojames.dev) / [Melanated in Tech](https://melanatedintech.com)

- [antoniojames.dev](https://antoniojames.dev)
- [melanatedintech.com](https://melanatedintech.com)
- [github.com/asjames18](https://github.com/asjames18)
