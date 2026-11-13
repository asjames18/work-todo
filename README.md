# Work To-Do

A local project to-do app for Windows, with a Today dashboard, desktop reminders, and optional email and AI helpers.

![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)
![Python](https://img.shields.io/badge/python-3.14-3776AB?logo=python&logoColor=white)
![Dependencies](https://img.shields.io/badge/dependencies-stdlib%20only-lightgrey)
![Status](https://img.shields.io/badge/status-private%20%7C%20active-2ea44f)

## Overview

Work To-Do runs a private HTTP server on `127.0.0.1` (default port `48231`) and opens in its own Microsoft Edge window. Projects, tasks, and notes are stored in `todo-data.json` next to the app. No third-party Python packages are required for the app itself.

The landing view is a **Today** dashboard: a built-in daily brief, overdue and due-today tasks, priorities by project, the schedule for today and the next 7 days, work email that is not yet a task, recent notes, and recent Live Transcriber meeting notes (read-only).

## Key Features

- **Today dashboard** — Open, due-today, overdue, and done-today counts, plus a built-in brief written on this PC. An optional AI brief is available when AI suggestions are turned on with an OpenAI-compatible model.
- **Projects, tasks, and notes** — Color-coded projects, priorities (normal, high, urgent), due dates and times, search, and smart views (Today, Upcoming, Overdue).
- **Desktop reminders** — A separate no-window helper (`reminder.pyw`) can show a morning summary of tasks due today and overdue tasks, a heads-up before a timed task, and an alert at the due time. Reminders work even when the main window is closed.
- **Keyboard shortcuts** — Navigation, task actions, undo, and settings. Press `?` in the app for the full list. A suggested desktop shortcut key is `Ctrl+Alt+T`.
- **Undo** — Deleting or completing a task, dismissing an email, or applying an AI change can be undone with `Ctrl+Z`.
- **Work email (optional)** — A read-only view of recent Inbox mail from classic Outlook, plus drag-and-drop or paste of `.eml`, `.msg`, or copied email text. Nothing is added to your list until you confirm it. Mail is never sent, moved, deleted, flagged, or marked read.
- **AI suggestions (optional, off by default)** — Built-in local rules (no network), or an OpenAI-compatible provider such as LM Studio, Ollama, or OpenRouter. Suggested changes (priority, due date, mark done, add or move a task, add a note) go through a confirm card first. You can also ask AI to add a task from free text or from a voice clip transcribed locally.
- **Live Transcriber (optional)** — Voice input uses the Live Transcriber's local Whisper model when that app is installed. The dashboard can also show recent meeting notes from that app, read-only.

## Screenshots

<!-- ![Screenshot](docs/screenshot.png) -->

## Requirements

- Windows
- Python 3.14 (standard library only; no pip packages for the app itself)
- Microsoft Edge (used to host the app window; the system browser is used as a fallback)
- Classic Outlook (optional), for the read-only work-email inbox
- [Live Transcriber](https://github.com/asjames18/live-transcriber) (optional), for voice input and meeting notes. By default the app looks for a sibling folder named `transcriber`. You can point it at a different location with the `WORKTODO_TX_DIR` environment variable.

## Installation

From this folder:

```bat
C:\Python314\pythonw.exe todo.pyw
```

Or double-click `todo.pyw` if `.pyw` is associated with Python 3.14.

Create a Desktop shortcut:

- **Target:** `C:\Python314\pythonw.exe`
- **Arguments:** `"C:\path\to\work-todo\todo.pyw"`
- **Start in:** the `work-todo` folder
- **Icon:** `todo.ico`
- **Suggested shortcut key:** `Ctrl+Alt+T`

The desktop reminder helper is a separate no-window process:

```bat
C:\Python314\pythonw.exe reminder.pyw
```

To start reminders with Windows, add a Startup-folder shortcut to `reminder.pyw` (same working directory and icon). In the app, open **Reminders & settings** to start the helper, send a test notification, or turn reminders off. To stop it starting with Windows, press `Win+R`, type `shell:startup`, and delete “Work To-Do Reminders”.

`todo-classic.pyw` is an older Tkinter version kept for reference.

## Configuration

Copy `settings.example.json` if you want a reference for the settings shape. Live settings (reminders, morning summary time, heads-up lead time, email, and AI) are stored inside `todo-data.json`, which is created on first run and is not committed.

AI features stay off until you turn them on in **Reminders & settings**. Built-in rules run on this PC with no network. For OpenRouter or another OpenAI-compatible provider, set the **name** of an environment variable in settings (`api_key_env`), not the key itself. For OpenRouter, create a User environment variable named:

```bat
OPENROUTER_API_KEY
```

The key is read from that environment variable at call time. It is never stored in `todo-data.json`, settings, or logs.

Optional environment variables:

| Variable | Purpose |
| --- | --- |
| `TODO_DATA` | Path to the data file (default: `todo-data.json` next to the app) |
| `TODO_PORT` | Main app port (default: `48231`) |
| `TODO_REMINDER_PORT` | Reminder helper port (default: `48232`) |
| `WORKTODO_TX_DIR` | Live Transcriber install folder |
| `WORKTODO_MEET_DIR` | Folder of meeting-note JSON files (default: `<TX_DIR>\transcripts\notes`) |

## Usage

1. Start `todo.pyw`. The Today dashboard opens first.
2. Add projects from the sidebar, then add tasks and notes. Priorities and due dates can be set from the task card or with keyboard shortcuts (`P`, `D`, `T`).
3. Turn on desktop reminders in **Reminders & settings** (`Ctrl+,`) and keep `reminder.pyw` running if you want notifications when the window is closed.
4. Optionally connect classic Outlook on the Email view, or drag/paste an email and confirm it as a task.
5. Optionally enable AI suggestions and review each proposed change before it is applied.

Press `?` at any time for keyboard shortcuts.

## Privacy & Safety

- The server binds to `127.0.0.1` only and rejects requests whose `Host` header is not `127.0.0.1` or `localhost` on the app port.
- All projects, tasks, and notes stay in `todo-data.json` on this PC.
- The reminder helper never changes your tasks. It only records which notifications it has already shown, in `reminder-state.json`.
- The Outlook helper is read-only. It never sends, moves, deletes, flags, or marks mail as read. Opening a message displays it in Outlook and changes nothing. The local email cache (`email-cache.json`) keeps headers and a short preview, not the full body.
- AI is off until you enable it. Built-in rules make no network calls. When an OpenAI-compatible provider is selected, the app sends a snapshot of to-do state (and, for notes or parse requests, the relevant text) to the endpoint you configure. The API key is read from an environment variable and is never stored or logged.
- Voice clips used by “Ask AI to add” are transcribed on this PC by the Live Transcriber's local Whisper model; that audio is not sent online. Windows voice typing (`Win+H`) can be used in any text box without this integration.
- Meeting notes from Live Transcriber are read only. This app does not modify those files.

## Project Structure

| File | Description |
| --- | --- |
| `todo.pyw` | Main app: local server, UI, dashboard, and routing |
| `reminder.pyw` | Background reminder helper (no window) and Windows toast notifications |
| `todo_ai.py` | Optional AI suggestions layer (local rules or OpenAI-compatible providers) |
| `todo_email.py` | Read-only Outlook inbox, `.eml`/`.msg`/paste parsing, and email-to-task helpers |
| `outlook_mail.ps1` | PowerShell helper used to read classic Outlook (and to display a message) |
| `voice_helper.py` | Local Whisper transcription helper, used with Live Transcriber's virtual environment |
| `todo-classic.pyw` | Earlier Tkinter version, kept for reference |
| `settings.example.json` | Example settings shape (reminders and AI). Not used at runtime |
| `todo.ico` | Application icon |
| `toast-icon.png` | Icon used in Windows reminder toasts |
| `.gitignore` | Excludes data files, caches, logs, and secrets |

User data created at runtime (`todo-data.json`, `email-cache.json`, `reminder-state.json`, and similar) is gitignored and is not part of this repository.

## Author

[Antonio James](https://github.com/asjames18) · Melanated in Tech
