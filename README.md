# Work To-Do

A local project to-do app for Windows. It runs a private server on `127.0.0.1` and opens in its own Microsoft Edge window. The landing view is a **Today** dashboard: overdue and due-today tasks, priorities by project, the next 7 days, work email that is not yet a task, recent notes, and recent Live Transcriber meeting notes (read-only).

Projects, tasks and notes are stored in `todo-data.json` next to the app (created on first run; not committed). Reminder state and the email cache are also local files and are gitignored.

## Requirements

- Windows
- Python 3.14 (standard library only; no pip packages for the app itself)
- Classic Outlook (optional), if you want the read-only work-email inbox
- [Live Transcriber](../live-transcriber) (optional), for voice input and meeting notes

## Run on Windows

From this folder:

```bat
C:\Python314\pythonw.exe todo.pyw
```

Or double-click `todo.pyw` if `.pyw` is associated with Python 3.14.

Create a Desktop shortcut:

- Target: `C:\Python314\pythonw.exe`
- Arguments: `"C:\path\to\work-todo\todo.pyw"`
- Start in: the `work-todo` folder
- Icon: `todo.ico`
- Suggested shortcut key: `Ctrl+Alt+T`

The desktop reminder helper (morning summary and due-time toasts) is a separate no-window process:

```bat
C:\Python314\pythonw.exe reminder.pyw
```

To start reminders with Windows, add a Startup-folder shortcut to `reminder.pyw` (same working directory and icon).

`todo-classic.pyw` is an older Tkinter version kept for reference.

## AI (optional)

AI features are off until you turn them on in Reminders & settings. Built-in rules run on this PC with no network. For OpenRouter or another OpenAI-compatible provider, set the **name** of an environment variable in settings (not the key itself). For OpenRouter, create a User environment variable:

```bat
OPENROUTER_API_KEY
```

The key is read at call time and is never stored in `todo-data.json` or logs. Copy `settings.example.json` if you want a reference for the settings shape; the live settings live inside `todo-data.json`.
