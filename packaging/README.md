# Work To-Do for Windows

A fast, private to-do app for your projects. It runs on your PC (nothing is uploaded) and opens in its own window.

## Install (about 1 minute)

1. Open the downloaded `WorkToDo-Setup-<version>.zip` and double-click **`Install-WorkToDo.cmd`**. (Extracting first with right-click, **Extract All...** also works. If you run it straight from the zip, it finds the download in Downloads, Desktop or Documents and unpacks it for you.)
   - Windows may say *"Windows protected your PC"* or *"unknown publisher"*, because the installer isn't code-signed. Click **More info**, then **Run anyway**.
2. The installer checks for **Python 3.10 or newer**. If it's missing, it opens the python.org download page. Install Python (tick **"Add python.exe to PATH"**), then run `Install-WorkToDo.cmd` again. It never installs anything without asking.
3. Work To-Do opens into a short **setup wizard**: your name, light or dark theme, (optional) Outlook, (optional) AI, and reminders.

What the installer does:
- Copies the app to `%LOCALAPPDATA%\WorkToDo\app` (no admin rights needed). Your data lives separately in `%LOCALAPPDATA%\WorkToDo\data`.
- Adds **Desktop** and **Start menu** shortcuts. Press **Ctrl+Alt+T** anywhere to open it.
- Optionally starts the small **reminder helper** when you sign in. It shows reminders and powers **Ctrl+Alt+N** quick add.

Requirements: Windows 10 or 11, Python 3.10+ (python.org), Microsoft Edge (built into Windows). Classic Outlook is optional.

## Highlights

| Feature | Shortcut |
| --- | --- |
| Dashboard: today at a glance, plus a brief | Ctrl+D, or G then D |
| New task / new note | N / Shift+N |
| Quick add from anywhere in Windows (even with the app closed) | Ctrl+Alt+N |
| Recurring tasks (daily, weekdays, weekly on a day, monthly on day N) | Y on a selected task, or the Repeat box when adding |
| Focus timer (25/5 Pomodoro, adjustable) with time logged per task and project | F on a selected task, Shift+F pause, Ctrl+Shift+F stop |
| Weekly review: done, slipped, next week, time per project, AI recap | G then W ([ and ] change week, A = AI recap) |
| Back up everything to one file / restore | Ctrl+Shift+E / Ctrl+Shift+O |
| Light / dark theme, accent colours | Ctrl+Shift+L / Ctrl+Shift+K |
| All shortcuts, and a replayable tour | ? (then T for the tour) |
| Settings | Ctrl+, |

## Optional extras

- **Outlook inbox:** shows emails sent to you from **classic Outlook**, read-only, so you can make them tasks in one click. Nothing is ever sent, moved, deleted or marked read.
- **AI (OpenRouter):** a daily brief, a weekly recap and smart suggestions. Create a key at openrouter.ai. Then either set the Windows user environment variable `OPENROUTER_API_KEY` yourself, or paste the key in the setup wizard, which saves it **only** in that variable for your Windows account, never in the app's files. Free models (ids ending in `:free`) cost nothing. Paid models bill your OpenRouter account.

## Your data and privacy

- Your tasks, notes, settings and focus logs are in `%LOCALAPPDATA%\WorkToDo\data\todo-data.json`, on your PC only.
- **Back up** (Ctrl+Shift+E) saves one JSON file to Downloads, plus a copy in `data\backups`. **Restore** (Ctrl+Shift+O) always saves a copy of what's there first.

## Updates

When update checks are on (Settings, then Updates), Work To-Do checks GitHub Releases once a day. If a newer version is out, the dashboard shows an **Update available** banner with the release notes. **Update now** downloads it, checks its SHA-256 checksum, backs up the current version and installs it. Your data is never touched, and if anything fails, the old version is put back. You can also click **Check for updates** in Settings at any time.

## Uninstall

Double-click **`Uninstall-WorkToDo.cmd`**, or use **Start menu, then Uninstall Work To-Do**. Your tasks are kept unless you choose to delete them.
