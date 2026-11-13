# Work To-Do for Windows

A fast, private to-do app for your projects. It runs on your PC (nothing is uploaded) and opens in its own window.

## Install (1-2 minutes, plus 3-10 minutes if you add voice input)

1. Open the downloaded `WorkToDo-Setup-<version>.zip` and double-click **`Install-WorkToDo.cmd`**. (Extracting first with right-click, **Extract All...** also works. If you run it straight from the zip, it finds the download in Downloads, Desktop or Documents and unpacks it for you.)
   - Windows may say *"Windows protected your PC"* (SmartScreen) or *"unknown publisher"*, because the installer isn't code-signed. Click **More info**, then **Run anyway**.
2. The installer sets up everything the app needs. It asks before it installs anything:
   - **Python** (if the PC doesn't have Python 3.10 or newer): it asks *"Python is needed, install it now? [Y/n]"*, then downloads the official **Python 3.13.17** from python.org (about 29 MB), checks its SHA-256 checksum and its Python Software Foundation signature, and installs it for your Windows account only (no admin needed).
   - **Voice input (optional, recommended)**: speak a task and it's typed for you, transcribed on this PC by Whisper (nothing is sent online). It asks *"Set up voice input now?"* and then downloads a private speech engine into `%LOCALAPPDATA%\WorkToDo\voice`: about 85 MB of Python packages plus the 465 MB `small.en` Whisper model, roughly 1 GB on disk. If Live Transcriber is already installed, Work To-Do uses its Whisper instead and downloads nothing.
3. Work To-Do opens into a short **setup wizard**: your name, light or dark theme, (optional) Outlook, (optional) AI, and reminders.

What the installer does:
- Copies the app to `%LOCALAPPDATA%\WorkToDo\app` (no admin rights needed). Your data lives separately in `%LOCALAPPDATA%\WorkToDo\data`.
- Adds **Desktop** and **Start menu** shortcuts. Press **Ctrl+Alt+T** anywhere to open it (skipped if another shortcut already uses that key).
- Optionally starts the small **reminder helper** when you sign in. It shows reminders and powers **Ctrl+Alt+N** quick add.
- Running the installer again is safe: it upgrades or repairs the app, keeps your data, and skips anything that's already installed.
- If a download fails or you're offline, it says so and stops safely (nothing half-installed). Run it again when you're back online.

Requirements: Windows 10 or 11 (64-bit), an internet connection for the first install, about 60 MB of disk (1.5 GB free with voice input). Microsoft Edge is built into Windows and hosts the app window. Classic Outlook is optional.

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

Double-click **`Uninstall-WorkToDo.cmd`**, or use **Start menu, then Uninstall Work To-Do**. It removes the app, its voice-input files and its shortcuts. Your tasks are kept unless you choose to delete them. Python stays installed (other apps may use it); remove it from Windows Settings > Apps if you want.
