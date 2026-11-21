# TODO

## Before the first public push (done in the initial commit)
- [x] Remove personal/company paths from `DEFAULTS` (OneDrive, user name) – default to
      `%USERPROFILE%\Documents\ShareX\Screenshots` and `%USERPROFILE%\Documents\ScreenshotSessions`
- [x] Make sure no company template, logo or sample screenshots are committed
- [x] Add `.gitignore` (`*.ini`, `*.bak`, `ScreenshotSessions/`, `*.pptx`, `*.docx`, `*.pdf`, `*.zip`)
- [x] Ship `settings.example.ini` with every setting documented
- [x] Rename script to match the repo (e.g. `SnapDeck.ahk`)
- [x] Choose a licence (MIT suggested)

## Documentation
- [x] README: what it does, requirements (Windows 10 2004+, AutoHotkey v2, ShareX, Office for PPTX/PDF/Word)
- [x] README: hotkey table, notes panel, finish dialog, settings reference
- [ ] README: GIF demo (start → capture → note → stop → deck)
- [x] README: "Run at login" (shell:startup) instructions
- [x] README: known limitations (new Outlook has no COM; PowerPoint can't save straight into OneDrive – handled)
- [x] Sample output screenshots (title slide, screenshot slide with note heading)

## Robustness
- [ ] Auto-detect ShareX screenshot folder from ShareX's `ApplicationConfig.json`
- [ ] Friendly first-run check: AutoHotkey v2, ShareX folder, PowerPoint/Word installed
- [ ] Fall back gracefully when Office isn't installed (ZIP + Notes.txt only)
- [ ] Log file for errors (`SnapDeck.log`)
- [ ] Recover an unfinished session after a crash/reboot (rebuild from the session folder)
- [ ] Handle multi-monitor badge/panel placement on the monitor with the active window

## Features (backlog)
- [ ] Undo last screenshot (Ctrl+Alt+Backspace)
- [ ] Rebuild outputs from an existing session folder (tray: "Rebuild session…")
- [ ] Reorder / delete screenshots in the finish dialog before building
- [ ] Settings window instead of editing the .ini
- [ ] Configurable hotkeys in settings
- [ ] Work with other capture tools (Snipping Tool, Greenshot) – just a watched folder
- [ ] Optional OCR text of each screenshot into speaker notes
- [ ] Markdown / HTML output
- [ ] Pick meeting name automatically from Teams/Outlook window title
- [ ] Stop automatically when Teams/Zoom meeting window closes

## Release
- [ ] Tag v1.0.0 with a release zip (script + example ini + README)
- [ ] Optional: compiled `.exe` via Ahk2Exe so users don't need AutoHotkey installed
- [ ] Add repo topics: autohotkey, sharex, screenshots, powerpoint, productivity, windows
