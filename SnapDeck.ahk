; SnapDeck.ahk  (AutoHotkey v2)  -  https://github.com/singhramandeep/SnapDeck
; MIT licence
; Collects ShareX screenshots taken during a session and turns them into
; PowerPoint / PDF / Word / ZIP.
;
;   Ctrl+Alt+S   start session
;   Ctrl+Alt+N   show/focus the notes panel; notes go to the latest screenshot's speaker notes
;   Ctrl+Alt+E   stop -> finish dialog (name, formats, optional Outlook email)
;   Ctrl+Alt+P / W / Z   same dialog with PowerPoint / Word / ZIP pre-ticked
;
; Settings live in an .ini file next to this script (created on first run)
; and can also be changed from the tray icon's Settings menu.
#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ---------------- settings (.ini) ----------------
INI := RegExReplace(A_ScriptFullPath, "\.ahk$", ".ini")
DEFAULTS := Map(
    "ShareXFolder",    A_MyDocuments "\ShareX\Screenshots",
    "OutputFolder",    A_MyDocuments "\SnapDeck",
    "Mode",            "Move",      ; Move or Copy (Copy keeps originals in ShareX)
    "AutoStopMinutes", "60",        ; 0 = never auto-stop
    "AutoStopFormats", "pptx",      ; comma list of pptx,pdf,docx,zip used when auto-stopping
    "Template",        "",          ; optional .potx/.pptx used as the PowerPoint design
    "Captions",        "1",         ; 1 = time + app caption on each slide/page
    "ShowBadge",       "1",         ; 1 = on-screen REC badge while recording
    "NotesPanel",      "1",         ; 1 = open the notes panel when recording starts
    "NotesOnSlide",    "1"          ; 1 = show notes as a heading on the slide (0 = speaker notes only)
)
for k, v in DEFAULTS
    if IniRead(INI, "Settings", k, "<none>") = "<none>"
        IniWrite v, INI, "Settings", k
Setting(k) => IniRead(INI, "Settings", k, DEFAULTS[k])

IMG_EXT := "png,jpg,jpeg,bmp,gif,webp"
POLL_MS := 700

global session := "", items := [], seen := Map(), startStamp := "", titleNote := ""
global badge := 0, badgeText := 0, stopGui := 0
global notesGui := 0, notesEdit := 0, notesLabel := 0, notesSkip := 0, skipNotes := false

if !DirExist(Setting("ShareXFolder")) {
    MsgBox "ShareX folder not found:`n" Setting("ShareXFolder") "`n`nFix ShareXFolder in:`n" INI
    ExitApp
}
DirCreate Setting("OutputFolder")

; ---------------- tray ----------------
L_START := "Start session`tCtrl+Alt+S"
L_STOP  := "Stop && finish…`tCtrl+Alt+E"
L_NOTE  := "Show notes panel`tCtrl+Alt+N"
settingsMenu := Menu()
settingsMenu.Add("Auto-stop time…", (*) => AskAutoStop())
settingsMenu.Add("Choose PowerPoint template…", (*) => ChooseTemplate())
settingsMenu.Add("Clear PowerPoint template", (*) => (IniWrite("", INI, "Settings", "Template"), TrayTip("Using plain 16:9 slides", "Template cleared")))
settingsMenu.Add("Edit settings file", (*) => Run('notepad.exe "' INI '"'))
tray := A_TrayMenu
tray.Delete()
tray.Add(L_START, (*) => StartSession())
tray.Add(L_STOP, (*) => StopDialog())
tray.Add(L_NOTE, (*) => AddNote())
tray.Add()
tray.Add("Open output folder", (*) => Run(Setting("OutputFolder")))
tray.Add("Settings", settingsMenu)
tray.Add()
tray.Add("Exit", (*) => ExitApp())
SetIdleTray()
TrayTip "Ctrl+Alt+S start · Ctrl+Alt+E stop", "SnapDeck ready"

^!s::StartSession()
^!e::StopDialog()
^!p::StopDialog("pptx")
^!w::StopDialog("docx")
^!z::StopDialog("zip")
^!n::AddNote()

SetIdleTray() {
    A_TrayMenu.Enable(L_START), A_TrayMenu.Disable(L_STOP), A_TrayMenu.Disable(L_NOTE)
    A_TrayMenu.Default := L_START
    TraySetIcon "shell32.dll", 304
    A_IconTip := "SnapDeck: idle"
}

; ---------------- session ----------------
StartSession() {
    global session, items, seen, startStamp, titleNote
    if session {
        TrayTip "Already recording", "SnapDeck"
        return
    }
    startStamp := A_Now
    session := Setting("OutputFolder") "\Session_" FormatTime(startStamp, "yyyy-MM-dd_HH-mm-ss")
    DirCreate session
    items := [], seen := Map(), titleNote := ""
    global skipNotes := false
    Loop Files Setting("ShareXFolder") "\*.*", "FR"      ; ignore screenshots that already exist
        seen[A_LoopFileFullPath] := true
    SetTimer Tick, POLL_MS
    A_TrayMenu.Disable(L_START), A_TrayMenu.Enable(L_STOP), A_TrayMenu.Enable(L_NOTE)
    A_TrayMenu.Default := L_STOP
    TraySetIcon "shell32.dll", 44
    if Setting("ShowBadge") = "1"
        ShowBadge()
    if Setting("NotesPanel") = "1"
        ShowNotesPanel()
    Tick()
    mins := Setting("AutoStopMinutes")
    TrayTip "Take screenshots with ShareX. Ctrl+Alt+E to stop." (mins > 0 ? "`nAuto-stops after " mins " min." : ""), "Recording started"
}

Tick() {
    if !session
        return
    Poll()
    elapsed := DateDiff(A_Now, startStamp, "Seconds")
    status := "since " FormatTime(startStamp, "HH:mm") " (" Duration(elapsed) ") · " items.Length " shot" (items.Length = 1 ? "" : "s")
    A_IconTip := "Recording " status
    if badgeText
        badgeText.Value := "● REC  " status
    mins := Setting("AutoStopMinutes")
    if mins > 0 && elapsed >= mins * 60 && !stopGui {
        TrayTip "Reached " mins " min – building output.", "Auto-stopped"
        Finish("Session (auto-stopped)", StrSplit(Setting("AutoStopFormats"), ",", " "), false)
    }
}

Duration(s) => Format("{:02}:{:02}:{:02}", s // 3600, Mod(s, 3600) // 60, Mod(s, 60))

Poll() {
    global items, seen
    if !session
        return
    Loop Files Setting("ShareXFolder") "\*.*", "FR" {
        p := A_LoopFileFullPath
        if seen.Has(p) || !InStr("," IMG_EXT ",", "," StrLower(A_LoopFileExt) ",")
            continue
        if !IsFileReady(p)          ; ShareX still writing; retry next tick
            continue
        seen[p] := true
        fname := Format("{:03}_", items.Length + 1) A_LoopFileName
        it := {path: session "\" fname, file: fname, time: FormatTime(A_LoopFileTimeModified, "HH:mm"),
               app: AppFromName(A_LoopFileName), note: ""}
        try {
            if Setting("Mode") = "Copy"
                FileCopy p, it.path, 1
            else
                FileMove p, it.path, 1
            FlushNote()                 ; notes typed so far belong to the previous screenshot
            items.Push(it)
            UpdateNotesLabel()
        } catch {
            TrayTip "Could not capture " A_LoopFileName, "SnapDeck"
        }
    }
}

; ShareX's default name is <process>_<random>.png
AppFromName(name) {
    SplitPath name, , , , &noExt
    return RegExMatch(noExt, "^(.+)_[A-Za-z0-9]{6,}$", &m) ? m[1] : ""
}

IsFileReady(p) {
    try {
        f := FileOpen(p, "r -rwd")   ; exclusive open fails while ShareX holds it
        if !f
            return false
        sz := f.Length, f.Close()
        return sz > 0
    }
    return false
}

; ---------------- live notes panel ----------------
; Whatever you type belongs to the latest screenshot (or the title slide before the
; first one). When the next screenshot arrives the text is saved to the previous
; slide and the box is cleared.
AddNote() {
    if !session {
        TrayTip "Start a session first (Ctrl+Alt+S).", "Not recording"
        return
    }
    global skipNotes := false            ; Ctrl+Alt+N / tray always brings the panel back
    if notesSkip
        notesSkip.Value := 0
    ShowNotesPanel(true)
}

ShowNotesPanel(activate := false) {
    global notesGui, notesEdit, notesLabel, notesSkip
    if !notesGui {
        notesGui := Gui("+AlwaysOnTop +ToolWindow", "Meeting notes")
        notesGui.SetFont("s9", "Segoe UI")
        notesLabel := notesGui.Add("Text", "w300", "")
        notesGui.SetFont("s10")
        notesEdit := notesGui.Add("Edit", "w300 r6 WantReturn")
        notesGui.SetFont("s9")
        notesGui.Add("Button", "w70", "OK").OnEvent("Click", (*) => SaveNoteNow())
        notesGui.Add("Button", "x+8 w70", "Cancel").OnEvent("Click", (*) => CancelNote())
        notesSkip := notesGui.Add("Checkbox", "x+12 yp+4", "Skip notes this session")
        notesSkip.OnEvent("Click", (*) => SkipNotesClicked())
        notesGui.SetFont("s8 c808080")
        notesGui.Add("Text", "xm w300", "OK (Ctrl+Enter) saves to the slide now. Unsaved text is also saved when you take the next screenshot. Cancel discards it. Ctrl+Alt+N reopens this panel.")
        notesGui.OnEvent("Close", (*) => notesGui.Hide())
        DllCall("SetWindowDisplayAffinity", "ptr", notesGui.Hwnd, "uint", 0x11)   ; keep out of screenshots
        UpdateNotesLabel()
        gap := 12
        if badge {                                    ; sit just above the REC badge
            WinGetPos , , , &bh, badge
            gap += bh + 8
        }
        PlaceBottomRight(notesGui, gap)
    } else if !activate
        notesGui.Show("NoActivate")
    if activate {
        notesGui.Show()
        notesEdit.Focus()
    }
}

NoteTarget() => items.Length ? "screenshot " items.Length " (slide " items.Length + 1 ")" : "the title slide"

UpdateNotesLabel() {
    if notesLabel
        notesLabel.Value := "Notes for " NoteTarget()
}

SaveNoteNow() {
    if !notesEdit || Trim(notesEdit.Value, " `t`r`n") = ""
        return
    target := NoteTarget()
    FlushNote()
    notesLabel.Value := "✓ Saved to " target
    SetTimer UpdateNotesLabel, -2500
    notesEdit.Focus()
}

CancelNote() {
    if notesEdit {
        notesEdit.Value := ""
        notesEdit.Focus()
    }
}

SkipNotesClicked() {
    global skipNotes
    if notesSkip.Value {
        FlushNote()                      ; keep anything already typed
        skipNotes := true
        notesGui.Hide()
        TrayTip "Notes panel hidden for this session. Ctrl+Alt+N brings it back.", "Notes skipped"
    }
}

#HotIf notesGui && WinActive("ahk_id " notesGui.Hwnd)
^Enter::SaveNoteNow()
#HotIf

FlushNote() {
    global titleNote
    if !notesEdit
        return
    text := Trim(notesEdit.Value, " `t`r`n")
    if text != "" {
        if items.Length
            items[items.Length].note .= (items[items.Length].note ? "`n" : "") text
        else
            titleNote .= (titleNote ? "`n" : "") text
    }
    notesEdit.Value := ""
}

HideNotesPanel() {
    global notesGui, notesEdit, notesLabel, notesSkip
    if notesGui
        notesGui.Destroy()
    notesGui := 0, notesEdit := 0, notesLabel := 0, notesSkip := 0
}

; ---------------- badge ----------------
ShowBadge() {
    global badge, badgeText
    badge := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")   ; click-through
    badge.BackColor := "B00020"
    badge.MarginX := 10, badge.MarginY := 5
    badge.SetFont("s9 cWhite bold", "Segoe UI")
    badgeText := badge.Add("Text", "w280", "● REC")
    PlaceBottomRight(badge, 12)
    WinSetTransparent 220, badge
    ; keep the badge out of screenshots (Windows 10 2004+)
    DllCall("SetWindowDisplayAffinity", "ptr", badge.Hwnd, "uint", 0x11)
}

; Puts a window in the bottom-right corner of the primary monitor's work area,
; using its real on-screen size (so display scaling can't push it onto another monitor).
PlaceBottomRight(g, gapBottom) {
    g.Show("Hide AutoSize")
    WinGetPos , , &w, &h, g
    MonitorGetWorkArea(MonitorGetPrimary(), , , &right, &bottom)
    g.Show("NoActivate x" (right - w - 12) " y" (bottom - h - gapBottom))
}

HideBadge() {
    global badge, badgeText
    if badge
        badge.Destroy()
    badge := 0, badgeText := 0
}

; ---------------- finish ----------------
StopDialog(preset := "pptx") {
    global stopGui
    if !session {
        TrayTip "Press Ctrl+Alt+S to start.", "Not recording"
        return
    }
    if stopGui {
        stopGui.Show()
        return
    }
    Poll()
    g := Gui("+AlwaysOnTop", "Finish screenshot session")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , items.Length " screenshots since " FormatTime(startStamp, "HH:mm")
        " (" Duration(DateDiff(A_Now, startStamp, "Seconds")) ")")
    g.Add("Text", "y+12", "Meeting name (used for the title slide and file names):")
    nameE := g.Add("Edit", "w380")
    g.Add("Text", "y+12", "Create:")
    fmts := ["pptx", "pdf", "docx", "zip"], labels := ["PowerPoint", "PDF", "Word", "ZIP"]
    cbs := Map()
    for i, f in fmts
        cbs[f] := g.Add("Checkbox", (i = 1 ? "xm" : "x+18") (f = preset ? " Checked" : ""), labels[i])
    emailCb := g.Add("Checkbox", "xm y+12", "Open an Outlook email with the files attached")
    g.Add("Button", "xm y+16 w120 Default", "Create").OnEvent("Click", Go)
    g.Add("Button", "x+10 w140", "Keep recording").OnEvent("Click", Close)
    g.OnEvent("Close", Close)
    g.OnEvent("Escape", Close)
    stopGui := g
    g.Show()
    nameE.Focus()

    Go(*) {
        chosen := []
        for f in fmts
            if cbs[f].Value
                chosen.Push(f)
        if !chosen.Length {
            MsgBox "Tick at least one format.", "Finish", "Owner" g.Hwnd
            return
        }
        name := nameE.Value, email := emailCb.Value
        Close()
        Finish(name, chosen, email)
    }
    Close(*) {
        global stopGui
        g.Destroy()
        stopGui := 0
    }
}

Has(arr, v) {
    for x in arr
        if x = v
            return true
    return false
}

Finish(name, fmts, email := false) {
    global session
    Poll()
    FlushNote()
    SetTimer Tick, 0
    HideBadge()
    HideNotesPanel()
    folder := session, imgs := items, started := startStamp, tnote := titleNote
    session := ""
    SetIdleTray()
    if !imgs.Length {
        TrayTip "No screenshots captured.", "Session stopped"
        try DirDelete folder
        return
    }
    title := Trim(name) != "" ? Trim(name) : "Screenshots"

    ; rename folder to the meeting name (date is on the title slide); add (2), (3)… if taken
    safe := Trim(RegExReplace(title, '[\\/:*?"<>|]', "-"))
    newFolder := Setting("OutputFolder") "\" safe, n := 1
    while DirExist(newFolder) || FileExist(newFolder ".zip")
        newFolder := Setting("OutputFolder") "\" safe " (" (++n) ")"
    try {
        DirMove folder, newFolder
        folder := newFolder
    }
    for it in imgs
        it.path := folder "\" it.file
    SplitPath folder, &base

    ; plain-text notes file (also ends up in the ZIP)
    notesTxt := ""
    if tnote
        notesTxt .= "Title slide:`r`n" tnote "`r`n`r`n"
    for i, it in imgs
        if it.note
            notesTxt .= "Screenshot " i " (" it.time (it.app ? " · " it.app : "") ") – " it.file ":`r`n" it.note "`r`n`r`n"
    if notesTxt
        try FileAppend title "`r`n" FormatTime(started, "dddd, d MMMM yyyy · HH:mm") "`r`n`r`n" notesTxt, folder "\Notes.txt", "UTF-8"

    TrayTip imgs.Length " screenshots…", "Building output"
    outputs := [], errors := ""
    if Has(fmts, "pptx") || Has(fmts, "pdf") {
        try {
            for p in BuildPptx(folder, base, imgs, title, started, tnote, Has(fmts, "pptx"), Has(fmts, "pdf"))
                outputs.Push(p)
        } catch as e
            errors .= "PowerPoint/PDF: " e.Message "`n"
    }
    if Has(fmts, "docx") {
        try outputs.Push(BuildDocx(folder, base, imgs, title, started, tnote))
        catch as e
            errors .= "Word: " e.Message "`n"
    }
    if Has(fmts, "zip") {
        try outputs.Push(BuildZip(folder))
        catch as e
            errors .= "ZIP: " e.Message "`n"
    }
    if errors
        MsgBox "Some outputs failed:`n" errors "`nScreenshots are in:`n" folder, "SnapDeck"
    if email && outputs.Length {
        try CreateEmail(title, started, outputs)
        catch as e
            MsgBox "Could not open an Outlook email (the new Outlook app has no automation support):`n" e.Message
    }
    if outputs.Length {
        TrayTip outputs.Length " file(s) in " base, "Done"
        Run 'explorer.exe /select,"' outputs[1] '"'
    } else
        Run folder
}

Caption(it, i, n) => it.time (it.app ? " · " it.app : "") "   (" i "/" n ")"

FindLayout(pres, name, fallback) {
    lays := pres.SlideMaster.CustomLayouts
    Loop lays.Count
        if InStr(lays.Item(A_Index).Name, name)
            return lays.Item(A_Index)
    return fallback ? lays.Item(fallback) : ""
}

; Adds the template's logo (taken from its title layout) and an orange footer line
; to the in-memory copy of the Blank layout, so every screenshot slide inherits them.
; The template file itself is never modified.
BrandLayout(lay, titleLay, W, H, footerH) {
    logo := ""
    Loop titleLay.Shapes.Count {
        sh := titleLay.Shapes.Item(A_Index)
        if (sh.Type = 28 || sh.Type = 13) && sh.Width < 300 && sh.Height < 100 {
            logo := sh
            break
        }
    }
    line := lay.Shapes.AddShape(1, 0, H - footerH, W, 1.5)
    line.Fill.ForeColor.RGB := 0x005CFF          ; accent orange #FF5C00 (BGR)
    line.Line.Visible := 0
    if !logo
        return
    saved := ClipboardAll()                       ; copy/paste needs the clipboard; restore it after
    try {
        logo.Copy()
        Sleep 150
        pasted := lay.Shapes.Paste().Item(1)
        pasted.LockAspectRatio := -1
        pasted.Height := 15
        pasted.Left := 12
        pasted.Top := H - footerH + (footerH - 15) / 2 + 1
    }
    A_Clipboard := saved
}

SetNotes(slide, text) {
    try slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text := text
}

BuildPptx(folder, base, imgs, title, started, tnote, wantPptx, wantPdf) {
    ppt := ComObject("PowerPoint.Application")
    hadOpen := ppt.Presentations.Count
    tmpFiles := []
    try {
        tpl := Setting("Template")
        if tpl && FileExist(tpl) {
            pres := ppt.Presentations.Open(tpl, -1, -1, 0)   ; read-only, untitled copy, no window
            while pres.Slides.Count
                pres.Slides.Item(1).Delete()
        } else {
            pres := ppt.Presentations.Add(0)
            pres.PageSetup.SlideSize := 15                  ; 16:9
        }
        W := pres.PageSetup.SlideWidth, H := pres.PageSetup.SlideHeight
        titleLay := FindLayout(pres, "Title Slide", 1)
        blankLay := FindLayout(pres, "Blank", pres.SlideMaster.CustomLayouts.Count)
        branded := tpl && FileExist(tpl)
        capH := branded ? 30 : (Setting("Captions") = "1" ? 22 : 0)
        if branded
            BrandLayout(blankLay, titleLay, W, H, capH)

        ; title slide
        s := pres.Slides.AddSlide(1, titleLay)
        try s.Shapes.Title.TextFrame.TextRange.Text := title
        onSlide := Setting("NotesOnSlide") = "1"
        try s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text :=
            (onSlide && tnote ? tnote "`n" : "") FormatTime(started, "dddd, d MMMM yyyy · HH:mm") "`n" imgs.Length " screenshots"
        if tnote
            SetNotes(s, tnote)

        for i, it in imgs {
            s := pres.Slides.AddSlide(i + 1, blankLay)
            while s.Shapes.Count                              ; drop layout placeholders
                s.Shapes.Item(1).Delete()
            ; note shown as a heading above the screenshot
            topH := 0
            if onSlide && it.note {
                long := StrLen(it.note) > 80 || InStr(it.note, "`n")
                topH := long ? 62 : 46
                hb := s.Shapes.AddTextbox(1, 24, 8, W - 48, topH - 10)
                hb.TextFrame.WordWrap := -1
                hb.TextFrame.AutoSize := 0
                hb.TextFrame.VerticalAnchor := 3              ; middle
                tr := hb.TextFrame.TextRange
                tr.Text := it.note
                tr.Font.Size := long ? 14 : 22
                tr.Font.Bold := !long
                tr.Font.Color.RGB := 0x000000
            }
            pic := s.Shapes.AddPicture(it.path, 0, -1, 0, 0)
            pic.LockAspectRatio := -1
            scale := Min(W / pic.Width, (H - capH - topH) / pic.Height)
            if scale < 1
                pic.Width := pic.Width * scale
            pic.Left := (W - pic.Width) / 2
            pic.Top  := topH + (H - capH - topH - pic.Height) / 2
            if Setting("Captions") = "1" {
                tr := s.Shapes.AddTextbox(1, 10, H - capH + (branded ? 4 : 0), W - 20, 22).TextFrame.TextRange
                tr.Text := Caption(it, i, imgs.Length)
                tr.Font.Size := 10
                tr.Font.Color.RGB := 0x808080
                tr.ParagraphFormat.Alignment := 3             ; right
            }
            if it.note
                SetNotes(s, it.note)
        }
        ; template closing slide (e.g. a "Closing" layout)
        if branded && (closeLay := FindLayout(pres, "Closing", 0)) {
            s := pres.Slides.AddSlide(pres.Slides.Count + 1, closeLay)
            shapes := []
            Loop s.Shapes.Count
                shapes.Push(s.Shapes.Item(A_Index))
            for sh in shapes {
                try {
                    if sh.Name = s.Shapes.Title.Name
                        sh.TextFrame.TextRange.Text := "Thank you"
                    else if sh.Type = 14 && !InStr(sh.Name, "Slide Number")
                        sh.Delete()
                }
            }
        }
        ; PowerPoint can't save straight into OneDrive folders: save to temp, then move
        if wantPptx {
            pres.SaveAs(A_Temp "\" base ".pptx", 24)
            tmpFiles.Push(".pptx")
        }
        if wantPdf {
            ; with notes: export "notes pages" (slide on top, notes underneath); else plain slides
            hasNotes := !onSlide && tnote != ""
            if !onSlide
                for it in imgs
                    hasNotes := hasNotes || it.note != ""
            ; (PowerPoint needs an explicit print range here, or it throws "Type mismatch")
            pres.ExportAsFixedFormat(A_Temp "\" base ".pdf", 2, 2, 0, 1, hasNotes ? 5 : 1, 0,
                pres.PrintOptions.Ranges.Add(1, pres.Slides.Count), 1)
            tmpFiles.Push(".pdf")
        }
        pres.Close()
    } finally {
        if !hadOpen && ppt.Presentations.Count = 0          ; don't close your open decks
            ppt.Quit()
    }
    out := []
    for ext in tmpFiles {
        FileMove A_Temp "\" base ext, folder "\" base ext, 1
        out.Push(folder "\" base ext)
    }
    return out
}

AddPara(doc, text, size, bold := false, italic := false, color := 0) {
    r := doc.Content
    r.Collapse(0)
    r.InsertAfter(text)
    r.Font.Size := size, r.Font.Bold := bold, r.Font.Italic := italic, r.Font.Color := color
    r.InsertParagraphAfter()
}

BuildDocx(folder, base, imgs, title, started, tnote) {
    path := folder "\" base ".docx"
    word := ComObject("Word.Application")
    try {
        doc := word.Documents.Add()
        doc.Content.Delete()                                  ; start empty even if Normal.dotm has content
        ps := doc.PageSetup
        usableW := ps.PageWidth - ps.LeftMargin - ps.RightMargin
        AddPara(doc, title, 24, true)
        AddPara(doc, FormatTime(started, "dddd, d MMMM yyyy · HH:mm") " · " imgs.Length " screenshots", 12, false, false, 0x808080)
        if tnote
            AddPara(doc, tnote, 11)
        for i, it in imgs {
            r := doc.Content, r.Collapse(0)
            r.InsertBreak(7)                                  ; page break
            if it.note                                        ; note as a heading above the image
                AddPara(doc, it.note, 16, true)
            r := doc.Content, r.Collapse(0)
            pic := doc.InlineShapes.AddPicture(it.path, 0, -1, r)
            if pic.Width > usableW {
                pic.LockAspectRatio := -1
                pic.Width := usableW
            }
            doc.Content.InsertParagraphAfter()
            if Setting("Captions") = "1"
                AddPara(doc, Caption(it, i, imgs.Length), 9, false, true, 0x808080)
        }
        doc.SaveAs2(path, 16)                                 ; 16 = docx
        doc.Close()
    } finally {
        if word.Documents.Count = 0
            word.Quit()
    }
    return path
}

BuildZip(folder) {
    SplitPath folder, &name, &parent
    path := parent "\" name ".zip"
    RunWait 'powershell -NoProfile -Command "Compress-Archive -Path ' "'" folder "\*'" ' -DestinationPath ' "'" path "'" ' -Force"', , "Hide"
    if !FileExist(path)
        throw Error("ZIP not created")
    return path
}

; Opens a draft (never sends) with the outputs attached
CreateEmail(title, started, files) {
    ol := ComObject("Outlook.Application")
    m := ol.CreateItem(0)
    m.Subject := title " – screenshots (" FormatTime(started, "d MMM yyyy") ")"
    for f in files
        m.Attachments.Add(f)
    m.Display()
}

; ---------------- settings UI ----------------
AskAutoStop() {
    r := InputBox("Stop recording automatically after how many minutes?`n(0 = never)", "Auto-stop", "w320 h140", Setting("AutoStopMinutes"))
    if r.Result = "OK" && RegExMatch(r.Value, "^\d+$") {
        IniWrite r.Value, INI, "Settings", "AutoStopMinutes"
        TrayTip r.Value = 0 ? "Auto-stop off" : "Auto-stop after " r.Value " min", "Saved"
    }
}

ChooseTemplate() {
    f := FileSelect(1, , "Choose PowerPoint template", "PowerPoint (*.potx; *.pptx)")
    if f {
        IniWrite f, INI, "Settings", "Template"
        TrayTip f, "Template saved"
    }
}
