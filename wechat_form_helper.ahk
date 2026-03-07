#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent
SetWorkingDir A_ScriptDir

global CONFIG := {
    dataFile: A_ScriptDir "\wechat_form_data.tsv",
    progressFile: A_ScriptDir "\wechat_form_progress.tsv",
    syncScript: A_ScriptDir "\sync_progress_to_excel.py",
    tabDelayMs: 90,
    pasteDelayMs: 70,
    skipTabsAfterDeadline: 0
}

global records := []
global currentIndex := 1

Init()

F5::FillTextFieldsOnce()
F8::PasteDescriptionOnly()
F12::FillPhotoInFileDialog()
F10::MarkCurrentDoneAndNext()
F9::GoNextPendingRow()
F6::GoPrevPendingRow()
F3::IncreaseSkipTabs()
F2::DecreaseSkipTabs()
F4::ShowCurrentRowValues()
F1::SaveAndExit()
^Esc::ExitApp

Init() {
    global records, CONFIG

    if !FileExist(CONFIG.dataFile) {
        MsgBox "未找到数据文件:`n" CONFIG.dataFile "`n`n请先运行 run_wechat_helper.cmd。"
        ExitApp
    }

    records := LoadRecords()
    if records.Length = 0 {
        MsgBox "数据文件为空，请检查 wechat_form_data.tsv。"
        ExitApp
    }

    JumpToFirstPending()
    ShowHelp()
    OnExit(SaveProgress)
}

ShowHelp() {
    helpText := "微信小程序填报助手热键：`n"
    helpText .= "F5  一键填入4个文本项（地址/路段/截止/描述）`n"
    helpText .= "F8  仅粘贴问题描述（手动点到描述框后用）`n"
    helpText .= "F12 在文件选择框自动填入照片路径`n"
    helpText .= "F10 标记当前记录已填并跳下一条`n"
    helpText .= "F9  跳到下一条待填记录`n"
    helpText .= "F6  回到上一条待填记录`n"
    helpText .= "F3  +1 描述前Tab跳过数（当前: " CONFIG.skipTabsAfterDeadline "）`n"
    helpText .= "F2  -1 描述前Tab跳过数`n"
    helpText .= "F4  查看当前记录全部字段`n"
    helpText .= "F1  保存进度并退出`n"
    helpText .= "Ctrl+Esc 退出助手"
    MsgBox helpText
}

LoadRecords() {
    global CONFIG

    progressMap := LoadProgressMap()
    arr := []

    text := FileRead(CONFIG.dataFile, "UTF-8")
    lines := StrSplit(text, "`n", "`r")
    if lines.Length <= 1 {
        return arr
    }

    Loop lines.Length - 1 {
        raw := Trim(lines[A_Index + 1], "`r`n")
        if raw = "" {
            continue
        }

        cols := StrSplit(raw, "`t")
        while cols.Length < 8 {
            cols.Push("")
        }

        srcRow := SafeInt(cols[1], A_Index + 1)
        rec := {
            sourceRow: srcRow,
            address: cols[2],
            section: cols[3],
            deadlineHours: cols[4],
            category: cols[5],
            description: cols[6],
            photoPath: cols[7],
            disposal: cols[8],
            done: false,
            submitTime: ""
        }

        if progressMap.Has(srcRow) {
            p := progressMap[srcRow]
            rec.done := (p.status = "已填")
            rec.submitTime := p.submitTime
        }

        arr.Push(rec)
    }

    return arr
}

LoadProgressMap() {
    global CONFIG

    progressMap := Map()
    if !FileExist(CONFIG.progressFile) {
        return progressMap
    }

    text := FileRead(CONFIG.progressFile, "UTF-8")
    lines := StrSplit(text, "`n", "`r")
    if lines.Length <= 1 {
        return progressMap
    }

    Loop lines.Length - 1 {
        raw := Trim(lines[A_Index + 1], "`r`n")
        if raw = "" {
            continue
        }

        cols := StrSplit(raw, "`t")
        while cols.Length < 3 {
            cols.Push("")
        }

        rowNum := SafeInt(cols[1], 0)
        if rowNum <= 0 {
            continue
        }

        progressMap[rowNum] := {status: cols[2], submitTime: cols[3]}
    }

    return progressMap
}

SaveProgress(*) {
    global records, CONFIG

    text := "source_row`t状态`t提交时间`n"
    for rec in records {
        status := rec.done ? "已填" : ""
        text .= rec.sourceRow "`t" status "`t" rec.submitTime "`n"
    }

    try FileDelete(CONFIG.progressFile)
    FileAppend(text, CONFIG.progressFile, "UTF-8")
}

JumpToFirstPending() {
    global currentIndex, records

    currentIndex := 1
    for idx, rec in records {
        if !rec.done {
            currentIndex := idx
            Toast("当前源行: " rec.sourceRow "（待填）")
            return
        }
    }

    Toast("全部记录已填，按 F6 可回看")
}

MoveToPending(direction := 1) {
    global currentIndex, records

    idx := currentIndex + direction
    while idx >= 1 && idx <= records.Length {
        if !records[idx].done {
            currentIndex := idx
            Toast("当前源行: " records[idx].sourceRow "（待填）")
            return true
        }
        idx += direction
    }

    return false
}

GoNextPendingRow() {
    if !MoveToPending(1) {
        Toast("后面没有待填记录")
    }
}

GoPrevPendingRow() {
    if !MoveToPending(-1) {
        Toast("前面没有待填记录")
    }
}

FillTextFieldsOnce() {
    global records, currentIndex, CONFIG

    rec := records[currentIndex]

    ; 先点到“问题地址”输入框，再按 F5。
    if !PasteValue(rec.address) {
        return
    }
    Send "{Tab}"
    Sleep CONFIG.tabDelayMs

    if !PasteValue(rec.section) {
        return
    }
    Send "{Tab}"
    Sleep CONFIG.tabDelayMs

    if !PasteValue(rec.deadlineHours) {
        return
    }

    ; 跳过“问题类别”到“问题描述”。
    Send "{Tab}"
    Sleep CONFIG.tabDelayMs
    if CONFIG.skipTabsAfterDeadline > 0 {
        Send "{Tab " CONFIG.skipTabsAfterDeadline "}"
        Sleep CONFIG.tabDelayMs
    }

    if !PasteValue(rec.description) {
        return
    }

    Toast("源行 " rec.sourceRow " 四项文本已填（跳过Tab=" CONFIG.skipTabsAfterDeadline "）")
}

PasteDescriptionOnly() {
    global records, currentIndex

    rec := records[currentIndex]
    if !PasteValue(rec.description) {
        return
    }
    Toast("源行 " rec.sourceRow " 问题描述已粘贴")
}

IncreaseSkipTabs() {
    global CONFIG
    CONFIG.skipTabsAfterDeadline += 1
    Toast("描述前Tab跳过数: " CONFIG.skipTabsAfterDeadline)
}

DecreaseSkipTabs() {
    global CONFIG
    if CONFIG.skipTabsAfterDeadline > 0 {
        CONFIG.skipTabsAfterDeadline -= 1
    }
    Toast("描述前Tab跳过数: " CONFIG.skipTabsAfterDeadline)
}

PasteValue(value) {
    global CONFIG

    A_Clipboard := value
    if !ClipWait(0.5) {
        Toast("复制到剪贴板失败")
        return false
    }

    Send "^v"
    Sleep CONFIG.pasteDelayMs
    return true
}

FillPhotoInFileDialog() {
    global records, currentIndex

    rawPath := records[currentIndex].photoPath
    fullPath := ResolvePhotoPath(rawPath)
    if fullPath = "" {
        Toast("照片不存在或路径为空: " rawPath)
        return
    }

    if !WaitForAnyFileDialog(2500) {
        MsgBox "未检测到文件选择框。`n请先点击上传照片 +，确保文件框处于前台后再按 F12。"
        return
    }

    if TryFillFileDialogByControl(fullPath) {
        Toast("已按数据路径提交文件")
        return
    }

    A_Clipboard := fullPath
    if !ClipWait(0.5) {
        Toast("复制照片路径失败")
        return
    }

    Send "!n"
    Sleep 80
    Send "^a^v{Enter}"
    Sleep 220

    if IsAnyFileDialogActive() {
        MsgBox "已尝试自动输入路径，但文件框仍未关闭。`n请检查图片路径是否有效: `n" fullPath
        return
    }

    Toast("已按数据路径提交文件")
}

TryFillFileDialogByControl(fullPath) {
    try {
        ControlFocus "Edit1", "A"
        Sleep 60
        ControlSetText fullPath, "Edit1", "A"
        Sleep 80
        ControlSend "{Enter}", "Edit1", "A"
        Sleep 220
        return !IsAnyFileDialogActive()
    } catch {
        return false
    }
}

IsAnyFileDialogActive() {
    return WinActive("ahk_class #32770")
        || WinActive("ahk_class CabinetWClass")
        || WinActive("ahk_class ExploreWClass")
        || WinActive("打开")
        || WinActive("Open")
}

WaitForAnyFileDialog(timeoutMs := 2000) {
    deadline := A_TickCount + timeoutMs
    while A_TickCount < deadline {
        if IsAnyFileDialogActive() {
            return true
        }
        Sleep 80
    }
    return false
}

MarkCurrentDoneAndNext() {
    global records, currentIndex

    records[currentIndex].done := true
    records[currentIndex].submitTime := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    SaveProgress()
    SyncProgressToExcel()
    Toast("源行 " records[currentIndex].sourceRow " 已标记为已填")

    if !MoveToPending(1) {
        Toast("全部记录已处理完成，脚本即将退出")
        SetTimer(() => ExitApp(), -900)
    }
}

SyncProgressToExcel() {
    global CONFIG

    if !FileExist(CONFIG.syncScript) {
        return
    }

    cmd := "python `"" CONFIG.syncScript "`""
    try {
        RunWait cmd, A_ScriptDir, "Hide"
    } catch {
        ; 忽略同步异常，避免打断填报。
    }
}

OpenPhotoPath() {
    global records, currentIndex

    rawPath := records[currentIndex].photoPath
    fullPath := ResolvePhotoPath(rawPath)
    if fullPath = "" {
        Toast("照片不存在或路径为空: " rawPath)
        return
    }

    Run fullPath
}

ShowCurrentRowValues() {
    global records, currentIndex

    rec := records[currentIndex]
    content := "当前记录索引: " currentIndex "`n"
    content .= "Excel源行: " rec.sourceRow "`n`n"
    content .= "问题地址: " rec.address "`n"
    content .= "问题路段: " rec.section "`n"
    content .= "截止时间(小时数): " rec.deadlineHours "`n"
    content .= "问题类别: " rec.category "`n"
    content .= "问题描述: " rec.description "`n"
    content .= "上传照片路径: " rec.photoPath "`n"
    content .= "处置方式: " rec.disposal "`n"
    content .= "状态: " (rec.done ? "已填" : "待填")
    MsgBox content
}

ResolvePhotoPath(rawPath) {
    path := Trim(rawPath)
    if path = "" {
        return ""
    }

    fullPath := path
    if !FileExist(fullPath) {
        fullPath := A_ScriptDir "\" path
    }

    if !FileExist(fullPath) {
        return ""
    }
    return fullPath
}

SaveAndExit() {
    SaveProgress()
    ExitApp
}

SafeInt(value, defaultValue := 0) {
    try {
        return Integer(value)
    } catch {
        return defaultValue
    }
}

Toast(message) {
    ToolTip message
    SetTimer(() => ToolTip(), -1200)
}
