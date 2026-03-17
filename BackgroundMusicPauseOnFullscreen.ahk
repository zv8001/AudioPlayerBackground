#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

global MusicPlayer := ComObject("WMPlayer.OCX")
global ConfigFilePath := A_ScriptDir "\BackgroundMusicConfig.ini"
global MusicFilePath := ""
global WasPausedByFullscreen := false
global LastFullscreenState := false

MusicPlayer.settings.setMode("loop", true)
MusicPlayer.settings.volume := 50

InitializeTray()
LoadSettings()
SetTimer(CheckFullscreenState, 500)

if MusicFilePath != "" and FileExist(MusicFilePath)
{
    MusicPlayer.URL := MusicFilePath
    MusicPlayer.controls.play()
}
else
{
    TrayTip("Background Music", "No saved audio file found. Right click the tray icon and select Choose Audio File.", 3)
}

^!o::SelectMusicFile()
^!p::TogglePlayPause()
^!s::StopMusic()
^!Up::VolumeUp()
^!Down::VolumeDown()
^!q::ExitScript()

InitializeTray()
{
    TraySetIcon("shell32.dll", 138)
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Choose Audio File", SelectMusicFile)
    A_TrayMenu.Add("Play or Resume", PlayMusic)
    A_TrayMenu.Add("Pause", PauseMusic)
    A_TrayMenu.Add("Stop", StopMusic)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Volume Up", VolumeUp)
    A_TrayMenu.Add("Volume Down", VolumeDown)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", ExitScript)
}

LoadSettings()
{
    global ConfigFilePath
    global MusicFilePath
    global MusicPlayer

    if !FileExist(ConfigFilePath)
        return

    try
    {
        SavedMusicFilePath := IniRead(ConfigFilePath, "Settings", "MusicFilePath", "")
        SavedVolume := IniRead(ConfigFilePath, "Settings", "Volume", "50")

        MusicFilePath := SavedMusicFilePath

        SavedVolume := Integer(SavedVolume)
        SavedVolume := ClampValue(SavedVolume, 0, 100)
        MusicPlayer.settings.volume := SavedVolume
    }
}

SaveSettings()
{
    global ConfigFilePath
    global MusicFilePath
    global MusicPlayer

    IniWrite(MusicFilePath, ConfigFilePath, "Settings", "MusicFilePath")
    IniWrite(MusicPlayer.settings.volume, ConfigFilePath, "Settings", "Volume")
}

SelectMusicFile(*)
{
    global MusicFilePath
    global WasPausedByFullscreen
    global MusicPlayer

    SelectedFile := FileSelect(1, , "Select an Audio File", "Audio Files (*.mp3; *.wav; *.flac; *.m4a; *.aac; *.wma)")

    if SelectedFile = ""
        return

    if !FileExist(SelectedFile)
    {
        MsgBox("That file does not exist.")
        return
    }

    MusicFilePath := SelectedFile
    WasPausedByFullscreen := false

    MusicPlayer.URL := MusicFilePath
    SaveSettings()
    MusicPlayer.controls.play()

    TrayTip("Background Music", "Loaded:`n" MusicFilePath, 2)
}

PlayMusic(*)
{
    global MusicPlayer
    global MusicFilePath

    if MusicFilePath = ""
    {
        MsgBox("No audio file has been selected yet.")
        return
    }

    if !FileExist(MusicFilePath)
    {
        MsgBox("The saved audio file could not be found. Please choose a new file.")
        return
    }

    if MusicPlayer.URL != MusicFilePath
        MusicPlayer.URL := MusicFilePath

    MusicPlayer.controls.play()
}

PauseMusic(*)
{
    global MusicPlayer
    MusicPlayer.controls.pause()
}

StopMusic(*)
{
    global MusicPlayer
    global WasPausedByFullscreen

    WasPausedByFullscreen := false
    MusicPlayer.controls.stop()
}

TogglePlayPause()
{
    global MusicPlayer
    global MusicFilePath

    if MusicFilePath = ""
    {
        MsgBox("No audio file has been selected yet.")
        return
    }

    CurrentState := MusicPlayer.playState

    if CurrentState = 3
        MusicPlayer.controls.pause()
    else
        PlayMusic()
}

VolumeUp(*)
{
    global MusicPlayer

    NewVolume := ClampValue(MusicPlayer.settings.volume + 5, 0, 100)
    MusicPlayer.settings.volume := NewVolume
    SaveSettings()
    TrayTip("Background Music", "Volume: " NewVolume "%", 1)
}

VolumeDown(*)
{
    global MusicPlayer

    NewVolume := ClampValue(MusicPlayer.settings.volume - 5, 0, 100)
    MusicPlayer.settings.volume := NewVolume
    SaveSettings()
    TrayTip("Background Music", "Volume: " NewVolume "%", 1)
}

ClampValue(Value, Minimum, Maximum)
{
    if Value < Minimum
        return Minimum
    if Value > Maximum
        return Maximum
    return Value
}

CheckFullscreenState()
{
    global MusicPlayer
    global WasPausedByFullscreen
    global LastFullscreenState
    global MusicFilePath

    if MusicFilePath = ""
        return

    if !FileExist(MusicFilePath)
        return

    IsFullscreenNow := IsActiveWindowFullscreen()

    if IsFullscreenNow and !LastFullscreenState
    {
        if MusicPlayer.playState = 3
        {
            MusicPlayer.controls.pause()
            WasPausedByFullscreen := true
        }
    }
    else if !IsFullscreenNow and LastFullscreenState
    {
        if WasPausedByFullscreen
        {
            PlayMusic()
            WasPausedByFullscreen := false
        }
    }

    LastFullscreenState := IsFullscreenNow
}

IsActiveWindowFullscreen()
{
    ActiveWindowId := WinExist("A")
    if !ActiveWindowId
        return false

    ActiveProcessName := ""
    try ActiveProcessName := WinGetProcessName("ahk_id " ActiveWindowId)

    if ActiveProcessName = "explorer.exe"
        return false

    try
    {
        WindowMinMax := WinGetMinMax("ahk_id " ActiveWindowId)
        if WindowMinMax = -1
            return false
    }
    catch
    {
        return false
    }

    try
    {
        WindowStyle := WinGetStyle("ahk_id " ActiveWindowId)
        if !(WindowStyle & 0x10000000)
            return false
    }
    catch
    {
        return false
    }

    WindowX := 0
    WindowY := 0
    WindowW := 0
    WindowH := 0

    try WinGetPos(&WindowX, &WindowY, &WindowW, &WindowH, "ahk_id " ActiveWindowId)

    if WindowW <= 0 or WindowH <= 0
        return false

    MonitorIndex := GetMonitorFromWindowCenter(WindowX, WindowY, WindowW, WindowH)

    MonitorLeft := 0
    MonitorTop := 0
    MonitorRight := 0
    MonitorBottom := 0
    MonitorGet(MonitorIndex, &MonitorLeft, &MonitorTop, &MonitorRight, &MonitorBottom)

    MonitorWidth := MonitorRight - MonitorLeft
    MonitorHeight := MonitorBottom - MonitorTop
    Tolerance := 2

    CoversMonitor :=
    (
        Abs(WindowX - MonitorLeft) <= Tolerance
        and Abs(WindowY - MonitorTop) <= Tolerance
        and Abs(WindowW - MonitorWidth) <= Tolerance
        and Abs(WindowH - MonitorHeight) <= Tolerance
    )

    return CoversMonitor
}

GetMonitorFromWindowCenter(WindowX, WindowY, WindowW, WindowH)
{
    CenterX := WindowX + (WindowW // 2)
    CenterY := WindowY + (WindowH // 2)
    MonitorCount := MonitorGetCount()

    Loop MonitorCount
    {
        MonitorLeft := 0
        MonitorTop := 0
        MonitorRight := 0
        MonitorBottom := 0
        MonitorGet(A_Index, &MonitorLeft, &MonitorTop, &MonitorRight, &MonitorBottom)

        if CenterX >= MonitorLeft and CenterX < MonitorRight and CenterY >= MonitorTop and CenterY < MonitorBottom
            return A_Index
    }

    return 1
}

ExitScript(*)
{
    global MusicPlayer
    SaveSettings()
    MusicPlayer.controls.stop()
    ExitApp()
}