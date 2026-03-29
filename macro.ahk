#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, Off

; DPI awareness — окно не будет размываться на экранах с масштабированием
DllCall("Shcore.dll\SetProcessDpiAwareness", "Int", 2)

if !A_IsAdmin {
    Run('*RunAs "' . A_ScriptFullPath . '"')
    ExitApp
}

; ==========================================================
; RUCOY MACRO PRO — All-in-one (WebView2 direct messaging)
; ==========================================================

Global WV         := ""   ; WebView2 CoreWebView2 object
Global WVC        := ""   ; WebView2 Controller
Global Profiles   := []
Global ActiveTab  := 1
Global CurrentCaptureType := ""

; --- ПУТИ ---
rootDir     := A_ScriptDir . "\.."
profilesDir := rootDir . "\profiles"
htmlFile    := A_ScriptDir . "\rucoy_launcher.html"
icoFile     := A_ScriptDir . "\rucoy_icon.ico"
libDir      := rootDir . "\lib"
wv2File     := libDir . "\WebView2.ahk"

; --- ПАПКИ ---
try {
    DirCreate(profilesDir)
}

; --- ТРЕЙ ---
A_TrayMenu.Delete()
A_TrayMenu.Add("Rucoy Macro Pro", (*) => 0)
A_TrayMenu.Disable("Rucoy Macro Pro")
A_TrayMenu.Add()
A_TrayMenu.Add("Показать",     (*) => (IsSet(MainGui) ? MainGui.Show() : 0))
A_TrayMenu.Add("Скрыть",      (*) => (IsSet(MainGui) ? MainGui.Hide() : 0))
A_TrayMenu.Add("Перезапустить", (*) => Reload())
A_TrayMenu.Add()
A_TrayMenu.Add("Выход",       (*) => ExitApp())
if FileExist(icoFile) {
    TraySetIcon(icoFile)
} else {
    TraySetIcon("shell32.dll", 18)
}

; --- ЗАГРУЗКА WEBVIEW2 ---
if !FileExist(wv2File) {
    MsgBox("Файл lib\WebView2.ahk не найден!`nЗапусти launcher_app.ahk для установки.", "Ошибка", "Icon!")
    ExitApp
}
#Include "..\lib\WebView2.ahk"

; --- ПРОФИЛИ ---
Profiles := [LoadProfile(1), LoadProfile(2)]

; --- РАЗМЕР ОКНА ---
baseW := 510
baseH := 670
winW  := baseW
winH  := baseH

; --- ОКНО ---
Global MainGui := Gui("+Resize -MaximizeBox +MinSize400x500 -DPIScale", "Rucoy Macro Pro")
MainGui.BackColor := "0a0a12"
MainGui.MarginX := 0
MainGui.MarginY := 0
MainGui.OnEvent("Close", (*) => ExitApp())
MainGui.Show("w" . winW . " h" . winH . " Center")

WVC := WebView2.CreateControllerAsync(MainGui.Hwnd).await()
WV  := WVC.CoreWebView2
WV.Settings.AreDefaultContextMenusEnabled := false
WV.Settings.IsStatusBarEnabled := false
WV.Settings.IsZoomControlEnabled := false
WV.add_WebMessageReceived(OnMessage)

; Размер
SetBounds(WVC, winW, winH)
MainGui.OnEvent("Size", (g, mm, w, h) => (mm != -1 ? SetBounds(WVC, w, h) : 0))

WV.Navigate("file:///" . StrReplace(htmlFile, "\", "/"))

; --- ТАЙМЕРЫ ---
SetTimer(CheckBoth,  100)
SetTimer(CheckLoot,  100)
SetTimer(WriteState, 500)

; ==========================================================
; HELPER
; ==========================================================
SetBounds(ctrl, w, h) {
    r := Buffer(16, 0)
    NumPut("Int", 0, r, 0), NumPut("Int", 0, r, 4)
    NumPut("Int", w, r, 8), NumPut("Int", h, r, 12)
    ctrl.Bounds := r
}

SafeInt(val, def := 0) {
    try {
        return Integer(val)
    } catch {
        return def
    }
}

EscJSON(s) {
    s := StrReplace(s, "\",  "\\")
    s := StrReplace(s, '"',  '\"')
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "\r")
    return s
}

Send2JS(msg) {
    Global WV
    if (WV != "") {
        try {
            WV.PostWebMessageAsString(msg)
        }
    }
}

HotkeyConvert(hk) {
    if (hk = "") {
        return ""
    }
    if RegExMatch(hk, "^[!^+#]") {
        return hk
    }
    result := ""
    parts  := StrSplit(hk, "+")
    key    := Trim(parts[parts.Length])
    Loop parts.Length - 1 {
        mod := Trim(parts[A_Index])
        if (mod = "Alt") {
            result .= "!"
        }
        if (mod = "Ctrl") {
            result .= "^"
        }
        if (mod = "Shift") {
            result .= "+"
        }
        if (mod = "Win") {
            result .= "#"
        }
    }
    result .= key
    return result
}

; ==========================================================
; ПРОФИЛИ
; ==========================================================
GetIni(id) {
    Global profilesDir
    return profilesDir . "\settings_p" . id . ".ini"
}

LoadProfile(id) {
    ini := GetIni(id)
    p   := Map()
    p["id"]  := id
    p["ini"] := ini

    p["ActionKeyMana"]    := IniRead(ini, "Keys", "Mana",    "")
    p["ActionKeyHP"]      := IniRead(ini, "Keys", "HP",      "")
    p["ActionKeyWarrior"] := IniRead(ini, "Keys", "Warrior", "")
    p["ActionKeyLoot"]    := IniRead(ini, "Keys", "Loot",    "")

    defStart := ""
    defTag   := ""
    p["HotkeyStart"] := IniRead(ini, "Hotkeys", "Start", defStart)
    p["HotkeyTag"]   := IniRead(ini, "Hotkeys", "Tag",   defTag)
    p["SwHotkey"]    := IniRead(ini, "Switch",  "Hotkey", "")
    p["ComboHotkey"] := IniRead(ini, "Combo",   "Hotkey", "")

    p["SwClass1Key"]   := IniRead(ini, "Switch", "Class1Key",   "")
    p["SwClass1Delay"] := SafeInt(IniRead(ini, "Switch", "Class1Delay", "30"),   30)
    p["SwUltKey"]      := IniRead(ini, "Switch", "UltKey",      "")
    p["SwUltDelay"]    := SafeInt(IniRead(ini, "Switch", "UltDelay",    "30"),   30)
    p["SwClass2Key"]   := IniRead(ini, "Switch", "Class2Key",   "")
    p["SwClass2Delay"] := SafeInt(IniRead(ini, "Switch", "Class2Delay", "1000"), 1000)

    p["ComboUltKey"]   := IniRead(ini, "Combo", "UltKey",  "")
    p["ComboUltDelay"] := SafeInt(IniRead(ini, "Combo", "UltDelay", "1540"), 1540)

    p["ManaX"]     := SafeInt(IniRead(ini, "Mana", "X",     0))
    p["ManaY"]     := SafeInt(IniRead(ini, "Mana", "Y",     0))
    p["ManaColor"] := IniRead(ini, "Mana", "Color", "")
  
    p["HpX"]       := SafeInt(IniRead(ini, "HP", "X",   0))
    p["HpY"]       := SafeInt(IniRead(ini, "HP", "Y",   0))
    p["HpColor"]   := IniRead(ini, "HP", "Color",  "")
    p["HpX2"]      := SafeInt(IniRead(ini, "HP", "X2",  0))
    p["HpY2"]      := SafeInt(IniRead(ini, "HP", "Y2",  0))
    p["HpColor2"]  := IniRead(ini, "HP", "Color2", "")
    p["Hp2Enabled"] := IniRead(ini, "HP", "Hp2Enabled", "1") = "1"
    p["HealDelay"]     := SafeInt(IniRead(ini, "Heal", "Delay",     "600"), 600)
    p["HealTolerance"] := SafeInt(IniRead(ini, "Heal", "Tolerance", "30"),  30)
    p["LootTolerance"] := SafeInt(IniRead(ini, "Loot", "Tolerance", "20"),  20)

    p["LootX"]           := SafeInt(IniRead(ini, "Loot", "X",           0))
    p["LootY"]           := SafeInt(IniRead(ini, "Loot", "Y",           0))
    p["LootColor"]       := IniRead(ini, "Loot", "Color",   "")
    p["LootEnabled"]     := IniRead(ini, "Loot", "Enabled", "1") = "1"
    
    p["LootPresses"]     := SafeInt(IniRead(ini, "Loot", "Presses",    "3"),  3)
    p["LootPressDelay"]  := SafeInt(IniRead(ini, "Loot", "PressDelay", "80"), 80)
    p["LootWaitDelay"]   := SafeInt(IniRead(ini, "Loot", "WaitDelay",  "300"),300)

    p["TargetHwnd"] := SafeInt(IniRead(ini, "Window", "Hwnd", 0))
    if (p["TargetHwnd"] != 0 && !WinExist("ahk_id " . p["TargetHwnd"])) {
        p["TargetHwnd"] := 0
    }

    ; Тег
    p["TagClass1"]      := IniRead(ini, "Tag", "Class1",    "")
    p["TagClass1Delay"] := SafeInt(IniRead(ini, "Tag", "Class1Delay", "100"), 100)
    p["TagUlt1"]        := IniRead(ini, "Tag", "Ult1",      "")
    p["TagUlt1Delay"]   := SafeInt(IniRead(ini, "Tag", "Ult1Delay",   "100"), 100)
    p["TagClass2"]      := IniRead(ini, "Tag", "Class2",    "")
    p["TagClass2Delay"] := SafeInt(IniRead(ini, "Tag", "Class2Delay", "1000"),1000)
    p["TagUlt2"]        := IniRead(ini, "Tag", "Ult2",      "")
    p["TagUlt2Delay"]   := SafeInt(IniRead(ini, "Tag", "Ult2Delay",   "100"), 100)
    p["Mob1"]           := IniRead(ini, "Tag", "Mob1", "")
    p["Mob2"]           := IniRead(ini, "Tag", "Mob2", "")
    p["Mob3"]           := IniRead(ini, "Tag", "Mob3", "")
    p["Mob4"]           := IniRead(ini, "Tag", "Mob4", "")
 
    p["MobDelay"]       := SafeInt(IniRead(ini, "Tag", "MobDelay",    "300"), 300)
    p["MageMode"]       := SafeInt(IniRead(ini, "Tag", "MageMode",    "1"),   1)
    p["MageAimKey"]     := IniRead(ini, "Tag", "MageAim",    "")
    p["MageAimDelay"]   := SafeInt(IniRead(ini, "Tag", "MageAimDelay","100"), 100)

    p["IsRunning"]       := false
    p["WarriorMode"]     := false
    p["IsSwitchRunning"] := false
    p["IsComboRunning"]  := false
    p["SwitchStep"]      := 1
    p["SwitchTimerFn"]   := RunSwitchStep.Bind(p)
    p["ComboTimerFn"]    := RunComboStep.Bind(p)
    p["LastManaPress"]   := 0
    p["LastHpPress"]     := 0
    p["LastWarriorPress"]:= 0
    return p
}

; ==========================================================
; ОБРАБОТКА СООБЩЕНИЙ ИЗ HTML
; ==========================================================
OnMessage(sender, args) {
    Global ActiveTab, Profiles, WV

    msg   := args.TryGetWebMessageAsString()
    parts := StrSplit(msg, "|")
    cmd   := parts[1]

    if (cmd = "READY") {
        InitHotkeys()
        SendProfileToJS(ActiveTab)
        SendStateToJS()
        SendWindowsToJS()
    }
    else if (cmd = "SWITCH_TAB") {
        ActiveTab := Integer(parts[2])
        SendProfileToJS(ActiveTab)
        SendStateToJS()
    }
    else if (cmd = "GET_WINDOWS") {
        SendWindowsToJS()
    }
    else if (cmd = "SET_WINDOW") {
        hwnd := Integer(parts[2])
        Profiles[ActiveTab]["TargetHwnd"] := hwnd
        IniWrite(hwnd, Profiles[ActiveTab]["ini"], "Window", "Hwnd")
        SendProfileToJS(ActiveTab)
    }
    else if (cmd = "TOGGLE_START") {
        p := Profiles[ActiveTab]
        if p["IsRunning"] {
            StopProfile(p)
        } else {
            StartProfile(p)
        }
        SendStateToJS()
    }
    else if (cmd = "STOP_ACTIVE") {
        StopProfile(Profiles[ActiveTab])
        SendStateToJS()
    }
    else if (cmd = "TOGGLE_HP2") {
        p := Profiles[ActiveTab]
        p["Hp2Enabled"] := !p["Hp2Enabled"]
        IniWrite(p["Hp2Enabled"] ? "1" : "0", p["ini"], "HP", "Hp2Enabled")
        SendProfileToJS(ActiveTab)
    }
    else if (cmd = "TOGGLE_LOOT") {
        p := Profiles[ActiveTab]
        p["LootEnabled"] := !p["LootEnabled"]
        IniWrite(p["LootEnabled"] ? "1" : "0", p["ini"], "Loot", "Enabled")
        SendProfileToJS(ActiveTab)
    }
    else if (cmd = "SAVE") {
        SaveProfile(parts)
        SendProfileToJS(ActiveTab)
        Send2JS("TOAST|💾 Сохранено!")
    }
    else if (cmd = "CAPTURE") {
        StartCapture(parts[2])
    }
    else if (cmd = "CAPTURE_CANCEL") {
        StopCapture()
        Send2JS("CAPTURE_DONE|cancel")
    }
    else if (cmd = "RESET_PIXEL") {
        SetTimer(() => ResetPixel(parts[2]), -1)
    }
    else if (cmd = "RELOAD_PROFILE") {
        ReloadProfile(ActiveTab)
        SendProfileToJS(ActiveTab)
    }
}

; ==========================================================
; СОХРАНЕНИЕ ПРОФИЛЯ
; ==========================================================
SaveProfile(parts) {
    Global ActiveTab, Profiles
    p   := Profiles[ActiveTab]
    ini := p["ini"]

    data := parts[2]

    WriteField(data, ini, "Keys",    "Mana",        "keyMana",        p, "ActionKeyMana")
    WriteField(data, ini, "Keys",    "HP",           "keyHp",          p, "ActionKeyHP")
    WriteField(data, ini, "Keys",    "Warrior",      "keyWarrior",     p, "ActionKeyWarrior")
    WriteField(data, ini, "Keys",    "Loot",         "keyLoot",        p, "ActionKeyLoot")
    WriteField(data, ini, "Hotkeys", "Start",        "hotkeyStart",    p, "HotkeyStart")
    WriteField(data, ini, "Hotkeys", "Tag",          "hotkeyTag",      p, "HotkeyTag")
    WriteField(data, ini, "Switch",  "Hotkey",       "swHotkey",       p, "SwHotkey")
    WriteField(data, ini, "Switch",  "Class1Key",    "swClass1",       p, "SwClass1Key")
    WriteFieldInt(data, ini, "Switch", "Class1Delay","swClass1Delay",  p, "SwClass1Delay")
    WriteField(data, ini, "Switch",  "UltKey",       "swUlt",          p, "SwUltKey")
    WriteFieldInt(data, ini, "Switch","UltDelay",    "swUltDelay",     p, "SwUltDelay")
    WriteField(data, ini, "Switch",  "Class2Key",    "swClass2",       p, "SwClass2Key")
    WriteFieldInt(data, ini, "Switch","Class2Delay", "swClass2Delay",  p, "SwClass2Delay")
    WriteField(data, ini, "Combo",   "Hotkey",       "comboHotkey",    p, "ComboHotkey")
    WriteField(data, ini, "Combo",   "UltKey",       "comboUlt",       p, "ComboUltKey")
    WriteFieldInt(data, ini, "Combo","UltDelay",     "comboUltDelay",  p, "ComboUltDelay")
    WriteFieldInt(data, ini, "Heal", "Delay",        "healDelay",      p, "HealDelay")
    WriteFieldInt(data, ini, "Heal", "Tolerance",    "healTolerance",  p, "HealTolerance")
    WriteFieldInt(data, ini, "Loot", "Tolerance",    "lootTolerance",  p, "LootTolerance")
    WriteFieldInt(data, ini, "Loot", "Presses",      "lootPresses",    p, "LootPresses")
    WriteFieldInt(data, ini, "Loot", "PressDelay",   "lootPressDelay", p, "LootPressDelay")
    WriteFieldInt(data, ini, "Loot", "WaitDelay",    "lootWaitDelay",  p, "LootWaitDelay")
    WriteField(data, ini, "Tag",    "Class1",        "tagClass1",      p, "TagClass1")
    WriteFieldInt(data, ini, "Tag", "Class1Delay",   "tagClass1Delay", p, "TagClass1Delay")
    WriteField(data, ini, "Tag",    "Ult1",          "tagUlt1",        p, "TagUlt1")
    WriteFieldInt(data, ini, "Tag", "Ult1Delay",     "tagUlt1Delay",   p, "TagUlt1Delay")
    WriteField(data, ini, "Tag",    "Class2",        "tagClass2",      p, "TagClass2")
    WriteFieldInt(data, ini, "Tag", "Class2Delay",   "tagClass2Delay", p, "TagClass2Delay")
    WriteField(data, ini, "Tag",    "Ult2",          "tagUlt2",        p, "TagUlt2")
    WriteFieldInt(data, ini, "Tag", "Ult2Delay",     "tagUlt2Delay",   p, "TagUlt2Delay")
    WriteField(data, ini, "Tag",    "Mob1",          "mob1",           p, "Mob1")
    WriteField(data, ini, "Tag",    "Mob2",          "mob2",           p, "Mob2")
    WriteField(data, ini, "Tag",    "Mob3",          "mob3",           p, "Mob3")
    WriteField(data, ini, "Tag",    "Mob4",          "mob4",           p, "Mob4")
    WriteFieldInt(data, ini, "Tag", "MobDelay",      "mobDelay",       p, "MobDelay")
    WriteFieldInt(data, ini, "Tag", "MageMode",      "mageMode",       p, "MageMode")
    WriteField(data, ini, "Tag",    "MageAim",       "mageAim",        p, "MageAimKey")
    WriteFieldInt(data, ini, "Tag", "MageAimDelay",  "mageAimDelay",   p, "MageAimDelay")

    ReloadProfileHotkeys(p)
}

WriteField(data, ini, sec, key, jsonKey, p, pKey) {
    val := JSONGet(data, jsonKey)
    if (val != "__MISSING__") {
        IniWrite(val, ini, sec, key)
        p[pKey] := val
    }
}

WriteFieldInt(data, ini, sec, key, jsonKey, p, pKey) {
    val := JSONGet(data, jsonKey)
    if (val != "__MISSING__") {
        IniWrite(val, ini, sec, key)
        p[pKey] := SafeInt(val, p[pKey])
    }
}

JSONGet(json, key) {
    pattern := '"' . key . '"\s*:\s*(?:"([^"]*)"|([\d.]+)|(true|false))'
    if RegExMatch(json, pattern, &m) {
        if (m.Pos(1) > 0) {
            return m[1]
        }
        if (m.Pos(2) > 0) {
            return m[2]
        }
        if (m.Pos(3) > 0) {
            return m[3]
        }
    }
    return "__MISSING__"
}

; ==========================================================
; ОТПРАВКА ДАННЫХ В JS
; ==========================================================
SendProfileToJS(id) {
    Global Profiles
    p := Profiles[id]
    winTitle := ""
    if (p["TargetHwnd"] && WinExist("ahk_id " . p["TargetHwnd"])) {
        winTitle := SubStr(WinGetTitle("ahk_id " . p["TargetHwnd"]), 1, 50)
    }
    j := "{"
    j .= '"id":' . id . ","
    j .= '"keyMana":"' . EscJSON(p["ActionKeyMana"]) . '",'
    j .= '"keyHp":"' . EscJSON(p["ActionKeyHP"]) . '",'
    j .= '"keyWarrior":"' . EscJSON(p["ActionKeyWarrior"]) . '",'
    j .= '"keyLoot":"' . EscJSON(p["ActionKeyLoot"]) . '",'
    j .= '"hotkeyStart":"' . EscJSON(p["HotkeyStart"]) . '",'
    j .= '"hotkeyTag":"' . EscJSON(p["HotkeyTag"]) . '",'
    j .= '"swHotkey":"' . EscJSON(p["SwHotkey"]) . '",'
    j .= '"swClass1":"' . EscJSON(p["SwClass1Key"]) . '",'
    j .= '"swClass1Delay":' . p["SwClass1Delay"] . ","
    j .= '"swUlt":"' . EscJSON(p["SwUltKey"]) . '",'
    j .= '"swUltDelay":' . p["SwUltDelay"] . ","
    j .= '"swClass2":"' . EscJSON(p["SwClass2Key"]) . '",'
    j .= '"swClass2Delay":' . p["SwClass2Delay"] . ","
    j .= '"comboHotkey":"' . EscJSON(p["ComboHotkey"]) . '",'
    j .= '"comboUlt":"' . EscJSON(p["ComboUltKey"]) . '",'
    j .= '"comboUltDelay":' . p["ComboUltDelay"] . ","
    j .= '"healDelay":' . p["HealDelay"] . ","
    j .= '"lootPresses":' . p["LootPresses"] . ","
    j .= '"lootPressDelay":' . p["LootPressDelay"] . ","
    j .= '"lootWaitDelay":' . p["LootWaitDelay"] . ","
    j .= '"hp2Enabled":' . (p["Hp2Enabled"] ? "true" : "false") . ","
    j .= '"lootEnabled":' . (p["LootEnabled"] ? "true" : "false") . ","
    j .= '"healTolerance":' . p["HealTolerance"] . ","
    j .= '"lootTolerance":' . p["LootTolerance"] . ","
    j .= '"manaOk":' . (p["ManaColor"] != "" ? "true" : "false") . ","
    j .= '"manaColor":"' . EscJSON(p["ManaColor"]) . '",'
    j .= '"hp1Ok":' . (p["HpColor"] != "" ? "true" : "false") . ","
    j .= '"hp1Color":"' . EscJSON(p["HpColor"]) . '",'
    j .= '"hp2Ok":' . (p["HpColor2"] != "" ? "true" : "false") . ","
    j .= '"hp2Color":"' . EscJSON(p["HpColor2"]) . '",'
    j .= '"lootOk":' . (p["LootColor"] != "" ? "true" : "false") . ","
    j .= '"lootColor":"' . EscJSON(p["LootColor"]) . '",'
    j .= '"hwnd":"' . p["TargetHwnd"] . '",'
    j .= '"winTitle":"' . EscJSON(winTitle) . '",'
    j .= '"tagClass1":"' . EscJSON(p["TagClass1"]) . '",'
    j .= '"tagClass1Delay":' . p["TagClass1Delay"] . ","
    j .= '"tagUlt1":"' . EscJSON(p["TagUlt1"]) . '",'
    j .= '"tagUlt1Delay":' . p["TagUlt1Delay"] . ","
    j .= '"tagClass2":"' . EscJSON(p["TagClass2"]) . '",'
    j .= '"tagClass2Delay":' . p["TagClass2Delay"] . ","
    j .= '"tagUlt2":"' . EscJSON(p["TagUlt2"]) . '",'
    j .= '"tagUlt2Delay":' . p["TagUlt2Delay"] . ","
    j .= '"mob1":"' . EscJSON(p["Mob1"]) . '",'
    j .= '"mob2":"' . EscJSON(p["Mob2"]) . '",'
    j .= '"mob3":"' . EscJSON(p["Mob3"]) . '",'
    j .= '"mob4":"' . EscJSON(p["Mob4"]) . '",'
    j .= '"mobDelay":' . p["MobDelay"] . ","
    j .= '"mageMode":' . p["MageMode"] . ","
    j .= '"mageAim":"' . EscJSON(p["MageAimKey"]) . '",'
    j .= '"mageAimDelay":' . p["MageAimDelay"] . ","
    j .= '"isSw":' . (p["IsSwitchRunning"] ? "true" : "false") . ","
    j .= '"isCb":' . (p["IsComboRunning"] ? "true" : "false")
    j .= "}"
    Send2JS("PROFILE|" . j)
}

SendStateToJS() {
    Global Profiles, ActiveTab
    p1 := Profiles[1]
    p2 := Profiles[2]
    pa := Profiles[ActiveTab]

    j := "{"
    j .= '"p1_run":'     . (p1["IsRunning"] ? "true" : "false") . ","
    j .= '"p2_run":'     . (p2["IsRunning"] ? "true" : "false") . ","
    j .= '"active_run":' . (pa["IsRunning"] ? "true" : "false") . ","
    j .= '"active_sw":'  . (pa["IsSwitchRunning"] ? "true" : "false") . ","
    j .= '"active_cb":'  . (pa["IsComboRunning"]  ? "true" : "false") . ","
    j .= '"active_warrior":' . (pa["WarriorMode"] ? "true" : "false")
    j .= "}"

    Send2JS("STATE|" . j)
}

SendWindowsToJS() {
    wins := WinGetList()
    j := "["
    count := 0
    for hwnd in wins {
        try {
            title := WinGetTitle("ahk_id " . hwnd)
            class := WinGetClass("ahk_id " . hwnd)
            if (title = "" || title = "Program Manager" || title = "Rucoy Macro Pro") {
                continue
            }
            WinGetPos(&x, &y, &w, &h, "ahk_id " . hwnd)
            if (w < 100 || h < 100) {
                continue
            }
            pid := WinGetPID("ahk_id " . hwnd)
            emu := ""
            if (InStr(title, "Nox") || InStr(class, "Qt5QWindow")) {
                emu := "nox"
            } else if InStr(title, "LDPlayer") {
                emu := "ld"
            } else if InStr(title, "BlueStacks") {
                emu := "bs"
            } else if (InStr(title, "MEmu") || InStr(title, "MEmу") || InStr(class, "Qt5152QWindowIcon")) {
                emu := "memu"
            }
            if (count > 0) {
                j .= ","
            }
            j .= '{"hwnd":' . hwnd . ',"title":"' . EscJSON(SubStr(title,1,60)) . '","emu":"' . emu . '","pid":' . pid . '}'
            count++
        } catch {
            continue
        }
    }
    j .= "]"
    Send2JS("WINDOWS|" . j)
}

WriteState() {
    SendStateToJS()
}

; ==========================================================
; ХОТКЕИ
; ==========================================================
InitHotkeys() {
    Global Profiles
    Loop 2 {
        p := Profiles[A_Index]
        BindHotkey("", p["HotkeyStart"], OnStartHotkey.Bind(p["id"]))
        BindHotkey("", p["HotkeyTag"],   OnTagHotkey.Bind(p["id"]))
        BindHotkey("", p["SwHotkey"],    OnSwHotkey.Bind(p["id"]))
        BindHotkey("", p["ComboHotkey"], OnComboHotkey.Bind(p["id"]))
    }
}

ReloadProfileHotkeys(p) {
    BindHotkey("", p["HotkeyStart"], OnStartHotkey.Bind(p["id"]))
    BindHotkey("", p["HotkeyTag"],   OnTagHotkey.Bind(p["id"]))
    BindHotkey("", p["SwHotkey"],    OnSwHotkey.Bind(p["id"]))
    BindHotkey("", p["ComboHotkey"], OnComboHotkey.Bind(p["id"]))
}

ReloadProfile(id) {
    Global Profiles
    wasRunning := Profiles[id]["IsRunning"]
    wasSw      := Profiles[id]["IsSwitchRunning"]
    wasCb      := Profiles[id]["IsComboRunning"]
    for _, hk in [Profiles[id]["HotkeyStart"], Profiles[id]["HotkeyTag"],
                  Profiles[id]["SwHotkey"],    Profiles[id]["ComboHotkey"]] {
        if (hk != "") {
            try {
                Hotkey(HotkeyConvert(hk), "Off")
            }
        }
    }
    Profiles[id] := LoadProfile(id)
    Profiles[id]["IsRunning"] := wasRunning
    if wasSw {
        Profiles[id]["IsSwitchRunning"] := true
        Profiles[id]["SwitchStep"] := 1
        SetTimer(Profiles[id]["SwitchTimerFn"], -1)
    }
    if wasCb {
        Profiles[id]["IsComboRunning"] := true
        SetTimer(Profiles[id]["ComboTimerFn"], -1)
    }
    ReloadProfileHotkeys(Profiles[id])
}

BindHotkey(oldKey, newKey, callback) {
    if (oldKey != "") {
        try {
            Hotkey(HotkeyConvert(oldKey), "Off")
        }
    }
    if (newKey != "") {
        try {
            Hotkey(HotkeyConvert(newKey), callback, "On")
        }
    }
}

OnStartHotkey(id, *) {
    Global ActiveTab
    ActiveTab := id
    p := Profiles[id]
    if p["IsRunning"] {
        StopProfile(p)
    } else {
        StartProfile(p)
    }
    Send2JS("SWITCH_TAB|" . id)
    SendStateToJS()
}

OnTagHotkey(id, *) {
    TagMobs(Profiles[id])
    Send2JS("TAG_FIRED")
}

OnSwHotkey(id, *) {
    p := Profiles[id]
    if p["IsSwitchRunning"] {
        p["IsSwitchRunning"] := false
    } else {
        p["IsSwitchRunning"] := true
        p["SwitchStep"] := 1
        SetTimer(p["SwitchTimerFn"], -1)
    }
    SendStateToJS()
    SendProfileToJS(id)
}

OnComboHotkey(id, *) {
    p := Profiles[id]
    if p["IsComboRunning"] {
        p["IsComboRunning"] := false
    } else {
        p["IsComboRunning"] := true
        SetTimer(p["ComboTimerFn"], -1)
    }
    SendStateToJS()
    SendProfileToJS(id)
}

StartProfile(p) {
    if (!p["TargetHwnd"] || !WinExist("ahk_id " . p["TargetHwnd"])) {
        Send2JS("TOAST|⚠ Выберите целевое окно!")
        return
    }
    p["IsRunning"]    := true
    p["WarriorMode"]  := false
}

StopProfile(p) {
    p["IsRunning"]       := false
    p["WarriorMode"]     := false
    p["IsSwitchRunning"] := false
    p["IsComboRunning"]  := false
}

ResetPixel(type) {
    Global ActiveTab, Profiles
    p := Profiles[ActiveTab]
    if (type = "мана") {
        p["ManaX"] := 0, p["ManaY"] := 0, p["ManaColor"] := ""
        IniDelete(p["ini"], "Mana", "X"), IniDelete(p["ini"], "Mana", "Y"), IniDelete(p["ini"], "Mana", "Color")
        Send2JS("TOAST|🗑 Пиксель маны сброшен")
    } else if (type = "hp1") {
        p["HpX"] := 0, p["HpY"] := 0, p["HpColor"] := ""
        IniDelete(p["ini"], "HP", "X"), IniDelete(p["ini"], "HP", "Y"), IniDelete(p["ini"], "HP", "Color")
        Send2JS("TOAST|🗑 Пиксель HP сброшен")
    } else if (type = "hp2") {
        p["HpX2"] := 0, p["HpY2"] := 0, p["HpColor2"] := ""
        IniDelete(p["ini"], "HP", "X2"), IniDelete(p["ini"], "HP", "Y2"), IniDelete(p["ini"], "HP", "Color2")
        Send2JS("TOAST|🗑 Пиксель HP2 сброшен")
    } else if (type = "лут") {
        p["LootX"] := 0, p["LootY"] := 0, p["LootColor"] := ""
        IniDelete(p["ini"], "Loot", "X"), IniDelete(p["ini"], "Loot", "Y"), IniDelete(p["ini"], "Loot", "Color")
        Send2JS("TOAST|🗑 Пиксель лута сброшен")
    }
    SendProfileToJS(ActiveTab)
}

; ==========================================================
; ЛУПА
; ==========================================================
SendLoupeData() {
    Global CaptureHwnd
    try {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mx, &my)
        CoordMode("Pixel", "Screen")
        size := 4
        data := ""
        Loop (size*2+1) {
            row := A_Index - 1 - size
            Loop (size*2+1) {
                col := A_Index - 1 - size
                try {
                    c   := PixelGetColor(mx + col, my + row)
                    hex := Format("{:06X}", c & 0xFFFFFF)
                } catch {
                    hex := "000000"
                }
                data .= hex . ","
            }
            data .= ";"
        }
        Send2JS("LOUPE|" . mx . "|" . my . "|" . data)
    } catch {
    }
}

; ==========================================================
; ЗАХВАТ ПИКСЕЛЯ — через динамический хоткей Space
; ==========================================================
Global CaptureType := ""
Global CaptureHwnd := 0

StartCapture(type) {
    Global CaptureType, CaptureHwnd, ActiveTab, Profiles

    p    := Profiles[ActiveTab]
    hwnd := p["TargetHwnd"]

    if (!hwnd || !WinExist("ahk_id " . hwnd)) {
        Send2JS("TOAST|⚠ Выберите целевое окно!")
        Send2JS("CAPTURE_DONE|error")
        return
    }

    CaptureType := type
    CaptureHwnd := hwnd

    Send2JS("CAPTURE_WAIT|" . type)
    SetTimer(SendLoupeData, 50)

    ; Регистрируем Space как хоткей — он сработает в основном потоке
    try Hotkey("Space", OnCaptureSpace, "On")
}

StopCapture() {
    Global CaptureType, CaptureHwnd
    try Hotkey("Space", "Off")
    SetTimer(SendLoupeData, 0)
    Send2JS("LOUPE_STOP")
    CaptureType := ""
    CaptureHwnd := 0
}

OnCaptureSpace(*) {
    Global CaptureType, CaptureHwnd, ActiveTab, Profiles

    ; Сохраняем до StopCapture который сбросит глобалы
    type := CaptureType
    hwnd := CaptureHwnd
    p    := Profiles[ActiveTab]

    StopCapture()

    if (!hwnd) {
        hwnd := p["TargetHwnd"]
    }

    if (!hwnd || !WinExist("ahk_id " . hwnd)) {
        Send2JS("TOAST|⚠ Окно недоступно")
        Send2JS("CAPTURE_DONE|error")
        return
    }

    try {
        WinGetPos(&cx, &cy, &cw, &ch, "ahk_id " . hwnd)
    } catch {
        Send2JS("TOAST|⚠ Окно недоступно")
        Send2JS("CAPTURE_DONE|error")
        return
    }

    CoordMode "Mouse", "Screen"
    MouseGetPos(&sx, &sy)
    relX := sx - cx
    relY := sy - cy
    CoordMode "Pixel", "Screen"
    col := PixelGetColor(sx, sy)

    if (type = "мана") {
        p["ManaX"] := relX, p["ManaY"] := relY, p["ManaColor"] := col
        IniWrite(relX, p["ini"], "Mana", "X")
        IniWrite(relY, p["ini"], "Mana", "Y")
        IniWrite(col,  p["ini"], "Mana", "Color")
    } else if (type = "hp1") {
        p["HpX"] := relX, p["HpY"] := relY, p["HpColor"] := col
        IniWrite(relX, p["ini"], "HP", "X")
        IniWrite(relY, p["ini"], "HP", "Y")
        IniWrite(col,  p["ini"], "HP", "Color")
    } else if (type = "hp2") {
        p["HpX2"] := relX, p["HpY2"] := relY, p["HpColor2"] := col
        IniWrite(relX, p["ini"], "HP", "X2")
        IniWrite(relY, p["ini"], "HP", "Y2")
        IniWrite(col,  p["ini"], "HP", "Color2")
    } else if (type = "лут") {
        p["LootX"] := relX, p["LootY"] := relY, p["LootColor"] := col
        IniWrite(relX, p["ini"], "Loot", "X")
        IniWrite(relY, p["ini"], "Loot", "Y")
        IniWrite(col,  p["ini"], "Loot", "Color")
    }

    toastNames := Map("мана","Мана","hp1","HP 1","hp2","HP 2","лут","Лут")
    toastName  := toastNames.Has(type) ? toastNames[type] : type
    Send2JS("TOAST|✅ Пиксель захвачен: " . toastName)
    Send2JS("CAPTURE_DONE|" . type . "|" . col . "|" . relX . "|" . relY)
    SendProfileToJS(ActiveTab)
}

; ==========================================================
; БОЕВОЙ ЦИКЛ
; ==========================================================
ColorsMatch(c1, c2, tol := 30) {
    if (c1 = "" || c2 = "") {
        return false
    }
    try {
        c1n := Integer(c1), c2n := Integer(c2)
        r1 := (c1n >> 16) & 0xFF, g1 := (c1n >> 8) & 0xFF, b1 := c1n & 0xFF
        r2 := (c2n >> 16) & 0xFF, g2 := (c2n >> 8) & 0xFF, b2 := c2n & 0xFF
        return (Abs(r1-r2) <= tol && Abs(g1-g2) <= tol && Abs(b1-b2) <= tol)
    } catch {
        return false
    }
}

SendKeyToTarget(key, hwnd) {
    if (!hwnd || !WinExist("ahk_id " . hwnd) || key = "") {
        return
    }
    converted := ConvertKeyForSend(key)
    try {
        ControlSend(converted, , "ahk_id " . hwnd)
    }
}

ConvertKeyForSend(key) {
    if (key = "") {
        return ""
    }
    ; Если уже в фигурных скобках — не трогаем
    if (SubStr(key, 1, 1) = "{") {
        return key
    }
    ; Обрабатываем комбинации типа Ctrl+Q, Shift+3, Alt+F
    mods := ""
    k := key
    Loop {
        if (SubStr(k, 1, 5) = "Ctrl+") {
            mods .= "^", k := SubStr(k, 6)
        } else if (SubStr(k, 1, 4) = "Alt+") {
            mods .= "!", k := SubStr(k, 5)
        } else if (SubStr(k, 1, 6) = "Shift+") {
            mods .= "+", k := SubStr(k, 7)
        } else {
            break
        }
    }
    ; Одиночная буква — оборачиваем в {}
    if (StrLen(k) = 1 && k ~= "i)[a-z]") {
        return mods . "{" . k . "}"
    }
    ; Всё остальное (цифры, F-клавиши, Space и т.д.) — как есть
    return mods . k
}

CheckBoth() {
    Global Profiles
    Loop 2 {
        ProcessProfile(Profiles[A_Index])
    }
}

ProcessProfile(p) {
    hwnd := p["TargetHwnd"]
    if (!p["IsRunning"] || !hwnd || !WinExist("ahk_id " . hwnd)) {
        return
    }
    try {
        WinGetPos(&cx, &cy, &cw, &ch, "ahk_id " . hwnd)
    } catch {
        return
    }

    CoordMode("Pixel", "Screen")
    delay := p["HealDelay"]

    ; HP2 / Воин
    if (p["HpColor2"] != "" && p["Hp2Enabled"] && !p["WarriorMode"]) {
        try {
            if ColorsMatch(PixelGetColor(cx + p["HpX2"], cy + p["HpY2"]), p["HpColor2"], p["HealTolerance"]) {
                if (A_TickCount - p["LastWarriorPress"] > delay) {
                    p["WarriorMode"] := true
                    SendKeyToTarget(p["ActionKeyWarrior"], hwnd)
                    p["LastWarriorPress"] := A_TickCount
                }
            }
        } catch {
        }
    }

    ; HP1 (Срабатывает при появлении цвета)
    if (p["HpColor"] != "") {
        try {
            if ColorsMatch(PixelGetColor(cx + p["HpX"], cy + p["HpY"]), p["HpColor"], p["HealTolerance"]) {
                if (A_TickCount - p["LastHpPress"] > delay) {
                    SendKeyToTarget(p["ActionKeyHP"], hwnd)
                    p["LastHpPress"] := A_TickCount
                }
            }
        } catch {
        }
    }

    ; Выход из режима воина
    if (p["WarriorMode"] && p["HpColor2"] != "" && p["Hp2Enabled"]) {
        try {
            if !ColorsMatch(PixelGetColor(cx + p["HpX2"], cy + p["HpY2"]), p["HpColor2"], p["HealTolerance"]) {
                p["WarriorMode"] := false
            }
        } catch {
        }
    }

    ; Мана (Срабатывает при появлении цвета)
    if (!p["WarriorMode"] && p["ManaColor"] != "") {
        try {
            if ColorsMatch(PixelGetColor(cx + p["ManaX"], cy + p["ManaY"]), p["ManaColor"], p["HealTolerance"]) {
                if (A_TickCount - p["LastManaPress"] > delay) {
                    SendKeyToTarget(p["ActionKeyMana"], hwnd)
                    p["LastManaPress"] := A_TickCount
                }
            }
        } catch {
        }
    }
}

CheckLoot() {
    Global Profiles
    Loop 2 {
        p    := Profiles[A_Index]
        hwnd := p["TargetHwnd"]
        if (!p["IsRunning"] || p["LootColor"] = "" || !p["LootEnabled"] || !hwnd || !WinExist("ahk_id " . hwnd)) {
            continue
        }
        try {
            WinGetPos(&cx, &cy, &cw, &ch, "ahk_id " . hwnd)
            CoordMode("Pixel", "Screen")
            if ColorsMatch(PixelGetColor(cx + p["LootX"], cy + p["LootY"]), p["LootColor"], p["LootTolerance"]) {
                Loop p["LootPresses"] {
                    SendKeyToTarget(p["ActionKeyLoot"], hwnd)
                    Sleep(p["LootPressDelay"])
                }
                Sleep(p["LootWaitDelay"])
            }
        } catch {
        }
    }
}

RunSwitchStep(p) {
    if !p["IsSwitchRunning"] {
        return
    }
    hwnd := p["TargetHwnd"]
    if (!hwnd || !WinExist("ahk_id " . hwnd)) {
        p["IsSwitchRunning"] := false
        SendStateToJS()
        return
    }
    step := p["SwitchStep"]
    if (step = 1) {
        SendKeyToTarget(p["SwClass1Key"], hwnd)
        p["SwitchStep"] := 2
        SetTimer(p["SwitchTimerFn"], -p["SwClass1Delay"])
    } else if (step = 2) {
        SendKeyToTarget(p["SwUltKey"], hwnd)
        p["SwitchStep"] := 3
        SetTimer(p["SwitchTimerFn"], -p["SwUltDelay"])
    } else {
        SendKeyToTarget(p["SwClass2Key"], hwnd)
        p["SwitchStep"] := 1
        SetTimer(p["SwitchTimerFn"], -p["SwClass2Delay"])
    }
}

RunComboStep(p) {
    if !p["IsComboRunning"] {
        return
    }
    hwnd := p["TargetHwnd"]
    if (!hwnd || !WinExist("ahk_id " . hwnd)) {
        p["IsComboRunning"] := false
        SendStateToJS()
        return
    }
    SendKeyToTarget(p["ComboUltKey"], hwnd)
    SetTimer(p["ComboTimerFn"], -p["ComboUltDelay"])
}

TagMobs(p) {
    hwnd := p["TargetHwnd"]
    if (!hwnd || !WinExist("ahk_id " . hwnd)) {
        return
    }

    if p["TagClass1"] != "" {
        SendKeyToTarget(p["TagClass1"], hwnd)
        Sleep(p["TagClass1Delay"])
    }
    if p["TagUlt1"] != "" {
        SendKeyToTarget(p["TagUlt1"], hwnd)
        Sleep(p["TagUlt1Delay"])
    }
    if (p["MageMode"] = 1 && p["MageAimKey"] != "") {
        SendKeyToTarget(p["MageAimKey"], hwnd)
        Sleep(p["MageAimDelay"])
    }
    for _, k in [p["Mob1"], p["Mob2"], p["Mob3"], p["Mob4"]] {
        if (k != "") {
            SendKeyToTarget(k, hwnd)
            Sleep(p["MobDelay"])
        }
    }
    if p["TagClass2"] != "" {
        SendKeyToTarget(p["TagClass2"], hwnd)
        Sleep(p["TagClass2Delay"])
    }
    if p["TagUlt2"] != "" {
        SendKeyToTarget(p["TagUlt2"], hwnd)
        Sleep(p["TagUlt2Delay"])
    }
    if (p["MageMode"] = 2 && p["MageAimKey"] != "") {
        SendKeyToTarget(p["MageAimKey"], hwnd)
        Sleep(p["MageAimDelay"])
    }
}
