#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, Off

; ==========================================================
; RUCOY MACRO PRO — Launcher
; Просто запускает macro.exe (который сам открывает окно)
; ==========================================================

rootDir   := A_ScriptDir
coreDir   := rootDir . "\core"
macroFile := coreDir . "\macro.exe"


if !FileExist(macroFile) {
    MsgBox("Файл core\macro.exe не найден!", "Ошибка", "Icon!")
    ExitApp
}

; Проверяем что уже не запущен
statusFile := A_Temp . "\rucoy_macro_running.txt"
if FileExist(statusFile) {
    try {
        age := DateDiff(A_Now, FileGetTime(statusFile, "M"), "Seconds")
        if (age < 5) {
            MsgBox("Rucoy Macro Pro уже запущен!", "Инфо", "Icon!")
            ExitApp
        }
    }
}

; Запускаем напрямую .exe файл без интерпретатора
Run('"' . macroFile . '"')
ExitApp