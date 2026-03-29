#Requires AutoHotkey v2.0
#SingleInstance Force

exePath := A_ScriptDir . "\macro.exe"
RepoUrl := "https://raw.githubusercontent.com/dobbitaylor-dot/ponosdeda111/main/"

try {
    ; Скачиваем новый скомпилированный EXE
    Download(RepoUrl . "macro.exe", exePath)
    
    ; Запускаем командную строку, которая подождет, удалит этот старый .ahk и запустит .exe
    cmd := A_ComSpec ' /c ping 127.0.0.1 -n 2 > nul & del /f /q "' A_ScriptFullPath '" & start "" "' exePath '"'
    Run(cmd, , "Hide")
} catch {
    MsgBox("Произошла ошибка при переходе на новую версию. Пожалуйста, попросите у создателя новый файл macro.exe")
}
ExitApp()
