; Extract and call the phone number from a text selection
; - copies text
; - extracts first phone number
; - formats number (+ => 00, no spaces)
; - launches external programs
; - starts the call

; Author: Anton Pusch
; Last update: 2021-03-10

/* Build command: 
start "$env:ProgramFiles\AutoHotkey\Compiler\Ahk2Exe.exe" "/in source\LaunchWorkspace.ahk /out bin\LaunchWorkspace.exe /icon resources\LaunchWorkspace.ico"
*/  


#Include, <utilities>
#SingleInstance, force ; Override existing instance when lauched again
SetWorkingDir, % A_ScriptDir ; Ensures a consistent starting directory
Menu, Tray, Icon, % A_WinDir "\system32\SHELL32.dll", 3 ; Setup a keyboard as taskbar icon:

EnvGet, ProgramFiles_x86, % "ProgramFiles(x86)"

; Wait for successful VPN connection
Menu, Tray, Tip, % "Reconnect VPN"
RunWait, % getFile("Reconnect-VPN.ps1.bat", [".", "..\bin", "C:\Dev\studit"])

; Launch LAPS PWA
Menu, Tray, Tip, % "Launch LAPS"
WinClose, % "Log in ahk_exe chrome.exe" ; Kill broken windows
if (!WinExist("LAPS ahk_exe chrome.exe")) { ; Process not running jet
    Run, % ProgramFiles_x86 "\Google\Chrome\Application\chrome_proxy.exe  --profile-directory=Default --app-id=mfkfhplmaddecclcdbckohbbkcnlepej",, MAX
}

; Launch Cherwell PWA
Menu, Tray, Tip, % "Launch Cherwell"
if (!WinExist("ahk_exe Trebuchet.App.exe")) { ; Process not running jet
    Run, % ProgramFiles_x86 "\Cherwell Software\Cherwell Service Management\Trebuchet.App.exe",, MAX
}

; Kill & launch Baramundi Managment Agent
Menu, Tray, Tip, % "Launch Baramundi"
Process, Close, % "bMC.exe"
Run, % ProgramFiles_x86 "\baramundi\Management Center\bMC.exe"
WinWait, % "baramundi Management Center 2019 ahk_exe bMC.exe"
WinActivate, % "baramundi Management Center 2019 ahk_exe bMC.exe"
Send, {Enter}

; Kill & launch Avaya one-X Communicator
Menu, Tray, Tip, % "Launch Avaya"
Process, Close, % "onexcui.exe"
WinWaitClose, % "ahk_exe onexcui.exe"
Run, % ProgramFiles_x86 "\Avaya\Avaya one-X Communicator\onexcui.exe"
WinWait, % "WindowSplash ahk_exe onexcui.exe"
WinWait, % "Avaya one-X ahk_exe onexcui.exe"
Sleep, 500
WinWait, % "Avaya one-X ahk_exe onexcui.exe"
WinActivate, % "Avaya one-X ahk_exe onexcui.exe"
Send, {Enter}
WinWait, % "WindowMessageBox onexcui.exe",, 5
if (!ErrorLevel) { ; Dialog was found
    WinActivate, % "WindowMessageBox onexcui.exe"
    Send, {Tab}{Enter}
}

; Launch Microsoft Outlook
Menu, Tray, Tip, % "Launch Outlook"
Process, Exist, % "Outlook.exe"
if (!ErrorLevel) { ; Process not running jet
    Run, % ProgramFiles_x86 "\Microsoft Office\Office16\OUTLOOK.EXE"
}

;ahk_class HwndWrapper[onexcui.exe;;1dbdd2c3-2b80-42c6-8c21-c533463b5d73]
;ahk_class HwndWrapper[onexcui.exe;;9643fcd4-1653-4329-8ef0-ada96c0ef216]

Menu, Tray, Tip, % "Workspace ready!"
MsgBox,, % "Launch Workspace", % "Workspace ready!", 5