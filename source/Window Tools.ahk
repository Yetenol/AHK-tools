; Main control program to 
; - managed keyboard shortcuts
; - launches external programs

; Author: Anton Pusch
; Last update: 2021-02-10

#SingleInstance, force ; Override existing instance when lauched again
Menu, Tray, Icon, % A_WinDir "\system32\imageres.dll", 174 ; Setup a keyboard as taskbar icon:
Menu, Tray, Add ; Creates a separator line.
Menu, Tray, Add, Send Pause, SendPause
Menu, Tray, Add, Send Ctrl+Pause, SendCtrlBreak

; Always use digits on NumPad
SetNumLockState, AlwaysOn
return

; ===== Global shortcuts: =====
; Modifier keys:    # Win    ^ Ctrl    + Shift    ! Alt
; Notifications flags:


; -=-=-=-=-=- Windows Media API -=-=-=-=-=-
; Enables remote media control for Netflix, PrimeVideo

; Is active window a media player?
MediaPlayerActive() {
    return (WinActive("Netflix ahk_class ApplicationFrameWindow") ;  Netflix
        || WinActive("Amazon Prime Video for Windows ahk_class ApplicationFrameWindow")) ; PrimeVideo
}

; Play/Pause media (Netflix, PrimeVideo)
Media_Play_Pause::
    media_is_winding := false ; stop media winding
    if (MediaPlayerActive())
    {
        Send, {Space}
    }
return

#MaxThreadsPerHotkey 2; allow 2 threads so that these hotkeys can "turn themselves off"
; Fast Forwards media until key pressed again or paused (Netflix, PrimeVideo)
Media_Next::
    MediaWind("fast_forward")
return

; Rewind media until key pressed again or paused (Netflix, PrimeVideo)
Media_Prev::
    MediaWind("rewind")
return
#MaxThreadsPerHotkey 1 ; default

MediaWind(direction) 
{
    global media_is_winding
    if (MediaPlayerActive()) 
    {
        media_is_winding := !media_is_winding ; start/stop winding (stop kills over thread)
        while (media_is_winding && MediaPlayerActive())
        {
            Send, % (direction="fast_forward") ? "{Right}" : "{Left}" ; forward media / rewind media
            Sleep, % WinActive("Netflix") ? WinActive("Amazon Prime Video for Windows") ? 1500 : 1000 : 1000
        }
    }
    return
}

; Bring the Rainmeter widgets to the foreground (Win + Shift + R)
; - can be used as screensaver
#+r::
    Run, restartRainmeter.ps1.bat
return

; Convert screen region to text (Optical character recognition)
; - Uses external program Capture2Text.exe
!+PrintScreen:: ; Reuse last region: (Shift + Alt + PrintScreen)
!PrintScreen::  ; Select new region          (Alt + PrintScreen)
    PressingShiftKey := GetKeyState("Shift", "P")
    
    CoordMode, mouse, screen
    MouseGetPos, MouseX, MouseY

    Process, Exist, Capture2Text.exe ; Check whether AutoHotkey.exe is running
    If (ErrorLevel = 0) 
    { ; Capture2Text isn't running
        Run, % ProgramFiles "\Capture2Text\Capture2Text.exe"
        toast("Capture2Text wasn't running", "Launching it..", "S M")
        
        ; Launch Capture2Text
        ErrorLevel := 0
        while (ErrorLevel = 0)
        {
            Process, Exist, Capture2Text.exe ; Check whether AutoHotkey.exe is running
        }

        Sleep, 1000
        MouseMove, % MouseX, % MouseY ; Restore previous mouse position
        Send % (PressingShiftKey) ? "!+{PrintScreen}" : "!{PrintScreen}"
    }
return

; Close active window/tab
; Activated by touchpad (internal shortcut)
+Pause::
Pause::
    SetTitleMatchMode, 3 ; Window Title must match exactly (used for Adobe's last tab)
    if WinActive("ahk_exe firefox.exe") ; If active window is a browser
        || WinActive("ahk_exe msedge.exe") 
    || WinActive("ahk_exe brave.exe") 
    || WinActive("ahk_exe code.exe")
    || WinActive("ahk_exe gitkraken.exe")
    || (WinActive("ahk_exe AcroRd32.exe") && !WinActive("Adobe Acrobat Reader DC")) ; Adobe has no tab open
    Send, ^w ; Close active tab
    else
        Send, !{F4} ; Close active program
return

; Open new tab / Open action center
; Activated by touchpad (internal shortcut)
CtrlBreak::
    if WinActive("ahk_exe firefox.exe") ; If active window is a browser
        || WinActive("ahk_exe msedge.exe") 
    || WinActive("ahk_exe brave.exe") 
    Send, ^t ; Open new tab
    else
        Send, #a ; Open action center
return

; Pin active window always on top (Win + Numpad-)
#NumpadSub::
    Run, nircmd win settopmost foreground 1
return

; Unpin active window always on top (Win + Shift + Numpad-)
#+NumpadSub::
    Run, nircmd win settopmost foreground 0
return

; Restart StartMenu process (Win + F5)
#F5::
    Run, powershell -Command "Stop-Process -ProcessName StartMenuExperienceHost"
    Sleep, 1000
    Send, {LWin}
return

; Restart Explorer process (Win + Shift + F5)
#+F5::
    Run, powershell -Command "Stop-Process -ProcessName Explorer"
    Sleep, 2000
    Send, {LWin}
return

; Make active window transparent
; - Uses external program nircmd in path location

; Clear active window's transparency (Win + Numpad0)
#Numpad0::
    Run, nircmd win trans foreground 255
return
; Set active window's transparency to 90% (Win + Numpad1)
#Numpad1::
    Run, nircmd win trans foreground 227
return
; Set active window's transparency to 78% (Win + Numpad2)
#Numpad2::
    Run, nircmd win trans foreground 198
return
; Set active window's transparency to 67% (Win + Numpad3)
#Numpad3::
    Run, nircmd win trans foreground 170
return
; Set active window's transparency to 56% (Win + Numpad4)
#Numpad4::
    Run, nircmd win trans foreground 142
return
; Set active window's transparency to 44% (Win + Numpad5)
#Numpad5::
    Run, nircmd win trans foreground 113
return
; Set active window's transparency to 33% (Win + Numpad6)
#Numpad6::
    Run, nircmd win trans foreground 85
return
; Set active window's transparency to 22% (Win + Numpad7)
#Numpad7::
    Run, nircmd win trans foreground 57
return
; Set active window's transparency to 11% (Win + Numpad8)
#Numpad8::
    Run, nircmd win trans foreground 28
return
; Make active window's transparency invisible (Win + Numpad9)
#Numpad9::
    Run, nircmd win trans foreground 0
return

; Send PAUSE
SendPause:
    toast("Send PAUSE", "Sending PAUSE in 2s", "S")
    Sleep, 2000
    Send, % "{Pause}"
return

; Send CTRL + PAUSE
SendCtrlBreak:
    toast("Send PAUSE", "Sending CTRL + PAUSE in 2s", "S")
    Sleep, 2000
    Send, % "{CtrlBreak}"
return
 
/* Show a notification
 * @param (optional) title: Text header
 * @param (optional) message: Text body
 * @param (optional) options: space seperated string of flags
 * ->  "" Default    "I" Info    "W" Warning    "E" Error
 * ->  "S" Silent
 * ->  "M" Use AHK's message dialog, pauses the script
 * -> e.g: "S M" means silent Message Dialog
 */
toast(title := "", message := "", options := "")
{
    if (title = "" && message = "")
    { ; Empty toast => Hide previous toast
        if (!InStr(options, "M"))
        { ; Default windows notification
            Menu, Tray, NoIcon ; Kills the notification ballon
            Sleep, 10
            Menu, Tray, Icon
        }
    }
    else
    { ; Valid notification text
        message := (message = "") ? " " :  message  ; Message can't be empty

        flags := 0   ; Read options
        if (InStr(options, "M"))
        { ; Use MsgBox Dialog (pauses script)
            flags += (InStr(options, "I")) ? 0x40 : 0
            flags += (InStr(options, "W")) ? 0x30 : 0
            flags += (InStr(options, "E")) ? 0x10 : 0

            MsgBox, flags, title, text
        }
        else
        { ; Default windows notification
            flags += (InStr(options, "I")) ? 0x1 : 0
            flags += (InStr(options, "W")) ? 0x2 : 0
            flags += (InStr(options, "E")) ? 0x3 : 0
            flags += (InStr(options, "S")) ? 0x10 : 0

            TrayTip, % title, % message,, % flags
        }
    }


}