; Main control program to 
; - managed keyboard shortcuts
; - launches external programs

; Author: Anton Pusch
; Last update: 2021-02-10

/* Build command: 
start "$env:ProgramFiles\AutoHotkey\Compiler\Ahk2Exe.exe" "/in source\WindowTools.ahk /out bin\WindowTools.exe /icon resources\WindowTools.ico"
*/

#SingleInstance, force ; Override existing instance when lauched again
SetWorkingDir, % A_ScriptDir ; Ensures a consistent starting directory
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

^Pause:: ; Close all windows of that process    (Ctrl + Three finger gesture down)
+Pause:: ; Close window                        (Shift + Three finger gesture down)
Pause:: ; Close tab if existing otherwise close window (Three finger gesture down)
    if (GetKeyState("Ctrl", "P")) ; Is Ctrl pressed?
    { ; Close active window group
        ; Retrive information about active window group
        WinGet, windowExe, ProcessName, % "A"
        WinGetClass, windowClass, % "A"

        windowClosable := false

        if (windowExe = ahk_exe "Explorer.EXE" && windowClass = "Shell_TrayWnd")
        {} ; Taskbar
        else if (windowExe = ahk_exe "Explorer.EXE" && windowClass = "WorkerW")
        {} ; Desktop
        else if (windowExe = ahk_exe "Explorer.EXE" && windowClass = "Progman")
        {} ; Desktop
        else if (windowExe = "ApplicationFrameHost.exe" && windowClass = "ApplicationFrameWindow")
        {} ; Windows Store Apps
        else if (windowExe = "Rainmeter.exe" && windowClass = "RainmeterMeterWindow") 
        {} ; Rainmeter widget
        else
        { ; Valid window found
            windowClosable := true
        }

        if (windowClosable)
        {
            ; Close all windows of that process

            ; Get all candidates for windows of the same window group
            GroupAdd, % "activeGroup", % "ahk_exe " . windowExe . " ahk_class " . windowClass
            WinGet, windowList, List, % "ahk_group activeGroup"

            ; Double check all candidates
            loop, % windowList    
            {
                ; Only close candidates of same ProcessName and WindowClass
                WinClose, % "ahk_id " . windowList%A_Index% . " ahk_exe " . windowExe . " ahk_class " . windowClass
            }
            killTarget := "WindowGroup" ; Prevent further kills
        }
        else 
        {
            killTarget := "Window" ; Close active window
        }

    } 
    else if (GetKeyState("Shift", "P"))
    { ; Close window
        killTarget := "Window"
    }
    else
    { ; Close tab (if existing), otherwise close window
        killTarget := "Window"
        if (WinActive("ahk_exe firefox.exe") || WinActive("ahk_exe msedge.exe"))
        { ; A browser is active
            killTarget := "Tab"
        }
        else if (WinActive("ahk_exe code.exe")) ; Visual Studio Code
        { ; A tab bases program is active
            killTarget := "Tab"
        } 
        else if (WinActive("ahk_exe AcroRd32.exe") && !WinActive("Adobe Acrobat Reader DC (32-bit)"))
        { ; Adobe Acrobat Reader DC is active
            SetTitleMatchMode, 3 ; Window Title must be exactly matched
            if (!WinActive("Adobe Acrobat Reader DC (32-bit)"))
            { ; Adobe Reader has no tab open
                killTarget := "Tab"
            }
        }
        else if (WinActive("ahk_exe gitkraken.exe"))
        { ; GitKraken is active but no tab is open
            ; Find Gitkraken window
            CoordMode, pixel, screen
            WinGetPos, gitkrakenX, gitkrakenY, gitkrakenWidth, gitkrakenHeight, % "ahk_exe gitkraken.exe"
            
            ; Is the close tab cross visible? = Multiple tabs open?
            tabCrossImage := getFile("GitKraken NoTabCross.png", [".", "..\resources"])
            ImageSearch, imageX, imageY, gitkrakenX, gitkrakenY, % gitkrakenX + gitkrakenWidth, % gitkrakenY + gitkrakenHeight, % tabCrossImage
            if (ErrorLevel)
            { ; At least one tab open
                killTarget := "Tab"
            }
        }
    }
    if (killTarget = "Window")
    { ; Close window
        Send, !{F4}
    } 
    else if (killTarget = "Tab")
    { ; Close tab
        Send, ^w
    }
return

#o::
    Run, % "explorer shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}"
return

; Open new tab / Open action center
; Activated by touchpad (internal shortcut)
CtrlBreak::
    if (WinActive("ahk_exe firefox.exe") || WinActive("ahk_exe msedge.exe") || WinActive("ahk_exe gitkraken.exe"))
    { ; Browser(like) window is active
        Send, ^t ; Open new tab
    }
    else
    {
        Send, #a ; Open action center
    }
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
 * ->  "K" Kill the windows notification
 * -> e.g: "S M" means silent Message Dialog
 */
toast(title := "", message := "", options := "")
{
    if (InStr(options, "K"))
    { ; Kill previous toast
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

            MsgBox, % flags, % title, % message
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

getFile(filename, validLocations) {
    path := 0
    for i, location in validLocations
    { ; Check all locations and use the first valid one
        if (FileExist(location . "\" . filename))
        { ; Location is valid
            path := location . "\" . filename
            return path
        }
    }
    if (!path)
    { ; No valid location found
        toast("File missing", A_ScriptDir . "\" . filename, "E")
        return
    }
}