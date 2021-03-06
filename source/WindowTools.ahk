; Main control program to 
; - managed keyboard shortcuts
; - launches external programs

; Author: Anton Pusch
; Last update: 2021-02-10

/* Build command: 
start "$env:ProgramFiles\AutoHotkey\Compiler\Ahk2Exe.exe" "/in source\WindowTools.ahk /out bin\WindowTools.exe /icon resources\WindowTools.ico"
*/

#Include, <utilities>
#SingleInstance, force ; Override existing instance when lauched again
SetWorkingDir, % A_ScriptDir ; Ensures a consistent starting directory
Menu, Tray, Icon, % A_WinDir "\system32\imageres.dll", 174 ; Setup a keyboard as taskbar icon:
Menu, Tray, Add ; Creates a separator line.
Menu, Tray, Add, Send Pause, SendPause
Menu, Tray, Add, Send Ctrl+Pause, SendCtrlBreak

; Always use digits on NumPad
SetNumLockState, AlwaysOn
return

; ==================== Window Lists: ====================

BrowserActive() { ; Is the active window a browser?
    return WinActive("ahk_exe firefox.exe") || WinActive("ahk_exe msedge.exe") || WinActive("ahk_exe chrome.exe")
}

MediaPlayerActive() { ; Is active window a media player?
    return (WinActive("Netflix ahk_class ApplicationFrameWindow") ;  Netflix
        || WinActive("Amazon Prime Video for Windows ahk_class ApplicationFrameWindow")) ; PrimeVideo
}

; ==================== Windows Media API ====================
; Enables remote media control for Netflix, PrimeVideo

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

; ==================== Touchpad gestures ====================
; ===== Close gesture =====
^Pause:: ; Close all windows of that process    (Ctrl + Three finger down)
+Pause:: ; Close window                        (Shift + Three finger down)
Pause:: ; Close tab if existing otherwise close window (Three finger down)
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
            GroupAdd, % "activeGroup", % "ahk_exe " windowExe " ahk_class " windowClass
            WinGet, windowList, List, % "ahk_group activeGroup"

            ; Double check all candidates
            loop, % windowList    
            {
                ; Only close candidates of same ProcessName and WindowClass
                WinClose, % "ahk_id " windowList%A_Index% " ahk_exe " windowExe " ahk_class " windowClass
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
        if (BrowserActive())
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
        else if (WinActive("ahk_exe bMC.exe"))
        { ; Baramundi Managment Center
            ; Is the active tab a client?
            image := getFile("BMC Active client.png", [".", "..\resources"])
            if (locateImageInWindow("ahk_exe bMC.exe", image))
            { ; Found image! => Active tab is a client
                killTarget := "Tab"
            }
            else
            { ; Cannot find image! => No client is active
                ; Is there an inactive client open?
                image := getFile("BMC Inactive client.png", [".", "..\resources"])
                if (clickImageInWindow("ahk_exe bMC.exe", image))
                { ; Found image! => Switched to inactive client
                    killTarget := "Tab"
                }
                else
                { ; Cannot find image! => No clients are open
                    ; Protected tab shoundn't be closed
                    image := getFile("BMC Expand client.png", [".", "..\resources"])
                    if (clickImageInWindow("ahk_exe bMC.exe", image))
                    { ; Found image! => At least one client open
                        killTarget := "none"
                    }
                    else
                    { ; Cannot find image! => No client open
                        killTarget := "none"
                        toastQuestion("No more clients found!", "Do you want to exit?", 3, false, 0x4)
                        IfMsgBox, Yes
                        {
                            killTarget := "Window"
                        }
                    }
                }
            }
        }
        else if (WinActive("ahk_exe gitkraken.exe"))
        { ; GitKraken is active but no tab is open
            image := getFile("GitKraken single empty tab.png", [".", "..\resources"])
            if (!locateImageInWindow("ahk_exe gitkraken.exe", image))
            { ; Cannot find image! => At least one tab open
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

; ===== Open gesture =====
CtrlBreak:: ; Open new tab / Open action center (Three finger tap)
    if (WinActive("ahk_exe bMC.exe"))
    { ; Baramundi Managment Center
        ; Open the Environment tab?
        image := getFile("BMC environment tab.png", [".", "..\resources"])
        if (!clickImageInWindow("ahk_exe bMC.exe", image))
        { ; Cannot find image!
            toastError("Cannot find Environment tab")
        }
    }
    else if (BrowserActive() || WinActive("ahk_exe gitkraken.exe"))
    { ; Browser(like) window is active
        Send, ^t ; Open new tab
    }
    else
    {
        Send, #a ; Open action center
    }
return

; ==================== Window shortcuts ====================

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

; ==================== Transparency shortcuts ====================

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
    toastInfo("Send PAUSE", "Sending PAUSE in 2s",, false)
    Sleep, 2000
    Send, % "{Pause}"
return

; Send CTRL + PAUSE
SendCtrlBreak:
    toastInfo("Send PAUSE", "Sending CTRL + PAUSE in 2s",, false)
    Sleep, 2000
    Send, % "{CtrlBreak}"
return