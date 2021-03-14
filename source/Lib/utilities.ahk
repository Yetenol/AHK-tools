; Main AHK library 
; - various functions

; Author: Anton Pusch
; Last update: 2021-02-14

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
    for i, location in validLocations
    { ; Check all locations and use the first valid one
        if (attributes := FileExist(location))
        { ; Path is valid
            if (InStr(attributes, "D"))
            { ; Path is a directory
                if (FileExist(location "\" filename))
                { ; Location contains a valid file
                    return location "\" filename
                }
            }
            else
            { ; Location is a valid file
                return location
            }
        }
    }
    ; No valid location found
    toast("File missing", A_ScriptDir "\" filename, "E")
    return   
}

; Is the active window a browser?
BrowserActive() {
    return WinActive("ahk_exe firefox.exe") || WinActive("ahk_exe msedge.exe") || WinActive("ahk_exe chrome.exe")
}