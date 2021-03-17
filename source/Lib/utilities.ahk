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
;error() {
;    toast
;}
toastError(title := "", message := "", timeout := -1, options := 0) {
    toast(title, message, options,, "Error")
}
toastInfo(title := "", message := "", timeout := -1, options := 0) {
    toast(title, message, options,, "Info")
}
toastWarning(title := "", message := "", timeout := -1, options := 0) {
    toast(title, message, options,, "Warning")
}

toast(title := "", message := "", timeout := -1, options := 0, styleIcon = "")
{
    ; Deside whether to use native Balloon notifications or a Message Box
    EnvGet, domain, USERDOMAIN
    design := (options|| domain = "TUV") ? "Message Box" : "Balloon"
    
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
        if (InStr(options, "M") || domain = "TUV")
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

; Handle external resource
; - Get the correct filepath from multiple posibilities
; - Display error messages
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

; Find a specific image inside a window
locateImageInWindow(window, imagePath) {
    Coordmode, % "pixel", % "screen"
    if(!WinExist(window))
    { ; Window doesn't exist
        toast("Targeted window doesn't exist", "Window:`t" window "`nImage:`t" imagePath, "E")
    }
    WinGetPos, x, y, width, height, % window
    
    ImageSearch, imageX, imageY, x, y, % x + width, % y + height, % imagePath
    if (ErrorLevel = 2)
    { ; PROBLEM that prevented the command from conducting the search (such as failure to open the image file or a badly formatted option)
        toast("Failure to execute imageSearch", "Window:`t" window "`nImage:`t" imagePath, "E")
        return {}
    }
    else if (ErrorLevel = 1)
    { ; Cannot find image! => At least one tab open
        return false
    }
    else
    { ; Image was found!
        return {x: imageX, y: imageY}
    }
}

; Click a specific image inside a window
clickImageInWindow(window, imagePath) {
    image := locateImageInWindow(window, imagePath)
    if (!image)
    { ; Cannot find image!
        return false
    }
    else
    { ; Image was found!
        ; Click the image
        Coordmode, % "mouse", % "screen"
        MouseGetPos, mouseX, mouseY
        Coordmode, % "pixel", % "screen"

        Click, % image.x " " image.y
        
        ; Keep previous mouse position
        MouseMove, % mouseX, % mouseY
        return true
    }
}