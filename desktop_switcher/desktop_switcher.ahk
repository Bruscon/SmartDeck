#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%
SetBatchLines -1  ; Make script run faster

; Toggle debug mode (set to true to show debug window, false to hide)
DEBUG := false

; Create debug GUI if DEBUG is enabled
if (DEBUG) {
    Gui, +AlwaysOnTop
    Gui, Add, Text,, Debug Information:
    Gui, Add, Edit, w400 h200 vDebugBox
    Gui, Show, w420 h230, Desktop Switcher Debug
}

; Function to add debug messages
AddDebug(message) {
    if (!DEBUG)
        return
        
    GuiControlGet, currentText,, DebugBox
    newText := currentText . message . "`n"
    GuiControl,, DebugBox, %newText%
}

; Log startup information
AddDebug("Script started: " . A_Now)

; Load the VirtualDesktopAccessor.dll
dllPath := "C:\Users\Nick Brusco\Documents\projs\scripts\desktop_switcher\VirtualDesktopAccessor.dll"
AddDebug("Attempting to load DLL from: " . dllPath)

hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", dllPath, "Ptr")
if (hVirtualDesktopAccessor = 0) {
    AddDebug("ERROR: Failed to load DLL!")
    MsgBox, Failed to load the VirtualDesktopAccessor.dll! Check the path.
    ExitApp
} else {
    AddDebug("DLL loaded successfully, handle: " . hVirtualDesktopAccessor)
}

; Get function address
GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
if (GoToDesktopNumberProc = 0) {
    AddDebug("ERROR: Failed to get GoToDesktopNumber procedure!")
    MsgBox, Failed to find the GoToDesktopNumber function in the DLL!
    ExitApp
} else {
    AddDebug("Found GoToDesktopNumber procedure at: " . GoToDesktopNumberProc)
}

; Function to switch desktop with error handling
GoToDesktop(desktopNumber) {
    global GoToDesktopNumberProc
    AddDebug("Attempting to go to desktop: " . desktopNumber)
    
    if (GoToDesktopNumberProc != 0) {
        try {
            result := DllCall(GoToDesktopNumberProc, Int, desktopNumber)
            AddDebug("Function call completed with result: " . result)
        } catch e {
            AddDebug("ERROR in DllCall: " . e.message)
        }
    } else {
        AddDebug("ERROR: Cannot call function, address is invalid")
    }
}

; Define hotkeys for 8 virtual desktops
F13::
    AddDebug("F13 pressed - Switching to Desktop 1")
    GoToDesktop(0)
    return

F14::
    AddDebug("F14 pressed - Switching to Desktop 2")
    GoToDesktop(1)
    return

F15::
    AddDebug("F15 pressed - Switching to Desktop 3")
    GoToDesktop(2)
    return

F16::
    AddDebug("F16 pressed - Switching to Desktop 4")
    GoToDesktop(3)
    return

F17::
    AddDebug("F17 pressed - Switching to Desktop 5")
    GoToDesktop(4)
    return

F18::
    AddDebug("F18 pressed - Switching to Desktop 6")
    GoToDesktop(5)
    return

F19::
    AddDebug("F19 pressed - Switching to Desktop 7")
    GoToDesktop(6)
    return

F20::
    AddDebug("F20 pressed - Switching to Desktop 8")
    GoToDesktop(7)
    return

; Emergency exit key (keep this for troubleshooting)
^!Escape::
    AddDebug("Emergency exit triggered")
    ExitApp
    return