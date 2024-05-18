Option Explicit
Dim apptorun, objShell
Set objShell = CreateObject("WScript.Shell")

apptorun= InputBox ("Type the name of a program, folder, document, or internet resource, and SPRINT will open it for you.","SPRINT (Spinoff of Run)")

' Set the compatibility layer for the script
objShell.Environment("PROCESS")("__COMPAT_LAYER") = "RunAsInvoker"

' Start msconfig
objShell.Run apptorun



