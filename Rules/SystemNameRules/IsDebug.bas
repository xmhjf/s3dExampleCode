Attribute VB_Name = "IsDebug"
Option Explicit

Public Function IsInDebugMode() As Boolean

    On Error GoTo ErrorHandler
    
    'If the program is compiled, the following Debug statement
    ' has been removed so it will not generate an error.
    Debug.Print 1 / 0
    IsInDebugMode = False
    Exit Function
    
ErrorHandler:
    'We got an error so the Debug.Print statement must be working.
    IsInDebugMode = True
    Err.Clear
End Function
