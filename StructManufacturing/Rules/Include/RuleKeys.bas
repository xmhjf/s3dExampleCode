Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_PLATE As String = "Plate"
Public Const INPUT_PROFILE As String = "Profile"
Public Const INPUT_TUBE As String = "Tube"
Public Const INPUT_EDGE As String = "Edge"

'Global string constants for Plate Stress Level
'Public Const gsHigh = "High"
'Public Const gsMedium = "Medium"
'Public Const gsLow = "Low"

'Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
'  MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
'    & " - " & "::" & sFunctionName & " - " & sErrorName
'End Sub

Public Enum InputIndex
    igTemplateIndex = 1
    igPlateIndex = 1
    igProfileIndex = 1
End Enum

 
