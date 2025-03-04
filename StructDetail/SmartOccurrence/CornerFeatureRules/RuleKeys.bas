Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_PORT1FACE = "LogicalFace"
Public Const INPUT_PORT2EDGE = "Support1"
Public Const INPUT_PORT3EDGE = "Support2"

'Global string constants for questions and answers
Public Const gsDrainage = "Drainage on Part"
Public Const gsCornerFlip = "Corner Feature Orientation"
Public Const gsPlacement = "Corner Placement"
Public Const gsCrackArrest = "Arrest Stress Cracking"
Public Const gsApplyTreatment = "ApplyTreatment"


'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function




