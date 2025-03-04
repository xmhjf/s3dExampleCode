Attribute VB_Name = "RuleKeys"
Option Explicit


Public Const INPUT_PLANE As String = "Plane"
Public Const INPUT_BRACKETPLANE As String = "Bracket Plane"
Public Const INPUT_BRACKETPLATE As String = "Bracket Plate System"
Public Const INPUT_UPOINT As String = "U Point"
Public Const INPUT_VPOINT As String = "V Point"
Public Const INPUT_SUPPORTS As String = "Supports"
Public Const INPUT_SUPPORT1 As String = "Support1"
Public Const INPUT_SUPPORT2 As String = "Support2"
Public Const INPUT_SUPPORT3 As String = "Support3"
Public Const INPUT_SUPPORT4 As String = "Support4"
Public Const INPUT_SUPPORT5 As String = "Support5"
Public Const INPUT_DIRECTION As String = "Direction"


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



