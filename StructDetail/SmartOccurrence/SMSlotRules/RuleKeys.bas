Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_PENETRATING = "Penetrating"
Public Const INPUT_PENETRATED = "Penetrated"


' Assembly Method Constants
Public Const gsDrop = "Drop"
Public Const gsSlide = "Slide"


'Library for Question Custom Methods
Public Const LIBRARY_SOURCE_ID = m_sProjectName + ".SlotSelCM"
Public Const CMLIBRARY_SLOTDEFCM As String = m_sProjectName + ".SlotDefCM"


Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
  MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
    & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub


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


