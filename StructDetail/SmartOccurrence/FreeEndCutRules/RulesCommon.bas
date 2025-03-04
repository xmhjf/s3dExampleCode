Attribute VB_Name = "RulesCommon"
Option Explicit
'
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\FreeEndCutRules\RulesCommon.bas"

'Global string constants for questions and answers
Public Const INPUT_BOUNDED = "FreeEndPort"
Public Const INPUT_BOUNDING = "BoundingObject"
Public Const CMLIBRARY_FREEENDCUTDEFCM As String = "FreeEndCutRules.FreeEndCutDefCM"
'
Public Const QUES_ENDCUTTYPE As String = "EndCutType"
Public Const CL_ENDCUTTYPE As String = "EndCutTypeCodeList"
'
' EndCut Type Constants
Public Const gsW = "W"
Public Const gsC = "C"
Public Const gsF = "F"
Public Const gsS = "S"
Public Const gsFV = "FV"
Public Const gsR = "R"
Public Const gsRV = "RV"
'

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'********************************************************************
' ' Routine: SetQuestionEndCutType
'
' Description:
'********************************************************************
Public Sub SetQuestionEndCutType(ByRef pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
     
    Dim strError As String
    
    strError = "Setting question."
    pQH.SetQuestion QUES_ENDCUTTYPE, gsF, CL_ENDCUTTYPE
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetQuestionEndCutType", strError).Number
End Sub
