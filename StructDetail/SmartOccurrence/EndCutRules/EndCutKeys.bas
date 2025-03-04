Attribute VB_Name = "EndCutKeys"
Option Explicit

Public Const INPUT_BOUNDING = "Bounding"
Public Const INPUT_BOUNDED = "Bounded"

Public Const QUES_ENDCONDITION As String = "EndCondition"
Public Const CL_ENDCONDITION As String = "EndConditionCodeList"

Public Const QUES_ENDCUTTYPE As String = "EndCutType"
Public Const CL_ENDCUTTYPE As String = "EndCutTypeCodeList"

' End cut EndCondition Constants
'Public Const gsFixed = "Fixed"
'Public Const gsFree = "Free"
'Public Const gsFlangeFree = "FlangeFree"

' EndCut Type Constants
Public Const gsW = "W"
Public Const gsC = "C"
Public Const gsF = "F"
Public Const gsS = "S"
Public Const gsFV = "FV"
Public Const gsR = "R"
Public Const gsRV = "RV"

Public Const gsWeldPart = "WeldPart"
Public Const gsBottomFlange = "TheBottomFlange"
Public Const gsApplyTreatment = "ApplyTreatment"

Public Const CMLIBRARY_ENDCUTRULES As String = "EndCutRules.WebCutDefCM"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\EndCutKeys.bas"


Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'Public Sub SetQuestionEndCondition(ByRef pQH As IJDQuestionsHelper)
'    On Error GoTo ErrorHandler
'
'    Dim strError As String
''    Dim colAnswers As New Collection
''
''    colAnswers.Add "Fixed"
''    colAnswers.Add "Free"
''    colAnswers.Add "FlangeFree"
''    colAnswers.Add "Undefined"
''
''    strError = "Defining codelist."
''    pQH.DefineCodeList CL_ENDCONDITION, colAnswers
'
'    strError = "Setting question."
'    pQH.SetQuestion QUES_ENDCONDITION, gsFixed, CL_ENDCONDITION
'
'    Exit Sub
'
'ErrorHandler:
'    pQH.ReportError strError, "SetQuestionEndCondition"
'End Sub

Public Sub SetQuestionEndCutType(ByRef pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
     
    Dim strError As String
'    Dim colAnswers As New Collection
'
'    colAnswers.Add "W"
'    colAnswers.Add "C"
'    colAnswers.Add "F"
'    colAnswers.Add "S"
'    colAnswers.Add "FV"
''    colAnswers.Add "R"
''    colAnswers.Add "RV"
'
'    strError = "Defining codelist."
'    pQH.DefineCodeList CL_ENDCUTTYPE, colAnswers
    
    strError = "Setting question."
    pQH.SetQuestion QUES_ENDCUTTYPE, gsW, CL_ENDCUTTYPE
zMsgBox "   pQH.SetQuestion..." & QUES_ENDCUTTYPE & "..." & gsW
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetQuestionEndCutType", strError).Number
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




