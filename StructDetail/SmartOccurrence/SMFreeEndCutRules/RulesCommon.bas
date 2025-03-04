Attribute VB_Name = "RulesCommon"
Option Explicit
'
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "FreeEndCutRules\RulesCommon.bas"

'Global string constants for questions and answers
Public Const FE_INPUT_BOUNDED = "FreeEndPort"
Public Const FE_INPUT_BOUNDING = "BoundingObject"
Public Const CMLIBRARY_FREEENDCUTDEFCM As String = CUSTOMERID + "FreeEndCutRules.FreeEndCutDefCM"
'
Public Const QUES_ENDCUTTYPE As String = "EndCutType"
Public Const CL_ENDCUTTYPE As String = "EndCutTypeCodeList"
'

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
    pQH.SetQuestion QUES_ENDCUTTYPE, "Snip", CL_ENDCUTTYPE
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetQuestionEndCutType", strError).Number
End Sub

Public Sub MigrateTheAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)

    'kinda cool name huh ?  Like: Luke, the Alien !

    Const METHOD = "MigrateTheAggregator"
    On Error GoTo ErrorHandler
    MigrateFreeEndCut pAggregatorDescription, pMigrateHelper

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

