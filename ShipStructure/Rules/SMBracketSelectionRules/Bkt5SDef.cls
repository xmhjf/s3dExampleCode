VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt5SDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "BktSelRules.Bkt5SDef"
Const m_ItemName As String = "BktSelRules.Bkt5SDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktSelRules\Bkt5SDef.cls"

Implements IJDUserSymbolServices

' ** Start CM Methods implementation**
' ********************************************************************************************


Public Sub CMFinalConstructBracket(ByVal oAggregatorDescription As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    Const sMETHOD = "CMFinalConstructBracket"
    
    
    Exit Sub
ErrorHandler:
    'If reinforcement delete failed then it is a warning
    'Raise warning so that bound happens.
    LogToDoRaiseWaring oAggregatorDescription.CAO, MODULE, sMETHOD
End Sub


Public Sub CMFinalConstructBracketReinforcement(pMemberDesc As IJDMemberDescription)
    Const sMETHOD = "CMFinalConstructBracketReinforcement"
    On Error GoTo ErrorHandler
     
 
    Exit Sub
ErrorHandler:
    'If reinforcement delete failed then it is a warning
    'Raise warning so that bound happens.
    LogToDoRaiseWaring pMemberDesc.CAO, MODULE, sMETHOD
End Sub



Public Sub CMConstructBracketReinforcement(ByVal pMemberDescription As IJDMemberDescription, _
                                            ByVal oResourceManager As IUnknown, _
                                            ByRef oFlange As Object)
    Const sMETHOD = "CMConstructBracketReinforcement"
    On Error GoTo ErrorHandler
    
    Dim clsBracketCM As New BracketDefCM
    clsBracketCM.ConstructBracketReinforcement pMemberDescription, oResourceManager, oFlange
    
    If oFlange Is Nothing Then
        GoTo ErrorHandler
    End If
    
    Set clsBracketCM = Nothing
    
    Exit Sub
ErrorHandler:
    'If reinforcement creation failed then it is a warning
    'Raise warning so that bound happens.    'No executable statement before log error, don't add error handleing
    LogToDoRaiseWaring pMemberDescription.CAO, MODULE, sMETHOD
End Sub

Public Sub CMBracketReinforcementCondition(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const sMETHOD = "CMBracketReinforcementCondition"
    On Error GoTo ErrorHandler
    
    Dim clsBracketCM As New BracketDefCM
    clsBracketCM.IsBracketReinforcementNeeded pMD, bIsNeeded
    Set clsBracketCM = Nothing
    
    Exit Sub
ErrorHandler:
    'If reinforcement creation failed then it is a warning
    'Raise warning so that bound happens.    'No executable statement before log error, don't add error handleing
    LogToDoRaiseWaring pMD.CAO, MODULE, sMETHOD
End Sub

Public Sub CMDeleteBracketReinforcement(ByVal pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Const sMETHOD = "CMDeleteBracketReinforcement"
    
    Dim clsBracketCM As New BracketDefCM
    clsBracketCM.DeleteBracketReinforcement pMemberDescription
    Set clsBracketCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    'If reinforcement delete failed then it is a warning
    'Raise warning so that bound happens.
    LogToDoRaiseWaring pMemberDescription.CAO, MODULE, sMETHOD
End Sub
Public Sub CMEvaluateBracketBoundaries(oPropertyDescriptions As IJDPropertyDescription, oObject As Object)
    Const sMETHOD = "CMEvaluateBracketBoundaries"
    
    On Error GoTo ErrorHandler
    
    EvaluateBracketBoundaries oPropertyDescriptions, oObject
    
    Exit Sub
    
ErrorHandler:
    'If Trim (i.e. apply boudaries) fails then it is an error
    'No need to execute the bound if trim fails, so raise E_FAIL
    LogToDoRaiseError oObject, MODULE, sMETHOD
End Sub

' ** End CM **


' ********************************************************************************************
'         !!!!! Start USS Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
    InitializeBracketSymbolDefinition pDefinition
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Set IJDUserSymbolServices_InstanciateDefinition = InstanciateBracketDefinition(m_ItemProgid, m_ItemName, CodeBase, defParams, ActiveConnection)
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

