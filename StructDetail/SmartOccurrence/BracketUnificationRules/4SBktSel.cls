VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt4SSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgid As String = "BktUnifiedRules.Bkt4SSel"
Const m_SelectorName As String = "BktUnifiedRules.Bkt4SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BracketRules\4SBktSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
    pIH.SetInput INPUT_SUPPORTS
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    ' TODO - Add your question below
    pQH.SetQuestion QUESTION_BracketSupport3Type, _
                    DEFAULT_BracketSupport4or5Type_TRIMMED, _
                    CODELIST_BracketSupport4or5Type
    
    pQH.SetQuestion QUESTION_BracketSupport4Type, _
                    DEFAULT_BracketSupport4or5Type_TRIMMED, _
                    CODELIST_BracketSupport4or5Type
                    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
 
    ' Get the answers to the questions...
    Dim sBracketSupport3Type As String
    Dim sBracketSupport4Type As String
    sBracketSupport3Type = pSL.Answer(QUESTION_BracketSupport3Type)
    sBracketSupport4Type = pSL.Answer(QUESTION_BracketSupport4Type)
    
    'Add brackets that are always valid...
    pSL.Add "BKT_4SLCL"
            
    If (sBracketSupport3Type = "SetBack") Then
        pSL.Add "BKT_4SLCL_SBS3"
    End If
    
    If (sBracketSupport4Type = "SetBack") Then
        pSL.Add "BKT_4SLCL_SBS4"
    End If
    
    '******************************************************************
    ' If supports 3 and 4 are profiles, then add offset from flange bracket...
    '
    ' Get Bracket-by-Plane Smart Occurrence...
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = pSL.SmartOccurrence
      
    ' Create helper for bracket...
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    
    'Get the supports from BracketContourSO....
    Dim nSupports As Long
    Dim oSupportsColl As IJElements
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupportsColl, nSupports
        
    'Get Supports types
    Dim oSupport3 As Object
    Dim oSupport4 As Object

    If Not oSupportsColl Is Nothing Then
        If oSupportsColl.Count > 3 Then
            Set oSupport3 = oSupportsColl.Item(3)
            Set oSupport4 = oSupportsColl.Item(4)
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
        
    'Check support  3 and 4 is plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    Dim oProfileSup3 As StructDetailObjects.ProfileSystem
    Dim oProfileSup4 As StructDetailObjects.ProfileSystem
    
    Select Case pHelper.ObjectType(oSupport3)
        Case SDOBJECT_STIFFENERSYSTEM
            Select Case pHelper.ObjectType(oSupport4)
                Case SDOBJECT_STIFFENERSYSTEM
                    pSL.Add "BKT_4SCLT_OT"
            End Select
            
    End Select
    
    ' Clean up...
    Set oSupportsColl = Nothing
    Set oSupport3 = Nothing
    Set oSupport4 = Nothing
    Set oBracketByPlaneSO = Nothing
    Set oBracketUtils = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub

  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **



' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

