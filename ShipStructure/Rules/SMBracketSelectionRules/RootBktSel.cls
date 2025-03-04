' DI-CP-235026  MF: Error in Partclass View creation during bulkloading on Oracle DB  
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootBktSel"
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

Const m_SelectorProgid As String = "BktSelRules.RootBktSel"
Const m_SelectorName As String = "BktSelRules.RootBktSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktSelRules\RootBktSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    
    pIH.SetInput INPUT_SUPPORTS
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    'get the number of supports selected
    Dim nSupports As Long
    Dim oSupports As IJElements
    
    ' Get Bracket-by-Plane Smart Occurrence:
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = pSL.SmartOccurrence
    
    ' Create helper for final construction of bracket:
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
   
    
    
    Dim oBracketPlateSys As IJBracketPlateSystem
    If TypeOf oBracketByPlaneSO Is IJBracketPlateSystem Then
        Set oBracketPlateSys = oBracketByPlaneSO
        nSupports = oBracketPlateSys.NumberOfSupports
    Else
        'Get the number of supports from BracketContourSO.
        oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupports, nSupports
        ' The nSupports include the leaf supports also, take only root supports count.
        nSupports = nSupports / 2
    End If
    Set oBracketByPlaneSO = Nothing
    
    
        
' ********************* Selection code - TODO  ****************
    'select class based on number of supports
    Select Case nSupports
        Case 1
            'This class should implement the param and def rules
            pSL.Add "BPS_1S_Trimmed"
        Case 2
            pSL.Add "BPS_2S_Trimmed"
        Case 3
            pSL.Add "BPS_3S_Trimmed"
        Case 4
            'This class should implement the param and def rules
            pSL.Add "BPS_4S_Trimmed"
        Case 5
            'This class should implement the param and def rules Add5SBracketSelectorClass pSL
            pSL.Add "BPS_5S_Trimmed"
    End Select

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
    pIH.definition = pSelector
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
Private Sub Add5SBracketSelectorClass(pSL As IJDSelectorLogic)
    On Error Resume Next
    
    Dim eBracketSupport1Type As ShpStrBktSupportConnectionType
    Dim eBracketSupport2Type As ShpStrBktSupportConnectionType
    Dim eBracketSupport3Type As ShpStrBktSupportConnectionType
    Dim eBracketSupport4Type As ShpStrBktSupportConnectionType
    Dim eBracketSupport5Type As ShpStrBktSupportConnectionType

    On Error Resume Next
    eBracketSupport1Type = ConnType_None
    eBracketSupport2Type = ConnType_None
    eBracketSupport3Type = ConnType_None
    eBracketSupport4Type = ConnType_None
    eBracketSupport5Type = ConnType_None
    
    eBracketSupport1Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 1)
    eBracketSupport2Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 2)
    eBracketSupport3Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 3)
    eBracketSupport4Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 4)
    eBracketSupport5Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 5)
    
    
    'Add the class (sheet) name
    If (eBracketSupport1Type = Trimmed) And _
        (eBracketSupport2Type = Trimmed) And _
        (eBracketSupport3Type = Trimmed) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_Trimmed"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Trimmed) And _
        (eBracketSupport3Type = Trimmed) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapSup1"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Lapped) And _
        (eBracketSupport3Type = Trimmed) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapSup1Sup2"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Trimmed) And _
        (eBracketSupport3Type = Lapped) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapSup1Sup3"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Trimmed) And _
        (eBracketSupport3Type = Trimmed) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapSup1"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Lapped) And _
        (eBracketSupport3Type = Lapped) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapSup1Sup2Sup3"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Lapped) And _
        (eBracketSupport3Type = Lapped) And _
        (eBracketSupport4Type = Lapped) And _
        (eBracketSupport5Type = Trimmed) Then
        pSL.Add "BPS_5S_LapS1S2S3S4"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Lapped) And _
        (eBracketSupport3Type = Lapped) And _
        (eBracketSupport4Type = Trimmed) And _
        (eBracketSupport5Type = Lapped) Then
        pSL.Add "BPS_5S_LapS1S2S3S5"
        
    ElseIf (eBracketSupport1Type = Lapped) And _
        (eBracketSupport2Type = Lapped) And _
        (eBracketSupport3Type = Lapped) And _
        (eBracketSupport4Type = Lapped) And _
        (eBracketSupport5Type = Lapped) Then
        pSL.Add "BPS_5S_LapAllSupports"
        
    Else
        pSL.Add "BPS_5S_Root"
    End If
    
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
