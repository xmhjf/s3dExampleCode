VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VariableEdgeSel"
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
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "CornerFeatRules.VariableEdgeSel"
Const m_SelectorName As String = CUSTOMERID + "CornerFeatRules.VariableEdgeSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\VariableEdgeSel.cls"
Const PI As Double = 3.14159265

Implements IJDUserSymbolServices

Private Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
     
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Private Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Private Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    ' Get corner feature
    Dim oCornerFeature As IJSDOCornerFeature
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    Set oCornerFeature.object = pSL.SmartOccurrence
    
    Dim oCFChild As IJDesignChild
    Dim oParentSlot As Object
    
    Set oCFChild = oCornerFeature.object
    Set oParentSlot = oCFChild.GetParent
    
    Dim bIsBaseCF As Boolean, bIsInsideCF As Boolean, bIsOuterCF As Boolean
    
    Dim oSDFeatureUtils As IJSDFeatureAttributes
    Set oSDFeatureUtils = New SDFeatureUtils
    Dim oFacePort As Object, oEdgePort1 As Object, oEdgePort2 As Object
    
    oSDFeatureUtils.get_CornerCutInputsEx pSL.SmartOccurrence, oFacePort, oEdgePort1, oEdgePort2
    
    If TypeOf oParentSlot Is IJStructFeature Then
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oParentSlot
        If oStructFeature.get_StructFeatureType = SF_Slot Then
            DetermineCFPositionOnSlot oEdgePort1, oEdgePort2, bIsBaseCF, bIsInsideCF, bIsOuterCF
            If bIsBaseCF Then
                pSL.Add "VariableEdgeLongScallop"
                pSL.Add "VariableEdgeSnipe"
            ElseIf bIsOuterCF Then
                pSL.Add "VariableEdgeMouseEar"
            ElseIf bIsInsideCF Then
                pSL.Add "VariableEdgeSnipe"
                pSL.Add "VariableEdgeLongScallop"
            End If
        End If
        Exit Sub
    Else
        pSL.Add "VariableEdgeLongScallop"
        pSL.Add "VariableEdgeSnipe"
    End If
    
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

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

