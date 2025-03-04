VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootEdgeSel"
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

Const m_SelectorProgid As String = "EdgeFeatureRules.RootEdgeSel"
Const m_SelectorName As String = "EdgeFeatureRules.RootEdgeSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EdgeFeatureRules\RootEdgeSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_EDGE
    pIH.SetInput INPUT_POINT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    pQH.SetQuestion gsPlacement, "Manual", "PlacementCol"   ' NOTE: This default is only used if not created by Assembly connection
    pQH.SetQuestion gsOffEdge, "No", "BooleanCol"
    pQH.SetQuestion gsDrainage, "Yes", "BooleanCol"
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub
Private Sub EdgePlacement(pSL As IJDSelectorLogic)
On Error GoTo ErrorHandler
'  ------------------------------------------------------------
' This selection rule selects a edge feature based on whether it has been
' placed manually or by an assembly connection

'Get answers
Dim strPlacement As String
strPlacement = pSL.Answer(gsPlacement)
    
Select Case strPlacement
    Case "Manual"   'Always have all edge feature items available for manual placement
       
   
        pSL.Add "CutEdge"
        pSL.Add "OffEdge"
'        pSL.Add "Scallop R"
'        pSL.Add "Scallop 40"
'        pSL.Add "WideScallop HxWxR"
'        pSL.Add "WideScallop 140x200x50"
'        pSL.Add "WideScallop 150x250x60"
'        pSL.Add "FlatOval DxLxWxOxA"
'        pSL.Add "FlatOval 40x100x80x30x0"
'        pSL.Add "Bean HxRxO"
'        pSL.Add "Bean 50x15x10"
'        pSL.Add "Bean 50x15x10"
'        pSL.Add "Bean 60x20x10"
'        pSL.Add "Bean 80x25x15"
'        pSL.Add "Round DxO"
'        pSL.Add "Round 140x100"
'        pSL.Add "Round 150x250"
'        pSL.Add "Grave HxRxO"
'        pSL.Add "Grave 100x25x25"
'        pSL.Add "Grave 100x30x25"
'        pSL.Add "Knob WxHxLxGxE"
'        pSL.Add "Knob 40x70x100x25x25"
   

    Case "AssyConn"
        'No Corner Feature
End Select

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "EdgePlacement").Number
End Sub
Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
        
    Dim oSO As IJSmartOccurrence
    Set oSO = pSL.SmartOccurrence
    
    Dim oEdgeFeatureWrapper As EdgeFeature
    Dim eFeatureType As StructFeatureTypes
    
    Set oEdgeFeatureWrapper = New EdgeFeature
    Set oEdgeFeatureWrapper.object = oSO
    eFeatureType = oEdgeFeatureWrapper.EdgeFeatureType
    Set oSO = Nothing
    Set oEdgeFeatureWrapper = Nothing
    
    If eFeatureType = SF_WaterStop Then
       pSL.Add ("WaterStop")
       Exit Sub
    End If
        
    'Get answers
    
    Dim strOffEdge As String
    strOffEdge = pSL.Answer(gsOffEdge)
    
    Dim strPlacement As String
    strPlacement = pSL.Answer(gsPlacement)
    
     Select Case strPlacement
        Case "Manual"
            Call EdgePlacement(pSL)
        Case "AssyConn"
            Select Case strOffEdge
                Case "Yes"
                    pSL.Add "OffEdge"
                Case "No"
                    pSL.Add "CutEdge"
            End Select
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

