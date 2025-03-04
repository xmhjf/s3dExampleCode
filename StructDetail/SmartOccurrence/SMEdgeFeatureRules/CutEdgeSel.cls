VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CutEdgeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMEdgeFeatureRules
'  File        : CutEdgeSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    16/NOV/2012 - CR-CP-222498 - Added New SnipeAtNotch SmartItem and is selected only when Feature is Placed at Notch
'*************************************************************************************************************************
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

Const m_SelectorProgid As String = CUSTOMERID + "EdgeFeatureRules.CutEdgeSel"
Const m_SelectorName As String = CUSTOMERID + "EdgeFeatureRules.CutEdgeSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "EdgeFeatureRules\CutEdgeSel.cls"

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
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler


'        pSL.Add "Scallop R"
'        pSL.Add "Scallop 40"
'        pSL.Add "WideScallop HxWxR"
'        pSL.Add "WideScallop 140x200x50"
'        pSL.Add "WideScallop 150x250x60"
'        pSL.Add "Grave HxRxO"
'        pSL.Add "Grave 100x25x25"
'        pSL.Add "Grave 100x30x25"
'        pSL.Add "Knob WxHxLxGxE"
'        pSL.Add "Knob 40x70x100x25x25"

        ' Half Round R
'

        Dim eFeaturePositionAttr  As SD_FeaturePositionAttr
        Dim bIsRightThicker As Boolean
        
        'Check Feature position is at Seam or Notch
        GetFeaturePositionAttr pSL.SmartOccurrence, eFeaturePositionAttr
        
        'Select Snipe Feature when Feature is Placed at Notch
        If eFeaturePositionAttr = SD_FPA_ChamferedSeam Or eFeaturePositionAttr = SD_FPA_NotchedSeam Then
            pSL.Add "SnipeAtNotch"
        End If
        
        pSL.Add "Half Round R"
        pSL.Add "Half Round 30"
        pSL.Add "Half Round 35"
        pSL.Add "Half Round 50"
        pSL.Add "Half Round 75"
        pSL.Add "Half Round 100"
        pSL.Add "Half Round 150"
        pSL.Add "Half Round along Edge R"
        pSL.Add "Half Round along Edge 30"
        pSL.Add "Half Round along Edge 35"
        pSL.Add "Half Round along Edge 50"
        pSL.Add "Half Round along Edge 75"
        pSL.Add "Half Round along Edge 100"
        pSL.Add "Half Round along Edge 150"
        
        ' Scallop 150xD
        Dim oEdgeFeatureWrapper As New StructDetailObjects.EdgeFeature
        Dim oPartWithFeature As Object
        
        Set oEdgeFeatureWrapper.object = pSL.SmartOccurrence
        Set oPartWithFeature = oEdgeFeatureWrapper.GetPartObject
        
        If TypeOf oPartWithFeature Is IJProfilePart Then
            Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
            Dim dProfilePartHeight As Double
            
            Set oProfilePartWrapper.object = oPartWithFeature
            dProfilePartHeight = oProfilePartWrapper.Height
            dProfilePartHeight = Round(dProfilePartHeight, 3)
            
            If dProfilePartHeight <= 0.15 Then
                pSL.Add "Scallop 150x30"
                pSL.Add "Scallop along Edge 150x30"
            ElseIf dProfilePartHeight > 0.15 And dProfilePartHeight <= 0.18 Then
                pSL.Add "Scallop 150x45"
                pSL.Add "Scallop along Edge 150x45"
            ElseIf dProfilePartHeight < 0.3 Then
                pSL.Add "Scallop 150x50"
                pSL.Add "Scallop along Edge 150x50"
            Else
                pSL.Add "Scallop 150x75"
                pSL.Add "Scallop along Edge 150x75"
            End If
        Else
            pSL.Add "Scallop 150xD"
            pSL.Add "Scallop 150x30"
            pSL.Add "Scallop 150x45"
            pSL.Add "Scallop 150x50"
            pSL.Add "Scallop 150x75"
            pSL.Add "Scallop along Edge 150xD"
            pSL.Add "Scallop along Edge 150x30"
            pSL.Add "Scallop along Edge 150x45"
            pSL.Add "Scallop along Edge 150x50"
            pSL.Add "Scallop along Edge 150x75"
        End If
        
        ' Mushroom
        pSL.Add "Mushroom WxHxG"
        pSL.Add "Mushroom 320x105x200"
        pSL.Add "Mushroom along Edge WxHxG"
        pSL.Add "Mushroom along Edge 320x105x200"
        
        ' Ellipse
        pSL.Add "Ellipse WxHxG"
        pSL.Add "Ellipse 320x100x200"
        pSL.Add "Ellipse along Edge WxHxG"
        pSL.Add "Ellipse along Edge 320x100x200"
        
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

