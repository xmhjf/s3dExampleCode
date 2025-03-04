VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
' Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
' Project     : SMMbrACStd
' File        : FlangeCutSel.cls
'
' Description :
'
' Author      : Alligators
'
'*********************************************************************************************

Const m_sClassName As String = "FlangeCutSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    pQH.SetQuestion "BottomFlange", "No", "BooleanCol"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate centre cut
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)

    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
       
    ' ------------------
    ' Get the flange cut
    ' ------------------
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = pSelectorLogic.SmartOccurrence
    
    ' --------------------
    ' Get the bounded port
    ' --------------------
    sMsg = "Set the Bounded input"
    
    Dim oBoundedPort As IJPort
    
    Set oBoundedPort = oSDO_FlangeCut.BoundedPort
    If (oBoundedPort Is Nothing) Then
        sMsg = "Bounded Port is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding input"
    
    ' ---------------------
    ' Get the bounding port
    ' ---------------------
    Dim oBoundingPort As IJPort
    Set oBoundingPort = oSDO_FlangeCut.BoundingPort
    If (oBoundingPort Is Nothing) Then
        sMsg = "Bounding Port is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    ' ---------------------
    ' Get the input web cut
    ' ---------------------
    Dim oWebCut As New StructDetailObjects.WebCut
    Set oWebCut.object = oSDO_FlangeCut.WebCut
    
    ' --------------------------------------
    ' Get the Edge Mapping and Section Alias
    ' --------------------------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Set oEdgeMapColl = New Collection

    Set oEdgeMapColl = GetEdgeMap(pSelectorLogic.SmartOccurrence, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    ' ----------------------------
    ' Determine what flanges exist
    ' ----------------------------
    Dim bTopFL As Boolean
    Dim bBtmFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFR As Boolean
    
    CrossSection_Flanges oSDO_FlangeCut.Bounded, bTopFL, bBtmFL, bTopFR, bBtmFR
    
     ' -----------------
    ' Get the parent AC
    ' -----------------
    Dim oParentObj As Object
    Dim sParentItemName As String
    
    AssemblyConnection_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
         
    ' ---------------------------------------------------------------
    ' Determine where the bounded faces intersect the bounding object
    ' ---------------------------------------------------------------
    Dim oAppConnection As IJAppConnection
    Dim lStatus As Long
    
    Set oAppConnection = oParentObj
    
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBtmOrWR As ConnectedEdgeInfo
    Dim cTFIOrFL As ConnectedEdgeInfo
    Dim cBFIOrFR As ConnectedEdgeInfo
    
    Dim cWLorTFL As ConnectedEdgeInfo
    Dim eWRorTFR As ConnectedEdgeInfo
    
    ' --------------------
    ' If flange-penetrated
    ' --------------------
    If Not bPenetratesWeb Then
            
            GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, cTopOrWL, cBtmOrWR, cTFIOrFL, cBFIOrFR
            
            ' -----------------
            ' Pick a center cut
            ' -----------------
            If StrComp(pSelectorLogic.Answer("BottomFlange"), "Yes", vbTextCompare) = 0 Then
                If (bBtmFL) Then
                    cWLorTFL.IntersectingEdge = cTFIOrFL.IntersectingEdge
                    cWLorTFL.CoplanarEdge = cTFIOrFL.CoplanarEdge
                Else
                    cWLorTFL.IntersectingEdge = cTopOrWL.IntersectingEdge
                    cWLorTFL.CoplanarEdge = cTopOrWL.CoplanarEdge
                End If
                
                If (bBtmFR) Then
                    eWRorTFR.IntersectingEdge = cBFIOrFR.IntersectingEdge
                    eWRorTFR.CoplanarEdge = cBFIOrFR.CoplanarEdge
                Else
                    eWRorTFR.IntersectingEdge = cBtmOrWR.IntersectingEdge
                    eWRorTFR.CoplanarEdge = cBtmOrWR.CoplanarEdge
                End If
            Else
                If (bTopFL) Then
                    cWLorTFL.IntersectingEdge = cTFIOrFL.IntersectingEdge
                    cWLorTFL.CoplanarEdge = cTFIOrFL.CoplanarEdge
                Else
                    cWLorTFL.IntersectingEdge = cTopOrWL.IntersectingEdge
                    cWLorTFL.CoplanarEdge = cTopOrWL.CoplanarEdge
                End If
                
                If (bTopFR) Then
                    eWRorTFR.IntersectingEdge = cBFIOrFR.IntersectingEdge
                    eWRorTFR.CoplanarEdge = cBFIOrFR.CoplanarEdge
                Else
                    eWRorTFR.IntersectingEdge = cBtmOrWR.IntersectingEdge
                    eWRorTFR.CoplanarEdge = cBtmOrWR.CoplanarEdge
               End If
            End If
            
            Select Case lSectionAlias
                
                Case 0, 8, 9, 10, 11, 12, 17
                    'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
                    'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
                       pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                Case 1, 6
                    'WebTopFlangeRight, WebTopFlange
                     If cWLorTFL.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                     Else
                       pSelectorLogic.Add "AxisFlangeStd_TE_F" 'Add UA Case
                     End If
                Case 2
                    'WebBuiltUpTopFlangeRight
                     If cWLorTFL.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                     ElseIf cWLorTFL.IntersectingEdge = Top_Flange_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_TE_F" 'Add UA case
                     Else
                       pSelectorLogic.Add "AxisFlangeStd_TF_TE_F" 'Add BUTop
                     End If
                Case 3, 7
                     'WebBottomFlangeRight, WebBottomFlange
                     If eWRorTFR.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                     Else
                       pSelectorLogic.Add "AxisFlangeStd_F_BE" 'Add L Case
                     End If
                Case 4
                    'WebBuiltUpBottomFlangeRight
                     If eWRorTFR.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                     ElseIf eWRorTFR.IntersectingEdge = Bottom_Flange_Right Then
                       pSelectorLogic.Add "AxisFlangeStd_F_BE" 'Add L case
                     Else
                       pSelectorLogic.Add "AxisFlangeStd_F_BE_BF" 'Add BUBottom
                     End If
                Case 5, 13, 19
                    'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
                    If cWLorTFL.IntersectingEdge = Web_Right Then
                       If eWRorTFR.IntersectingEdge = Web_Right Then
                          pSelectorLogic.Add "AxisFlangeStd_F" 'Add FB case
                       Else
                          pSelectorLogic.Add "AxisFlangeStd_F_BE" 'Add L case
                       End If
                    ElseIf cWLorTFL.IntersectingEdge = Top_Flange_Right Or _
                        cWLorTFL.IntersectingEdge = Above Then
                        If eWRorTFR.IntersectingEdge = Web_Right Then
                            pSelectorLogic.Add "AxisFlangeStd_TE_F" 'Add UA case
                        Else
                            pSelectorLogic.Add "AxisFlangeStd_TE_F_BE" 'Add I case
                        End If
                    Else
                        pSelectorLogic.Add "AxisFlangeStd_TE_F_BE" 'Add I case
                    End If
                Case 14
                    'FlangeLeftAndRightBottomWebs
                    pSelectorLogic.Add "AxisFlangeStd_B_F"
                Case 15
                    'FlangeLeftAndRightTopWebs
                    pSelectorLogic.Add "AxisFlangeStd_T_F"
                Case 16, 18
                    'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
                    pSelectorLogic.Add "AxisFlangeStd_TB_F"
                Case 20
                    'Tube/Circular Cross Section
                    pSelectorLogic.Add ""
                Case Else
                    'Unknown Section Alias
                    
            End Select
      
    End If
    
    Exit Sub

ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

