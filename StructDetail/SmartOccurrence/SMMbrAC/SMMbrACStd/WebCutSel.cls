VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  File        : WebCutSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'*********************************************************************************************

Const m_sClassName As String = "WebCutSel"
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
    sMsg = "Defining/Initializing Selector Questions/Answers"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the proper web center cut
'
' Notes: Top and bottom web cut items and non-penetrated web cut items should be specified in the bulkload spreadsheet.
'        There is only one appropriate center web cut per bounding condition, and this selector is
'        be used to choose it.
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)

    Const METHOD = m_SelectorProgid & "::SelectorLogic"

    On Error GoTo ErrorHandler

    Dim sMsg As String

    ' ------------------------------
    ' Get bounded and bounding ports
    ' ------------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Set oBoundedPort = oSL.InputObject(INPUT_BOUNDED)
    Set oBoundingPort = oSL.InputObject(INPUT_BOUNDING)

    ' ---------------------
    ' Get the section alias
    ' ---------------------
    Dim sectionAlias As Long
    Dim bPenetratesWeb As Boolean
    
    GetEdgeMap oSL.SmartOccurrence, oBoundingPort, oBoundedPort, sectionAlias, bPenetratesWeb
    
    
''' The commented out code is from the Q&A version of member endcuts
''' Some portions will be needed later when new cases, such as bouned-by-tube, are added to the standard AC rules
    
'    Dim bIsBoundedTube As Boolean
'    Dim bIsBoundingTube As Boolean
'
'    ' --------------------
'    ' Determine if tubular
'    ' --------------------
'    bIsBoundedTube = IsTubularMember(oBoundedPort.Connectable)
'    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
        
'    ' --------------------------------------------------------------------------------
'    ' If both bounded and bounding are tubes, there is only one item (select and exit)
'    ' --------------------------------------------------------------------------------
'    If bIsBoundedTube And bIsBoundingTube Then
'        oSL.Add "Axis_Tube_00"
'        Exit Sub
'    ' ---------------------------------
'    ' If only the bounding is a tube...
'    ' ---------------------------------
'    ElseIf bIsBoundingTube And (Not bIsBoundedTube) Then
'        ' -----------------------------------------------
'        ' If outside-and-outside or face-and-outside case
'        ' -----------------------------------------------
'        Dim bIsFaceAndOutSideNoEdge As Boolean
'        Dim bIsOutsideAndOutsideNoEdge As Boolean
'        bIsFaceAndOutSideNoEdge = IIf(StrComp(sParentItemName, gsMbrAxisToFaceAndOutSideNoEdge, vbTextCompare) = 0, True, False)
'        bIsOutsideAndOutsideNoEdge = IIf(StrComp(sParentItemName, gsMbrAxisToOutSideAndOutSideNoEdge, vbTextCompare) = 0, True, False)
'
'        If bIsFaceAndOutSideNoEdge Or bIsOutsideAndOutsideNoEdge Then
'            ' -----------------
'            ' If web-penetrated
'            ' -----------------
'
'            If bPenetratesWeb Then
'                ' --------------------------------------------------------------------
'                ' Determined distance from bounded top and bottom to the bounding tube
'                ' --------------------------------------------------------------------
'                Dim oSDO_MbrBounded As StructDetailObjects.MemberPart
'                Dim oSDO_StiffBounded As StructDetailObjects.ProfilePart
'                Dim dBoundedFlangeThickness As Double
'
'                If TypeOf oSDO_WebCut.Bounded Is ISPSMemberPartCommon Then
'                    Set oSDO_MbrBounded = New MemberPart
'                    Set oSDO_MbrBounded.object = oSDO_WebCut.Bounded
'                    dBoundedFlangeThickness = oSDO_MbrBounded.flangeThickness
'                ElseIf TypeOf oSDO_WebCut.Bounded Is IJProfile Then
'                    Set oSDO_StiffBounded = New ProfilePart
'                    Set oSDO_StiffBounded.object = oSDO_WebCut.Bounded
'                    dBoundedFlangeThickness = oSDO_StiffBounded.flangeThickness
'                Else
'                    'Error Case!!!!. Unknown Bounded Type.
'                    'Need to yet support such cases.
'                End If
'
'                Dim dBoundedTopToTube As Double
'                Dim dBoundedBottomToTube As Double
'
'                dBoundedTopToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_TOP)
'                dBoundedBottomToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_BOTTOM)
'
'                ' ----------------------------------------------
'                ' If at least 5mm of the web intersects the tube
'                ' ----------------------------------------------
'                Dim dTopClearance As Double
'                Dim dBtmClearance As Double
'                dTopClearance = dBoundedTopToTube - dBoundedFlangeThickness
'                dBtmClearance = dBoundedBottomToTube - dBoundedFlangeThickness
'                If bIsOutsideAndOutsideNoEdge Then
'                    'Outside and Outside Case
'                    If (dTopClearance > 0.005 And dBtmClearance > 0.005) Then
'                        oSL.Add "Axis_Web_Tube"
'                    Else
'                        oSL.Add "Axis_Tube_OSAndOS" 'Straight
'                    End If
'                ElseIf bIsFaceAndOutSideNoEdge Then
'                    'Face and Outside Case
'                    If Sgn(dTopClearance) = 1 And dTopClearance < 0.005 Then
'                        oSL.Add "Axis_Tube_FCAndOS_TOP"
'                    ElseIf Sgn(dBtmClearance) = 1 And dBtmClearance < 0.005 Then
'                        oSL.Add "Axis_Tube_FCAndOS_BTM"
'                    Else
'                        oSL.Add "Axis_Web_Tube"
'                    End If
'                End If
'
'                    ' -------------------------------------------------------------
'                    ' Pick a cut that will remove follow the tube shape and no more
'                    ' An outside cut will remove the remaining material
'                    ' -------------------------------------------------------------
'                    oSL.Add "Axis_Web_Tube"
'
'            ' --------------------
'            ' If flange-penetrated
'            ' --------------------
'            Else
'                ' --------------------------------------------------------------------
'                ' Determined distance from bounded top and bottom to the bounding tube
'                ' --------------------------------------------------------------------
'                Dim dBoundedLeftToTube As Double
'                Dim dBoundedRightToTube As Double
'
'                dBoundedLeftToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_WEB_LEFT)
'                dBoundedRightToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT)
'
'                ' -------------------------
'                ' If there is a left flange
'                ' -------------------------
'                Dim bTFL As Boolean
'                Dim bTFR As Boolean
'                Dim bBFL As Boolean
'                Dim bBFR As Boolean
'                CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
'
'                If (bTFL) Or (bBFL) Then
'                    ' ----------------------------------------------------------
'                    ' If orthogonal and web or flange overlaps by at least 0.1mm
'                    ' ----------------------------------------------------------
'                    Dim lStatus As Long
'                    Dim oBoundedData As MemberConnectionData
'                    Dim oBoundingData As MemberConnectionData
'                    Dim bIsOrthogonal As Boolean
'
'                    InitMemberConnectionData oParentObj, oBoundedData, oBoundingData, lStatus, sMsg
'
'                    bIsOrthogonal = IsBoundingOrthogonalToBoundedFlg(oBoundingData, oBoundedData, JXSEC_TOP)
'
'
'                    If bIsOrthogonal And dBoundedLeftToTube < 0.0001 And dBoundedRightToTube < 0.0001 Then
'                        ' -------------------------------------------------------------
'                        ' Pick a cut that will remove follow the tube shape and no more
'                        ' An outside cut will remove the remaining material
'                        ' -------------------------------------------------------------
'                        oSL.Add "Axis_Web_Tube"
'                    ' ----------------------------------------------------------------------------------
'                    ' Otherwise pick a cut that follows the tube shape and then extends straight outside
'                    ' ----------------------------------------------------------------------------------
'                    Else
'                        oSL.Add "Axis_Tube_00"
'                    End If
'                ' --------------------------------------------------------------------------------------------
'                ' If not left flange, pick a cut that follows the tube shape and then extends straight outside
'                ' --------------------------------------------------------------------------------------------
'                Else
'                    oSL.Add "Axis_Tube_00"
'                End If
'            End If
'        ' -----------------------------------------------------------------------------------------------------------------------------
'        ' If not outside-and-outside or face-and-outside case, pick a cut that follows the tube shape and then extends straight outside
'        ' -----------------------------------------------------------------------------------------------------------------------------
'        Else
'            'ToCenter Case, or OnMember Case
'            'Identify border Web-Penetrated cases and choose appropriate smart item
'            Dim sSelector As String
'            Dim iACFlag As Integer
'
'            GetBoundingCaseForTube oSDO_WebCut.Bounded, oSDO_WebCut.Bounding, bPenetratesWeb, iACFlag
'
'            Select Case iACFlag
'             Case BorderAC_OSOS
'                oSL.Add "Axis_Tube_OSAndOS"
'             Case BorderAC_FCOS_TOP
'                oSL.Add "Axis_Tube_FCAndOS_TOP"
'             Case BorderAC_FCOS_BTM
'                oSL.Add "Axis_Tube_FCAndOS_BTM"
'             Case BorderAC_ToCenter, 0 'Zero => Intersections of curves confirm that it is ToCenter Case
'                oSL.Add "Axis_Tube_ToCenter"
'             Case Else
'                If bPenetratesWeb And StrComp(sParentItemName, "MbrAxis_ToOnTubeMember", vbTextCompare) = 0 Then
'                    'Handle Web-Penetrated and also On-TubeMember Case
'                    oSL.Add "Axis_Tube_00"
'                Else
'                    'Need to handle flange penetrated ToCenter case
'                    oSL.Add "Axis_Web_Tube"
'                End If
'            End Select
'        End If
'
'        ' --------------------------------------------------
'        ' Exit - done processing open member bounded to tube
'        ' --------------------------------------------------
'        Exit Sub
'
'    End If
    
    ' ---------------------------------------
    ' If non-tube to non-tube, web-penetrated
    ' ---------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrWL As ConnectedEdgeInfo
    Dim eIBFOrWR As ConnectedEdgeInfo
            
    GetConnectedEdgeInfo oSL.SmartOccurrence, oBoundedPort, oBoundingPort, eTopOrWL, eBtmOrWR, eITFOrWL, eIBFOrWR
        
    If bPenetratesWeb Then
        
        ' -----------------------------------------------------------------------
        ' Pick center cut based on section alias and where bounded meets bounding
        ' -----------------------------------------------------------------------
        Select Case sectionAlias
            
            Case 0, 8, 9, 10, 11, 12, 17
                'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
                'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
                   oSL.Add "AxisWebStd_F" 'Add FB case
            Case 1, 6
                'WebTopFlangeRight, WebTopFlange
                 If eTopOrWL.IntersectingEdge = Web_Right Then
                   oSL.Add "AxisWebStd_F" 'Add FB case
                 Else
                   oSL.Add "AxisWebStd_TE_F" 'Add UA Case
                 End If
            Case 2
                'WebBuiltUpTopFlangeRight
                 If eTopOrWL.IntersectingEdge = Web_Right Then
                   oSL.Add "AxisWebStd_F" 'Add FB case
                 ElseIf eTopOrWL.IntersectingEdge = Top_Flange_Right Then
                   oSL.Add "AxisWebStd_TE_F" 'Add UA case
                 Else
                   oSL.Add "AxisWebStd_TF_TE_F" 'Add BUTop
                 End If
            Case 3, 7
                 'WebBottomFlangeRight, WebBottomFlange
                 If eBtmOrWR.IntersectingEdge = Web_Right Then
                   oSL.Add "AxisWebStd_F" 'Add FB case
                 Else
                   oSL.Add "AxisWebStd_F_BE" 'Add L Case
                 End If
            Case 4
                'WebBuiltUpBottomFlangeRight
                 If eBtmOrWR.IntersectingEdge = Web_Right Then
                   oSL.Add "AxisWebStd_F" 'Add FB case
                 ElseIf eBtmOrWR.IntersectingEdge = Bottom_Flange_Right Then
                   oSL.Add "AxisWebStd_F_BE" 'Add L case
                 Else
                   oSL.Add "AxisWebStd_F_BE_BF" 'Add BUBottom
                 End If
            Case 5, 13, 19
                'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
                If eTopOrWL.IntersectingEdge = Web_Right Then
                   If eBtmOrWR.IntersectingEdge = Web_Right Then
                      oSL.Add "AxisWebStd_F" 'Add FB case
                   Else
                      oSL.Add "AxisWebStd_F_BE" 'Add L case
                   End If
                ElseIf eTopOrWL.IntersectingEdge = Top_Flange_Right Or _
                    eTopOrWL.IntersectingEdge = Above Then
                   If eBtmOrWR.IntersectingEdge = Web_Right Then
                      oSL.Add "AxisWebStd_TE_F" 'Add UA case
                   Else
                      oSL.Add "AxisWebStd_TE_F_BE" 'Add I case
                   End If
                Else
                   oSL.Add "AxisWebStd_TE_F_BE" 'Add I case
                End If
            Case 14
                'FlangeLeftAndRightBottomWebs
                oSL.Add "AxisWebStd_B_F"
            Case 15
                'FlangeLeftAndRightTopWebseactype
                oSL.Add "AxisWebStd_T_F"
            Case 16, 18
                'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
                oSL.Add "AxisWebStd_TB_F"
            Case 20
                'Tube/Circular Cross Section
                oSL.Add ""
            Case Else
                'Unknown Section Alias
    
        End Select
    
    End If
    
    Exit Sub

ErrorHandler:
    oSL.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************

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
    
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
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


