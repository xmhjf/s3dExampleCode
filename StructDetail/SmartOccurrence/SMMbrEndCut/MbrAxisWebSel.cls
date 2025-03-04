VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : MbrAxisWebSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Web Center/Standard Cut case
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    22/Aug/2011 - svsmylav
'           B-25298(CR-179926): corrected logic that selects web-cut items for standard
'           member bounded to tube case.
'    25/Aug/2011 - pnalugol
'           Made changes to account for Inset Brace
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    15Sep2011  -   pnalugol
'               Modified/Added methods to support Braces on Flange penetrated cases
'    16Sep2011  -   pnalugol
'    Modified/Added methods to support Flush Braces
'    11/Jul/2012 - svsmylav
'            TR-206472(DM-216590): Updated selector logic to use new smart items
'            that handle border cases of 'To-Center' - these are used to get Top/Bottom
'            WebCut points correctly for other than border cases as well.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'   03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
'*********************************************************************************************

Const m_sClassName As String = "MbrAxisWebSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

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
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
        
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)

    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Unknown Error"
    
    ' ------------------------------
    ' Get bounded and bounding ports
    ' ------------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Set oBoundedPort = oSL.InputObject(INPUT_BOUNDED)
    Set oBoundingPort = oSL.InputObject(INPUT_BOUNDING)
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    
    ' ----------
    ' Get the AC
    ' ----------
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj
    
    ' --------------------
    ' Determine if tubular
    ' --------------------
    bIsBoundedTube = IsTubularMember(oBoundedPort.Connectable)
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
        
    ' --------------------------
    ' If it is a free end cut...
    ' --------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oSL.SmartOccurrence
        
    Dim bFreeEndCut As Boolean
    bFreeEndCut = oSDO_WebCut.IsFreeEndCut
    
    If bFreeEndCut Then
        ' ------------------------------------------------
        ' If it is a tube, there is only one cut available
        ' ------------------------------------------------
        If bIsBoundedTube Then
            oSL.Add "MbrTube_FreeEndCut"
            Exit Sub
        End If
        

        ' -------------------------------------------------------------------
        ' Otherwise, pick an angled cut if welded, and a sniped cut otherwise
        ' -------------------------------------------------------------------
        Dim strEndCutTypeAnswer As String
        
        Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj
        
        GetSelectorAnswer oParentObj, "EndCutType", strEndCutTypeAnswer
        
        If LCase(Trim(strEndCutTypeAnswer)) = LCase(Trim("Welded")) Then
            oSL.Add "MbrWeb_FreeAngledCut"
        Else
            Dim oBounded As ISPSMemberPartPrismatic
            Dim eBoundedType As Integer
            Set oBounded = oBoundedPort.Connectable
            eBoundedType = oBounded.MemberType.Type
            
            If eBoundedType = MemberCategoryAndType.LRail Then
                oSL.Add "MbrWeb_FreeSnipeTopExtension"
            End If
            
            oSL.Add "MbrWeb_FreeSnipeTop"
            oSL.Add "MbrWeb_FreeSnipeBottom"
        End If
        
        ' ---------------------------------------
        ' Exit now - done processing free end cut
        ' ---------------------------------------
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------
    ' Get the section alias and determine if web or flange penetrated
    ' ---------------------------------------------------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As New Collection
    Set oEdgeMapColl = GetEdgeMap(oSL.SmartOccurrence, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    
    ' --------------------------------------------------------------------------------
    ' If both bounded and bounding are tubes, there is only one item (select and exit)
    ' --------------------------------------------------------------------------------
    If bIsBoundedTube And bIsBoundingTube Then
        oSL.Add "Axis_Tube_00"
        Exit Sub
    ' ---------------------------------
    ' If only the bounding is a tube...
    ' ---------------------------------
    ElseIf bIsBoundingTube And (Not bIsBoundedTube) Then
        ' -----------------------------------------------
        ' If outside-and-outside or face-and-outside case
        ' -----------------------------------------------
        Dim bIsFaceAndOutSideNoEdge As Boolean
        Dim bIsOutsideAndOutsideNoEdge As Boolean
        bIsFaceAndOutSideNoEdge = IIf(StrComp(sParentItemName, gsMbrAxisToFaceAndOutSideNoEdge, vbTextCompare) = 0, True, False)
        bIsOutsideAndOutsideNoEdge = IIf(StrComp(sParentItemName, gsMbrAxisToOutSideAndOutSideNoEdge, vbTextCompare) = 0, True, False)
                                    
        If bIsFaceAndOutSideNoEdge Or bIsOutsideAndOutsideNoEdge Then
            ' -----------------
            ' If web-penetrated
            ' -----------------

            If bPenetratesWeb Then
                ' --------------------------------------------------------------------
                ' Determined distance from bounded top and bottom to the bounding tube
                ' --------------------------------------------------------------------
                Dim oSDO_MbrBounded As StructDetailObjects.MemberPart
                Dim oSDO_StiffBounded As StructDetailObjects.ProfilePart
                Dim dBoundedFlangeThickness As Double
                
                If TypeOf oSDO_WebCut.Bounded Is ISPSMemberPartCommon Then
                    Set oSDO_MbrBounded = New MemberPart
                    Set oSDO_MbrBounded.object = oSDO_WebCut.Bounded
                    dBoundedFlangeThickness = oSDO_MbrBounded.flangeThickness
                ElseIf TypeOf oSDO_WebCut.Bounded Is IJProfile Then
                    Set oSDO_StiffBounded = New ProfilePart
                    Set oSDO_StiffBounded.object = oSDO_WebCut.Bounded
                    dBoundedFlangeThickness = oSDO_StiffBounded.flangeThickness
                Else
                    'Error Case!!!!. Unknown Bounded Type.
                    'Need to yet support such cases.
                End If
                    
                Dim dBoundedTopToTube As Double
                Dim dBoundedBottomToTube As Double
                
                dBoundedTopToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_TOP)
                dBoundedBottomToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_BOTTOM)
                
                ' ----------------------------------------------
                ' If at least 5mm of the web intersects the tube
                ' ----------------------------------------------
                Dim dTopClearance As Double
                Dim dBtmClearance As Double
                dTopClearance = dBoundedTopToTube - dBoundedFlangeThickness
                dBtmClearance = dBoundedBottomToTube - dBoundedFlangeThickness
                If bIsOutsideAndOutsideNoEdge Then
                    'Outside and Outside Case
                    If (dTopClearance > 0.005 And dBtmClearance > 0.005) Then
                        oSL.Add "Axis_Web_Tube"
                    Else
                        oSL.Add "Axis_Tube_OSAndOS" 'Straight
                    End If
                ElseIf bIsFaceAndOutSideNoEdge Then
                    'Face and Outside Case
                    If Sgn(dTopClearance) = 1 And dTopClearance < 0.005 Then
                        oSL.Add "Axis_Tube_FCAndOS_TOP"
                    ElseIf Sgn(dBtmClearance) = 1 And dBtmClearance < 0.005 Then
                        oSL.Add "Axis_Tube_FCAndOS_BTM"
                    Else
                        oSL.Add "Axis_Web_Tube"
                    End If
                End If
                    
                    ' -------------------------------------------------------------
                    ' Pick a cut that will remove follow the tube shape and no more
                    ' An outside cut will remove the remaining material
                    ' -------------------------------------------------------------
                    oSL.Add "Axis_Web_Tube"

            ' --------------------
            ' If flange-penetrated
            ' --------------------
            Else
                ' --------------------------------------------------------------------
                ' Determined distance from bounded top and bottom to the bounding tube
                ' --------------------------------------------------------------------
                Dim dBoundedLeftToTube As Double
                Dim dBoundedRightToTube As Double
                
                dBoundedLeftToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_WEB_LEFT)
                dBoundedRightToTube = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT)
                
                ' -------------------------
                ' If there is a left flange
                ' -------------------------
                Dim bTFL As Boolean
                Dim bTFR As Boolean
                Dim bBFL As Boolean
                Dim bBFR As Boolean
                CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

                If (bTFL) Or (bBFL) Then
                    ' ----------------------------------------------------------
                    ' If orthogonal and web or flange overlaps by at least 0.1mm
                    ' ----------------------------------------------------------
                    Dim lStatus As Long
                    Dim oBoundedData As MemberConnectionData
                    Dim oBoundingData As MemberConnectionData
                    Dim bIsOrthogonal As Boolean
                    
                    InitMemberConnectionData oParentObj, oBoundedData, oBoundingData, lStatus, sMsg
                                         
                    bIsOrthogonal = IsBoundingOrthogonalToBoundedFlg(oBoundingData, oBoundedData, JXSEC_TOP)
                    
                    
                    If bIsOrthogonal And dBoundedLeftToTube < 0.0001 And dBoundedRightToTube < 0.0001 Then
                        ' -------------------------------------------------------------
                        ' Pick a cut that will remove follow the tube shape and no more
                        ' An outside cut will remove the remaining material
                        ' -------------------------------------------------------------
                        oSL.Add "Axis_Web_Tube"
                    ' ----------------------------------------------------------------------------------
                    ' Otherwise pick a cut that follows the tube shape and then extends straight outside
                    ' ----------------------------------------------------------------------------------
                    Else
                        oSL.Add "Axis_Tube_00"
                    End If
                ' --------------------------------------------------------------------------------------------
                ' If not left flange, pick a cut that follows the tube shape and then extends straight outside
                ' --------------------------------------------------------------------------------------------
                Else
                    oSL.Add "Axis_Tube_00"
                End If
            End If
        ' -----------------------------------------------------------------------------------------------------------------------------
        ' If not outside-and-outside or face-and-outside case, pick a cut that follows the tube shape and then extends straight outside
        ' -----------------------------------------------------------------------------------------------------------------------------
        Else
            'ToCenter Case, or OnMember Case
            'Identify border Web-Penetrated cases and choose appropriate smart item
            Dim sSelector As String
            Dim iACFlag As Integer

            GetBoundingCaseForTube oSDO_WebCut.Bounded, oSDO_WebCut.Bounding, bPenetratesWeb, iACFlag
            
            Select Case iACFlag
             Case BorderAC_OSOS
                oSL.Add "Axis_Tube_OSAndOS"
             Case BorderAC_FCOS_TOP
                oSL.Add "Axis_Tube_FCAndOS_TOP"
             Case BorderAC_FCOS_BTM
                oSL.Add "Axis_Tube_FCAndOS_BTM"
             Case BorderAC_ToCenter, 0 'Zero => Intersections of curves confirm that it is ToCenter Case
                oSL.Add "Axis_Tube_ToCenter"
             Case Else
                If bPenetratesWeb And StrComp(sParentItemName, "MbrAxis_ToOnTubeMember", vbTextCompare) = 0 Then
                    'Handle Web-Penetrated and also On-TubeMember Case
                    oSL.Add "Axis_Tube_00"
                Else
                    'Need to handle flange penetrated ToCenter case
                    oSL.Add "Axis_Web_Tube"
                End If
            End Select
        End If
        
        ' --------------------------------------------------
        ' Exit - done processing open member bounded to tube
        ' --------------------------------------------------
        Exit Sub
    
    End If
    
    ' ---------------------------------------
    ' If non-tube to non-tube, web-penetrated
    ' ---------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrWL As ConnectedEdgeInfo
    Dim eIBFOrWR As ConnectedEdgeInfo
            
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, eTopOrWL, eBtmOrWR, eITFOrWL, eIBFOrWR
        
    If bPenetratesWeb Then
        
        ' -----------------------------------------------------------------------
        ' Pick center cut based on section alias and where bounded meets bounding
        ' -----------------------------------------------------------------------
        Select Case lSectionAlias
            
            Case 0, 8, 9, 10, 11, 12, 17
                'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
                'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
                   oSL.Add "Axis_Web_F" 'Add FB case
            Case 1, 6
                'WebTopFlangeRight, WebTopFlange
                 If eTopOrWL.IntersectingEdge = Web_Right Then
                   oSL.Add "Axis_Web_F" 'Add FB case
                 Else
                   oSL.Add "Axis_Web_TE_F" 'Add UA Case
                 End If
            Case 2
                'WebBuiltUpTopFlangeRight
                 If eTopOrWL.IntersectingEdge = Web_Right Then
                   oSL.Add "Axis_Web_F" 'Add FB case
                 ElseIf eTopOrWL.IntersectingEdge = Top_Flange_Right Then
                   oSL.Add "Axis_Web_TE_F" 'Add UA case
                 Else
                   oSL.Add "Axis_Web_TF_TE_F" 'Add BUTop
                 End If
            Case 3, 7
                 'WebBottomFlangeRight, WebBottomFlange
                 If eBtmOrWR.IntersectingEdge = Web_Right Then
                   oSL.Add "Axis_Web_F" 'Add FB case
                 Else
                   oSL.Add "Axis_Web_F_BE" 'Add L Case
                 End If
            Case 4
                'WebBuiltUpBottomFlangeRight
                 If eBtmOrWR.IntersectingEdge = Web_Right Then
                   oSL.Add "Axis_Web_F" 'Add FB case
                 ElseIf eBtmOrWR.IntersectingEdge = Bottom_Flange_Right Then
                   oSL.Add "Axis_Web_F_BE" 'Add L case
                 Else
                   oSL.Add "Axis_Web_F_BE_BF" 'Add BUBottom
                 End If
            Case 5, 13, 19
                'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
                If eTopOrWL.IntersectingEdge = Web_Right Then
                   If eBtmOrWR.IntersectingEdge = Web_Right Then
                      oSL.Add "Axis_Web_F" 'Add FB case
                   Else
                      oSL.Add "Axis_Web_F_BE" 'Add L case
                   End If
                ElseIf eTopOrWL.IntersectingEdge = Top_Flange_Right Or _
                    eTopOrWL.IntersectingEdge = Above Then
                   If eBtmOrWR.IntersectingEdge = Web_Right Then
                      oSL.Add "Axis_Web_TE_F" 'Add UA case
                   Else
                      oSL.Add "Axis_Web_TE_F_BE" 'Add I case
                   End If
                Else
                   oSL.Add "Axis_Web_TE_F_BE" 'Add I case
                End If
            Case 14
                'FlangeLeftAndRightBottomWebs
                oSL.Add "Axis_WebB_F"
            Case 15
                'FlangeLeftAndRightTopWebseactype
                oSL.Add "Axis_WebT_F"
            Case 16, 18
                'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
                oSL.Add "Axis_WebTB_F"
            Case 20
                'Tube/Circular Cross Section
                oSL.Add ""
            Case Else
                'Unknown Section Alias
    
        End Select
    ' ------------------------------------------
    ' If non-tube to non-tube, flange-penetrated
    ' ------------------------------------------
    Else
        ' ------------------------------------------
        ' If this item is a brace for another member
        ' ------------------------------------------
        Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj
        
        Dim bIsConnBTWBraceAndBounded As Boolean
        Dim bIsConnBTWBraceAndBounding As Boolean
        
        If TypeOf oParentObj Is StructAssemblyConnection Then
            IsCutOnInsetBrace oParentObj, bIsConnBTWBraceAndBounded, bIsConnBTWBraceAndBounding
        End If
        
        If bIsConnBTWBraceAndBounded Then
            
            ' ------------------
            ' Get the brace type
            ' ------------------
            Dim bIsTopBrace As Boolean
            Dim sBraceType As String
            Dim oStructPort As IJStructPort
            Set oStructPort = oBoundingPort
            
            If oStructPort.SectionID = 514 Then
                bIsTopBrace = True
            ElseIf oStructPort.SectionID = 513 Then
                bIsTopBrace = False
            End If
            
            Dim oParentAC As Object
            Dim oBraceChild As IJDesignChild
            Set oBraceChild = oSDO_WebCut.Bounded
            Set oParentAC = oBraceChild.GetParent
            
            If bIsTopBrace Then
                GetSelectorAnswer oParentAC, "TopBraceType", sBraceType
            Else
                GetSelectorAnswer oParentAC, "BottomBraceType", sBraceType
            End If
            
            ' --------------------------------------------------
            ' If it is an inset member, there is only one option
            ' --------------------------------------------------
            If sBraceType = "InsetMember" Then
                oSL.Add "Axis_Web_Cope_Ext"
                Exit Sub
            End If
        End If
        
        ' ----------------------------------------------------
        ' Ask if the web should be shaped (e.g. sniped, coped)
        ' ----------------------------------------------------
        Dim strAnswer As String
        GetSelectorAnswer oParentObj, "TopFlangeOr(Left)Web", strAnswer
               
        ' ----------------------------------------------------
        ' If no special cases handled above, add standard cuts
        ' ----------------------------------------------------
'        If StrComp(strAnswer, "Sniped", vbTextCompare) = 0 Then
'           oSL.Add "" 'Add Web Cut which is sniped at one End
'        Else
           oSL.Add "Axis_Web_Straight"
           oSL.Add "Axis_Web_Flush"
'        End If
    
    End If

    Exit Sub
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
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
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
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
