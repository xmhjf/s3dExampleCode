VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'
'Copyright (C) 2011-12 Intergraph Corporation. All rights reserved.
'
'File : MbrAxisFlangeSel.cls
'
'Author : Alligators
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Flange Center/Standard Cuts
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   available Selector Items:
'
'History:
'    22/Aug/2011 - svsmylav
'           B-25298(CR-179926): standard member bounded to tube case:
'           modified selector item names for web penetrated case; also added condition to select
'           smart item for flange penetrated orthogonal case.
'    25/Aug/2011 - pnalugol
'           Made changes to account for Inset Brace B-52232
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    15Sep2011  -   pnalugol
'               Modified/Added methods to support Braces on Flange penetrated cases
'    16Sep2011  -   pnalugol
'    Modified/Added methods to support Flush Braces
'    07/Feb/2012 - svsmylav
'           TR-CP-209564: Added 'Axis_Flange_Sniped_LR' and 'Axis_Flange_Sniped_R'
'            sniped-welded new FlangeCut smart items.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'********************************************************************

Const m_sClassName As String = "MbrAxisFlangeSel"
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
' Description : Select the appropriate standard/normal cut
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Unknown Error"
    
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
    
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    ' --------------------------------------------------------------
    ' Determine if this cut is for the top or bottom flange
    ' Determine if the desired shape for the flange (web-penetrated)
    ' --------------------------------------------------------------
    Dim strAnswer As String
    Dim bBottomFlange As Boolean
    
    If StrComp(pSelectorLogic.answer("BottomFlange"), "Yes", vbTextCompare) = 0 Then
        GetSelectorAnswer oParentObj, "BottomFlangeOrRightWeb", strAnswer
        bBottomFlange = True
    Else
       GetSelectorAnswer oParentObj, "TopFlangeOr(Left)Web", strAnswer
       bBottomFlange = False
    End If
    
    ' --------------------------------------------------------
    ' Determine if there are top and bottom "outside" web cuts
    ' --------------------------------------------------------
    'dispid 2/3 creates Top/Btm WebCut, respectively.  ItemByDispid will return  Nothing if the object was not created
    Dim oMemberObjects As IJDMemberObjects
    Dim oCtrWebCut As Object
    Dim oTopWebCut As Object
    Dim oBtmWebCut As Object
    
    Set oMemberObjects = oParentObj
    Set oCtrWebCut = oMemberObjects.ItemByDispid(1)
    Set oTopWebCut = oMemberObjects.ItemByDispid(2)
    Set oBtmWebCut = oMemberObjects.ItemByDispid(3)
    
    ' -------------------------------------------------------------
    ' If a non-tubular member is web-penetrated by a tubular member
    ' -------------------------------------------------------------
    Dim dDistanceFromBounding1 As Double
    Dim dDistanceFromBounding2 As Double
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    
    bIsBoundedTube = IsTubularMember(oBoundedPort)
    bIsBoundingTube = IsTubularMember(oBoundingPort)
    
    If bPenetratesWeb And bIsBoundingTube And (Not bIsBoundedTube) Then
        ' -------------------------------------------------------------------------------
        ' Check how far above or below the bounding object the bounded top and bottom are
        ' -------------------------------------------------------------------------------
        dDistanceFromBounding1 = GetDistanceFromBounding(oSDO_FlangeCut.Bounding, oSDO_FlangeCut.Bounded, JXSEC_TOP)
        dDistanceFromBounding2 = GetDistanceFromBounding(oSDO_FlangeCut.Bounding, oSDO_FlangeCut.Bounded, JXSEC_BOTTOM)
    
        ' -----------------------------
        ' If the flange is to be sniped
        ' -----------------------------
        If StrComp(strAnswer, "Sniped", vbTextCompare) = 0 Then
            ' ----------------------------------------------------------------------------------
            ' If there is an outside web cut or the flange does not intesect the bounding object
            ' ----------------------------------------------------------------------------------
            If (Not bBottomFlange And (Not oTopWebCut Is Nothing Or dDistanceFromBounding1 > 0.0005)) Or _
               (bBottomFlange And (Not oBtmWebCut Is Nothing Or dDistanceFromBounding2 > 0.0005)) Or _
               oCtrWebCut Is Nothing Then
                ' ----------------------------------------
                ' A set back flange cut is the only option
                ' ----------------------------------------
                ' Choose further based on symmetrical or asymmetrical flange
                If (bTopFL And bTopFR) Then
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_Tube_LR"
                Else
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_Tube_R"
                End If
            ' -------------------------------------------------------
            ' Otherwise, default to welded with set back as an option
            ' -------------------------------------------------------
            ' Note: we currently lack an item that is welded and sniped, like the stiffener "Connected" type,
            ' which uses SM_FlangeCutSnipWeld.sym.  That is why the default is commented out.  The result is all
            ' selections are the same, but the logic is left here so that we can easily add it later.
            Else
                ' Choose further based on symmetrical or asymmetrical flange
                If (bTopFL And bTopFR) Then
                    'pSelectorLogic.Add ""
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_Tube_LR"
                Else
                    'pSelectorLogic.Add ""
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_Tube_R"
                End If
            End If
        ' -------------------------------
        ' If the flange is to be straight
        ' -------------------------------
        Else
            ' ----------------------------------------------------------------------------------
            ' If there is an outside web cut or the flange does not intesect the bounding object
            ' ----------------------------------------------------------------------------------
            If (Not bBottomFlange And (Not oTopWebCut Is Nothing Or dDistanceFromBounding1 > 0.0005)) Or _
               (bBottomFlange And (Not oBtmWebCut Is Nothing Or dDistanceFromBounding2 > 0.0005)) Or _
               oCtrWebCut Is Nothing Then
                ' ----------------------------------------
                ' A set back flange cut is the only option
                ' ----------------------------------------
                ' Choose further based on symmetrical or asymmetrical flange
                If (bTopFL And bTopFR) Then
                    pSelectorLogic.Add "Axis_Flange_StraightSetback_Tube_LR"
                Else
                    pSelectorLogic.Add "Axis_Flange_StraightSetback_Tube_R"
                End If
            ' -------------------------------------------------------
            ' Otherwise, default to welded with set back as an option
            ' -------------------------------------------------------
            Else
                ' Choose further based on symmetrical or asymmetrical flange
                If (bTopFL And bTopFR) Then
                    pSelectorLogic.Add "Axis_Flange_Flush_LR"
                    pSelectorLogic.Add "Axis_Flange_StraightSetback_Tube_LR"
                Else
                    pSelectorLogic.Add "Axis_Flange_Flush_R"
                    pSelectorLogic.Add "Axis_Flange_StraightSetback_Tube_R"
                End If
            End If
        End If
        
        Exit Sub '*** Exit ***

    End If
        
    ' -------------------------------------------------------------------------------------------
    ' Create a setback cut, if the parent AC has created an inset brace to this side (top/bottom)
    ' -------------------------------------------------------------------------------------------
    Dim sTopBraceType As String
    Dim sBtmBraceType As String
    
    If bBottomFlange Then
        GetSelectorAnswer oParentObj, "BottomBraceType", sBtmBraceType
        If sBtmBraceType = "InsetMember" Then
            If (bTopFL And bTopFR) Then
                pSelectorLogic.Add "Axis_Flange_StraightSetback_LR"
            Else
                pSelectorLogic.Add "Axis_Flange_StraightSetback_R"
            End If
            
            Exit Sub
        End If
    Else
        GetSelectorAnswer oParentObj, "TopBraceType", sTopBraceType
        If sTopBraceType = "InsetMember" Then
            If (bTopFL And bTopFR) Then
                pSelectorLogic.Add "Axis_Flange_StraightSetback_LR"
            Else
                pSelectorLogic.Add "Axis_Flange_StraightSetback_R"
            End If
            
            Exit Sub
        End If
    End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' If this flange cut is on the brace at the bounded member, place a flange cut that is perpendicular to the bounded
    ' -----------------------------------------------------------------------------------------------------------------
    Dim bIsConnBTWBraceAndBounded As Boolean
    Dim bIsConnBTWBraceAndBounding As Boolean
    
    bIsConnBTWBraceAndBounded = False
    bIsConnBTWBraceAndBounding = False
    
    If TypeOf oParentObj Is StructAssemblyConnection Then
        IsCutOnInsetBrace oParentObj, bIsConnBTWBraceAndBounded, bIsConnBTWBraceAndBounding
    End If
    
    If bIsConnBTWBraceAndBounded Then
            
        Dim bIsTopBrace As Boolean
        Dim sBraceType As String
        bIsTopBrace = False
        
        'see if its a top/bottom brace
        Dim oStructPort As IJStructPort
        Set oStructPort = oWebCut.BoundingPort
        If oStructPort.SectionID = 514 Then 'Top brace
            bIsTopBrace = True
        End If
        
        Dim oParentAC As Object
        Dim oBraceChild As IJDesignChild
        Set oBraceChild = oSDO_FlangeCut.Bounded
        Set oParentAC = oBraceChild.GetParent
        
        If bIsTopBrace Then
            GetSelectorAnswer oParentAC, "TopBraceType", sBraceType
        Else
            GetSelectorAnswer oParentAC, "BottomBraceType", sBraceType
        End If
        
        If sBraceType = "InsetMember" Then 'Cut between InsetBrace and Bounded of Parent AC.
            pSelectorLogic.Add "Axis_Flange_Variable_Bevel"
            Exit Sub
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' Determine where the bounded faces intersect the bounding object
    ' ---------------------------------------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Set oAppConnection = oParentObj
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
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
        ' ---------------------------------------------------------------
        ' If flange is orthogonal to a tubular member, pick a special cut
        ' ---------------------------------------------------------------
        If bIsBoundingTube Then
            Dim eCode As JXSEC_CODE
            
            If bBottomFlange Then
                eCode = JXSEC_BOTTOM
            Else
                eCode = JXSEC_TOP
            End If
            
            If IsBoundingOrthogonalToBoundedFlg(oBoundingData, oBoundedData, eCode) Then
                pSelectorLogic.Add "Axis_Flange_Tube"
            End If
        ' ------------------------
        ' If not bounded by a tube
        ' ------------------------
        Else
            
            GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, cTopOrWL, cBtmOrWR, cTFIOrFL, cBFIOrFR
            
            ' -----------------
            ' Pick a center cut
            ' -----------------
            If StrComp(pSelectorLogic.answer("BottomFlange"), "Yes", vbTextCompare) = 0 Then
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
                       pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                Case 1, 6
                    'WebTopFlangeRight, WebTopFlange
                     If cWLorTFL.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                     Else
                       pSelectorLogic.Add "Axis_Flange_TE_F" 'Add UA Case
                     End If
                Case 2
                    'WebBuiltUpTopFlangeRight
                     If cWLorTFL.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                     ElseIf cWLorTFL.IntersectingEdge = Top_Flange_Right Then
                       pSelectorLogic.Add "Axis_Flange_TE_F" 'Add UA case
                     Else
                       pSelectorLogic.Add "Axis_Flange_TF_TE_F" 'Add BUTop
                     End If
                Case 3, 7
                     'WebBottomFlangeRight, WebBottomFlange
                     If eWRorTFR.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                     Else
                       pSelectorLogic.Add "Axis_Flange_F_BE" 'Add L Case
                     End If
                Case 4
                    'WebBuiltUpBottomFlangeRight
                     If eWRorTFR.IntersectingEdge = Web_Right Then
                       pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                     ElseIf eWRorTFR.IntersectingEdge = Bottom_Flange_Right Then
                       pSelectorLogic.Add "Axis_Flange_F_BE" 'Add L case
                     Else
                       pSelectorLogic.Add "Axis_Flange_F_BE_BF" 'Add BUBottom
                     End If
                Case 5, 13, 19
                    'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
                    If cWLorTFL.IntersectingEdge = Web_Right Then
                       If eWRorTFR.IntersectingEdge = Web_Right Then
                          pSelectorLogic.Add "Axis_Flange_F" 'Add FB case
                       Else
                          pSelectorLogic.Add "Axis_Flange_F_BE" 'Add L case
                       End If
                    ElseIf cWLorTFL.IntersectingEdge = Top_Flange_Right Or _
                        cWLorTFL.IntersectingEdge = above Then
                        If eWRorTFR.IntersectingEdge = Web_Right Then
                            pSelectorLogic.Add "Axis_Flange_TE_F" 'Add UA case
                        Else
                            pSelectorLogic.Add "Axis_Flange_TE_F_BE" 'Add I case
                        End If
                    Else
                        pSelectorLogic.Add "Axis_Flange_TE_F_BE" 'Add I case
                    End If
                Case 14
                    'FlangeLeftAndRightBottomWebs
                    pSelectorLogic.Add "Axis_FlangeB_F"
                Case 15
                    'FlangeLeftAndRightTopWebs
                    pSelectorLogic.Add "Axis_FlangeT_F"
                Case 16, 18
                    'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
                    pSelectorLogic.Add "Axis_FlangeTB_F"
                Case 20
                    'Tube/Circular Cross Section
                    pSelectorLogic.Add ""
                Case Else
                    'Unknown Section Alias
            End Select
        End If
    ' -----------------
    ' If web-penetrated
    ' -----------------
    Else
        
        GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, cTopOrWL, cBtmOrWR, cTFIOrFL, cBFIOrFR
    
        ' ------------------------------------------------------------------
        ' Identify cases where the flange is set back without an outside cut
        ' ------------------------------------------------------------------
        Dim sBottomShape As String
        Dim sTopShape As String
        Dim sShape As String
        
        Dim bFlangeReliefByFlangeCut As Boolean
        bFlangeReliefByFlangeCut = False
        
        GetMemberACTopAndBottomShape pSelectorLogic.SmartOccurrence, , sBottomShape, , sTopShape
        
        If (bBottomFlange) And (sBottomShape = gsEdgeToOutside) And (oBtmWebCut Is Nothing) Then
            bFlangeReliefByFlangeCut = True
        ElseIf (Not bBottomFlange) And (sTopShape = gsEdgeToOutside) And (oTopWebCut Is Nothing) Then
            bFlangeReliefByFlangeCut = True
        End If
        
        ' --------------------------------------------------------------------------------------------------
        ' If the flange is to be sniped, pick a cut appropriate for the flange (symmetrical or asymmetrical)
        ' --------------------------------------------------------------------------------------------------
        Dim isFlangeOutside As Boolean
        isFlangeOutside = False
        
        If (Not bBottomFlange) Then
            If (cTopOrWL.IntersectingEdge = above And cTFIOrFL.IntersectingEdge = above) Or _
               (cTopOrWL.IntersectingEdge = Below And cTFIOrFL.IntersectingEdge = Below) Then
                isFlangeOutside = True
            End If
        ElseIf bBottomFlange Then
            If (cBtmOrWR.IntersectingEdge = above And cBFIOrFR.IntersectingEdge = above) Or _
               (cBtmOrWR.IntersectingEdge = Below And cBFIOrFR.IntersectingEdge = Below) Then
                isFlangeOutside = True
            End If
        End If
        
        If StrComp(strAnswer, "Sniped", vbTextCompare) = 0 Then
        
            ' When the initial R1 schema was first used by internal customers (e.g. business development, QTPs, etc), we lacked
            ' an item that was welded and sniped, like the stiffener "Connected" type, which uses SM_FlangeCutSnipWeld.sym.
            ' A schema changes was approved to add the item, but with the stipulation that bulkloading the change be optional.
            ' So we must check that the item exists.
            
            Dim bSnipWeldItemsFlag As Boolean
            Dim eSCType As SmartClassType
            Dim eSCSubType As SmartClassSubType
            Dim sClassName As String
            
            eSCType = SMARTTYPE_FLANGECUT
            eSCSubType = 264
            sClassName = "MbrAxisFlangeSel"
            If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "Axis_Flange_Sniped_LR") And _
              CheckSmartItemExists(eSCType, eSCSubType, sClassName, "Axis_Flange_Sniped_R") Then
              bSnipWeldItemsFlag = True
            Else
              bSnipWeldItemsFlag = False
            End If
            
            ' -----------------------------------------------------------------------------------------
            ' If there is an outside web cut, the flange is free. Pick a sniped cut that is not welded.
            ' -----------------------------------------------------------------------------------------
            If (Not bBottomFlange And Not oTopWebCut Is Nothing) Or (bBottomFlange And Not oBtmWebCut Is Nothing) Or _
               (oCtrWebCut Is Nothing) Or isFlangeOutside Or bFlangeReliefByFlangeCut Then
                ' Choose further based on symmetrical or asymmetrical flange
                If (Not bBottomFlange And bTopFL And bTopFR) Or (bBottomFlange And bBtmFL And bBtmFR) Then
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_LR"
                Else
                    pSelectorLogic.Add "Axis_Flange_SnipedSetback_R"
                End If
            ' ---------------------------------------------------------------------
            ' Otherwise, pick a welded snipe cut with a setback one as an alternate
            ' ---------------------------------------------------------------------
            Else
                ' Choose further based on symmetrical or asymmetrical flange
                If (Not bBottomFlange And bTopFL And bTopFR) Or (bBottomFlange And bBtmFL And bBtmFR) Then
                    If bSnipWeldItemsFlag Then
                        pSelectorLogic.Add "Axis_Flange_Sniped_LR"
                    Else
                        pSelectorLogic.Add "Axis_Flange_SnipedSetback_LR"
                    End If
                Else
                    If bSnipWeldItemsFlag Then
                        pSelectorLogic.Add "Axis_Flange_Sniped_R"
                    Else
                        pSelectorLogic.Add "Axis_Flange_SnipedSetback_R"
                    End If
                End If
            End If
        ' -------------------------------
        ' If the flange is to be straight
        ' -------------------------------
        Else
             ' ------------------------------------------------------------------------------------------
             ' If there is an outside web cut, the flange is free. Pick a straight cut that is not welded
             ' ------------------------------------------------------------------------------------------
             If (Not bBottomFlange And Not oTopWebCut Is Nothing) Or (bBottomFlange And Not oBtmWebCut Is Nothing) Or _
                (oCtrWebCut Is Nothing) Or isFlangeOutside Or bFlangeReliefByFlangeCut Then
                 ' Choose further based on symmetrical or asymmetrical flange
                 If (Not bBottomFlange And bTopFL And bTopFR) Or (bBottomFlange And bBtmFL And bBtmFR) Then
                     pSelectorLogic.Add "Axis_Flange_StraightSetback_LR"
                 Else
                     pSelectorLogic.Add "Axis_Flange_StraightSetback_R"
                 End If
             ' ------------------------------------------------------------------------
             ' Otherwise, pick a welded straight cut with a setback one as an alternate
             ' ------------------------------------------------------------------------
             Else
                 ' Choose further based on symmetrical or asymmetrical flange
                 If (Not bBottomFlange And bTopFL And bTopFR) Or (bBottomFlange And bBtmFL And bBtmFR) Then
                     pSelectorLogic.Add "Axis_Flange_Flush_LR"
                     pSelectorLogic.Add "Axis_Flange_StraightSetback_LR"
                 Else
                     pSelectorLogic.Add "Axis_Flange_Flush_R"
                     pSelectorLogic.Add "Axis_Flange_StraightSetback_R"
                 End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
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
