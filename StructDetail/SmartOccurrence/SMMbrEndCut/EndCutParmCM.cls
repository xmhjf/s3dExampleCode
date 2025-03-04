VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndCutParmCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : EndCutParmCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'    11/Apr/2012 - svsmylav
'       DM-213229: Updated logic to handle stiffener as bounded object.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same

'*********************************************************************************************
Const m_sClassName As String = "EndCutParmCM"
Const m_ParamProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Public Sub ComputeFaceSnipeFlangeCutParameters(oPRL As IJDParameterLogic, _
                                               bIsBottomCut As Boolean, _
                                               faceOffset As Double, _
                                               nose As Double, _
                                               slope As Double, _
                                               flangeOffset As Double, _
                                               edgeInsideOffset As Double, _
                                               isBottom As Long, _
                                               cutDepth As Double)
                                      
    Const METHOD = m_ParamProgid & "::ComputeFaceSnipeFlangeCutParameters"
                                      
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    ' -----------------------------------------------------------------------------------
    ' Determine if we are on the top or bottom flange and set the "BottomFlage" parameter
    ' -----------------------------------------------------------------------------------
    sMsg = "Determine if we are on the top or bottom flange"
    
    Dim sBottomFlange As String
    GetSelectorAnswer oPRL, "BottomFlange", sBottomFlange

    If sBottomFlange = "No" Then
        isBottom = 0
    Else
        isBottom = 1
    End If
    
    ' --------------------------------------------------
    ' Determine what flanges exist on the bounded member
    ' --------------------------------------------------
    sMsg = "Determine what flanges exist on the bounded member"
    
    Dim oPortBounded As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)

    Dim oBoundedObject As Object
    Set oBoundedObject = oPortBounded.Connectable
    
    Dim dFlgLen As Double
    Dim dWebThickness As Double
    
    If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBoundedObject
        dFlgLen = oBoundedMemberPart.FlangeLength
        dWebThickness = oBoundedMemberPart.WebThickness
    ElseIf TypeOf oBoundedObject Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New ProfilePart
        Set oBoundedProfilePart.object = oBoundedObject
        dFlgLen = oBoundedProfilePart.FlangeLength
        dWebThickness = oBoundedProfilePart.WebThickness
    Else
        'Special case need to be handled
    End If
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    CrossSection_Flanges oBoundedObject, bTFL, bBFL, bTFR, bBFR
    
    ' ----------------------------------------------------------------------------
    ' Calculate the distance from the web to the flange edge, if the flange exists
    ' ----------------------------------------------------------------------------
    sMsg = "Calculating the flange offset"
    ' If for a top cut
    If Not bIsBottomCut Then
        If (sBottomFlange = "No" And bTFL) Then
            If bTFL And bTFR Then
                flangeOffset = (dFlgLen - dWebThickness) / 2
            Else
                flangeOffset = dFlgLen - dWebThickness
            End If
        ElseIf (sBottomFlange = "Yes" And bBFL) Then
            If bBFL And bBFR Then
                flangeOffset = (dFlgLen - dWebThickness) / 2
            Else
                flangeOffset = dFlgLen - dWebThickness
            End If
        Else
            flangeOffset = 0.00001
        End If
    ' If for a bottom cut
    Else
        If (sBottomFlange = "No" And bTFR) Then
            If bTFR And bTFL Then
                flangeOffset = (dFlgLen - dWebThickness) / 2
            Else
                flangeOffset = dFlgLen - dWebThickness
            End If
        ElseIf (sBottomFlange = "Yes" And bBFR) Then
            If bBFR And bBFL Then
                flangeOffset = (dFlgLen - dWebThickness) / 2
            Else
                flangeOffset = dFlgLen - dWebThickness
            End If
        Else
            flangeOffset = 0.00001
        End If
    End If

    ' ----------------------------------------
    ' Set the nose to a static value of 0.01mm
    ' ----------------------------------------
    sMsg = "Calculating the nose value"
    
    nose = 0.00001
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    
    GetConnectedEdgeInfo oPRL.SmartOccurrence, oPRL.InputObject(INPUT_BOUNDED), oPRL.InputObject(INPUT_BOUNDING), eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, oMeasurements

    ' ---------------------------------------------------------------------------------------
    ' If there is no flange outside the web, set a vertical offset equal to the web thickness
    ' If the web is above/below the bounding object then set the vertical offset to be the
    ' distance to the top/bottom of the bounding object + 15mm
    ' ---------------------------------------------------------------------------------------
    sMsg = "Calculating the vertical offset"

    If Not bIsBottomCut Then
        If eTopOrWL.IntersectingEdge = Above And eBottomOrWR.IntersectingEdge = Above Then
            edgeInsideOffset = oMeasurements.Item("DimPt15ToFL") + 0.015
        Else
            If flangeOffset = 0.00001 Then
                edgeInsideOffset = dWebThickness
            Else
                edgeInsideOffset = dFlgLen / 2 + dWebThickness / 2
            End If
        End If
    Else
        If eBottomOrWR.IntersectingEdge = Below And eTopOrWL.IntersectingEdge = Below Then
            edgeInsideOffset = oMeasurements.Item("DimPt23ToFR") + 0.015
        Else
            If flangeOffset = 0.00001 Then
                edgeInsideOffset = dWebThickness
            Else
                edgeInsideOffset = dFlgLen / 2 + dWebThickness / 2
            End If
        End If
    End If
    
    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    sMsg = "Calculating the cutting depth"
    
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(oPRL)
    cutDepth = dCuttingDepth

    ' -------------
    ' Set the slope
    ' -------------
    sMsg = "Calculating the cutting depth"

    Dim dPI As Double
    dPI = Atn(1) * 4#
    
    Dim sShapeOutSide As String
    Dim sParentItemName As String
    Dim oParentObj As Object

    Parent_SmartItemName oPRL.SmartOccurrence, sParentItemName, oParentObj

    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    If sShapeOutSide = "" Then
        If bIsBottomCut Then
            GetSelectorAnswer oParentObj, "BottomShapeOutside", sShapeOutSide
        Else
            GetSelectorAnswer oParentObj, "TopShapeOutside", sShapeOutSide
        End If
    End If
    
    If sShapeOutSide = gsStraight Then
        slope = 90 * dPI / 180
    ElseIf sShapeOutSide = gsSniped Then
        slope = 65 * dPI / 180
    Else
        slope = 45 * dPI / 180
    End If
    
    ' ----------------------------------------------------------
    ' Set the horizontal face offset such that the slope will intersect with the
    ' bottom of the near side of the web.
    ' ----------------------------------------------------------
    sMsg = "Calculating the horizontal offset"
    
    If sShapeOutSide = gsStraight Then
        faceOffset = dFlgLen / 4
    Else
        faceOffset = Tan(slope) * (edgeInsideOffset - dFlgLen - dWebThickness)
    End If
        
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ComputeEdgeHookFlangeReliefWebCutParameters(oPRL As IJDParameterLogic, _
                                                       bIsBottomCut As Boolean, _
                                                       dFlangeThicknessOrOffset As Double, _
                                                       dInsideOffset As Double, _
                                                       dRadius As Double, _
                                                       dEdgeOffset As Double, _
                                                       dCuttingDepth As Double)
    
    Const METHOD = m_sClassName & "::ComputeEdgeHookFlangeReliefWebCutParameters"
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oPortBounded As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)

    Dim oPortBounding As IJPort
    Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
    
    ' ---------------------------------------------------------
    ' Set the flange thickness (web cut) or offset (flange cut)
    ' ---------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = GetFlangeOffset(oPRL, bIsBottomCut)
    End If
    
    ' -------------------------------------------------
    ' Get the relative position of bounding and bounded
    ' -------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrFL As ConnectedEdgeInfo
    Dim eITFOrFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    Dim bWebPenetrated As Boolean
    
    GetConnectedEdgeInfo oPRL.SmartOccurrence, oPortBounded, oPortBounding, eTopOrWL, eBtmOrWR, eITFOrFL, eITFOrFR, oMeasurements, bWebPenetrated

    ' ------------------------------------------
    ' Determine the overlap and inside clearance
    ' ------------------------------------------
    Dim dOutsideOverlap As Double
    Dim dInsideClearance As Double
    Dim dEdgeLength As Double
        
    GetEdgeOverlapAndClearance oPRL.SmartOccurrence, Not bIsBottomCut, isBottomFlange(oPRL), , dOutsideOverlap, dInsideClearance, , dEdgeLength
    
    ' ------------------------------------------------------------------
    ' If the overlap is negative, calculate the flange (not web) overlap
    ' ------------------------------------------------------------------
    Dim dBoundedThickness As Double
    dBoundedThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    
    If dOutsideOverlap < 0# Then
        If dBoundedThickness - dOutsideOverlap > dEdgeLength Then
            dOutsideOverlap = dEdgeLength
        Else
            dOutsideOverlap = dBoundedThickness - dOutsideOverlap
        End If
    End If
    
    ' -----------------------------------------------------
    ' If the feature should start in the middle of the edge
    ' -----------------------------------------------------
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oPRL.SmartOccurrence, sACItemName, oACObject
    
    Dim sAnswer As String
    GetSelectorAnswer oACObject, "ShapeAtEdge", sAnswer
    
    If sAnswer = vbNullString Then
        
        GetSelectorAnswer oACObject, "ShapeAtEdgeOverlap", sAnswer
    End If
    
    If LCase(sAnswer) = LCase("EdgeToOutside") Or LCase(sAnswer) = LCase("EdgeToFlange") Then
        ' -------------------------------------------------------------------------------------------------------
        ' Set the bottom offset to so the cut starts half way between the bounding outside and the bounded flange
        ' -------------------------------------------------------------------------------------------------------
        dInsideOffset = dEdgeLength - (dOutsideOverlap / 2#)
    ' -------------------------------
    ' If it is to start at the corner
    ' -------------------------------
    Else
        ' ----------------------------------------
        ' Set the bottom offset to the edge length
        ' ----------------------------------------
        dInsideOffset = dEdgeLength
    End If

    ' --------------------------------------------------------------------------------------------------------------------------
    ' Web cuts:
    '   Set the radius equal to distance from the offset point to the inside of the bottom flange, 5mm increments, 50mm max
    ' Flange cuts:
    '   Set the radius equal to distance from the offset point to flange, 5mm increments, 50mm max
    ' --------------------------------------------------------------------------------------------------------------------------
    ' For flange cuts, radius limited to distance from start of feature to edge of the flange (right) unless the web is in the way (left):
    '                         +---------------------
    ' |                       |                    -            ' |                          +----------------
    ' |                       |                                 ' |                          |   ^
    ' |                       |                                 ' |                          |   |
    ' |                       |                                 ' |                          |   |
    ' |                       |          Bounded                ' +------------------+       |   |
    ' |                       |                                 '           ^    ^   |       |   |
    ' +------------------+    |                                 '           |    C   |       /   A
    '             ^   ^  |    |----------------------           '           |    v   |      /    |
    '             |   C  |    |----------------------           '           B   ---- |-----/     |
    '  Bounding   |   v  |   /        ^                         ' Bounding  |        |           |
    '             B  --- |--/         |                         '           |        |           |
    '             |      |            A                         '           v        |           v
    '             |      |            |                         ' -------------------+ -------------
    '             v      |            v                         '          |                            Bounded
    ' -------------------+ -------------------                  '          |
    '          |    A = outside corner to web right             '          |-------------------------------------------
    '          |    B = edge length                             '          |-------------------------------------------
    '          |    c = inside offset                           '          |
    '          +----------------------------------------        '          |
    
    
    
    If bIsBottomCut Then
        If eFeatureType = SF_WebCut Then
            dRadius = oMeasurements.Item("DimPt15ToBottomInside") - dEdgeLength + dInsideOffset
        Else
            If eBtmOrWR.IntersectingEdge = Above Or eBtmOrWR.IntersectingEdge = Top Or _
               eBtmOrWR.IntersectingEdge = Top_Flange_Right_Top Or eBtmOrWR.IntersectingEdge = Web_Right_Top Then
            
                dRadius = oMeasurements.Item("DimPt15ToFR") - dEdgeLength + dInsideOffset
            Else
                dRadius = oMeasurements.Item("DimPt15ToWL") - dEdgeLength + dInsideOffset
            End If
        End If
    Else
        If eFeatureType = SF_WebCut Then
            dRadius = oMeasurements.Item("DimPt23ToTopInside") - dEdgeLength + dInsideOffset
        Else
            If eTopOrWL.IntersectingEdge = Below Or eTopOrWL.IntersectingEdge = Bottom Or _
               eTopOrWL.IntersectingEdge = Bottom_Flange_Right_Bottom Or eTopOrWL.IntersectingEdge = Web_Right_Bottom Then
            
                dRadius = oMeasurements.Item("DimPt23ToFL") - dEdgeLength + dInsideOffset
            Else
                dRadius = oMeasurements.Item("DimPt23ToWR") - dEdgeLength + dInsideOffset
            End If
        End If
    End If
    
    Dim lMult As Long
    lMult = Int(dRadius / 0.005)
    
    dRadius = 0.05
    
    If lMult < 10 Then
        dRadius = lMult * 0.005
    End If
    
    If Not dRadius > 0# Then
        dRadius = dOutsideOverlap / 2#
    End If
    
    ' -----------------------------------------------------------------------
    ' Set the lateral offset from the bounding edge equal to the radius + 5mm
    ' -----------------------------------------------------------------------
    ' At least 0.01mm is required to keep a ProfileSketchCurve line from becoming zero-length
    dEdgeOffset = dRadius + 0.005

    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    dCuttingDepth = EndCut_GetCutDepth(oPRL)

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

Public Sub ComputeOutsideEdgeHookParameters(oPRL As IJDParameterLogic, _
                                            bIsBottomCut As Boolean, _
                                            dFlangeThicknessOrOffset As Double, _
                                            dRadius As Double, _
                                            dOffset As Double, _
                                            dEdgeOffset As Double, _
                                            dCutDepth As Double)

    Const METHOD = m_sClassName & "::ComputeOutsideEdgeHookParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ---------------------------------------------------------
    ' Set the flange thickness (web cut) or offset (flange cut)
    ' ---------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = GetFlangeOffset(oPRL, bIsBottomCut)
    End If
    
    ' ------------------------
    ' Get the overlap distance
    ' ------------------------
    Dim bIsBottomEdge As Boolean
    Dim bIsBottomFlange As Boolean
    Dim dOverlap As Double
    Dim dOutsideClearance As Double
    Dim dEdgeLength As Double
    
    bIsBottomFlange = IsBottomFlange(oPRL)
    
    GetEdgeOverlapAndClearance oPRL.SmartOccurrence, bIsBottomCut, bIsBottomFlange, dOverlap, , , dOutsideClearance, dEdgeLength
    
    ' ------------------------------------------------------------------
    ' If the overlap is negative, calculate the flange (not web) overlap
    ' ------------------------------------------------------------------
    Dim dBoundedThickness As Double
    dBoundedThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    
    If dOverlap < 0# Then
        If dBoundedThickness - dOverlap > dEdgeLength Then
            dOverlap = dEdgeLength
        Else
            dOverlap = dBoundedThickness - dOverlap
        End If
    End If

    ' ----------------------------------
    ' Set the radius to half the overlap
    ' ----------------------------------
    ' The 0.01mm tolerance is to avoid ACIS failures where the radius ends directly on an edge
    Dim dOverall As Double
    
    dOverall = Abs(dOverlap) / 2# + dOutsideClearance
    
    If dOverall >= 0.100001 Then
        dRadius = 0.05
    ElseIf dOverall >= 0.070001 Then
        dRadius = 0.035
    ElseIf dOverall >= 0.050001 Then
        dRadius = 0.025
    ElseIf dOverall >= 0.030001 Then
        dRadius = 0.015
    ElseIf dOverall >= 0.020001 Then
        dRadius = 0.01
    ElseIf dOverall >= 0.010001 Then
        dRadius = 0.005
    Else
        dRadius = dOverall / 2# + 0.000002 ' dRadius must be a minimum of 0.02mm.  It seems that 0.01mm should work, but it does not.
    End If

    ' --------------------------------------------------------
    ' Set the vertical offset (along the edge) from the bottom
    ' --------------------------------------------------------
    If Abs(dOverlap) > dEdgeLength Then
        dOffset = dEdgeLength / 2#
    Else
        dOffset = dEdgeLength - Abs(dOverlap) / 2#
    End If
    
    ' --------------------------------------------
    ' Set the horizontal offset away from the edge
    ' --------------------------------------------
    dEdgeOffset = dRadius + 0.01 'EdgeOffset should be greater than or equal to the radius

    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(oPRL)
    dCutDepth = dCuttingDepth

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub ComputeHookFlangReliefParameters(oPRL As IJDParameterLogic, _
                                            bIsBottomCut As Boolean, _
                                            bWebIsOutside As Boolean, _
                                            dEdgeOffset As Double, _
                                            dRadius As Double, _
                                            dEdgeSetback As Double, _
                                            dEdgeInsideOffset As Double, _
                                            dFlangeThicknessOrOffset As Double, _
                                            dCutDepth As Double)

    Const METHOD = m_sClassName & "::ComputeHookFlangReliefParameters"

    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ------------------------
    ' Get the flange thickness
    ' ------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = GetFlangeOffset(oPRL, bIsBottomCut)
    End If
    
    ' ----------------------------------------------------------------------------------------
    ' Set the edge setback to 15mm for InsideToOutside, or near-zero value for CornerToOutside
    ' ----------------------------------------------------------------------------------------
    Dim sBottomShape As String
    Dim sTopShape As String
    Dim sShape As String
    
    GetMemberACTopAndBottomShape oPRL.SmartOccurrence, , sBottomShape, , sTopShape
    
    If bIsBottomCut Then
        sShape = sBottomShape
    Else
        sShape = sTopShape
    End If
    
    If sShape = gsCornerToOutside Then
        dEdgeSetback = 0.00001
    Else
        dEdgeSetback = 0.015
    End If
    
    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    dCutDepth = EndCut_GetCutDepth(oPRL)
    
    ' ---------------------------------------
    ' The remaining parameters are hard-coded
    ' ---------------------------------------
    dRadius = 0.015
    dEdgeOffset = 2 * dRadius + dEdgeSetback + 0.00001
    dEdgeInsideOffset = 0.015

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

Public Sub GetSquareFlangeReliefParameters(oPRL As IJDParameterLogic, _
                                           bIsBottomCut As Boolean, _
                                           dRadius As Double, _
                                           dFlangeThicknessOrOffset As Double, _
                                           dEdgeInsideOffset As Double, _
                                           dEdgeOffset As Double, _
                                           dCutDepth As Double, _
                                           Optional dEdgeLength As Double, _
                                           Optional dMaxNose As Double)
                                            
    Const METHOD = m_sClassName & "::GetSquareFlangeReliefParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ---------------------------------------------------------
    ' Set the flange thickness (web cut) or offset (flange cut)
    ' ---------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = GetFlangeOffset(oPRL, bIsBottomCut)
    End If
    
    ' -----------------------------
    ' Get the overlap and clearance
    ' -----------------------------
    Dim dOverlap As Double
    Dim dOutsideClearance As Double
    Dim dInsideClearance As Double
    Dim dOverall As Double
    
    GetEdgeOverlapAndClearance oPRL.SmartOccurrence, bIsBottomCut, IsBottomFlange(oPRL), dOverlap, , dInsideClearance, dOutsideClearance, dEdgeLength
       
    Select Case oPRL.SmartItem.Name
        ' ------------------------------------
        ' If there is clearance for the flange
        ' ------------------------------------
        Case "SquareFlangeReliefBottom", "SquareFlangeReliefTop", _
             "SquareFlangeReliefWebBottomCutSnipe", "SquareFlangeReliefWebTopCutSnipe", _
             "FlgSqrFlangeReliefTop", "FlgSqrFlangeReliefBtm", "FlgSqrFlgRlfOutsideSnpTop_WL", "FlgSqrFlgRlfOutsideSnpBtm_FR"
    
            ' If there is not a large inside clearance the results will not look good, as the bounded will barely be connected
            ' to the bounding.  But to avoid ToDo list errors even with poor user settings, make sure the inside offset is not
            ' more than the clearance.  Make it 1/4 of the clearance, using standard values, up to 35mm
            If dInsideClearance > 0.14 Then
                dEdgeInsideOffset = 0.035
            ElseIf dInsideClearance > 0.1 Then
                dEdgeInsideOffset = 0.025
            ElseIf dInsideClearance > 0.06 Then
                dEdgeInsideOffset = 0.015
            ElseIf dInsideClearance > 0.04 Then
                dEdgeInsideOffset = 0.01
            ElseIf dInsideClearance > 0.02 Then
                dEdgeInsideOffset = 0.005
            Else
                dEdgeInsideOffset = dInsideClearance - 0.00001
            End If
            
            If dEdgeInsideOffset < 0.025 Then
                dRadius = dEdgeInsideOffset
            Else
                dRadius = 0.025
            End If
            
            dEdgeOffset = 0.05
            
        ' -------------------------------------------
        ' If tight to the inside of the bounding edge
        ' -------------------------------------------
        Case "SquareAlongFlangeReliefBottom", "SquareAlongFlangeReliefTop", _
             "SquareAlongFlangeReliefWebTopCutSnipe", "SquareAlongFlangeReliefWebBottomCutSnipe", _
             "FlgSqrAlgFlangeReliefTop", "FlgSqrAlgFlangeReliefBtm", "FlgSqrAlgFlgRlfOutsideSnpTop_WL", "FlgSqrAlgFlgRlfOutsideSnpBtm_FR"
            
            ' -------------------------------
            ' Get the shape at top and bottom
            ' -------------------------------
            Dim sBottomShape As String
            Dim sTopShape As String
            Dim sShape As String
     
            GetMemberACTopAndBottomShape oPRL.SmartOccurrence, , sBottomShape, , sTopShape
            
            If bIsBottomCut Then
                sShape = sBottomShape
            Else
                sShape = sTopShape
            End If
                        
            ' ------------------------------------------------------------------
            ' If the overlap is negative, calculate the flange (not web) overlap
            ' ------------------------------------------------------------------
            Dim dBoundedWebOrFlangeThickness As Double
            dBoundedWebOrFlangeThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
            
            If dOverlap < 0# Then
                If dBoundedWebOrFlangeThickness - dOverlap > dEdgeLength Then
                    dOverlap = dEdgeLength
                Else
                    dOverlap = dBoundedWebOrFlangeThickness - dOverlap
                End If
            End If
            
            ' -----------------------------------------------------------------
            ' Make sure the radius is not too large for the available clearance
            ' -----------------------------------------------------------------
            ' Target radius to take half the available clearance
            dOverall = dOverlap + dOutsideClearance
            
            If dOverall >= 0.200001 Then
                dRadius = 0.1
            ElseIf dOverall >= 0.150001 Then
                dRadius = 0.075
            ElseIf dOverall >= 0.100001 Then
                dRadius = 0.05
            ElseIf dOverall >= 0.050001 Then
                dRadius = 0.025
            ElseIf dOverall >= 0.030001 Then
                dRadius = 0.015
            ElseIf dOverall >= 0.020001 Then
                dRadius = 0.01
            ElseIf dOverall >= 0.010001 Then
                dRadius = 0.005
            ElseIf dOverall > 0.00001 Then
                dRadius = dOverlap
            Else
                dRadius = 0.00002 ' a radius of 0.01mm resulted in ACIS errors in some cases.  Changing our near-zero value to 0.02mm.
            End If
            
            ' ---------------------------------------------
            ' The inside offset is a static near-zero value
            ' ---------------------------------------------
            ' A value of 0.01mm should work, but failed.  Raising to 0.02mm resolved the problem, but it is not known why
            ' There is a static 0.01mm "jog" where the contour layer changes, but this is not affected by the value of the offset
            ' (i.e. no edge length are driven to zero as a result of the offset = jog = 0.01mm)
            dEdgeInsideOffset = 0.00002
            
            ' --------------------------------------------------------------------------------
            ' Set the edge offset
            ' Do not extend radius away from bounding edge if the inside of the edge is sloped
            ' --------------------------------------------------------------------------------
            Dim sSectionName As String
    
            Dim oPortBounding As IJPort
            Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
 
            If TypeOf oPortBounding.Connectable Is ISPSMemberPartCommon Then
                Dim oSDOMemberPart As New StructDetailObjects.MemberPart
                Set oSDOMemberPart.object = oPortBounding.Connectable
                sSectionName = oSDOMemberPart.SectionName
            Else
                Dim oSDOProfilePart As New StructDetailObjects.ProfilePart
                Set oSDOProfilePart.object = oPortBounding.Connectable
                sSectionName = oSDOProfilePart.SectionName
            End If
            
            If UCase(Left(sSectionName, 2)) = "ST" Or _
               UCase(Left(sSectionName, 1)) = "S" Then
                
                dEdgeOffset = dRadius
            Else
                dEdgeOffset = dRadius + 0.015
            End If
            
    End Select
    
    ' -------------------------------------------------------------------
    ' Define the cutting depth to completely cut the Member cross Section
    ' -------------------------------------------------------------------
    dCutDepth = EndCut_GetCutDepth(oPRL)

    ' -------------------------------
    ' Calculate the maximum nose size
    ' -------------------------------
    dMaxNose = dOutsideClearance

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

Public Sub GetOutsideCutSnipeParameters(oPRL As IJDParameterLogic, _
                                        bIsBottomCut As Boolean, _
                                        dSlope As Double, _
                                        dSlopeOutside As Double, _
                                        dClearance As Double, _
                                        dNose As Double)
         
    Const METHOD = m_sClassName & "::GetOutsideCutSnipeParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oPRL.InputObject(INPUT_BOUNDED)
    
    ' ------------------------------------------------
    ' Set the slope based on the "ShapeOutside" answer
    ' ------------------------------------------------
    Dim dPI As Double
    dPI = Atn(1) * 4#
    
    Dim sShapeOutSide As String
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oPRL.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    
    If sShapeOutSide = "" Then
        If bIsBottomCut Then
            GetSelectorAnswer oParentObj, "BottomShapeOutside", sShapeOutSide
        Else
            GetSelectorAnswer oParentObj, "TopShapeOutside", sShapeOutSide
        End If
    End If
    
    If sShapeOutSide = gsStraight Then
        dSlope = dPI / 2#
    Else
        dSlope = dPI / 3#
    End If
    
    ' --------------------------------------------------------------------------------
    ' If there is a web/flange at the outside of the cut, set the outside slope to 90°
    ' Otherwise have it match the main slope
    ' --------------------------------------------------------------------------------
    Dim bIsBottomFlange As Boolean
    bIsBottomFlange = IsBottomFlange(oPRL)
    
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

    If (eFeatureType = SF_WebCut And bIsBottomCut And bBFR) Or _
       (eFeatureType = SF_WebCut And Not bIsBottomCut And bTFR) Or _
       (eFeatureType = SF_FlangeCut And Not bIsBottomCut And bIsBottomFlange And Not bBFL) Or _
       (eFeatureType = SF_FlangeCut And Not bIsBottomCut And Not bIsBottomFlange And Not bTFL) Then
        
        dSlopeOutside = dPI / 2#
    Else
        dSlopeOutside = dSlope
    End If
    
    ' For reasons yet-to-be-determined, both angles cannot be equal if both are 90
    If dSlope = dPI / 2# And dSlope = dSlopeOutside Then
        dSlope = 89.99 / 180 * dPI
    End If
    
    ' --------------------------------------------
    ' Standard flange clearance is near-zero value
    ' --------------------------------------------
    dClearance = 0.00001
        
    ' ------------------------------------------------------------
    ' Set the nose to allow for at least a 10 mm Snipe
    ' If there is insufficient clearance, set to a near-zero value
    ' ------------------------------------------------------------
    Dim dOutsideClearance As Double
    
    GetEdgeOverlapAndClearance oPRL.SmartOccurrence, bIsBottomCut, bIsBottomFlange, , , , dOutsideClearance
            
    If dOutsideClearance > 0.025 Then
        dNose = 0.015
    ElseIf dOutsideClearance > 0.02 Then
        dNose = 0.01
    ElseIf dOutsideClearance > 0.015 Then
        dNose = 0.005
    Else
        dNose = 0.00001
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' Get web thickness for flange cuts, flange thickness for web cuts
' For bottom web cuts, if there is no bottom flange, a near-zero value is returned
' For top web cuts, if there is no top flange, a near-zero value is returned
Public Function GetWebOrFlangeThickness(oPRL As IJDParameterLogic, bIsBottomCut As Boolean) As Double

    Const METHOD = m_sClassName & "::GetWebOrFlangeThickness"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Getting the flange thickness"

    ' ----------------------------------
    ' Determine if web cut or flange cut
    ' ----------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
        
    ' --------------------------
    ' Determine if bottom flange
    ' --------------------------
    Dim bIsBottomFlange As Boolean
    bIsBottomFlange = IsBottomFlange(oPRL)

    ' -------------------------------------------
    ' Get the thickness as defined by the section
    ' -------------------------------------------
    Dim oPortBounded As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)

    Dim oBoundedObject As Object
    Set oBoundedObject = oPortBounded.Connectable
    
    Dim dSectionThickness As Double
    
    If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBoundedObject
        If eFeatureType = SF_WebCut Then
            dSectionThickness = oBoundedMemberPart.flangeThickness
        Else
            dSectionThickness = oBoundedMemberPart.WebThickness
        End If
    ElseIf TypeOf oBoundedObject Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oBoundedObject
        If eFeatureType = SF_WebCut Then
            dSectionThickness = oBoundedProfilePart.flangeThickness
        Else
            dSectionThickness = oBoundedProfilePart.WebThickness
        End If
    End If

    ' --------------------------------------------------
    ' If a web cut and no flange, set to near-zero value
    ' --------------------------------------------------
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedObject, bTFL, bBFL, bTFR, bBFR
        
    If eFeatureType = SF_WebCut Then

        Dim dFlangeThickness As Double
        
        If ((bIsBottomCut) And (bBFL Or bBFR)) Or ((Not bIsBottomCut) And (bTFL Or bTFR)) Then
            GetWebOrFlangeThickness = dSectionThickness
        Else
            GetWebOrFlangeThickness = 0.00002
        End If
    Else
        GetWebOrFlangeThickness = dSectionThickness
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub GetExtendedOutsideCutParameters(oPRL As IJDParameterLogic, _
                                           bIsBottomCut As Boolean, _
                                           dThickness As Double, _
                                           dInsideOffset As Double, _
                                           dEdgeSetback As Double, _
                                           dCutDepth As Double)
         
    Const METHOD = m_sClassName & "::GetOutsideCutSnipeParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' -------------------------------
    ' Get the web or flange thickness
    ' -------------------------------
    dThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    
    ' --------------------------------------------
    ' Set the distance inside the bounding surface
    ' --------------------------------------------
    dInsideOffset = 0.00001
    
    ' ----------------------------------------------------
    ' Find the mapped flange length in the sketching plane
    ' ----------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Set oBoundedPort = oPRL.InputObject(INPUT_BOUNDED)
    Set oBoundingPort = oPRL.InputObject(INPUT_BOUNDING)
    
    GetConnectedEdgeInfo oPRL.SmartOccurrence, oBoundedPort, oBoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, oMeasurements
    
    Dim dFlangeLength As Double
    If bIsBottomCut Then
        dFlangeLength = oMeasurements.Item("DimPt3ToPt23")
    Else
        dFlangeLength = oMeasurements.Item("DimPt11ToPt15")
    End If
    
    ' ---------------------
    ' Get the section alias
    ' ---------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As New Collection

    Set oEdgeMapColl = GetEdgeMap(oPRL.SmartOccurrence, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    Dim boundingAlias As eBounding_Alias
    boundingAlias = GetBoundingAliasSimplified(lSectionAlias)
    
    ' ----------------------------------------------------
    ' Find the mapped web thickness in the sketching plane
    ' ----------------------------------------------------
    Dim dWebThickness As Double
    
    Select Case boundingAlias ' Web, WebBuiltUpTopFlangeRight,WebBuiltUpBottomFlangeRight and FlangeLeftAndRightWebs are not expected.
        Case WebTopFlangeRight, FlangeLeftAndRightBottomWebs
            dWebThickness = oMeasurements.Item("DimPt3ToPt23")
        Case WebBottomFlangeRight, FlangeLeftAndRightTopWebs
            dWebThickness = oMeasurements.Item("DimPt11ToPt15")
        Case WebTopAndBottomRightFlanges
            If bIsBottomCut Then
                dWebThickness = oMeasurements.Item("DimPt3ToPt23") - 2 * oMeasurements.Item("DimPt20ToPt21")
            Else
                dWebThickness = oMeasurements.Item("DimPt11ToPt15") - 2 * oMeasurements.Item("DimPt17ToPt18")
            End If
    End Select
    
    ' -------------------------------------------------------------------
    ' Determine if the bounding section has a left flange (top or bottom)
    ' -------------------------------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean

    CrossSection_Flanges oBoundingPort.Connectable, bTFL, bBFL, bTFR, bBFR
    
    ' ------------------------------------------------------
    ' Set the offset from the mappined bounding flange right
    ' ------------------------------------------------------
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oPRL.SmartOccurrence, sParentItemName, oParentObj
        
    Dim sExtendOrSetBack As String
    Dim dOffset As Double
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "Offset", dOffset
    
    Select Case sExtendOrSetBack
    
        Case gsExtendNearCorner
        
            dEdgeSetback = dOffset
        
        Case gsExtendFarCorner
            
            dEdgeSetback = dOffset + dFlangeLength
        
        Case gsExtendNearSide
            
            Select Case boundingAlias
                
                Case WebTopFlangeRight, WebBottomFlangeRight, WebTopAndBottomRightFlanges
                    
                    If bTFL Or bBFL Then
                        dEdgeSetback = dOffset + (dFlangeLength - dWebThickness) / 2
                    Else
                        dEdgeSetback = dOffset + dFlangeLength - dWebThickness
                    End If
                
                Case FlangeLeftAndRightBottomWebs, FlangeLeftAndRightTopWebs
                    
                    dEdgeSetback = dOffset + dWebThickness
            
            End Select
        
        Case gsExtendFarSide
            
            Select Case boundingAlias
                
                Case WebTopFlangeRight, WebBottomFlangeRight, WebTopAndBottomRightFlanges
                    
                    If bTFL Or bBFL Then
                        dEdgeSetback = dOffset + (dFlangeLength + dWebThickness) / 2
                    Else
                        dEdgeSetback = dOffset + dFlangeLength
                    End If
                
                Case FlangeLeftAndRightBottomWebs, FlangeLeftAndRightTopWebs
                    
                    dEdgeSetback = dOffset + dFlangeLength - dWebThickness
            
            End Select
        
        Case gsOffsetFarCorner
            
            dEdgeSetback = dFlangeLength - dOffset
        
        Case gsOffsetFarSide
            
            Select Case boundingAlias
                
                Case WebTopFlangeRight, WebBottomFlangeRight, WebTopAndBottomRightFlanges
                    
                    If bTFL Or bBFL Then
                        dEdgeSetback = (dFlangeLength + dWebThickness) / 2 - dOffset
                    Else
                        dEdgeSetback = dFlangeLength - dOffset
                    End If
                
                Case FlangeLeftAndRightBottomWebs, FlangeLeftAndRightTopWebs
                    
                    dEdgeSetback = dFlangeLength - dWebThickness - dOffset
            
            End Select
        
        Case gsOffsetNearSide
        
            Select Case boundingAlias
                
                Case WebTopFlangeRight, WebBottomFlangeRight, WebTopAndBottomRightFlanges
                    
                    If bTFL Or bBFL Then
                        dEdgeSetback = (dFlangeLength - dWebThickness) / 2 - dOffset
                    Else
                        dEdgeSetback = dFlangeLength - dWebThickness - dOffset
                    End If
                
                Case FlangeLeftAndRightBottomWebs, FlangeLeftAndRightTopWebs
                    
                    dEdgeSetback = dWebThickness - dOffset
            
            End Select
        
        Case gsOffsetNearCorner
            
            dEdgeSetback = dOffset
        
        Case Else
            
            dEdgeSetback = 0.00006
    
    End Select

    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    dCutDepth = EndCut_GetCutDepth(oPRL)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

Public Sub SetBottomFlangeParameter(oPRL As IJDParameterLogic)

    If IsBottomFlange(oPRL) Then
        oPRL.Add "BottomFlange", 1
    Else
        oPRL.Add "BottomFlange", 0
    End If

End Sub

Public Function IsBottomFlange(oPRL As IJDParameterLogic) As Boolean

    Dim sBottomFlange As String
    GetSelectorAnswer oPRL, "BottomFlange", sBottomFlange

    If sBottomFlange = gsYes Then
        IsBottomFlange = True
    Else
        IsBottomFlange = False
    End If

End Function

Public Sub GetHookFaceReliefParameters(oPRL As IJDParameterLogic, _
                                       bIsBottomCut As Boolean, _
                                       dFaceOffset As Double, _
                                       dRadius As Double, _
                                       dOutsideOffset As Double, _
                                       dFlangeThicknessOrOffset As Double, _
                                       dCuttingDepth As Double)
    
    Const METHOD = m_sClassName & "::GetHookFaceReliefParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ---------------------------------------------------------
    ' Set the flange thickness (web cut) or offset (flange cut)
    ' ---------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = GetFlangeOffset(oPRL, bIsBottomCut)
    End If
    
    ' ----------------------------------
    ' Get the parent assembly connection
    ' ----------------------------------
    Dim sParentName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oPRL.SmartOccurrence, sParentName, oParentObj
    
    Do While Not TypeOf oParentObj Is IJAppConnection
        Parent_SmartItemName oParentObj, sParentName, oParentObj
    Loop

    ' --------------------------------------
    ' If member axis or stiffener bounded AC
    ' --------------------------------------
    If GetMbrAssemblyConnectionType(oParentObj) = ACType_Axis Or _
       GetMbrAssemblyConnectionType(oParentObj) = ACType_Bounded Then
        
        ' --------------------------
        ' If there is an inset brace
        ' --------------------------
        Dim sBraceType As String

        If bIsBottomCut Then
            If GetMbrAssemblyConnectionType(oParentObj) = ACType_Axis Then
                GetSelectorAnswer oParentObj, "BottomBraceType", sBraceType
            End If
        Else
            If GetMbrAssemblyConnectionType(oParentObj) = ACType_Axis Then
                GetSelectorAnswer oParentObj, "TopBraceType", sBraceType
            End If
        End If
        
        If sBraceType = "InsetMember" Then
            
            ' ------------------------------------------------------
            ' Get the answers that determine the offsets at each end
            ' ------------------------------------------------------
            Dim dBraceBoundedEndValue As Double
            Dim sInsetBraceBoundedEndType As String
            Dim dBraceBoundingEndValue As Double
        
            If GetMbrAssemblyConnectionType(oParentObj) = ACType_Axis Then
                If bIsBottomCut Then
                    GetSelectorAnswer oParentObj, "BottomBraceBoundedEndType", sInsetBraceBoundedEndType
                    GetSelectorAnswer oParentObj, "BottomBraceBoundedEndValue", dBraceBoundedEndValue
                    GetSelectorAnswer oParentObj, "BottomBraceBoundingEndValue", dBraceBoundingEndValue
                Else
                    GetSelectorAnswer oParentObj, "TopBraceBoundedEndType", sInsetBraceBoundedEndType
                    GetSelectorAnswer oParentObj, "TopBraceBoundedEndValue", dBraceBoundedEndValue
                    GetSelectorAnswer oParentObj, "TopBraceBoundingEndValue", dBraceBoundingEndValue
                End If
            End If
            
            ' -----------------------------------------------------------------------------------------------------
            ' If located as an offset from the bounding along the bounded member, set the face offset to this value
            ' -----------------------------------------------------------------------------------------------------
            If sInsetBraceBoundedEndType = "Offset" Then
                dFaceOffset = dBraceBoundedEndValue
                
            ' -----------------------------------------------------------------------------------------------------------------
            ' If located as a point along the bounding and a slope toward the bounded, calculate the pointon the bounded member
            ' -----------------------------------------------------------------------------------------------------------------
            ' Here we presume the bounding end type is "offset", since they cannot both be "slope"
            ElseIf sInsetBraceBoundedEndType = "Slope" Then
                dFaceOffset = dBraceBoundedEndValue * dBraceBoundingEndValue
            End If
            
            ' -------------------------------------------
            ' Static values for radius and outside offset
            ' -------------------------------------------
            dRadius = 0.05
            dOutsideOffset = 0.1
        
        ' -------------------------
        ' If not for an inset brace
        ' -------------------------
        Else
            ' ---------------------------------------------------
            ' Determine where the bounded outside face intersects
            ' ---------------------------------------------------
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            
            Set oBoundedPort = oPRL.InputObject(INPUT_BOUNDED)
            Set oBoundingPort = oPRL.InputObject(INPUT_BOUNDED)
            
            Dim eTop As ConnectedEdgeInfo
            Dim eBottom As ConnectedEdgeInfo
            Dim eInsideTop As ConnectedEdgeInfo
            Dim eInsideBottom As ConnectedEdgeInfo
            Dim oMeasurements As Collection
            
            GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, eTop, eBottom, eInsideTop, eInsideBottom, oMeasurements
            
            ' ---------------------------------------------------------------
            ' If the bounded outside surface goes outside the bounding member
            ' ---------------------------------------------------------------
            If (Not bIsBottomCut And eTop.IntersectingEdge = Above) Or (bIsBottomCut And eBottom.IntersectingEdge = Below) Then
            
                ' ------------------------------------------------------------------------
                ' Outside offset is distance to outside + 10mm and other values are static
                ' ------------------------------------------------------------------------
                If bIsBottomCut Then
                    If eFeatureType = SF_WebCut Then
                        dOutsideOffset = oMeasurements.Item("DimPt23ToBottom") + 0.01
                    Else
                        dOutsideOffset = oMeasurements.Item("DimPt23ToWR") + 0.01
                    End If
                Else
                    If eFeatureType = SF_WebCut Then
                        dOutsideOffset = oMeasurements.Item("DimPt15ToTop") + 0.01
                    Else
                        dOutsideOffset = oMeasurements.Item("DimPt15ToFL") + 0.01
                    End If
                End If
                
                dFaceOffset = 0.065
                dRadius = 0.025
            ' -------------------------------------------------------------
            ' If the outside intersects the bounding, we have static values
            ' -------------------------------------------------------------
            Else
                dFaceOffset = 0.065
                dRadius = 0.025
                dOutsideOffset = 0.065
            End If
        End If
    ' --------------------------------------
    ' If a generic AC, all values are static
    ' --------------------------------------
    Else
        dFaceOffset = 0.065
        dRadius = 0.025
        dOutsideOffset = 0.065
    End If
   
    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    dCuttingDepth = EndCut_GetCutDepth(oPRL)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

Public Sub ComputeFaceOffsetSnipeParameters(oPRL As IJDParameterLogic, _
                                            bIsBottomCut As Boolean, _
                                            dAngle As Double, _
                                            dNose As Double, _
                                            dFaceOffset As Double, _
                                            dFlangeThickness As Double, _
                                            dFlangeClearance As Double, _
                                            dCuttingDepth As Double)
    
    Const METHOD = m_sClassName & "::ParameterRuleOutputs"
    
    Dim dPI As Double
    dPI = Atn(1) * 4#
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
    
    ' ----------------------------
    ' Get the web/flange thickness
    ' ----------------------------
    dFlangeThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    
    ' -----------------------------------------------------
    ' Get the port that the outside bounded edge intersects
    ' -----------------------------------------------------
    Dim eTop As ConnectedEdgeInfo
    Dim eBottom As ConnectedEdgeInfo
    Dim eInsideTop As ConnectedEdgeInfo
    Dim eInsideBottom As ConnectedEdgeInfo
    Dim oMeasurements As Collection

    GetConnectedEdgeInfo oPRL.SmartOccurrence, oPortBounded, oPortBounding, eTop, eBottom, eInsideTop, eInsideBottom, oMeasurements
    
    Dim intersectedEdge As eBounding_Edge
    
    If bIsBottomCut Then
        intersectedEdge = eBottom.IntersectingEdge
    Else
        intersectedEdge = eTop.IntersectingEdge
    End If
    
    ' --------------------------------------
    ' Set the angle to a constant 60 degrees
    ' --------------------------------------
    dAngle = dPI / 3#
    
    ' ---------------------------
    ' Set nose to a constant 15mm
    ' ---------------------------
    dNose = 0.015 '15mm
    
    ' -----------------------------------------------------------------
    ' Set the flange clearance to a constant 1mm
    ' -----------------------------------------------------------------
    dFlangeClearance = 0.00001
    
    ' ------------------------------------------------------
    ' If the bounded object goes outside the bounding object
    ' ------------------------------------------------------
    If (bIsBottomCut And intersectedEdge = Below) Or (Not bIsBottomCut And intersectedEdge = Above) Then
        
        ' ----------------------------------------------------------------------------------
        ' Compute a face offset that should leave 25mm vertical offset from the bounding top
        ' ----------------------------------------------------------------------------------
        Dim dDistToOutside As Double
        
        If bIsBottomCut Then
            dDistToOutside = oMeasurements.Item("DimPt23ToBottomInside")
        Else
            dDistToOutside = oMeasurements.Item("DimPt15ToTopInside")
        End If
        
        dFaceOffset = (dDistToOutside / Tan((dPI / 2) - dAngle)) + 0.025
    ' ------------------------------------
    ' Otherwise set it to a constant 100mm
    ' ------------------------------------
    Else
        dFaceOffset = 0.1
    End If
    
    ' ---------------------
    ' Get the cutting depth
    ' ---------------------
    dCuttingDepth = EndCut_GetCutDepth(oPRL)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

Public Sub ComputeFaceSnipeParametersTopOffset(oPRL As IJDParameterLogic, _
                                               bIsBottomCut As Boolean, _
                                               dFaceOffset As Double, _
                                               dEdgeSetback As Double, _
                                               dWebOrFlangeThickness As Double, _
                                               dClearance As Double, _
                                               dSlope As Double, _
                                               dSlopeOutside As Double, _
                                               dCuttingDepth As Double)

    Const METHOD = m_sClassName & "::ComputeFaceSnipeParametersTopOffset"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
    
    ' ----------------------------
    ' Get the web/flange thickness
    ' ----------------------------
    dWebOrFlangeThickness = GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    
    ' ---------------------------------------------
    ' Set the flange clearance to a constant 0.01mm
    ' ---------------------------------------------
    dClearance = 0.00001
    
    ' -----------------------------------------------------------
    ' Get the port that the outside bounded edge intersects
    ' For flange cuts, also find the port that the web intersects
    ' -----------------------------------------------------------
    ' For web cuts, this is inside the flange, if one exists
    ' For flange cuts, this is the flange edge (or Web Left, if there is no left flange)
    
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eITFOrFL As ConnectedEdgeInfo
    Dim eIBFOrFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection

    GetConnectedEdgeInfo oPRL.SmartOccurrence, oPortBounded, oPortBounding, eTopOrWL, eBottomOrWR, eITFOrFL, eIBFOrFR, oMeasurements
    
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    Dim outsideIntersectedEdge As eBounding_Edge
    Dim farWebIntersectedEdge As eBounding_Edge ' used for flange cuts only
    
    If bIsBottomCut Then
        outsideIntersectedEdge = eBottomOrWR.IntersectingEdge
        farWebIntersectedEdge = eTopOrWL.IntersectingEdge
    Else
        outsideIntersectedEdge = eTopOrWL.IntersectingEdge
        farWebIntersectedEdge = eBottomOrWR.IntersectingEdge
    End If
    
    ' ---------------------------------------------------------------------
    ' Measure the distance from the bounding outside to the bounded outside
    ' For flange cuts, also measure to the near side of the web
    ' ---------------------------------------------------------------------
    ' For web cuts, measure to the inside of the flange, if one exists
    ' For flange cuts, measure to the flange edge (or Web Left, if there is no left flange)
    
    Dim distToFlangeInsideOrEdge As Double
    Dim distToWeb As Double
    
    If eFeatureType = SF_WebCut Then
        If bIsBottomCut Then
            distToFlangeInsideOrEdge = oMeasurements.Item("DimPt23ToBottomInside")
        Else
            distToFlangeInsideOrEdge = oMeasurements.Item("DimPt15ToTopInside")
        End If
    Else
        If bIsBottomCut Then
            distToFlangeInsideOrEdge = oMeasurements.Item("DimPt23ToFR")
            distToWeb = oMeasurements.Item("DimPt23ToWR")
        Else
            distToFlangeInsideOrEdge = oMeasurements.Item("DimPt15ToFL")
            distToWeb = oMeasurements.Item("DimPt15ToWL")
        End If
    End If
    
    ' ----------------------------------------------------------------------------------
    ' If a web cut, default offset is twice the flange thickness, or 25mm if not flanged
    ' ----------------------------------------------------------------------------------
    Dim defaultEdgeSetback As Double
    
    If eFeatureType = SF_WebCut Then
        If HasBottomFlange(oPortBounded.Connectable) Or HasTopFlange(oPortBounded.Connectable) Then
            defaultEdgeSetback = 2 * GetWebOrFlangeThickness(oPRL, bIsBottomCut)
        Else
            defaultEdgeSetback = 0.025
        End If
    
    ' ----------------------------------------------------------
    ' If a flange cut, default offset is twice the web thickness
    ' ----------------------------------------------------------
    Else
        defaultEdgeSetback = 2 * GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    End If
    
    dEdgeSetback = defaultEdgeSetback
    
    ' ------------
    ' For web cuts
    ' ------------
    If eFeatureType = SF_WebCut Then
        ' -----------------------------------------------------------------------------------------------------
        ' If flanged and not outside the bounding object, make the default offset from the inside of the flange
        ' -----------------------------------------------------------------------------------------------------
        If bIsBottomCut And HasBottomFlange(oPortBounded.Connectable) And Not outsideIntersectedEdge = Below Then
            dEdgeSetback = distToFlangeInsideOrEdge + defaultEdgeSetback + dClearance
        ElseIf (Not bIsBottomCut) And HasTopFlange(oPortBounded.Connectable) And Not outsideIntersectedEdge = Above Then
            dEdgeSetback = distToFlangeInsideOrEdge + defaultEdgeSetback + dClearance
        ' ----------------------------------------------------------------------------------------------------------
        ' If not flanged, and not outside the bounding object, make it the distance to the top + the default setback
        ' ----------------------------------------------------------------------------------------------------------
        ElseIf bIsBottomCut And Not HasBottomFlange(oPortBounded.Connectable) And Not outsideIntersectedEdge = Below Then
            dEdgeSetback = distToFlangeInsideOrEdge + defaultEdgeSetback
        ElseIf Not bIsBottomCut And Not HasTopFlange(oPortBounded.Connectable) And Not outsideIntersectedEdge = Above Then
            dEdgeSetback = distToFlangeInsideOrEdge + defaultEdgeSetback
        End If
    ' ---------------
    ' For flange cuts
    ' ---------------
    Else
        ' -----------------------------------------------------------------------------------------------------------------------
        ' If the web is inside by less than the default amount, make it the distance to the far side of the web + default setback
        ' -----------------------------------------------------------------------------------------------------------------------
        If (bIsBottomCut And Not farWebIntersectedEdge = Below) Or (Not bIsBottomCut And Not farWebIntersectedEdge = Above) Then
            If distToWeb < defaultEdgeSetback Then
                dEdgeSetback = distToWeb + dWebOrFlangeThickness + defaultEdgeSetback
            End If
        End If
    End If
    
    ' ------------------------
    ' Set slope outside to 90°
    ' ------------------------
    Dim dPI As Double
    dPI = Atn(1) * 4#

    dSlopeOutside = dPI / 2
    
    ' ---------------------------------------------
    ' Set offset from the face to a constant 0.01mm
    ' ---------------------------------------------
    dFaceOffset = 0.00001
    
    ' ---------------------------------------------------------
    ' Set the slope to 90° or 60°, depending on selector answer
    ' ---------------------------------------------------------
    Dim sShapeOutSide As String
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oPRL.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    
    If sShapeOutSide = "" Then
        If bIsBottomCut Then
            GetSelectorAnswer oParentObj, "BottomShapeOutside", sShapeOutSide
        Else
            GetSelectorAnswer oParentObj, "TopShapeOutside", sShapeOutSide
        End If
    End If
    
    If sShapeOutSide = gsStraight Then
        dSlope = 89.99 * dPI / 180 'with 90 degrees noticed asserts, so gave very close value
    Else
        dSlope = dPI / 4#
    End If

    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    dCuttingDepth = EndCut_GetCutDepth(oPRL)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' Amount to offset input web curve to represent the flange edge on the same side
' Example: if input curve is web left, offset is to flange left
'          if no left flange, offset is set to near-zero value
' Intended for top flange cuts that have web left as input and have an offset construction curve to represent flange left, and
'              bottom flange cuts that have web right as input and have an offset construction curve to represent flange right
Private Function GetFlangeOffset(oPRL As IJDParameterLogic, bIsBottomCut As Boolean) As Double

    Const METHOD = m_sClassName & "::GetFlangeOffset"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Getting the flange width"

    ' ---------------------------------------------
    ' Determine if this is the top or bottom flange
    ' ---------------------------------------------
    Dim bIsBottomFlange As Boolean
    Dim sBottomFlange As String
    
    GetSelectorAnswer oPRL, "BottomFlange", sBottomFlange
    
    If sBottomFlange = gsYes Then
        bIsBottomFlange = True
    Else
        bIsBottomFlange = False
    End If

    ' ---------------------------------------
    ' Get the width as defined by the section
    ' ---------------------------------------
    Dim oPortBounded As IJPort
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)

    Dim oBoundedObject As Object
    Set oBoundedObject = oPortBounded.Connectable
    
    Dim dSectionWidth As Double
    Dim dWebThickness As Double
    
    If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBoundedObject
        dSectionWidth = oBoundedMemberPart.FlangeLength
        dWebThickness = oBoundedMemberPart.WebThickness
    ElseIf TypeOf oBoundedObject Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oBoundedObject
        dSectionWidth = oBoundedProfilePart.FlangeLength
        dWebThickness = oBoundedProfilePart.WebThickness
    End If

    ' ----------------------------------
    ' If no flange, set to web thickness
    ' ----------------------------------
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedObject, bTFL, bBFL, bTFR, bBFR
    
    If bIsBottomFlange Then
        If bIsBottomCut Then
            If Not bBFR Then
                GetFlangeOffset = 0.00001
            Else
                If bBFL Then
                    GetFlangeOffset = (dSectionWidth - dWebThickness) / 2
                Else
                    GetFlangeOffset = dSectionWidth - dWebThickness
                End If
            End If
        Else
            If Not bBFL Then
                GetFlangeOffset = 0.00001
            Else
                If bBFR Then
                    GetFlangeOffset = (dSectionWidth - dWebThickness) / 2
                Else
                    GetFlangeOffset = dSectionWidth - dWebThickness
                End If
            End If
        End If
    Else
        If bIsBottomCut Then
            If Not bTFR Then
                GetFlangeOffset = 0.00001
            Else
                If bTFL Then
                    GetFlangeOffset = (dSectionWidth - dWebThickness) / 2
                Else
                    GetFlangeOffset = dSectionWidth - dWebThickness
                End If
            End If
        Else
            If Not bTFL Then
                GetFlangeOffset = 0.00001
            Else
                If bTFR Then
                    GetFlangeOffset = (dSectionWidth - dWebThickness) / 2
                Else
                    GetFlangeOffset = dSectionWidth - dWebThickness
                End If
            End If
        End If
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

