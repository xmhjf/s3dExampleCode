VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TubeService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'    Tube Template service for Creation of BasePlane, ControlLine, TemplateOutContour etc.
'    which can be customized by shipyard.
'
'Description:
'History :
'   Siva              20th Oct 2011      Creation
'******************************************************************************************************************
Option Explicit
Implements IJDMfgTemplateService
Implements IJDMfgTemplateService2

' Method: - How the extension is measured;
Private Enum EnumExtnMethod
    Linear = 0
    Perpendicular
    AlongEdge
End Enum

' Measure: - how to measure the extension offset to the tube edge.
Private Enum EnumExtnMeasure
    AlongGirth = 0
    AlongAxis
End Enum

Private Const MAX_LENGTH = 1000

Private Const MODULE = "StrMfgTemplateProcessTube.TubeService"

Private Sub Class_Initialize()
 Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'********************************************************************************************************************************
' Function Name:    CreateBasePlane
' Interface:        IJDMfgTemplateService
' Inputs:           TubePart -- selected part
'                   oProcessSettings -- Template settings, type, orientation, direction, side etc.
' Output:           Object - Baseplane IJPlane
' Assumption:
' Algorithm:        1.
'                   2.
'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateBasePlane(ByVal oMemberPart As Object, ByVal oProcessSettings As Object, ByVal oTemplateSet As Object, Optional bUserDefined As Boolean = False) As Object
 Const METHOD = "IJDMfgTemplateService_CreateBasePlane"
    On Error GoTo ErrorHandler

    ' Create the needed services
    Dim oGeomHelper As GSCADStrMfgUtilities.MfgGeomHelper
    Set oGeomHelper = New GSCADStrMfgUtilities.MfgGeomHelper
    
    bUserDefined = False
    
    Dim oTemplateHelper As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    ' Create the Template set data and cache the data
    Dim oTemplSetData As TemplSetData
    Set oTemplSetData = New TemplSetData
    oTemplSetData.InitSettings oMemberPart, oProcessSettings, oTemplateSet

    ' Get the tube surface of specific side
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = GetTubeSurfaceInfo(oMemberPart, oTemplSetData.TemplateSide)
    
    ' get the tube short and long points info for creating base plane with min distance
    Dim oShortPointLC As IJDPosition, oLongPointLC As IJDPosition, oShortPointFeature As IJDPosition, oLongPointFeature As IJDPosition
    GetTubeShortAndLongPointsInfo oMemberPart, oTemplSetData.TemplateSide, oTemplSetData.TemplateSideEnd, oShortPointLC, oLongPointLC, oShortPointFeature, oLongPointFeature
    
    ' Find end points of the LandingCurve
    Dim oStartPos       As IJDPosition
    Dim oEndPos         As IJDPosition
    Dim oBasePlanePos   As IJDPosition
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    ' Get the minimum height
    Dim dAlong  As Double
    dAlong = oTemplSetData.MinHeight
        
    ' Get the reference vector(which always points to tube end) for constructing base plane
    Dim oRefVecForBPNormal As IJDVector
    
    Dim oIsoCurveAtBCL     As IJCurve
    Set oIsoCurveAtBCL = oTemplateHelper.GetIsoCurveAtPointOnSurface(oShortPointFeature, oSurfaceBody)
    
    Dim oRefWire    As IJWireBody
    Dim oRefCurve   As IJCurve
    Dim oRefPos     As IJDPosition
    
    ' Check if the isocurve is finite(i.e., curved tube) or infinite(i.e., straight tube)
    If oIsoCurveAtBCL.Length < MAX_LENGTH Then 'Curved tube
        Set oRefWire = oMfgRuleHelper.ComplexStringToWireBody(oIsoCurveAtBCL)
        Set oRefCurve = oIsoCurveAtBCL
        Set oRefPos = oShortPointFeature
    Else ' Straight tube
        ' Get the landing curve of the member
        Dim oProfileSupport As IJProfilePartSupport
        Set oProfileSupport = New ProfilePartSupport
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oProfileSupport
        Set oPartSupport.Part = oMemberPart
        
        Dim oLandingCurve As IJWireBody
        Dim oThicknessDirection As IJDVector
        oProfileSupport.GetProfilePartLandingCurve oRefWire, oThicknessDirection, False, SideA
        Set oRefPos = oShortPointLC
        Set oRefCurve = oMfgRuleHelper.WireBodyToComplexString(oRefWire)
    End If
        
    oRefWire.GetEndPoints oStartPos, oEndPos
    
    ' Check if short point is nearer to start point or end point
    If oRefPos.DistPt(oStartPos) < oRefPos.DistPt(oEndPos) Then
        Set oBasePlanePos = oMfgRuleHelper.GetPointAlongCurveAtDistance(oRefWire, oRefPos, dAlong, oEndPos)
        Set oRefVecForBPNormal = oBasePlanePos.Subtract(oStartPos)
    Else
        Set oBasePlanePos = oMfgRuleHelper.GetPointAlongCurveAtDistance(oRefWire, oRefPos, dAlong, oStartPos)
        Set oRefVecForBPNormal = oBasePlanePos.Subtract(oEndPos)
    End If
    
    oRefVecForBPNormal.Length = 1
    
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    
    Dim oTangentAtBPP    As IJDVector   ' Tangent at Base Plane Position
    Set oTangentAtBPP = oGeomHelper.GetTangentByPointOnCurve(oRefCurve, oBasePlanePos)
    
    oTangentAtBPP.Length = 1
        
    Dim strBasePlaneType     As String
    strBasePlaneType = oTemplSetData.TemplateBasePlane
    
    If strBasePlaneType = "BySystem" Or strBasePlaneType = "NormalToAxis" Then
        '''''''''''''''''''''''''''''''''''''''''''''
        '"System" / "NormalToAxis"'''
        '''''''''''''''''''''''''''''''''''''''''''''
        
        If oRefVecForBPNormal.Dot(oTangentAtBPP) < 0 Then
            oTangentAtBPP.Length = -1
        End If
        
        oBasePlane.DefineByPointNormal oBasePlanePos.X, oBasePlanePos.Y, oBasePlanePos.Z, oTangentAtBPP.X, oTangentAtBPP.Y, oTangentAtBPP.Z

    ElseIf strBasePlaneType = "Global" Then
        
        ' Construct base plane in Global axis
        Dim oXVec As IJDVector
        Set oXVec = New DVector

        Dim oYVec As IJDVector
        Set oYVec = New DVector

        Dim oZVec As IJDVector
        Set oZVec = New DVector

        oXVec.Set 1, 0, 0
        oYVec.Set 0, 1, 0
        oZVec.Set 0, 0, 1

        Dim dDPValueX     As Double
        dDPValueX = Abs(oTangentAtBPP.Dot(oXVec))

        Dim dDPValueY     As Double
        dDPValueY = Abs(oTangentAtBPP.Dot(oYVec))

        Dim dDPValueZ     As Double
        dDPValueZ = Abs(oTangentAtBPP.Dot(oZVec))

        Dim oNormal    As IJDVector

        If dDPValueX > dDPValueY And dDPValueX > dDPValueZ Then
           Set oNormal = oXVec
        ElseIf dDPValueY > dDPValueZ Then
           Set oNormal = oYVec
        Else
           Set oNormal = oZVec
        End If

        If oRefVecForBPNormal.Dot(oNormal) < 0 Then
            oNormal.Length = -1
        End If

        oBasePlane.DefineByPointNormal oBasePlanePos.X, oBasePlanePos.Y, oBasePlanePos.Z, oNormal.X, oNormal.Y, oNormal.Z
        
    ElseIf strBasePlaneType = "UserDefined" Then
            bUserDefined = True
            Set IJDMfgTemplateService_CreateBasePlane = Nothing
            Exit Function
    End If

    If Not oBasePlane Is Nothing Then
        
        Dim dRootX As Double, dRootY As Double, dRootZ As Double
        Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
        
        oBasePlane.GetRootPoint dRootX, dRootY, dRootZ
        oBasePlane.GetNormal dNormalX, dNormalY, dNormalZ
        
        Dim oStart As Object, oEnd As Object, oBPWire As Object
        oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBasePlane, oBPWire, oStart, oEnd
        
        Dim oGeomElements As IJElements
        Set oGeomElements = New JObjectCollection
        oGeomElements.Add oBPWire

        ' The base plane(oBasePlane) that we got before is infinite one, it needs to be bounded by range of the plate with some tolerance
        Set IJDMfgTemplateService_CreateBasePlane = oTemplateHelper.MakeTemplateBasePlane(oGeomElements, _
                                                                                          dRootX, dRootY, dRootZ, _
                                                                                          dNormalX, dNormalY, dNormalZ, False)
        
    Else
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FaliedToCreateBasePlane, , "RULES")
        Exit Function
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FaliedToCreateBasePlane, , "RULES")
End Function

'********************************************************************************************************************************
' Function Name:    CreateControlLine
' Interface:        IJDMfgTemplateService
' Inputs:           TubePart -- selected plate
'                   oProcessSettings -- Template settings, type, orientation, direction, side etc.
' Output:           Object - Baseplane IJPlane
' Assumption:
' Algorithm:        1. Check the Template Side which is BaseSide or OffsetSide

'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateControlLine(ByVal oMemberPart As Object, ByVal oProcessSettings As Object, ByVal oBasePlane As Object, ByVal oTemplateSet As Object) As Object
Const METHOD = "IJDMfgControlLineService_CreateControlLine"
    On Error GoTo ErrorHandler
    
    ' Create the needed services
    Dim oGeomHelper As GSCADStrMfgUtilities.MfgGeomHelper
    Set oGeomHelper = New GSCADStrMfgUtilities.MfgGeomHelper
    
    Dim oTemplateHelper As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    Dim oMfgGeomUtilwrapper As IJDMfgGeomUtilWrapper
    Set oMfgGeomUtilwrapper = New MfgGeomUtilWrapper
    
    Dim oMfgMGHelper As GSCADMathGeom.MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    ' Create template set data and cache the settings data
    Dim oTemplSetData   As TemplSetData
    Set oTemplSetData = New TemplSetData
    
    oTemplSetData.InitSettings oMemberPart, oProcessSettings, oTemplateSet
    
    Dim strTemplateSide As String
    Dim strTemplateSideEnd As String
    
    strTemplateSide = oTemplSetData.TemplateSide
    strTemplateSideEnd = oTemplSetData.TemplateSideEnd
    
    ' Get the tube surface and tube end cut contours info.
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = GetTubeSurfaceInfo(oMemberPart, strTemplateSide)
    
    Dim oEndCutElems    As IJElements
    Set oEndCutElems = GetTubeEndCutContoursInfo(oMemberPart, strTemplateSide, strTemplateSideEnd)
    
    Dim oModelBody  As IJDModelBody
    Set oModelBody = oEndCutElems.Item(1)
        
    Dim oRefPos             As IJDPosition
    Dim strTemplateType     As String
    strTemplateType = oTemplSetData.TemplateType
    
    ' If type is model origin then get the corresponding position on end cut
    If strTemplateType = "ModelOrigin" Then
        
        Dim oProfileSupport As IJProfilePartSupport
        Set oProfileSupport = New ProfilePartSupport
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oProfileSupport
        Set oPartSupport.Part = oMemberPart
    
        ' Find the LandingCurve and thickness direction of the web
        Dim oLandingCurve As IJWireBody
        Dim oThicknessDirection As IJDVector
        oProfileSupport.GetProfilePartLandingCurve oLandingCurve, oThicknessDirection, False, SideA
    
        ' Find endpoints of the LandingCurve
        Dim oStartPos   As IJDPosition
        Dim oEndPos     As IJDPosition
        oLandingCurve.GetEndPoints oStartPos, oEndPos
        
        Dim oStartClosePos  As IJDPosition, oEndClosePos As IJDPosition
        Dim dStartMinDist   As Double, dEndMinDist As Double
        
        oModelBody.GetMinimumDistanceFromPosition oStartPos, oStartClosePos, dStartMinDist
        oModelBody.GetMinimumDistanceFromPosition oEndPos, oEndClosePos, dEndMinDist
        
        If dStartMinDist < dEndMinDist Then
            Set oRefPos = oStartClosePos
        Else
            Set oRefPos = oEndClosePos
        End If
    
    ElseIf strTemplateType = "ShortDistanceBCL" Or strTemplateType = "LongDistanceBCL" Then
    
        Dim oShortPointLC As IJDPosition, oLongPointLC As IJDPosition, oShortPointFeature As IJDPosition, oLongPointFeature As IJDPosition
        GetTubeShortAndLongPointsInfo oMemberPart, oTemplSetData.TemplateSide, oTemplSetData.TemplateSideEnd, oShortPointLC, oLongPointLC, oShortPointFeature, oLongPointFeature
    
        If strTemplateType = "ShortDistanceBCL" Then
            Set oRefPos = oShortPointFeature    ' Short Point
        Else
            Set oRefPos = oLongPointFeature     ' Long Point
        End If
    Else
        ' logic for getting global min/max positions
        
        Dim oXVector                As IJDVector
        Dim oYVector                As IJDVector
        Dim oZVector                As IJDVector
    
        Dim oMinBoxPoints           As IJElements
        Dim Points(1 To 4)          As IJDPosition
                 
        Dim oVectorElems            As IJElements
        Set oVectorElems = New JObjectCollection
        
        'Create vectors of unit length
        Set oXVector = New DVector
        Set oYVector = New DVector
        Set oZVector = New DVector
        
        'Normalize the vectors
        oXVector.Set 1, 0, 0
        oYVector.Set 0, 1, 0
        oZVector.Set 0, 0, 1
        
        oXVector.Length = 1#
        oYVector.Length = 1#
        oZVector.Length = 1#
        
        oVectorElems.Add oXVector
        oVectorElems.Add oYVector
        oVectorElems.Add oZVector
        
        ' Construct min box for the surface in standard x, y, z directions
        Set oMinBoxPoints = oGeomHelper.GetGeometryMinBoxByVectors(oEndCutElems, oVectorElems)
        
        Set Points(1) = oMinBoxPoints.Item(1)
        Set Points(2) = oMinBoxPoints.Item(2)
        Set Points(3) = oMinBoxPoints.Item(3)
        Set Points(4) = oMinBoxPoints.Item(4)
        
        Dim oMinMaxRefPos  As IJDPosition
        Dim oMinMaxPlane    As IJPlane
        Set oMinMaxPlane = New Plane3d
        
        If strTemplateType = "Global_X_Max" Then
            If Points(1).X > Points(2).X And Points(1).X > Points(3).X And Points(1).X > Points(4).X Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).X > Points(3).X And Points(2).X > Points(4).X Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).X > Points(4).X Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 1, 0, 0
            
        ElseIf strTemplateType = "Global_X_Min" Then
            If Points(1).X < Points(2).X And Points(1).X < Points(3).X And Points(1).X < Points(4).X Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).X < Points(3).X And Points(2).X < Points(4).X Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).X < Points(4).X Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 1, 0, 0
            
        ElseIf strTemplateType = "Global_Y_Max" Then
            If Points(1).X > Points(2).Y And Points(1).Y > Points(3).Y And Points(1).Y > Points(4).Y Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).Y > Points(3).Y And Points(2).Y > Points(4).Y Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).Y > Points(4).Y Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 0, 1, 0
            
        ElseIf strTemplateType = "Global_Y_Min" Then
            If Points(1).Y < Points(2).Y And Points(1).Y < Points(3).Y And Points(1).Y < Points(4).Y Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).Y < Points(3).Y And Points(2).Y < Points(4).Y Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).Y < Points(4).Y Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 0, 1, 0
            
        ElseIf strTemplateType = "Global_Z_Max" Then
            If Points(1).Z > Points(2).Z And Points(1).Z > Points(3).Z And Points(1).Z > Points(4).Z Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).Z > Points(3).Z And Points(2).Z > Points(4).Z Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).Z > Points(4).Z Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 0, 0, 1
            
        ElseIf strTemplateType = "Global_Z_Min" Then
            If Points(1).Z < Points(2).Z And Points(1).Z < Points(3).Z And Points(1).Z < Points(4).Z Then
                Set oMinMaxRefPos = Points(1)
            ElseIf Points(2).Z < Points(3).Z And Points(2).Z < Points(4).Z Then
                Set oMinMaxRefPos = Points(2)
            ElseIf Points(3).Z < Points(4).Z Then
                Set oMinMaxRefPos = Points(3)
            Else
                Set oMinMaxRefPos = Points(4)
            End If
            
            oMinMaxPlane.DefineByPointNormal oMinMaxRefPos.X, oMinMaxRefPos.Y, oMinMaxRefPos.Z, 0, 0, 1
        End If
        
        Dim oPlaneSheetBody   As Object
        Set oPlaneSheetBody = oMfgGeomUtilwrapper.CreateFinitePlane(oMinMaxPlane)
            
        Dim oClosePos1  As IJDPosition, oClosePos2   As IJDPosition
        Dim dMinDist    As Double
        
        oModelBody.GetMinimumDistance oPlaneSheetBody, oClosePos1, oClosePos2, dMinDist
        
        Set oRefPos = oClosePos1
    End If
    
    ' get the base plane wire or top curve
    Dim oStart As Object, oEnd As Object, oBPWire As Object
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBasePlane, oBPWire, oStart, oEnd
    
    Dim oBPCurveCS  As IJComplexString
    Set oBPCurveCS = oMfgRuleHelper.WireBodyToComplexString(oBPWire)
    
    ' Create 3D control line
    Dim oControlLine As IJCurve
    Set oControlLine = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oRefPos, oBPCurveCS, oSurfaceBody)
    
    ' Make sure the BCL position is the start of the control line for later access
    Dim bAtStart As Boolean
    oMfgMGHelper.IsPositionedAtStart oControlLine, oRefPos, bAtStart
    
    If bAtStart = False Then
        Dim oRevCtrlLine    As IJComplexString
        oMfgMGHelper.ReverseComplexString oControlLine, oRevCtrlLine
        Set oControlLine = oRevCtrlLine
    End If
    
    ' Return the Base control curve
    If Not oControlLine Is Nothing Then
        Set IJDMfgTemplateService_CreateControlLine = oControlLine
    Else
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FaliedToCreateBaseControlLine, , "RULES"
        Err.Raise vbObjectError + TPL_CBCL_FaliedToCreateBaseControlLine
        Exit Function
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FaliedToCreateBaseControlLine, , "RULES")
End Function

Private Function IJDMfgTemplateService_CreateMarkingLine(ByVal oMemberPart As Object, ByVal oProcessSettings As Object, ByVal pBasePlane As Object) As Object

End Function

Private Function IJDMfgTemplateService_CreateSightLines(ByVal pTemplateSet As Object, ByVal numOfSightLine As Long, ByVal pMarkingSetting As Object) As Object
Const METHOD = "IJDMfgSightLineService_CreateSightLines"
On Error GoTo ErrorHandler

   Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************************************************************************
' Function Name:    CreateTemplateOutContour
'
' Interface:        IJDMfgTemplateService
'
' Inputs:           TubePart -- selected plate
'                   TemplateSet -- TemplateSet object created by semantic
'                   ProcessSetting -- contains the settings for TemplateSet ex: Type, Side etc
'                   SketchingLines -- Collection of lines sketched by user
'
' Output:           Collection of Geom3d Objects.
'
' Assumption:       Within the output collection, a set of 4 objects make one template. each object represent
'                   the bottom curve, side, top , side of the template
'
' Algorithm:        1. GetSettingValues From ProcessRules
'********************************************************************************************************************************

Private Function IJDMfgTemplateService_CreateTemplateOutContour(ByVal oMemberPart As Object, ByVal pUnkTemplateSet As Object, ByVal oProcessSettings As Object, ByVal pSketchingLines As Object) As Object
    Const METHOD = "IJDMfgTemplateService_CreateTemplateOutContour"
    On Error Resume Next
    
    ' Create the needed services
    Dim oTemplateInfo As IJElements
    Set oTemplateInfo = New JObjectCollection
    
    Dim oGeomHelper     As GSCADStrMfgUtilities.MfgGeomHelper
    Set oGeomHelper = New GSCADStrMfgUtilities.MfgGeomHelper
    
    Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oTemplateHelper As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    Dim oRuleHelper     As MfgRuleHelpers.Helper
    Set oRuleHelper = New MfgRuleHelpers.Helper
    
    ' Create template set data and cache the settings data
    Dim oTemplSetData   As TemplSetData
    Set oTemplSetData = New TemplSetData
    oTemplSetData.InitSettings oMemberPart, oProcessSettings, pUnkTemplateSet

    ' Create 3D bottom line
    Dim strTemplateSide As String
    Dim strTemplateSideEnd As String
    
    strTemplateSide = oTemplSetData.TemplateSide
    strTemplateSideEnd = oTemplSetData.TemplateSideEnd
    
    ' Get the tube end cut contour information
    Dim oEdgeContourElems   As IJElements
    Set oEdgeContourElems = GetTubeEndCutContoursInfo(oMemberPart, strTemplateSide, strTemplateSideEnd)

    Dim oBottomCurve    As IJComplexString
    Set oBottomCurve = oRuleHelper.WireBodyToComplexString(oEdgeContourElems.Item(1))
    oTemplateInfo.Add oBottomCurve

    ' Create 3D top line by intersecting the base plane with the tube surface
    Dim oSurfaceBody   As IJSurfaceBody
    Set oSurfaceBody = GetTubeSurfaceInfo(oMemberPart, oTemplSetData.TemplateSide)

    Dim oBasePlane As IJPlane
    Dim oTemplateSet As IJDMfgTemplateSet

    Set oTemplateSet = pUnkTemplateSet
    Set oBasePlane = oTemplateSet.GetBasePlane
    
    Dim oTopCurveWire   As IJWireBody
    Dim oTopCurve       As IJCurve
    Dim oDummyPos       As IJDPosition
    
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBasePlane, oTopCurveWire, oDummyPos, oDummyPos
    
    Set oTopCurve = oRuleHelper.WireBodyToComplexString(oTopCurveWire)
    
    oTemplateInfo.Add oTopCurve
    
    ' Create Side lines(iso lines on surface)
    
    ' Get the short and long points information
    Dim oShortPointLC As IJDPosition, oLongPointLC As IJDPosition, oShortPointFeature As IJDPosition, oLongPointFeature As IJDPosition
    GetTubeShortAndLongPointsInfo oMemberPart, oTemplSetData.TemplateSide, oTemplSetData.TemplateSideEnd, oShortPointLC, oLongPointLC, oShortPointFeature, oLongPointFeature
    
    Dim oControlLine    As IJCurve
    Set oControlLine = oTemplateSet.GetControlLine
    
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    oControlLine.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
    Dim oBCLPos As IJDPosition
    Set oBCLPos = New DPosition
        
    oBCLPos.Set dStartX, dStartY, dStartZ
    
'    Method: - How the extension is measured;
'            Linear extension - 0
'            Perpendicular - 1
'            Along Edge - 2
'    Measure: - how to measure the extension offset to the tube edge.
'            Along girth - 0
'            Along tube axis - 1

    Dim eMethod_Left As EnumExtnMethod
    Dim eMethod_Right As EnumExtnMethod
    
    ' Get the extension info from settings
    Dim eMeasure_Left As EnumExtnMeasure
    Dim eMeasure_Right As EnumExtnMeasure
        
    Dim dExtnValue1_Left As Double
    Dim dExtnValue1_Right As Double
    
    dExtnValue1_Left = oTemplSetData.TemplateExtnValueLeft
    dExtnValue1_Right = oTemplSetData.TemplateExtnValueRight
    
    eMeasure_Left = oTemplSetData.TemplateMeasureLeft
    eMeasure_Right = oTemplSetData.TemplateMeasureRight
    
    eMethod_Left = oTemplSetData.TemplateMethodLeft
    eMethod_Right = oTemplSetData.TemplateMethodRight
    
    Dim oIsoCurve       As IJCurve
    Dim oTempCurve      As IJCurve
    Dim oIsoCS          As IJComplexString
    Dim dMinDist        As Double
    
    Dim oRefBCLPos      As IJDPosition
    Set oRefBCLPos = New DPosition
    
    
    ' **** Create reference curve ****
    Dim oRefCurveCS   As IJComplexString
    Dim oRefCurveWB As Object
    
    Set oRefCurveWB = GetReferenceCurve(oMemberPart, oSurfaceBody, oBasePlane, oTopCurve)
    Set oRefCurveCS = oRuleHelper.WireBodyToComplexString(oRefCurveWB)
        
    ' check if the extension is ignore
    If Not dExtnValue1_Left = 0 And Not dExtnValue1_Right = 0 Then
        Dim oRefCurve       As IJCurve
        Dim oTrimPos1       As IJDPosition
        Dim oTrimPos2       As IJDPosition
        
        Dim dRefCurveLen    As Double
        Dim dLeftExtnVal    As Double
        Dim dRightExtnVal   As Double
        Dim bAtStart        As Boolean
        
        Dim oLeftPos        As IJDPosition
        Dim oLeftIsoCS      As IJComplexString
        Dim oRightPos       As IJDPosition
        Dim oRightIsoCS     As IJComplexString
        
        Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
        
        If Not dExtnValue1_Left = 0 Then
            
            ' Check measure left is along axis
            If eMeasure_Left = AlongAxis Then  ' Along axis
                Set oRefCurve = oRefCurveCS
                dRefCurveLen = oRefCurve.Length
                
                dLeftExtnVal = dRefCurveLen * (0.5 - (dExtnValue1_Left / 360))
                
                Set oIsoCS = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oBCLPos, oRefCurveCS, oSurfaceBody)
                Set oIsoCurve = oIsoCS
                
                oIsoCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        
                oMfgMGHelper.IsPositionedAtStart oIsoCS, oBCLPos, bAtStart
                
                If bAtStart = True Then
                    oRefBCLPos.Set dEndX, dEndY, dEndZ
                Else
                    oRefBCLPos.Set dStartX, dStartY, dStartZ
                End If
                
                Set oLeftPos = oGeomHelper.GetPointAtDistAlongCurve(oRefCurve, oRefBCLPos, dLeftExtnVal)
                
                If oLeftPos Is Nothing Then
                    Set oLeftPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oRefBCLPos, dLeftExtnVal, oRefBCLPos)
                End If

                Set oTempCurve = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oLeftPos, oBottomCurve, oSurfaceBody)
                
                oTempCurve.DistanceBetween oTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                
                Set oTrimPos1 = New DPosition
                oTrimPos1.Set dInX, dInY, dInZ
                
                oTempCurve.DistanceBetween oBottomCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                Set oTrimPos2 = New DPosition
                oTrimPos2.Set dInX, dInY, dInZ
                
                oMfgMGHelper.TrimCurveByPoints oTempCurve, oTrimPos1, oTrimPos2
                
                oTemplateInfo.Add oTempCurve
                
                Set oLeftIsoCS = oTempCurve
                'Dim FileName As String
                                'FileName = Environ("TEMP")
                                'If FileName = "" Or FileName = vbNullString Then
                                '       FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
                                'End If
                                'FileName = FileName & "\LeftSideLine.txt"
                'oGeomHelper.PrintComplexString oLeftIsoCS, FileName
                
            ElseIf eMeasure_Left = AlongGirth Then  ' along girth
                
                Set oRefCurve = oBottomCurve
                dRefCurveLen = oRefCurve.Length
                
                dLeftExtnVal = 0.5 * dRefCurveLen - dExtnValue1_Left
                Set oLeftPos = oGeomHelper.GetPointAtDistAlongCurve(oRefCurve, oBCLPos, dLeftExtnVal)
                                
                If oLeftPos Is Nothing Then
                    Set oRefCurveWB = oRuleHelper.ComplexStringToWireBody(oRefCurve)
                    Set oLeftPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oBCLPos, dLeftExtnVal, oBCLPos)
                End If
                
                Set oLeftIsoCS = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oLeftPos, oTopCurve, oSurfaceBody)
                oTemplateInfo.Add oLeftIsoCS
            End If
            
        End If
        
        If Not dExtnValue1_Right = 0 Then
            
            ' Check measure right is along axis
            If eMeasure_Right = AlongAxis Then  ' along axis
            
                Set oRefCurve = oRefCurveCS
                dRefCurveLen = oRefCurve.Length
                
                dRightExtnVal = dRefCurveLen * (0.5 + (dExtnValue1_Right / 360))
                
                Set oIsoCS = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oBCLPos, oRefCurveCS, oSurfaceBody)
                
                Set oIsoCurve = oIsoCS
                oIsoCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        
                oMfgMGHelper.IsPositionedAtStart oIsoCS, oBCLPos, bAtStart
                
                If bAtStart = True Then
                    oRefBCLPos.Set dEndX, dEndY, dEndZ
                Else
                    oRefBCLPos.Set dStartX, dStartY, dStartZ
                End If
                
                Set oRightPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oRefBCLPos, dRightExtnVal, oLeftPos)
                
                Set oTempCurve = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oRightPos, oBottomCurve, oSurfaceBody)
                
                oTempCurve.DistanceBetween oTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                
                Set oTrimPos1 = New DPosition
                oTrimPos1.Set dInX, dInY, dInZ
                
                oTempCurve.DistanceBetween oBottomCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                Set oTrimPos2 = New DPosition
                oTrimPos2.Set dInX, dInY, dInZ
                
                oMfgMGHelper.TrimCurveByPoints oTempCurve, oTrimPos1, oTrimPos2
                
                Set oRightIsoCS = oTempCurve
                
                oTemplateInfo.Add oRightIsoCS
                Dim FileName As String
                                FileName = Environ("TEMP")
                                If FileName = "" Or FileName = vbNullString Then
                                        FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
                                End If
                                FileName = FileName & "\RightSideLine.txt"
                'oGeomHelper.PrintComplexString oTempCurve, FileName
            
            ElseIf eMeasure_Right = AlongGirth Then  ' along girth
                
                Set oRefCurve = oBottomCurve
                dRefCurveLen = oRefCurve.Length
                
                dRightExtnVal = (0.5 * dRefCurveLen + dExtnValue1_Right)
                Set oRightPos = oGeomHelper.GetPointAtDistAlongCurve(oRefCurve, oBCLPos, dRightExtnVal)
                
                If oRightPos Is Nothing Then
                    Set oRefCurveWB = oRuleHelper.ComplexStringToWireBody(oRefCurve)
                    Set oRightPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oBCLPos, dRightExtnVal, oBCLPos)
                End If
                
                Set oRightIsoCS = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oRightPos, oTopCurve, oSurfaceBody)
                
                oTemplateInfo.Add oRightIsoCS
                
            End If
            
        End If
        
    Else
        ' Extension Ignore logic
        
        ' Get the reference curve length
        Set oRefCurve = oRefCurveCS
        dRefCurveLen = oRefCurve.Length
        
        ' Get the trimmed iso curve at BCL position
        Set oIsoCS = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oBCLPos, oRefCurveCS, oSurfaceBody)
        
        Set oIsoCurve = oIsoCS
        oIsoCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        oMfgMGHelper.IsPositionedAtStart oIsoCS, oBCLPos, bAtStart
        
        If bAtStart = True Then
            oRefBCLPos.Set dEndX, dEndY, dEndZ
        Else
            oRefBCLPos.Set dStartX, dStartY, dStartZ
        End If
        
        ' Get the point opposite to the BCL position as there is no extension
        Dim oMidPos As IJDPosition
        Set oMidPos = oGeomHelper.GetPointAtDistAlongCurve(oRefCurve, oRefBCLPos, 0.5 * dRefCurveLen)
        
        ' if the above routine(wrapper to core IM routine) fails then call alternative(G&T wrapper)
        If oMidPos Is Nothing Then
            Set oMidPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oRefBCLPos, 0.5 * dRefCurveLen, oRefBCLPos)
        End If
        
        ' Get the iso curve at mid point
        Set oTempCurve = oTemplateHelper.GetTrimmedIsoCurveAtPointOnSurface(oMidPos, oBottomCurve, oSurfaceBody)
        oTempCurve.DistanceBetween oTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
        
        Set oTrimPos1 = New DPosition
        oTrimPos1.Set dInX, dInY, dInZ
        
        oTempCurve.DistanceBetween oBottomCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
        Set oTrimPos2 = New DPosition
        oTrimPos2.Set dInX, dInY, dInZ
        
        oMfgMGHelper.TrimCurveByPoints oTempCurve, oTrimPos1, oTrimPos2
        oTemplateInfo.Add oTempCurve
        
    End If
    
    Dim oTemplateColl As IJElements
    Set oTemplateColl = New JObjectCollection
    
    'Add GroupIndex and Sketched information
    Dim oTempPropColl As Collection
    Set oTempPropColl = New Collection
    
    Dim lGroupIndex     As Long
    Dim strKey          As String
    
    ' Group index is 1 as there is only one template and key is empty
    lGroupIndex = 1
    strKey = ""
    
    oTempPropColl.Add lGroupIndex
    oTempPropColl.Add strKey
    
    oTemplateInfo.Add oTempPropColl
    oTemplateColl.Add oTemplateInfo

    If Not oTemplateColl Is Nothing Then
        Set IJDMfgTemplateService_CreateTemplateOutContour = oTemplateColl
    Else
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
        GoTo CleanUp
    End If

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
End Function

Private Function IJDMfgTemplateService_CreateTemplatePlane(ByVal oMemberPart As Object, ByVal pUnkPosition As Object, ByVal oProcessSettings As Object, ByVal pDispTemplateSet As Object, Optional ByVal lGroupNumber As Long = 1&) As Object
Const METHOD = "IJDMfgTemplateService_CreateTemplatePlane"
On Error GoTo ErrorHandler

    ' No Implementation
    Dim strErrorMsg As String
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, strErrorMsg, "SMCustomWarningMessages", TPL_TC_FailedToCreateBottomLines, , "RULES")
End Function

Private Function IJDMfgTemplateService_ValidateProcessSettings(ByVal oProcessSettings As Object) As Boolean
Const METHOD = "IJDMfgTemplateService_ValidateProcessSettings"
On Error GoTo ErrorHandler

    ' No Implementation
    IJDMfgTemplateService_ValidateProcessSettings = True
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToValidateProcessSettings, , "RULES")
End Function

' ***********************************************************************************
' Public Function GetTubeShortAndLongPointsInfo()
'
' Description:  Gets the tube end cut's short and long points info
'
' ***********************************************************************************
Private Sub GetTubeShortAndLongPointsInfo(ByVal oMemberPart As Object, strTemplateSide As String, strTemplateSideEnd As String, oShortPointLC As IJDPosition, oLongPointLC As IJDPosition, oShortPointFeature As IJDPosition, oLongPointFeature As IJDPosition)
    Const METHOD = "GetTubeShortAndLongPointsInfo"
    On Error GoTo ErrorHandler
    
    Dim oFeatureContour     As IJWireBody
    Dim oPortElems          As IJElements
    Set oPortElems = GetTubeEndCutContoursInfo(oMemberPart, strTemplateSide, strTemplateSideEnd)
    
    Dim lCount As Long
    lCount = oPortElems.Count
    
    If lCount < 1 Then
        Exit Sub
    ElseIf lCount > 1 Then
    
        Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
        Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
        
        Dim oMfgGeomHelper As MfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper
        
        Dim oMergeElems As IJElements
        Set oMergeElems = oMfgGeomHelper.OptimizedMergingOfInputCurves(oPortElems)
        
        oMfgMGHelper.ComplexStringToWireBody oMergeElems.Item(1), oFeatureContour
        
    Else
        Set oFeatureContour = oPortElems.Item(1)
    End If

    Dim oTemplateHelper As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    On Error Resume Next
    oTemplateHelper.GetShortDistPointAndLongDistPointOnEndCut oMemberPart, oFeatureContour, oShortPointLC, oLongPointLC, oShortPointFeature, oLongPointFeature
    On Error GoTo ErrorHandler
    
    ' In case of curved tubes that are bounded, CTX_3DCUTOUT is given to base ports
    ' So if above routine fails try with CTX_3DCUTOUT
    If oShortPointLC Is Nothing Or oLongPointLC Is Nothing Then
        oTemplateHelper.GetShortDistPointAndLongDistPointOnEndCut oMemberPart, oFeatureContour, oShortPointLC, oLongPointLC, oShortPointFeature, oLongPointFeature
    End If
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

' ***********************************************************************************
' Public Function GetTubeEndCutContoursInfo()
'
' Description:  Gets the tube end cur contours from the side and side end inputs
'
' ***********************************************************************************
Private Function GetTubeEndCutContoursInfo(oMemberPart As Object, strTemplateSide As String, strTemplateSideEnd As String) As IJElements
    Const METHOD = "GetTubeEndCutContoursInfo"
    On Error GoTo ErrorHandler
    
    Dim oTemplateHelper     As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    Dim eCTXFlag    As eUSER_CTX_FLAGS
    
    If strTemplateSideEnd = "Base" Then
        eCTXFlag = CTX_BASE
    Else
        eCTXFlag = CTX_OFFSET
    End If

    Dim oSurfaceElements        As IJElements
    Dim oMemberPartPrismatic    As ISPSMemberPartPrismatic
    Set oMemberPartPrismatic = oMemberPart
    
    Dim strSectionType          As String
    strSectionType = oMemberPartPrismatic.CrossSection.definition.Type
    
    If Not strSectionType = "HSSR" Then

        Dim lSectionID  As Long
    
        If strTemplateSide = "Outer" Then
            lSectionID = 3073  ' Outer Tube
        Else
            lSectionID = 3074  ' Inner Tube
        End If
    
        On Error Resume Next
        Dim oContourElements    As IJElements
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lSectionID)
        
        ' In case of curved tubes that are bounded, CTX_3DCUTOUT is given to base ports
        ' So if above routine fails try with CTX_3DCUTOUT
        If oContourElements Is Nothing Then
            eCTXFlag = CTX_3DCUTOUT
            Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lSectionID)
        End If
        On Error GoTo ErrorHandler
        
        Set GetTubeEndCutContoursInfo = oContourElements
        
    Else
    
        ' The below code is for HSSR(square/Rectangular cross sections)
        ' This is not a round crosssection so get the individual edges as we expect more than one edge
        
        'JXSEC_BOTTOM = 513 (&H201)
        'JXSEC_TOP = 514 (&H202)
        'JXSEC_WEB_LEFT = 257 (&H101)
        'JXSEC_WEB_RIGHT = 258 (&H102)
        
        'JXSEC_INNER_TOP = 3585 (&HE01)
        'JXSEC_INNER_BOTTOM = 3586 (&HE02)
        'JXSEC_INNER_WEB_LEFT = 3329 (&HD01)
        'JXSEC_INNER_WEB_RIGHT = 3330 (&HD02)

        Dim lWebLeft    As Long
        Dim lWebRight    As Long
        Dim lTop    As Long
        Dim lBottom    As Long
        
        If strTemplateSide = "Outer" Then
            lWebLeft = 257
            lWebRight = 258
            lTop = 514
            lBottom = 513
        Else
            lWebLeft = 3329
            lWebRight = 3330
            lTop = 3585
            lBottom = 3586
        End If
        
        Dim oTempWB         As IJWireBody
        Dim oTempElems      As IJElements
        Set oTempElems = New JObjectCollection
        
        Dim oMfgGeomHelper  As MfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper
        
        Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
        Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
        
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lWebLeft)
        oTempElems.AddElements oContourElements
        Set oContourElements = Nothing
        
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lTop)
        oTempElems.AddElements oContourElements
        Set oContourElements = Nothing
        
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lWebRight)
        oTempElems.AddElements oContourElements
        Set oContourElements = Nothing
        
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lBottom)
        oTempElems.AddElements oContourElements
        Set oContourElements = Nothing
        
        If oTempElems.Count < 4 And strTemplateSide = "Inner" Then
            ' Try with outer sections IDs. This work around is needed as section IDs on inner edge ports is not proper.
            lWebLeft = 257
            lWebRight = 258
            lTop = 514
            lBottom = 513
            
            Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lWebLeft)
            oTempElems.AddElements oContourElements
            Set oContourElements = Nothing
            
            Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lTop)
            oTempElems.AddElements oContourElements
            Set oContourElements = Nothing
            
            Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lWebRight)
            oTempElems.AddElements oContourElements
            Set oContourElements = Nothing
            
            Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LEDGES, eCTXFlag, lBottom)
            oTempElems.AddElements oContourElements
            Set oContourElements = Nothing
            
        End If
        
        Dim oMergedElems    As IJElements
        Set oMergedElems = oMfgGeomHelper.OptimizedMergingOfInputCurves(oTempElems)
        
        If oMergedElems.Count > 1 Then
            ' This is work around fix for HSSR c/s.
            ' This fix is needed as some members do not have proper section IDs
            Dim lMaxLength  As Long, lMaxInd As Long
            Dim lMinLength  As Long, lMinInd As Long
            Dim iIndex      As Long
            
            lMinLength = 10000
            
            For iIndex = 1 To oMergedElems.Count
                Dim oCurve As IJCurve
                Set oCurve = oMergedElems.Item(iIndex)
                
                Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
                oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
                
                ' Check if the curve is closed
                If Abs(dStartX - dEndX) < 0.001 And Abs(dStartY - dEndY) < 0.001 And Abs(dStartZ - dEndZ) < 0.001 Then
                    If lMaxLength < oCurve.Length Then
                        lMaxLength = oCurve.Length
                        lMaxInd = iIndex
                    End If
                    
                    If lMinLength > oCurve.Length Then
                        lMinLength = oCurve.Length
                        lMinInd = iIndex
                    End If
                End If
            Next
            
            If strTemplateSide = "Outer" Then
                oMfgMGHelper.ComplexStringToWireBody oMergedElems.Item(lMaxInd), oTempWB
            Else
                oMfgMGHelper.ComplexStringToWireBody oMergedElems.Item(lMinInd), oTempWB
            End If
        Else
            oMfgMGHelper.ComplexStringToWireBody oMergedElems.Item(1), oTempWB
        End If
        
        oTempElems.Clear
        If Not oTempWB Is Nothing Then
            oTempElems.Add oTempWB
        End If
        
        Set GetTubeEndCutContoursInfo = oTempElems
    End If
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function GetTubeSurfaceInfo()
'
' Description:  Gets the tube surface from side input
'
' ***********************************************************************************
Private Function GetTubeSurfaceInfo(ByVal oMemberPart As Object, strTemplateSide As String) As Object
    Const METHOD = "GetTubeSurfaceInfo"
    On Error GoTo ErrorHandler
    
    Dim oTemplateHelper         As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    Dim oSurfaceElements        As IJElements
    Dim oMemberPartPrismatic    As ISPSMemberPartPrismatic
    Set oMemberPartPrismatic = oMemberPart
    
    Dim strSectionType          As String
    strSectionType = oMemberPartPrismatic.CrossSection.definition.Type
    
    If Not strSectionType = "HSSR" Then
    
        Dim lSectionID  As Long
        
        If strTemplateSide = "Outer" Then
            lSectionID = 3073  ' Outer Tube
        Else
            lSectionID = 3074  ' Inner Tube
        End If
        
        On Error Resume Next
        Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lSectionID)
        On Error GoTo ErrorHandler
    
        If Not oSurfaceElements Is Nothing Then
            ' This is a round crosssection so we expect only one face
            Set GetTubeSurfaceInfo = oSurfaceElements.Item(1)
        End If
    
    Else
    
        ' The below code is for HSSR(square/Rectangular cross sections)
        ' This is not a round crosssection so get the individual faces as we expect more than one face
        
        'JXSEC_BOTTOM = 513 (&H201)
        'JXSEC_TOP = 514 (&H202)
        'JXSEC_WEB_LEFT = 257 (&H101)
        'JXSEC_WEB_RIGHT = 258 (&H102)
        
        'JXSEC_INNER_TOP = 3585 (&HE01)
        'JXSEC_INNER_BOTTOM = 3586 (&HE02)
        'JXSEC_INNER_WEB_LEFT = 3329 (&HD01)
        'JXSEC_INNER_WEB_RIGHT = 3330 (&HD02)

        Dim lWebLeft    As Long
        Dim lWebRight    As Long
        Dim lTop    As Long
        Dim lBottom    As Long
        
        If strTemplateSide = "Outer" Then
            lWebLeft = 257
            lWebRight = 258
            lTop = 514
            lBottom = 513
        Else
            lWebLeft = 3329
            lWebRight = 3330
            lTop = 3585
            lBottom = 3586
        End If
        
        Dim oSurfCollection     As Collection
        Set oSurfCollection = New Collection
        
        Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lWebLeft)
        oSurfCollection.Add oSurfaceElements.Item(1)
        Set oSurfaceElements = Nothing
        
        Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lTop)
        oSurfCollection.Add oSurfaceElements.Item(1)
        Set oSurfaceElements = Nothing
        
        Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lWebRight)
        oSurfCollection.Add oSurfaceElements.Item(1)
        Set oSurfaceElements = Nothing
        
        Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lBottom)
        oSurfCollection.Add oSurfaceElements.Item(1)
        Set oSurfaceElements = Nothing
        
        Dim oMfgGeomHelper As MfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper
        
        Dim oMergedSurface    As IJSurfaceBody
        Set oMergedSurface = oMfgGeomHelper.MergeSurfaceBodies(oSurfCollection, 0.001, 0.001)
        
        Set GetTubeSurfaceInfo = oMergedSurface
        
    End If
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function IJDMfgTemplateService2_CreateTemplateUnfoldOutContour()
'
' Description:  creates unfolded output for the template 3D geometries
'
' ***********************************************************************************
Private Function IJDMfgTemplateService2_CreateTemplateUnfoldOutContour(ByVal oPart As Object, ByVal oUnkTemplateSet As Object, ByVal oProcessSettings As Object) As Object
 Const METHOD = "IJDMfgTemplateService2_CreateTemplateUnfoldOutContour"
    On Error GoTo ErrorHandler
    
    ' Create the needed services
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oRuleHelper     As MfgRuleHelpers.Helper
    Set oRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oTemplateHelper     As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    ' Create template set data and cache the settings data
    Dim oTemplSetData As TemplSetData
    Set oTemplSetData = New TemplSetData
    
    oTemplSetData.InitSettings oPart, oProcessSettings, oUnkTemplateSet

    ' Get the template tube surface base on the input side
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = GetTubeSurfaceInfo(oPart, oTemplSetData.TemplateSide)
    
    Dim oTemplateSet As IJDMfgTemplateSet
    Set oTemplateSet = oUnkTemplateSet
    
    ' Get the base plane
    Dim oBasePlane As IJPlane
    Set oBasePlane = oTemplateSet.GetBasePlane
    
    ' Get the template
    Dim oTemplates  As IJElements
    Set oTemplates = oTemplateSet.GetTemplates
    
    Dim oMfgTemplate    As IJMfgTemplate
    Set oMfgTemplate = oTemplates.Item(1)
    
    ' Get the template 3D bottom curve
    Dim oBottomCurve As IJCurve
    Set oBottomCurve = oMfgTemplate.GetTemplateLocationMarkLine
    
    ' Get the template 3D top curve
    Dim oTopCurve As IJCurve
    Set oTopCurve = oMfgTemplate.TopLine
    
    ' Get the template 3D BCL curve
    Dim oBaseCtrlCurve As IJCurve
    Set oBaseCtrlCurve = oTemplateSet.GetControlLine
    
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    oBaseCtrlCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
    Dim oBCLPos As IJDPosition
    Set oBCLPos = New DPosition
    
    ' BCL position on end cut is start point of BCL
    oBCLPos.Set dStartX, dStartY, dStartZ
    
    ' Get the extension information
    Dim eMethod_Left As EnumExtnMethod
    Dim eMethod_Right As EnumExtnMethod
    
    Dim eMeasure_Left As EnumExtnMeasure
    Dim eMeasure_Right As EnumExtnMeasure
    
    Dim dExtnValue1_Left As Double
    Dim dExtnValue1_Right As Double
    
    dExtnValue1_Left = oTemplSetData.TemplateExtnValueLeft
    dExtnValue1_Right = oTemplSetData.TemplateExtnValueRight
    
    eMeasure_Left = oTemplSetData.TemplateMeasureLeft
    eMeasure_Right = oTemplSetData.TemplateMeasureRight
    
    eMethod_Left = oTemplSetData.TemplateMethodLeft
    eMethod_Right = oTemplSetData.TemplateMethodRight
    
    ' **** Get the reference curve needed for unfolding - Refer below picture****
        
'    ___'_________'___________
    '    '        '           |
   '      '       '           |
   '       '      '           |
    '       '     '           |
   ' '       '    '           |
   ' '        '   '           |
    '__________'__'___________|
                ' '
'  Endcut     BP  Ref curve
               
    Dim oRefCurve   As IJCurve
    Dim oRefCurveWB As Object

    Set oRefCurveWB = GetReferenceCurve(oPart, oSurfaceBody, oBasePlane, oTopCurve)
    Set oRefCurve = oRuleHelper.WireBodyToComplexString(oRefCurveWB)
    
    Dim dRefCurveLen As Double, dLeftExtnVal As Double, dRightExtnVal As Double, dAlongEdge_Left As Double, dAlongEdge_Right As Double
    dRefCurveLen = oRefCurve.Length
    
    If Not dExtnValue1_Left = 0 And Not dExtnValue1_Right = 0 Then
        If eMeasure_Left = AlongAxis Then
            dLeftExtnVal = dRefCurveLen * (dExtnValue1_Left / 360)
        Else
            dLeftExtnVal = dExtnValue1_Left
        End If
        
        If eMeasure_Right = AlongAxis Then
            dRightExtnVal = dRefCurveLen * (dExtnValue1_Right / 360)
        Else
            dRightExtnVal = dExtnValue1_Right
        End If
        
        If eMethod_Left = AlongEdge Then      ' Along Edge
            dAlongEdge_Left = dLeftExtnVal
            dLeftExtnVal = 0#
        End If
        
        If eMethod_Right = AlongEdge Then     ' Along Edge
            dAlongEdge_Right = dRightExtnVal
            dRightExtnVal = 0#
        End If
        
    End If
    
    ' Get the unfolded output of the bottom curve
    Dim oUnfoldElements As IJElements
    Set oUnfoldElements = oTemplateHelper.GetUnfoldedContourOfInputCurve(oBottomCurve, oRefCurve, oBCLPos, oSurfaceBody, dAlongEdge_Left, dAlongEdge_Right)
    
    Dim oTempCurve  As IJCurve
    Set oTempCurve = oUnfoldElements.Item(1)
    
    Dim oUnfoldBottomCurve  As IJCurve
    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    
    ' Apply extension(linear/perpendicular) to unfolded bottom curve
    If Not dLeftExtnVal = 0 Or Not dRightExtnVal = 0 Then
        
        ' [in]int extOpt   - 0: extend to point, 1: extend by distance
        ' [in]int extType  - 0: tangent line; 1: polynomial extension
        ' [in]int extEnd   - 0: extend the beginning, 1: extend the end, 2: extend both
        
         If Not dLeftExtnVal = 0 And Not dRightExtnVal = 0 Then
            
            If eMethod_Left = Linear And eMethod_Right = Linear Then        'Linear
                Set oUnfoldBottomCurve = oMfgGeomHelper.ExtraPolateCurve(oTempCurve, 1, 0, 2, Nothing, Nothing, dLeftExtnVal)
            ElseIf eMethod_Left = Perpendicular And eMethod_Right = Perpendicular Then        'Perpendicular
                oTempCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
                ' Apply the extn value to Y coordinate
                dStartY = dStartY - dLeftExtnVal
                dEndY = dEndY + dRightExtnVal
                
                Set oStartPos = New DPosition
                oStartPos.Set dStartX, dStartY, dStartZ
                
                Set oEndPos = New DPosition
                oEndPos.Set dEndX, dEndY, dEndZ
                
                Set oUnfoldBottomCurve = oMfgGeomHelper.ExtraPolateCurve(oTempCurve, 0, 0, 2, oStartPos, oEndPos, dLeftExtnVal)
                
                Set oStartPos = Nothing
                Set oEndPos = Nothing
            End If
        End If
    Else
        Set oUnfoldBottomCurve = oTempCurve
    End If
    
    ' Unfold the  top line
    Set oUnfoldElements = Nothing
    Set oUnfoldElements = oTemplateHelper.GetUnfoldedContourOfInputCurve(oTopCurve, oRefCurve, oBCLPos, oSurfaceBody, 0, 0)
    
    Dim oUnfoldTopCurve  As IJCurve
    Set oUnfoldTopCurve = oMfgGeomHelper.ExtraPolateCurve(oUnfoldElements.Item(1), 1, 0, 2, Nothing, Nothing, 5)
    
    ' Create sidelines
    oUnfoldBottomCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
    Set oStartPos = New DPosition
    oStartPos.Set dStartX, dStartY, dStartZ
    
    Set oEndPos = New DPosition
    oEndPos.Set dEndX, dEndY, dEndZ

    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, 1, 0
    
    Dim oSideLine1  As IJCurve
    Set oSideLine1 = CreateLineAtPosition(oStartPos, oMarkVec, 5)
    
    Dim oSideLine2  As IJCurve
    Set oSideLine2 = CreateLineAtPosition(oEndPos, oMarkVec, 5)
    
    Dim dMinDist As Double, dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
    oSideLine1.DistanceBetween oUnfoldTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oSidePos1   As IJDPosition
    Set oSidePos1 = New DPosition
    
    oSidePos1.Set dSrcX, dSrcY, dSrcZ
    oSideLine2.DistanceBetween oUnfoldTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oSidePos2   As IJDPosition
    Set oSidePos2 = New DPosition
    
    oSidePos2.Set dSrcX, dSrcY, dSrcZ
    
    Set oSideLine1 = Nothing
    Set oSideLine2 = Nothing
    
    Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Set oSideLine1 = CreateLineFromPoints(oStartPos, oSidePos1)
    oMfgMGHelper.TrimCurveByPoints oUnfoldTopCurve, oSidePos1, oSidePos2
    Set oSideLine2 = CreateLineFromPoints(oSidePos2, oEndPos)
    
        'Dim sUBCFileName As String
        'Dim sUTCFileName As String
        'Dim sUSC1FileName As String
        'Dim sUSC2FileName As String
        'sUBCFileName = Environ("TEMP")
        'sUTCFileName = Environ("TEMP")
        'sUSC1FileName = Environ("TEMP")
        'sUSC2FileName = Environ("TEMP")
        'If sUBCFileName = "" Or sUBCFileName = vbNullString Then
        '       sUBCFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'If sUTCFileName = "" Or sUTCFileName = vbNullString Then
        '       sUTCFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'If sUSC1FileName = "" Or sUSC1FileName = vbNullString Then
        '       sUSC1FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'If sUSC2FileName = "" Or sUSC2FileName = vbNullString Then
        '       sUSC2FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'sUBCFileName = sUBCFileName & "\UnfoldBottomCurve.txt"
        'sUTCFileName = sUTCFileName & "\UnfoldTopCurve.txt"
        'sUSC1FileName = sUSC1FileName & "\UnfoldSideCurve1.txt"
        'sUSC2FileName = sUSC2FileName & "\UnfoldSideCurve2.txt"
        
    'oMfgGeomHelper.PrintComplexString oUnfoldBottomCurve, sUBCFileName
    'oMfgGeomHelper.PrintComplexString oUnfoldTopCurve, sUTCFileName
    'oMfgGeomHelper.PrintComplexString oSideLine1, sUSC1FileName
    'oMfgGeomHelper.PrintComplexString oSideLine2, sUSC2FileName
    
    ' Create Geom 2Ds for unfolded geometries and add to Geom Coll
    Dim oGeomCol As IJElements
    Set oGeomCol = New JObjectCollection
    
    Dim oMfgGeom2d As IJMfgGeom2d
    Set oMfgGeom2d = CreateGeom2D(oUnfoldBottomCurve, STRMFG_TemplateLocationMarkLine)
    oGeomCol.Add oMfgGeom2d
    
    Set oMfgGeom2d = Nothing
    Set oMfgGeom2d = CreateGeom2D(oUnfoldTopCurve, STRMFG_TopLine)
    oGeomCol.Add oMfgGeom2d
    
    Set oMfgGeom2d = Nothing
    Set oMfgGeom2d = CreateGeom2D(oSideLine1, STRMFG_BoundaryLine3D)
    oGeomCol.Add oMfgGeom2d
    
    Set oMfgGeom2d = Nothing
    Set oMfgGeom2d = CreateGeom2D(oSideLine2, STRMFG_BoundaryLine3D)
    oGeomCol.Add oMfgGeom2d
    
    Set IJDMfgTemplateService2_CreateTemplateUnfoldOutContour = oGeomCol
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
End Function

' ***********************************************************************************
' Public Function IJDMfgTemplateService2_GetTemplateUnfoldRefCurve()
'
' Description:  Returns reference curve used for unfolding of template 3D geometries
'
' ***********************************************************************************
Private Function IJDMfgTemplateService2_GetTemplateUnfoldRefCurve(ByVal oPart As Object, ByVal oUnkTemplateSet As Object, ByVal oProcessSettings As Object) As Object
Const METHOD = "IJDMfgTemplateService2_GetTemplateUnfoldRefCurve"
On Error GoTo ErrorHandler
    
    ' Create the needed services
    Dim oRuleHelper     As MfgRuleHelpers.Helper
    Set oRuleHelper = New MfgRuleHelpers.Helper
    
    ' Create template set data and cache the settings data
    Dim oTemplSetData As TemplSetData
    Set oTemplSetData = New TemplSetData
    
    oTemplSetData.InitSettings oPart, oProcessSettings, oUnkTemplateSet
    
    ' Get the template tube surface base on the input side
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = GetTubeSurfaceInfo(oPart, oTemplSetData.TemplateSide)
    
    Dim oTemplateSet As IJDMfgTemplateSet
    Set oTemplateSet = oUnkTemplateSet
    
    ' Get the base plane
    Dim oBasePlane As IJPlane
    Set oBasePlane = oTemplateSet.GetBasePlane
    
    ' Get the template
    Dim oTemplates  As IJElements
    Set oTemplates = oTemplateSet.GetTemplates
    
    Dim oMfgTemplate    As IJMfgTemplate
    Set oMfgTemplate = oTemplates.Item(1)
    
    ' Get the template 3D top curve
    Dim oTopCurve As IJCurve
    Set oTopCurve = oMfgTemplate.TopLine
    
    ' Get the reference curve
    Dim oRefCurve   As IJCurve
    Dim oRefCurveWB As Object
    
    Set oRefCurveWB = GetReferenceCurve(oPart, oSurfaceBody, oBasePlane, oTopCurve)
    Set oRefCurve = oRuleHelper.WireBodyToComplexString(oRefCurveWB)
    Set IJDMfgTemplateService2_GetTemplateUnfoldRefCurve = oRefCurve
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
