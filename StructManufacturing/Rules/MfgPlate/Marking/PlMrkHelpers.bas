Attribute VB_Name = "PlMrkHelpers"
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation.  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    PlMrkHelpers for the MfgPlateMarking rules.
'
'  History:
'       TBH         april 8. 2002   created
'       MJV     2004.04.23      Include correct error handling
'******************************************************************

Option Explicit

Const MODULE = "PlMrkHelpers.bas"

Public m_oMfgRuleHelper As MfgRuleHelpers.Helper

Public m_oSystemMarkFactory As GSCADMfgSystemMark.MfgSystemMarkFactory
Public m_oGeom3dFactory As GSCADMfgGeometry.MfgGeom3dFactory
Public m_oGeom2dFactory As GSCADMfgGeometry.MfgGeom2dFactory
Public m_oGeom3dColFactory As GSCADMfgGeometry.MfgGeomCol3dFactory
Public m_oGeom2dColFactory As GSCADMfgGeometry.MfgGeomCol2dFactory

' Enumerator for the feature types.
Public Enum enumMarkFeature
    eEdgeFeature = 0        ' Only Edge Features
    eCornerFeature = 1      ' Only Edge Features
    eAllFeatures = 2        ' All features
    eIgnoreFeature = 3      ' Ignore features
End Enum

' Enumerator for the Seam Type.
Public Enum enumSeamType
    ePlanningSeam = 0       ' Only Planning Seams
    eDesignSeam = 1         ' Only Design Seams
    eAllSeams = 2           ' All Seams
End Enum

Public m_bDebug As Boolean
Private Const PI As Double = 3.14159265358979 'required for TR 53039
Private Const ER_OFFSET_DIST As Double = 0

'~~~~~~~~~~~~Roll Lines Marking Declarations~~~~~~~~~~~~~~
Private Enum Option_Number_of_Roll_Lines
    ShowAllGeneratedBySystem = -1
    DoNotShowRollLines = 0
    GenerateRollAnnotationOnly = 1
End Enum

Const ROLL_LINES_DISPLAY = ShowAllGeneratedBySystem

Const LARGEvsSMALL_CRITERION_1_ROLLRADIUS As Double = 0.075 ' 75 mm roll radius
Const LARGEvsSMALL_CRITERION_2_DISTBETROLLB As Double = 0.1   ' 100 mm distance between roll boundaries
Const FitTolerance As Double = 0.001
Dim m_dCount As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' ***********************************************************************************
' Public Sub Initialize
'
' Description: Helper to Initialize the most used objects
'
' ***********************************************************************************
Public Sub Initialize()
Const METHOD = "Initialize"
    On Error GoTo ErrorHandler

    m_bDebug = False

    Set m_oGeom3dFactory = New GSCADMfgGeometry.MfgGeom3dFactory
    Set m_oGeom2dFactory = New GSCADMfgGeometry.MfgGeom2dFactory
    Set m_oGeom3dColFactory = New GSCADMfgGeometry.MfgGeomCol3dFactory
    Set m_oGeom2dColFactory = New GSCADMfgGeometry.MfgGeomCol2dFactory

    Set m_oMfgRuleHelper = New MfgRuleHelpers.Helper
    Set m_oSystemMarkFactory = New GSCADMfgSystemMark.MfgSystemMarkFactory
    
    m_dCount = 0

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub UnInitialize()
    Set m_oMfgRuleHelper = Nothing
    Set m_oSystemMarkFactory = Nothing
    Set m_oGeom3dFactory = Nothing
    Set m_oGeom2dFactory = Nothing
    Set m_oGeom3dColFactory = Nothing
    Set m_oGeom2dColFactory = Nothing
End Sub

' ***********************************************************************************
' Public Function CurveLength
'
' Description:  Helper function to get length of a curve. Input is a Wirebody and
'               output is a double which is the length of the curve.
'
' ***********************************************************************************
Public Function CurveLength(ByVal oWB As IJWireBody) As Double
Const METHOD = "CurveLength"
    On Error GoTo ErrorHandler
    Dim oCurve As IJCurve
    Set oCurve = m_oMfgRuleHelper.WireBodyToComplexString(oWB)

    CurveLength = oCurve.Length
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    CurveLength = 0#
End Function



' ***********************************************************************************
' Public Function CreateShipDirectionMarkLine
'
' Description:  Method will create a small line as a ComplexString,
'               and the direction information will be in the MarkingInfo.Name ans Direction.
'               The line is only there to indicate a position on the plate.
'
' ***********************************************************************************
Public Function CreateShipDirectionMarkLine(oPos As IJDPosition, oSurfacePort As IJPort, oDirection As IJDVector, ByRef oNeutralSurface As IJSurfaceBody, ByRef oMark As IJComplexString) As Boolean

Const METHOD = "CreateShipDirectionMarkLine"
On Error GoTo ErrorHandler
    Dim dLength As Double
    dLength = 1000#
    Dim endX As Double
    Dim endY As Double
    Dim endZ As Double

    m_oMfgRuleHelper.ScaleVector oDirection, -1

    Dim oLine As IJLine
    Set oLine = New Line3d
    endX = oPos.x + (oDirection.x * dLength)
    endY = oPos.y + (oDirection.y * dLength)
    endZ = oPos.z + (oDirection.z * dLength)
    oLine.DefineBy2Points oPos.x, oPos.y, oPos.z, endX, endY, endZ

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True

    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oProjElemCol As IJElements
    oMfgMGHelper.ProjectCSToSurface oCS, oNeutralSurface, Nothing, oProjElemCol
            
    Set oCS = Nothing
    If Not oProjElemCol Is Nothing Then
        
        If oProjElemCol.Count = 1 Then
            Set oCS = oProjElemCol.Item(1)
        Else
            Dim lIdx As Long
            For lIdx = 1 To oProjElemCol.Count
                Set oCS = oProjElemCol.Item(lIdx)
                
                Dim oCurve As IJCurve
                Set oCurve = oCS
                
                Dim sX As Double
                Dim sY As Double
                Dim sZ As Double
                Dim dummy As Double
                
                oCurve.EndPoints sX, sY, sZ, dummy, dummy, dummy
                If Abs(oPos.x - sX) > FitTolerance Or Abs(oPos.y - sY) > FitTolerance Or Abs(oPos.z - sZ) > FitTolerance Then
                    Set oCS = Nothing
                    Set oCurve = Nothing
                Else
                    Exit For
                End If
            Next
        End If

    End If
    
    If Not oCS Is Nothing Then
        Set oMark = oCS
        CreateShipDirectionMarkLine = True
    Else
        Set oMark = Nothing
        CreateShipDirectionMarkLine = False
    End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    CreateShipDirectionMarkLine = False
End Function


' ***********************************************************************************
' Public Function CreateWebFrameChecjMarkFromSystem
'
' Description: Method will create a straight line spanding over all involved plateparts (system)
'              and project it on to the surface of the relevant part
'
'   Input    : Current platepart,upside
'
' ***********************************************************************************
Public Function CreateWebFrameCheckMarkFromParent(oPlatePart As IJPlatePart, UpSide As Long, oConObjsCol As Collection) As IJComplexString
Const METHOD = "CreateWebFrameCheckMarkFromParent"
On Error GoTo ErrorHandler
'   Steps :
'   1. Prerequisits (Type = Transverse, Connected to Hull, Has EdgeReinforcament)
'   2. Find approximate endpoint of line (endpoints of Hull boundary)
'   3. Create straight line from these endpoints
'   4. Project line on to surface
'   5. Displace line by dDistanceFromHull

    Dim dDistanceFromHull As Double
    dDistanceFromHull = WebFrameMarkDistanceFromHull

'   1. Prerequisits (Type = Transverse, Connected to Hull, Has EdgeReinforcament)
'   *******************************************************************************
    Dim oPlate As IJPlate
    Set oPlate = oPlatePart
    Dim ePlateType As StructPlateType
    ePlateType = oPlate.plateType

    If Not ePlateType = LBulkheadPlate And Not ePlateType = TBulkheadPlate Then Exit Function

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlatePart

    Dim oPartPort As IJStructPort, oSystemPort As IJStructPort

    Dim oSystem As IJSystem
    Dim i As Long
    Dim dummy1 As IJStructPort

    ' If more parts are involved in the conditions all parts must be checked
    Set oSystem = oPlateWrapper.GetRootSystem
    If Not oSystem Is Nothing Then
        Dim oSystemParts As Collection
        Set oSystemParts = m_oMfgRuleHelper.GetSystemDerivedParts(oSystem, True)

        Dim lCount As Long
        lCount = oSystemParts.Count

        Dim bHullPresent As Boolean
        Dim bEdgeREPresent As Boolean
        bHullPresent = False
        bEdgeREPresent = False

        ' Find relevant subplateparts
        ' And check if hull and edgereinforcement are present :
        ' - Hull connection on any transverse part in the system
        ' - Edge reinforcement on the selected part
        For i = 1 To lCount
            If TypeOf oSystemParts.Item(i) Is IJPlatePart Then
                Set oPlate = oSystemParts.Item(i)
                ePlateType = oPlate.plateType

                If ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                    Set oPlateWrapper.object = oPlate
                    ' bHullPresent will be true if any transverse bulkhead is connected to hull
                    If oPlateWrapper.IsConnectedToHull(dummy1, oSystemPort, oConObjsCol) Then bHullPresent = True
                End If
                Set oPlate = Nothing
            End If
        Next i
    Else
        Set oPlateWrapper.object = oPlatePart
        bHullPresent = oPlateWrapper.IsConnectedToHull(oPartPort, oSystemPort, oConObjsCol)
        bEdgeREPresent = oPlateWrapper.HasEdgeReenforcements
    End If

    Set dummy1 = Nothing

    Dim dummy2 As IJStructPort
    Set oPlateWrapper.object = oPlatePart
    ' bEdgeREPresent will be true if the selected part is connected to hull
    If oPlateWrapper.HasEdgeReenforcements Then bEdgeREPresent = True
    oPlateWrapper.IsConnectedToHull oPartPort, dummy2, oConObjsCol
    Set dummy2 = Nothing

    If oPartPort Is Nothing Or oSystemPort Is Nothing Then GoTo CleanUp

    If Not bHullPresent Then GoTo CleanUp
    If Not bEdgeREPresent Then GoTo CleanUp

'   2. Find approximate endpoint of line (endpoints of Hull boundary)
'   *******************************************************************************

    Dim oModelBody As IJDModelBody
    Dim oPort As IJPort
    Dim oHullWire As IJWireBody

    ' If systemport is available use that one
    If Not oSystemPort Is Nothing Then
        ' Check if it has more lumbs !! (only systems)
        Set oPort = oSystemPort
        Set oModelBody = oPort.Geometry

        Dim lNoOfLumbs As Long, lNoOfShells As Long, lNoOfWires As Long, lNoOfFaces As Long, lNoOfLoops As Long, lNoOfCooedges As Long, lNoOfEdges As Long, lNoOfVert As Long, bIsSane As Boolean
        oModelBody.CheckTopology "", lNoOfLumbs, lNoOfShells, lNoOfWires, lNoOfFaces, lNoOfLoops, lNoOfCooedges, lNoOfEdges, lNoOfVert, bIsSane

        If lNoOfWires > 1 Then
            ' Find the right one
            Dim oPos As IJDPosition, oPos2 As IJDPosition
            Set oPos = New DPosition
            oPos.Set 0, 0, 0

            Dim oVector As IJDVector
            Dim oPPort As IJPort
            Set oPPort = oPartPort
            Set oPos2 = m_oMfgRuleHelper.ProjectPointOnSurface(oPos, oPPort.Geometry, oVector)

            Set oHullWire = m_oMfgRuleHelper.GetClosestLumb(oModelBody, oPos2)
        Else
            Set oHullWire = oModelBody
        End If

        If lNoOfWires < 1 Then GoTo ErrorHandler

    Else
        Set oPort = m_oMfgRuleHelper.MapStructPort(False, oPartPort)
        Set oHullWire = oPort.Geometry
    End If
    Set oModelBody = Nothing

'   3. Create straight line from these endpoints

    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    oHullWire.GetEndPoints oStartPos, oEndPos

    Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
    oStartPos.Get x1, y1, z1
    oEndPos.Get x2, y2, z2

    Dim oLine As IJLine
    Set oLine = New Line3d
    oLine.DefineBy2Points x1, y1, z1, x2, y2, z2

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d

    oCS.AddCurve oLine, True
    Set oLine = Nothing

    Dim oInitWire As IJWireBody
    Set oInitWire = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)
    Set oCS = Nothing

'   4. Project line on to surface

    ' Find direction to project
    Dim oMidPos As IJDPosition
    Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oHullWire)

    Dim oSurfacePort As IJPort
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfacePort = oPlateWrapper.GetSurfacePort(UpSide)
    Set oSurfaceBody = oSurfacePort.Geometry

    Dim oProjectionDir As IJDVector
    m_oMfgRuleHelper.ProjectPointOnSurface oMidPos, oSurfaceBody, oProjectionDir

    Dim oUnboundedSurface As IJSurfaceBody
    Set oUnboundedSurface = oPlateWrapper.GetUnboundedSurface(UpSide)

    On Error Resume Next

    Dim oUnkWire As IUnknown
    If Not oProjectionDir Is Nothing Then
        Set oCS = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oUnboundedSurface, oInitWire, Nothing)
        Set oUnkWire = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)
        Set oCS = Nothing
    Else
        Set oUnkWire = oInitWire
    End If

'   5. Displace line

    ' Find displace direction
    Dim oDisplaceDir As IJDVector
    Set oPort = oPartPort
    Set oSurfaceBody = oPort.Geometry

'    Dim oLateralPos As IJDPosition
'    Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oInitWire)
'    Set oLateralPos = m_oMfgRuleHelper.ProjectPointOnSurface(oMidPos, oSurfaceBody, oDisplaceDir)
'    oSurfaceBody.GetNormalFromPosition oLateralPos, oDisplaceDir
'    m_oMfgRuleHelper.ScaleVector oDisplaceDir, -1

    Dim oWire As IJWireBody
    Set oWire = m_oMfgRuleHelper.OffsetCurve(oUnboundedSurface, oUnkWire, Nothing, dDistanceFromHull, True)

    ' Log the message when code failed to offset the wire
    If oWire Is Nothing Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1094, , "RULES"
        GoTo CleanUp
    End If

    'Since offset was done to an unbounded surface project wire back on bounded surface
    'Set oCS = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWire, oProjectionDir)
    Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWire)

    ' Log the message when code failed to project the wire
    If oCS Is Nothing Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1094, , "RULES"
        GoTo CleanUp
    End If

    ' Extend the complex string as it can be short after offsetting
    Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
    oMfgGeomUtilWrapper.ExtendWire oCS, 0.5

    Set CreateWebFrameCheckMarkFromParent = oCS
    Set oWire = Nothing
    Set oCS = Nothing

CleanUp:
    Set oPlate = Nothing
    Set oPlateWrapper = Nothing
    Set oSystemParts = Nothing
    Set oModelBody = Nothing
    Set oPort = Nothing
    Set oHullWire = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oLine = Nothing
    Set oCS = Nothing
    Set oMidPos = Nothing
    Set oInitWire = Nothing
    Set oProjectionDir = Nothing
    Set oSurfaceBody = Nothing
    Set oSurfacePort = Nothing
    Set oUnkWire = Nothing

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


' ***********************************************************************************
' Public Function CreateCrossLineForConnPlate()
'
' Description:  function creates  fitting mark line(s) on the this plate's surface.
'               Input arguments: marking point position, this plate surface, conn plate surtface.
'               Output argument: WB of mark line.
'
'
' Unresolved issue: GetConnectedTypeForContour() not exposed yet
'                   Automation Error: oThisPartSufaceBody.GetNormalFromPosition oMarkPointPos, oNormalVector
' ***********************************************************************************
Public Function CreateLocationFittingMark(oMarkPointPos As IJDPosition, _
    oThisPartSuface As IUnknown, oConnPartSurface As IUnknown) As IJComplexString

    Const METHOD = "CreateLocationFittingMark"

    On Error GoTo ErrorHandler

    'helper objects
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d

    Dim xStart As Double
    Dim yStart As Double
    Dim zStart As Double
    Dim xEnd As Double
    Dim yEnd As Double
    Dim zEnd As Double

    Dim oFaceSufaceBody As IJSurfaceBody
    Set oFaceSufaceBody = oConnPartSurface

    Dim oFaceNormalLineCS As IJComplexString
    Set oFaceNormalLineCS = New ComplexString3d

    'project oMarkPointPos on both surfaces
    Dim oVectorToThis As IJDVector
    Dim oVectorToConn As IJDVector
    Dim oMarkPointPosThisSurf As IJDPosition
    Dim oMarkPointPosConnSurf As IJDPosition

    'Project mark point onto the surfaces
    Set oMarkPointPosThisSurf = m_oMfgRuleHelper.ProjectPointOnSurface(oMarkPointPos, oThisPartSuface, oVectorToThis)
    Set oMarkPointPosConnSurf = m_oMfgRuleHelper.ProjectPointOnSurface(oMarkPointPos, oConnPartSurface, oVectorToConn)

    'create normal to conn plate surface
    Dim oFaceNormalVector As IJDVector
    oFaceSufaceBody.GetNormalFromPosition oMarkPointPosConnSurf, oFaceNormalVector
    oMarkPointPos.Get xStart, yStart, zStart
    m_oMfgRuleHelper.ScaleVector oFaceNormalVector, FittingMarkLength / 2
    oFaceNormalVector.Get xEnd, yEnd, zEnd

    Dim oFaceNormalLine As IngrGeom3D.Line3d
    Set oFaceNormalLine = New IngrGeom3D.Line3d
    oFaceNormalLine.DefineBy2Points xStart - xEnd, yStart - yEnd, zStart - zEnd, xStart + xEnd, yStart + yEnd, zStart + zEnd
    oFaceNormalLineCS.AddCurve oFaceNormalLine, False

    'ComplexStringAlongVectorOnToSurface fails if the location fitting mark is
    'outside the plate outer contour so "On Error Resume Next" is added below
    'to skip those failed cases.
    On Error Resume Next
    'm_oMfgRuleHelper.ScaleVector oVectorToThis, -1
    Set oCS = m_oMfgRuleHelper.ComplexStringAlongVectorOnToSurface(oThisPartSuface, oFaceNormalLineCS, oVectorToThis)
    If oCS.CurveCount = 0 Then
        Set CreateLocationFittingMark = Nothing
    Else
        Dim oCurve As IJCurve
        Dim lCrvLength As Double

        Set oCurve = oCS
        lCrvLength = oCurve.Length
        ' If the curve length is less than FittingMarkLength, extend it.
        If lCrvLength < FittingMarkLength Then
            Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
            oMfgGeomUtilWrapper.ExtendWire oCS, (FittingMarkLength - lCrvLength)
            Set oMfgGeomUtilWrapper = Nothing
        End If
        Set oCurve = Nothing

        'return complex string
        Set CreateLocationFittingMark = oCS
    End If

CleanUp:
    Set oCS = Nothing
    Set oFaceSufaceBody = Nothing
    Set oFaceNormalLineCS = Nothing
    Set oVectorToThis = Nothing
    Set oVectorToConn = Nothing
    Set oMarkPointPosThisSurf = Nothing
    Set oMarkPointPosConnSurf = Nothing
    Set oFaceNormalVector = Nothing
    Set oFaceNormalLine = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'*****************************************************************************************
'
'
'
'
'
'
'*****************************************************************************************

Public Function CreateEdgeCheckMark(ByVal oPlate As IJPlatePart, ByVal oPos As IJDPosition, ByVal eUpSide As Long) As IJComplexString
    Const METHOD = "CreateEdgeCheckMark"
    On Error GoTo ErrorHandler

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlate

    Dim oUpsidePort As IJPort
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(eUpSide)

    'Displace the position into the plate
    'Note that lines will be very long and x-value diplaced
    'This is in order to trim and project them correct later

    Dim y As Double, u As Double, v As Double
    u = EdgeCheckMarkUDistance
    v = EdgeCheckMarkVDistance

    If oPos.y > 0 Then
        y = oPos.y - u
    Else
        y = oPos.y + u
    End If

    Dim oMarkPos As IJDPosition
    Set oMarkPos = New DPosition
    oMarkPos.Set oPos.x + 0.2, y, oPos.z + v

    ' Find middlepoint of edge to determine how the mark should be oriented
    Dim oConnectable As IJConnectable
    Set oConnectable = oUpsidePort.Connectable
    Dim oElements As IJElements
    oConnectable.enumPorts oElements, PortEdge

    If oElements Is Nothing Then
        Exit Function
    ElseIf oElements.Count = 0 Then
        Exit Function
    End If

    Dim oTop As Double, oBottom As Double
    Dim oEdgePort As IJPort
    Dim oEdgeWire As IJWireBody
    Dim oSPos As IJDPosition, oEPos As IJDPosition
    oTop = 0
    oBottom = 100

    Dim i As Integer
    For i = 1 To oElements.Count
        Set oEdgePort = oElements.Item(i)
        Set oEdgeWire = oEdgePort.Geometry
        oEdgeWire.GetEndPoints oSPos, oEPos
        If oSPos.z > oTop Then oTop = oSPos.z
        If oSPos.z < oBottom Then oBottom = oSPos.z
        If oEPos.z > oTop Then oTop = oEPos.z
        If oEPos.z < oBottom Then oBottom = oEPos.z
    Next i

    Dim oMiddlePos As Double
    oMiddlePos = ((oTop - oBottom) / 2) + oBottom

    Dim oVerticalEnd As IJDPosition
    Set oVerticalEnd = New DPosition
    Dim oHorizontalEnd As IJDPosition
    Set oHorizontalEnd = New DPosition

    If oPos.z >= oMiddlePos Then
        ' Create Upwards/Outbound Mark
        oVerticalEnd.Set oMarkPos.x, oMarkPos.y, oMarkPos.z + 10

        If oPos.y > 0 Then
            y = oPos.y + 10
        Else
            y = oPos.y - 10
        End If

        oHorizontalEnd.Set oMarkPos.x, y, oMarkPos.z
    Else
        ' Create Downwards/Inbound Mark
        oVerticalEnd.Set oMarkPos.x, oMarkPos.y, oMarkPos.z - 10

        If oPos.y > 0 Then
            y = oPos.y + 10
        Else
            y = oPos.y - 10
        End If
        oHorizontalEnd.Set oMarkPos.x, y, oMarkPos.z
    End If

    Dim oVerticalLine As IJLine
    Set oVerticalLine = New Line3d
    Dim oHorisontalLine As IJLine
    Set oHorisontalLine = New Line3d

    oVerticalLine.DefineBy2Points oMarkPos.x, oMarkPos.y, oMarkPos.z, oVerticalEnd.x, oVerticalEnd.y, oVerticalEnd.z
    oHorisontalLine.DefineBy2Points oMarkPos.x, oMarkPos.y, oMarkPos.z, oHorizontalEnd.x, oHorizontalEnd.y, oHorizontalEnd.z

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oVerticalLine, False
    oCS.AddCurve oHorisontalLine, False

    'Project mark back on to plate
    '(will trim the length)

    Dim oDir As IJDVector
    Set oDir = New DVector
    oDir.Set -1, 0, 0

    Dim oProjectedMark As IJComplexString
    Set oProjectedMark = m_oMfgRuleHelper.ComplexStringAlongVectorOnToSurface(oUpsidePort.Geometry, oCS, oDir)

    Set CreateEdgeCheckMark = oProjectedMark

    On Error GoTo ErrorHandler
CleanUp:
    Set oPlateWrapper = Nothing
    Set oUpsidePort = Nothing
    Set oMarkPos = Nothing
    Set oEdgeWire = Nothing
    Set oVerticalEnd = Nothing
    Set oHorizontalEnd = Nothing
    Set oVerticalLine = Nothing
    Set oHorisontalLine = Nothing
    Set oCS = Nothing
    Set oDir = Nothing
    Set oProjectedMark = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' CreateEdgeCheckMarkLogical
'
' Description:  This function creates the logical edge check marks based on given inputs.
'
' ***********************************************************************************
Public Function CreateEdgeCheckMarkLogical(ByVal Part As Object, ByVal UpSide As Long, _
                                        ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                        Optional ByVal eMarkFeature As enumMarkFeature = eIgnoreFeature, _
                                        Optional ByVal eSeamType As enumSeamType = eAllSeams, _
                                        Optional ByVal bApplyMarginOffset As Boolean) As IJMfgGeomCol3d

    Const METHOD = "CreateEdgeCheckMarkLogical"
    On Error GoTo ErrorHandler

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_EDGE_CHECKLINES_MARK, ReferenceObjColl, oGeomCol3d
    Set CreateEdgeCheckMarkLogical = oGeomCol3d

    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)

    If oConObjsCol Is Nothing Then
        'No connected objects so we can leave
        GoTo CleanUp
    ElseIf oConObjsCol.Count = 0 Then
        'No connected objects so we can leave
        GoTo CleanUp
    End If

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oMfgPart As IJMfgPlatePart

    If Not oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Exit Function
    End If

    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Set oPlateCreation_AE = oMfgPart.ActiveEntity

    Dim oOuterContourGeomCol3d As IJMfgGeomCol3d
    Set oOuterContourGeomCol3d = oPlateCreation_AE.GeometriesBeforeUnfold
    If oOuterContourGeomCol3d Is Nothing Then
        'Since we do not find any geometries to be marked we can exit
        Exit Function
    ElseIf oOuterContourGeomCol3d.GetCount = 0 Then
        'zero elements count, so we can leave
        GoTo CleanUp
    End If

    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = Part

    Dim oFeatureColl As Collection

    oSDPartSupport.GetFeatures oFeatureColl

    Dim icount As Integer
    Dim eFeatureType As StructFeatureTypes

    ' Get the corner feature collection
    Dim oSDFeatureColl As Collection
    Set oSDFeatureColl = New Collection
    For icount = 1 To oFeatureColl.Count

        oSDPartSupport.GetFeatureType oFeatureColl.Item(icount), eFeatureType

        If eFeatureType = SF_CornerFeature And eMarkFeature = eCornerFeature Then
            oSDFeatureColl.Add oFeatureColl.Item(icount)
        ElseIf eFeatureType = SF_EdgeFeature And eMarkFeature = eEdgeFeature Then
            oSDFeatureColl.Add oFeatureColl.Item(icount)
        Else
            If (eFeatureType = SF_CornerFeature Or eFeatureType = SF_EdgeFeature) And eMarkFeature = eAllFeatures Then
                oSDFeatureColl.Add oFeatureColl.Item(icount)
            End If
        End If
    Next

    If Not eMarkFeature = eIgnoreFeature Then
        ' If features not exist on plate part then exit
        If oSDFeatureColl.Count = 0 Then
            GoTo CleanUp
        End If
    End If

    Dim oSeamCol As Collection
    Dim oPositionCol As Collection
    Dim oThisPartCol As Collection
    Dim oConnPartCol As Collection
    Dim oMarkVecColl As Collection

    ' Get all the connecting seams information
    oPlateWrapper.GetConnectingSeamInfo oOuterContourGeomCol3d, oConObjsCol, oSeamCol, oPositionCol, oThisPartCol, oConnPartCol, oMarkVecColl

    If (oPositionCol Is Nothing) Then
        'Since there are no seams to be taken into account we can leave the function
        GoTo CleanUp
    End If

    If (oPositionCol.Count = 0) Then
        'Since there are no seams to be taken into account we can leave the function
        GoTo CleanUp
    Else

        Dim oNewPositionCol As Collection
        Dim oNewThisPartCol As Collection
        Dim oNewConnPartCol As Collection
        Dim oNewMarkVecColl As Collection

        If eMarkFeature = eIgnoreFeature Then
            Set oNewPositionCol = oPositionCol
            Set oNewThisPartCol = oThisPartCol
            Set oNewConnPartCol = oConnPartCol
            Set oNewMarkVecColl = oMarkVecColl
        Else
            Set oNewPositionCol = New Collection
            Set oNewThisPartCol = New Collection
            Set oNewConnPartCol = New Collection
            Set oNewMarkVecColl = New Collection

            Dim i   As Long
            For i = 1 To oPositionCol.Count

                Dim oTempPos    As IJDPosition
                Set oTempPos = oPositionCol.Item(i)

                Dim oSeam   As IJDSeamType
                Set oSeam = oSeamCol.Item(i)


                If eSeamType = eAllSeams Or (oSeam.SeamType = sstPlanningSeam And eSeamType = ePlanningSeam) Or (oSeam.SeamType = sstDesignSeam And eSeamType = eDesignSeam) Then
                    ' Get the feature range and check if the seam position is with in feature range
                    For icount = 1 To oSDFeatureColl.Count
                        Dim pMinPos As IJDPosition, pMaxPos As IJDPosition
                        oSDPartSupport.GetFeatureRange oSDFeatureColl.Item(icount), pMinPos, pMaxPos

                        If ((pMinPos.x <= oTempPos.x Or (pMinPos.x - oTempPos.x < 0.02)) And (pMinPos.y <= oTempPos.y Or (pMinPos.y - oTempPos.y < 0.02)) And (pMinPos.z <= oTempPos.z Or (pMinPos.z - oTempPos.z < 0.02)) And _
                            (pMaxPos.x >= oTempPos.x Or (oTempPos.x - pMaxPos.x < 0.02)) And (pMaxPos.y >= oTempPos.y Or (oTempPos.y - pMaxPos.y < 0.02)) And (pMaxPos.z >= oTempPos.z Or (oTempPos.z - pMaxPos.z < 0.02))) Then

                            oNewPositionCol.Add oPositionCol.Item(i)
                            oNewThisPartCol.Add oThisPartCol.Item(i)
                            oNewConnPartCol.Add oConnPartCol.Item(i)
                            oNewMarkVecColl.Add oMarkVecColl.Item(i)

                            Exit For
                        End If
                    Next
                End If
            Next
        End If

        If (oNewPositionCol.Count = 0) Then
            GoTo CleanUp
        End If

        ' Check for applying margin for the marks
        If bApplyMarginOffset = True Then
            Dim oEntityHelper As IJMfgEntityHelper
            Set oEntityHelper = New MfgEntityHelper

            Dim oMfgRuleHelper As MfgRuleHelpers.Helper
            Set oMfgRuleHelper = New MfgRuleHelpers.Helper

            For i = 1 To oNewPositionCol.Count

                Dim oSeamPosition    As IJDPosition
                Set oSeamPosition = oNewPositionCol.Item(i)

                Dim oConnectable As IJConnectable
                Set oConnectable = oNewThisPartCol.Item(i)

                Dim bIsConnected    As Boolean
                Dim oConnections    As IJElements
                oConnectable.isConnectedTo oNewConnPartCol.Item(i), bIsConnected, oConnections

                'Get margin information for offsetting the seam position
                If bIsConnected = True Then
                    ' Get the PC and return
                    For icount = 1 To oConnections.Count
                        If TypeOf oConnections.Item(icount) Is IJStructPhysicalConnection Then
                            Dim oAppConn As IJAppConnection
                            Set oAppConn = oConnections.Item(icount)

                            Dim oPorts As IJElements
                            oAppConn.enumPorts oPorts

                            Dim dMarginValue As Double
                            dMarginValue = 0#

                            Dim jCount      As Long
                            For jCount = 1 To 2

                                Dim oMfgDefport     As IJPort

                                On Error Resume Next ' GetCorrectMfgDefPort can fail with errors

                                If jCount = 1 Then
                                    Set oMfgDefport = oEntityHelper.GetCorrectMfgDefPort(oPorts.Item(1))
                                Else
                                    Set oMfgDefport = oEntityHelper.GetCorrectMfgDefPort(oPorts.Item(2))
                                End If

                                On Error GoTo ErrorHandler

                                If Not oMfgDefport Is Nothing Then
                                    Dim oMfgDefCol      As Collection
                                    Set oMfgDefCol = oMfgRuleHelper.GetMfgDefinitions(oMfgDefport)
                                Else
                                    GoTo NextItem2
                                End If

                                Dim kCount      As Long
                                For kCount = 1 To oMfgDefCol.Count
                                    Dim dTempMarginValue    As Double
                                    If (TypeOf oMfgDefCol.Item(kCount) Is IJDFabMargin) Then

                                        If TypeOf oMfgDefCol.Item(kCount) Is IJConstMargin Then
                                            Dim oConstMargin    As IJConstMargin
                                            Set oConstMargin = oMfgDefCol.Item(kCount)
                                            dTempMarginValue = oConstMargin.Value
                                        Else
                                            'Compute the margin value at seam position
                                            Dim oObliqueMargin As IJObliqueMargin
                                            Set oObliqueMargin = oMfgDefCol.Item(kCount)
                                            If Not oObliqueMargin Is Nothing Then

                                                Dim oStartPos As IJDPosition
                                                Set oStartPos = oObliqueMargin.GetStartPosition(oMfgDefport)

                                                Dim dSeamPosMarginVal   As Double
                                                Dim dDistToStartPos     As Double
                                                Dim dPortLength         As Double
                                                Dim dDiffVal            As Double

                                                dDiffVal = Abs(oObliqueMargin.EndValue - oObliqueMargin.StartValue)

                                                Dim oTopoLocate         As IJTopologyLocate
                                                Set oTopoLocate = New TopologyLocate

                                                Dim oEdgePort As IJPort
                                                Set oEdgePort = oTopoLocate.GetEdgePortForFacePort(oMfgDefport)

                                                dPortLength = CurveLength(oEdgePort.Geometry)

                                                dDistToStartPos = oSeamPosition.DistPt(oStartPos)

                                                If oObliqueMargin.StartValue < oObliqueMargin.EndValue Then
                                                    dTempMarginValue = oObliqueMargin.StartValue + (dDiffVal * (dDistToStartPos / dPortLength))
                                                Else
                                                    dTempMarginValue = oObliqueMargin.StartValue - (dDiffVal * (dDistToStartPos / dPortLength))
                                                End If

                                            End If
                                        End If

                                        If jCount = 1 Then
                                            dMarginValue = dMarginValue + dTempMarginValue
                                        Else
                                            dMarginValue = dMarginValue - dTempMarginValue
                                        End If
                                    End If
                                Next
NextItem2:
                            Next

                            ' Offset the seam position according to margin value
                            If dMarginValue <> 0 Then
                                Dim oPortSurfaceNormal      As IJDVector
                                Dim oProjPosOnSurf          As IJDPosition

                                Dim oMGHelper       As IJMfgMGHelper
                                Set oMGHelper = New MfgMGHelper

                                ' Set the Mfg Def port as first port as the margin value taken is relative to first port
                                Set oMfgDefport = oEntityHelper.GetCorrectMfgDefPort(oPorts.Item(1))

                                Dim oPortSurfaceBody    As IJSurfaceBody
                                Set oPortSurfaceBody = oMfgDefport.Geometry

                                oMGHelper.ProjectPointOnSurfaceBody oPortSurfaceBody, oSeamPosition, oProjPosOnSurf, oPortSurfaceNormal
                                oPortSurfaceBody.GetNormalFromPosition oProjPosOnSurf, oPortSurfaceNormal

                                oPortSurfaceNormal.Length = dMarginValue

                                Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
                                dNormalX = oPortSurfaceNormal.x
                                dNormalY = oPortSurfaceNormal.y
                                dNormalZ = oPortSurfaceNormal.z

                                oSeamPosition.x = oSeamPosition.x + dNormalX
                                oSeamPosition.y = oSeamPosition.y + dNormalY
                                oSeamPosition.z = oSeamPosition.z + dNormalZ
                            End If

                        End If
                    Next
                End If
            Next
        End If
    End If

    For i = 1 To oNewPositionCol.Count
        Dim oSystemMark As IJMfgSystemMark
        Dim oObjSystemMark As IUnknown
        Dim oMoniker As IMoniker
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3d As IJMfgGeom3D
        Dim oPosition As IJDPosition
        Dim oCS As IJComplexString
        Dim oMarkVector As IJDVector

        Set oPosition = oNewPositionCol.Item(i)
        Set oMarkVector = oNewMarkVecColl.Item(i)

        ' Create the edge check mark geometry
        Set oCS = CreateLogicalEdgeCheckMarkGeometry(Part, oPosition, oMarkVector, UpSide)

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        'Set the marking side
        oSystemMark.SetMarkingSide UpSide

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark

        oMarkingInfo.Name = "EDGE_CHECK_MARK"

        Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_EDGE_CHECKLINES_MARK

        Set oObjSystemMark = oSystemMark

        oSystemMark.Set3dGeometry oGeom3d

        oGeomCol3d.AddGeometry 1, oGeom3d

NextItem:
        Set oPosition = Nothing
        Set oCS = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oObjSystemMark = Nothing
    Next i

    'Return the 3d collection
    Set CreateEdgeCheckMarkLogical = oGeomCol3d

CleanUp:
    Set oOuterContourGeomCol3d = Nothing
    Set oPlateWrapper = Nothing
    Set oPositionCol = Nothing
    Set oSeamCol = Nothing
    Set oPositionCol = Nothing
    Set oThisPartCol = Nothing
    Set oConnPartCol = Nothing
    Set oMarkVecColl = Nothing
    Set oNewPositionCol = Nothing
    Set oNewThisPartCol = Nothing
    Set oNewConnPartCol = Nothing
    Set oNewMarkVecColl = Nothing
    Set oSDFeatureColl = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateLogicalEdgeCheckMarkGeometry()
'
' Description:  it create the edge check mark of fixed length at the given position
'
' ***********************************************************************************
Public Function CreateLogicalEdgeCheckMarkGeometry(ByVal oPlate As IJPlatePart, ByVal oMarkPos As IJDPosition, ByVal oMarkVector As IJDVector, ByVal eUpSide As Long) As IJComplexString
    Const METHOD = "CreateLogicalEdgeCheckMarkGeometry"
    On Error GoTo ErrorHandler

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlate

    ' Get the plate upside surface
    Dim oUpsidePort As IJPort
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(eUpSide)

    ' Intially create a line with some length
    oMarkVector.Length = 0.2

    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d

    oMarkLine.DefineBy2Points oMarkPos.x - oMarkVector.x, oMarkPos.y - oMarkVector.y, oMarkPos.z - oMarkVector.z, oMarkPos.x + oMarkVector.x, oMarkPos.y + oMarkVector.y, oMarkPos.z + oMarkVector.z

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False

    'Project mark back on to plate
    '(will trim the length)

    Dim oProjCS As IJComplexString

    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oUpsidePort.Geometry

    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    ' Project the line created on the plate part surface
    oMfgMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

    If oProjCS Is Nothing Then
        GoTo CleanUp
    End If

    Dim oCurve As IJCurve
    Set oCurve = oProjCS

    Dim dCurveLength     As Double
    dCurveLength = oCurve.Length

    ' If the length of projects CS is more then 50mm then trim it.
    If dCurveLength > 0.05 Then

        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        Dim oStartPos As IJDPosition
        Set oStartPos = New DPosition

        Dim oEndPos As IJDPosition
        Set oEndPos = New DPosition

        oStartPos.Set dStartX, dStartY, dStartZ
        oEndPos.Set dEndX, dEndY, dEndZ

        ' we need edge check marks with length 50 mm
        If oStartPos.DistPt(oMarkPos) > oEndPos.DistPt(oMarkPos) Then
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - 0.05, oStartPos
        Else
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - 0.05, oEndPos
        End If
    End If

    Set CreateLogicalEdgeCheckMarkGeometry = oProjCS

CleanUp:
    Set oPlateWrapper = Nothing
    Set oUpsidePort = Nothing
    Set oSurfaceBody = Nothing
    Set oMarkPos = Nothing
    Set oMarkLine = Nothing
    Set oCS = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function GetMarginValueOnThePort(oPort As IJPort) As Double

Const METHOD = "GetMarginValueOnThePort"
On Error GoTo ErrorHandler

Dim oMfgDefCol As Collection
Dim oConstMargin As IJConstMargin
Dim oObliqueMargin As IJObliqueMargin

Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)

If oMfgDefCol.Count > 0 Then
    Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
    lFabMargin = 0
    lAssyMargin = 0
    lCustomMargin = 0
    Dim j As Integer
    For j = 1 To oMfgDefCol.Count
        If TypeOf oMfgDefCol.Item(j) Is IJAssyMarginChild Then
            Set oConstMargin = oMfgDefCol.Item(j)
            lAssyMargin = lAssyMargin + oConstMargin.Value
        ElseIf TypeOf oMfgDefCol.Item(j) Is IJObliqueMargin Then
            Set oObliqueMargin = oMfgDefCol.Item(j)
            If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                lFabMargin = lFabMargin + oObliqueMargin.EndValue
            Else
                lFabMargin = lFabMargin + oObliqueMargin.StartValue
            End If
        ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
            Set oConstMargin = oMfgDefCol.Item(j)
            lFabMargin = lFabMargin + oConstMargin.Value
        'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
        End If

    Next j
    If lAssyMargin <> 0 Or lFabMargin <> 0 Or lCustomMargin <> 0 Then
         Dim TotMargin As Double
         TotMargin = lAssyMargin + lFabMargin + lCustomMargin
    End If
End If
GetMarginValueOnThePort = TotMargin

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetPhysicalConnectionData(ByVal oThisPart As Object, ByVal oReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal bReturnAll As Boolean) As Collection

Const METHOD = "GetPhysicalConnectionData"
On Error GoTo ErrorHandler

    Dim index As Long
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oConnection As IJAppConnection
    Dim aConnectionData As ConnectionData

    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = oThisPart

    Set GetPhysicalConnectionData = New Collection

    For index = 1 To oReferenceObjColl.Count
        If TypeOf oReferenceObjColl.Item(index) Is IJStructPhysicalConnection Then
            Dim bIsCrossOfTee As Boolean
            Dim oConnType As ContourConnectionType

            Set oConnection = oReferenceObjColl.Item(index)
            oSDPartSupport.GetConnectionTypeForContour oConnection, _
                                                       oConnType, _
                                                       bIsCrossOfTee

            If bReturnAll = True Or (oConnType = PARTSUPPORT_CONNTYPE_LAP) Or _
               (oConnType = PARTSUPPORT_CONNTYPE_PROFILE_END) Or _
               (oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsCrossOfTee) Then

               Dim oPortElements As IJElements
               oConnection.enumPorts oPortElements

               Dim oPort1 As IJPort
               Dim oport2 As IJPort

               Set oPort1 = oPortElements.Item(1)
               Set oport2 = oPortElements.Item(2)

               If (oPort1.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oPort1
                    Set aConnectionData.ToConnectable = oport2.Connectable
                    Set aConnectionData.ToConnectedPort = oport2
                    GetPhysicalConnectionData.Add aConnectionData
               ElseIf (oport2.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oport2
                    Set aConnectionData.ToConnectable = oPort1.Connectable
                    Set aConnectionData.ToConnectedPort = oPort1
                    GetPhysicalConnectionData.Add aConnectionData
               End If
            End If
        End If

    Next index

    Set oSDPartSupport = Nothing
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function GetPlateNeutralSurfaceNormal(ByVal oPlatePart As Object) As IJDVector

Const METHOD = "GetPlateNeutralSurfaceNormal"
On Error GoTo ErrorHandler

    Set GetPlateNeutralSurfaceNormal = Nothing
    Dim oMfgPart As IJMfgPlatePart
    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlatePart
    Dim oNeutralSurface As IJSurfaceBody
    Dim oPoint As IJDPosition, oVector As IJDVector
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
        Set oNeutralSurface = oMfgPlateCreation_AE.NeutralSurface
        Dim oPlane As IJPlane
        If TypeOf oNeutralSurface Is IJPlane Then
        Set oPlane = oNeutralSurface
        If Not oPlane Is Nothing Then
            Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
            oPlane.GetNormal dNormalX, dNormalY, dNormalZ
            Set GetPlateNeutralSurfaceNormal = New DVector
            GetPlateNeutralSurfaceNormal.Set dNormalX, dNormalY, dNormalZ
        End If
        Else
        Dim oSGOSurfaceBodyUtil As GSCADShipGeomOps.SGOSurfaceBodyUtilities
        Set oSGOSurfaceBodyUtil = New SGOSurfaceBodyUtilities
        oSGOSurfaceBodyUtil.GetPointAndNormalNearestBoundingBoxCenter oNeutralSurface, oPoint, oVector

        Set GetPlateNeutralSurfaceNormal = New DVector
        GetPlateNeutralSurfaceNormal.Set oVector.x, oVector.y, oVector.z
        End If



        Set oMfgPlateCreation_AE = Nothing
    Else
        Exit Function
    End If

    Set oPlateWrapper = Nothing
    Set oMfgPart = Nothing
    Set oNeutralSurface = Nothing

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'********************************************************************
' Routine: CreateEndFittingMark
' Description: This function Creates the End Fitting Mark(s) at the Free End(s) of the
'               ... Profile Parts.
' Inputs:  Plate Part, Collection of Profile Location Marking Lines, Molded Side, UpSide,
'          ... Geom3D Collection, NormalVector to the Plate Part
' Outputs: Procedure creates the "End Fitting Marks"
' Notes:
'********************************************************************

'Public Sub CreateEndFittingMark(Part As Object, _
'                                    oCS As IJComplexString, _
'                                    oWB As IJWireBody, _
'                                    oPartSupport As GSCADSDPartSupport.ProfilePartSupport, _
'                                    oSDPlateWrapper As StructDetailObjects.PlatePart, _
'                                    oSDProfileWrapper As StructDetailObjects.ProfilePart, _
'                                    sMoldedSide As String, _
'                                    UpSide As Long, _
'                                    oConnectionData As ConnectionData, _
'                                    oGeomCol3D As IJMfgGeomCol3d, _
'                                    oPlateNormal As IJDVector)

Public Sub CreateEndFittingMark(Part As Object, _
                                    oCS As IJComplexString, _
                                    oWB As IJWireBody, _
                                    oPartSupport As Object, _
                                    oSDPlateWrapper As StructDetailObjects.PlatePart, _
                                    oSDProfileWrapper As Object, _
                                    sMoldedSide As String, _
                                    UpSide As Long, _
                                    oConnectionData As ConnectionData, _
                                    oGeomCol3d As IJMfgGeomCol3d, _
                                    oPlateNormal As IJDVector, _
                                    bFlip As Boolean)

        Const METHOD = "CreateEndFittingMark"
        On Error GoTo ErrorHandler

        '*** Declaration of Variables ***'
        Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
        Dim lNoOfObj As Long, i As Long
        Dim oObj As Object, oConnectedObj As Object
        Dim oPort1 As IJPort
        Dim oNamed As IJNamedItem
        Dim oPlateName As IJNamedItem
        Dim oPCWB As IJWireBody
        Dim oProjectedPt As IJDPosition
        Dim dDistFromStart As Double, dDistFromEnd As Double, dDepth As Double
        Dim bIsStartFreeEnd As Boolean, bIsEndFreeEnd As Boolean
        Dim oGeom3d As IJMfgGeom3D
        Dim oToolBox As IJDTopologyToolBox
        Dim oProfileWB As IJWireBody
        Dim oStartPos As New DPosition, oEndPos As New DPosition
        Dim oStartDir As New DVector, oEndDir As New DVector, oTDir As New DVector
        Dim oProfilePartSupport As IJProfilePartSupport
        Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

        Dim oPlatePartSupport As IJPlatePartSupport
        Dim dThickness As Double, dCriterion As Double

        '*** Initialization of Variables ***'
        Set oToolBox = New IMSModelGeomOps.DGeomOpsToolBox

        '*** For the Given Profile, Get collections of
        '... 1. Physical Connections
        '... 2. Connected Objects
        '... 3. Objects Connected at Base Port
        '... 4. Objects Connected at Offset Port
        oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol

        '*** Get the # of Physical Connections ***'
        lNoOfObj = oConObjCol.Count

        '*** Get the Start & End Points of the WireBody ***'
        oWB.GetEndPoints oStartPos, oEndPos

        bIsStartFreeEnd = True
        bIsEndFreeEnd = True

        '*** Get the Profile WireBody ***'
        Set oProfileWB = oMfgRuleHelper.ComplexStringToWireBody(oCS)

        '*** Get the start and end positions and their directions ***'
        oProfileWB.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir
        'The oEndDir is nothing but the Tangent Vector to the Profile Location Marking Line/Curve


        '*** Get the Profile Depth to set the distance tolerence to check if a physical connection
        '... exists at the end of the profile part ***'
        'Set oProfilePartSupport = oPartSupport
        'oProfilePartSupport.GetWebDepth dDepth
        If TypeOf oPartSupport Is IJPlatePartSupport Then
            Set oPlatePartSupport = oPartSupport
            oPlatePartSupport.GetThickness dThickness
            dCriterion = dThickness
        ElseIf TypeOf oPartSupport Is IJProfilePartSupport Then
            Set oProfilePartSupport = oPartSupport
            oProfilePartSupport.GetWebDepth dDepth
            dCriterion = dDepth
        End If


        '*** lNoOfObj > 1 means that there are other objects (other than base plate)
        '... having physical connection with the profile.
        '... So check which end of the profile is/are Free Ends, and place the marks accordingly.
        If lNoOfObj > 1 Then

            For i = 1 To lNoOfObj

                ' Get the object connected at the other port
                Set oPort1 = oOtherPortCol.Item(i)
                Set oObj = oPort1.Connectable

                ' This gives the name of the object connected at the other port
                Set oNamed = oObj

                ' This gives the name of the Base Plate
                Set oPlateName = oSDPlateWrapper.object

                'If the Connected Object is Base Plate, ignore it...
                '... and if not, then get the distance of PC from StartPt and EndPt.
                '... if the distance is greater than certain tolerance, that means the object is not
                '... connected, so this is a free end cut. Then place the End Fitting Mark.

                'If oNamed.Name <> oPlateName.Name Then
                If Not (oObj Is oSDPlateWrapper.object) Then

                    '*** Get the Geometry of the Connected Object from the Port ***'
                    Set oConnectedObj = oPort1.Geometry

                    '*** Get the WireBody object for the Location Marking Line ***'
                    Set oPCWB = oConnCol.Item(i)

                    '*** Find the closest point on the WireBody of Physical Connection
                    '... in order to find the minimum distance from Start Point and End Point
                    oToolBox.GetNearestPointOnWireBodyFromPoint oPCWB, oStartPos, Nothing, oProjectedPt

                    '*** Get the distance of the point on WB from StartPoint***'
                    dDistFromStart = oProjectedPt.DistPt(oStartPos)

                    oToolBox.GetNearestPointOnWireBodyFromPoint oPCWB, oEndPos, Nothing, oProjectedPt

                    '*** Get the distance of the point on WB from EndPoint***'
                    dDistFromEnd = oProjectedPt.DistPt(oEndPos)

                    '*** If the distance is less than WebDepth of the profile, then the object has
                    '... physical connection with the Profile End
'                    If dDistFromStart < dDepth Then bIsStartFreeEnd = False
'                    If dDistFromEnd < dDepth Then bIsEndFreeEnd = False
                    If dDistFromStart < dCriterion Then bIsStartFreeEnd = False
                    If dDistFromEnd < dCriterion Then bIsEndFreeEnd = False
                End If
            Next i

            'Checking if Connected object is a PlatePart
            If TypeOf oSDProfileWrapper.object Is IJPlatePart Then
                Dim oRootPartSystem                As IJSystem
                Dim oStructDetailHelper            As GSCADStructDetailUtil.StructDetailHelper

                Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
                oStructDetailHelper.IsPartDerivedFromSystem oSDProfileWrapper.object, oRootPartSystem, True

                If Not oRootPartSystem Is Nothing Then
                    Dim oPlateUtils As IJPlateAttributes
                    Set oPlateUtils = New PlateUtils

                    'Checking if the Connected PlatePart is a Bracket
                    If (IsFlangedBracket(oSDProfileWrapper.object) = True) Or (oPlateUtils.IsBracketByPlane(oRootPartSystem) = True) Or (oPlateUtils.IsTrippingBracket(oRootPartSystem) = True) Then
                        Dim oPlane                          As IJPlane
                        Dim oPoint1                         As IJPoint
                        Dim oPoint2                         As IJPoint
                        Dim strRootSelector                 As String
                        Dim oSupport1                       As Object
                        Dim oSupport2                       As Object
                        Dim oSupport3                       As Object
                        Dim oSupport4                       As Object
                        Dim oSupport5                       As Object
                        oPlateUtils.GetInput_BracketByPlane_IndividualSupports oRootPartSystem, oPlane, oPoint1, oPoint2, strRootSelector, oSupport1, oSupport2, oSupport3, oSupport4, oSupport5

                        Dim bFittingMarkRemoval As Boolean
                        Dim b2SBracket As Boolean

                        Dim oBracketSupportCol             As New Collection

                        If oSupport3 Is Nothing Then 'Type of Bracket is 2S Bracket
                            b2SBracket = True

                            'Checking the Supports fo 2S Brackets have Profile and Plate as support
                            If TypeOf oSupport1 Is IJPlate Then
                                If TypeOf oSupport2 Is IJProfile Then
                                    bFittingMarkRemoval = True
                                End If
                            ElseIf TypeOf oSupport1 Is IJProfile Then
                                If TypeOf oSupport2 Is IJPlate Then
                                    bFittingMarkRemoval = True
                                End If
                            End If
                        Else 'Type of Bracket is 3S Bracket
                            'Checking the Supports fo 2S Brackets have Profile and Plate and Profile support
                            Dim lCount As Long, lProfileCount As Long, lPlateCount As Long

                            'Checking for Object1 -> Support1
                            If TypeOf oSupport1 Is IJPlate Then
                                lPlateCount = lPlateCount + 1
                            ElseIf TypeOf oSupport1 Is IJProfile Then
                                lProfileCount = lProfileCount + 1
                                oBracketSupportCol.Add oSupport1
                                bFittingMarkRemoval = True
                            End If
                            'Checking for Object2 -> Support2
                            If TypeOf oSupport2 Is IJPlate Then
                                lPlateCount = lPlateCount + 1
                            ElseIf TypeOf oSupport2 Is IJProfile Then
                                lProfileCount = lProfileCount + 1
                                oBracketSupportCol.Add oSupport2
                                bFittingMarkRemoval = True
                            End If
                            'Checking for Object3 -> Support3
                            If TypeOf oSupport3 Is IJPlate Then
                                lPlateCount = lPlateCount + 1
                            ElseIf TypeOf oSupport3 Is IJProfile Then
                                lProfileCount = lProfileCount + 1
                                oBracketSupportCol.Add oSupport3
                                bFittingMarkRemoval = True
                            End If
                        End If

                        If bFittingMarkRemoval = True Then
                            Dim oProfileModelBody As IJDModelBody
                            Dim oClosestPos As IJDPosition
                            Dim dMinimumDistance1 As Double, dMinimumDistance2 As Double
                            If b2SBracket = True Then
                                If TypeOf oSupport1 Is IJProfile Then
                                    Set oProfileModelBody = oSupport1
                                ElseIf TypeOf oSupport2 Is IJProfile Then
                                    Set oProfileModelBody = oSupport2
                                End If
                                oProfileModelBody.GetMinimumDistanceFromPosition oStartPos, oClosestPos, dMinimumDistance1
                                oProfileModelBody.GetMinimumDistanceFromPosition oEndPos, oClosestPos, dMinimumDistance2

                                If dMinimumDistance1 < dMinimumDistance2 Then
                                    bIsStartFreeEnd = False
                                Else
                                    bIsEndFreeEnd = False
                                End If
                            Else
                                For lCount = 1 To oBracketSupportCol.Count
                                    Set oProfileModelBody = oBracketSupportCol.Item(lCount)

                                    oProfileModelBody.GetMinimumDistanceFromPosition oStartPos, oClosestPos, dMinimumDistance1
                                    oProfileModelBody.GetMinimumDistanceFromPosition oEndPos, oClosestPos, dMinimumDistance2

                                    If dMinimumDistance1 <> dMinimumDistance2 Then
                                        If dMinimumDistance1 < dMinimumDistance2 Then
                                            bIsStartFreeEnd = False
                                        Else
                                            bIsEndFreeEnd = False
                                        End If
                                    End If
                                Next
                            End If

                            Set oClosestPos = Nothing
                            Set oProfileModelBody = Nothing
                        End If
                        Set oSupport1 = Nothing
                        Set oSupport2 = Nothing
                        Set oSupport3 = Nothing
                        Set oSupport4 = Nothing
                        Set oSupport5 = Nothing
                        Set oPoint1 = Nothing
                        Set oPoint2 = Nothing
                        Set oPlane = Nothing
                        Set oBracketSupportCol = Nothing
                    End If
                    Set oPlateUtils = Nothing
                End If
                Set oRootPartSystem = Nothing
                Set oStructDetailHelper = Nothing
            End If


            '*** This means that the Start End is Free End, Place the Fitting Mark Here ***'
            If bIsStartFreeEnd Then
                '*** Create the End Fitting Mark at StartPoint***'
                Set oTDir = oStartDir
                Set oGeom3d = GetEndFittingMarkAtPoint(oStartPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oStartDir, oPlateNormal, oTDir, bFlip)
                '*** Add the Geom3d object to collection ***'
                oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
            End If

            '*** This means that the Other End is Free End, Place the Fitting Mark Here ***'
            If bIsEndFreeEnd Then
                '*** Create the End Fitting Mark at EndPoint***'
                Set oTDir = oEndDir
                oTDir.Length = -1#
                Set oGeom3d = GetEndFittingMarkAtPoint(oEndPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oEndDir, oPlateNormal, oTDir, bFlip)

                '*** Add the Geom3d object to collection ***'
                oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
            End If

        '*** lNoOfObj = 1 means that the only physical connection that is present is
        '... between the profile and the base plate.
        '... So place the fitting marks at both ends of the profile.
        Else
            Dim oFittingMarkVec As IJDVector

            If TypeOf Part Is IJPlatePart Then
                Set oFittingMarkVec = GetThicknessDirectionVector(oCS, oSDProfileWrapper, sMoldedSide)
            Else
            Dim oProfileSection As IJDProfileSection
            Set oProfileSection = oConnectionData.ToConnectable

            Dim eMountingFace As ProfileFaceName
            eMountingFace = oProfileSection.mountingFace

            ' if the profile mounting face is left web or right web then get web direction and
            ' use web direction as fitting mark vector.
            If eMountingFace = LeftWeb Or eMountingFace = RightWeb Then
                Set oFittingMarkVec = GetWebDirectionVector(oProfileSection)
            End If
            End If

            Set oTDir = oStartDir

            '*** Create the End Fitting Mark at StartPoint***'
            Set oGeom3d = GetEndFittingMarkAtPoint(oStartPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oStartDir, oPlateNormal, oTDir, bFlip, , oFittingMarkVec)

            '*** Add the Geom3d object to collection ***'
            oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d

            '*** Create the End Fitting Mark at EndPoint***'

            Set oTDir = oEndDir
            oTDir.Length = -1#
            Set oGeom3d = GetEndFittingMarkAtPoint(oEndPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oEndDir, oPlateNormal, oTDir, bFlip, , oFittingMarkVec)

            '*** Add the Geom3d object to collection ***'
            oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d

        End If


CleanUp:
        Set oConObjCol = Nothing
        Set oConnCol = Nothing
        Set oThisPortCol = Nothing
        Set oOtherPortCol = Nothing
        Set oObj = Nothing
        Set oConnectedObj = Nothing
        Set oPort1 = Nothing
        Set oNamed = Nothing
        Set oPlateName = Nothing
        Set oPCWB = Nothing
        Set oProjectedPt = Nothing
        Set oToolBox = Nothing
        Set oProfileWB = Nothing
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oStartDir = Nothing
        Set oEndDir = Nothing
        Set oProfilePartSupport = Nothing

        Exit Sub


ErrorHandler:

        Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'********************************************************************
' Routine: GetEndFittingMarkAtPoint
' Description: Creates the End Fitting Mark Line (of specified Length) normal to the
'              ... Location mark line.
' Inputs: Point at which Mark is to be placed, ProfileWrapper object, Molded Side, UpSide,
'              ... Tangent and Normal Vectors to the surface body
' Outputs:  End Fitting Mark Line
' Notes:
'********************************************************************
Public Function GetEndFittingMarkAtPoint(oStartPos As IJDPosition, _
                                            oSDProfileWrapper As Object, _
                                            sMoldedSide As String, _
                                            UpSide As Long, _
                                            oConnectionData As ConnectionData, _
                                            oTanVec As IJDVector, _
                                            oPlateNormal As IJDVector, _
                                            oTDir As IJDVector, _
                                            bFlip As Boolean, _
                                            Optional dLength As Double = 0, _
                                            Optional oFittingMarkVec As IJDVector) As IJMfgGeom3D


        Const METHOD = "GetEndFittingMarkAtPoint"
        On Error GoTo ErrorHandler

        '*** Declaration of Variables ***'
        Dim oThickVecDir2D As New DVector
        Dim oNewPos As New DPosition
        Dim oLineCS As IJComplexString
        Dim oLine1 As IJLine
        Dim oSystemMark As IJMfgSystemMark
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3d As IJMfgGeom3D
        Dim oMoniker As IMoniker
        Dim oNormalVec As New DVector
        Dim dDot As Double


        If dLength = 0 Then
             dLength = END_FITTING_MARK_LENGTH
        End If

        '*** Initialization of Variables ***'
        Set oLine1 = New Line3d
        Set oLineCS = New ComplexString3d

        '*** Get the thickness direction ***'
        Set oThickVecDir2D = GetThicknessDirectionVectorAtAGivenPos(oStartPos, oSDProfileWrapper, Nothing, sMoldedSide)
        oThickVecDir2D.Length = dLength

        If oFittingMarkVec Is Nothing Then
            '*** Get the Normal to the Plate surface in order to place the Fitting Mark ***'
            Set oNormalVec = oTanVec.Cross(oPlateNormal)

            dDot = oNormalVec.Dot(oThickVecDir2D)

            '*** Determine the thickness direction  ***'
            If dDot > 0 Then
                oNormalVec.Length = dLength
            Else
                oNormalVec.Length = dLength * -1#
            End If

            '*** Flip the Mark  ***'
            If bFlip Then
                oNormalVec.Length = dLength * -1#
            End If

            '*** Get the other co-ordinates of the mark ***'
            Set oNewPos = oStartPos.Offset(oNormalVec)
        Else
            oFittingMarkVec.Length = dLength
            Set oNewPos = oStartPos.Offset(oFittingMarkVec)
        End If

        '*** Construct the Mark ***'
        oLine1.DefineBy2Points oStartPos.x, oStartPos.y, oStartPos.z, oNewPos.x, oNewPos.y, oNewPos.z

        '*** Create the Complex String ***'
        oLineCS.AddCurve oLine1, True

        Dim oSystemMarkFactory As New GSCADMfgSystemMark.MfgSystemMarkFactory

        Set oSystemMark = oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        '*** Set the marking side ***'
        oSystemMark.SetMarkingSide UpSide

        '*** QI for the MarkingInfo object on the SystemMark ***'
        Set oMarkingInfo = oSystemMark

        '*** Set the name and thickness for marking info ***'
        oMarkingInfo.Name = ""

        'Set the Fitting Mark's thickness direction (i.e. towards Material or not)
        oMarkingInfo.ThicknessDirection = oTDir

        Dim oGeom3dFactory As New GSCADMfgGeometry.MfgGeom3dFactory
        Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

        Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        oGeom3d.PutGeometry oLineCS
        oGeom3d.FaceId = UpSide

        '*** Set the Type of the Mark ***'
        oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
        oGeom3d.PutSubGeometryType STRMFG_FET_MARK
        oSystemMark.Set3dGeometry oGeom3d

        Set oMoniker = oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
        oGeom3d.PutMoniker oMoniker


        '*** Return the Complex String ***'
        Set GetEndFittingMarkAtPoint = oGeom3d


CleanUp:

        Set oMarkingInfo = Nothing
        Set oThickVecDir2D = Nothing
        Set oNewPos = Nothing
        Set oLineCS = Nothing
        Set oLine1 = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oMoniker = Nothing
        Set oNormalVec = Nothing

        Exit Function

ErrorHandler:

        Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetThicknessDirectionVectorAtAGivenPos(oInputPos As IJDPosition, oSDConWrapper As Object, oConnObjNormal As IJDVector, sMoldedSide As String) As IJDVector
    Const METHOD = "GetThicknessDirectionVectorAtAGivenPos"

    On Error GoTo ErrorHandler
    Dim oProjPos As IJDPosition
    Dim dThickness As Double

    'we must adjust the returned TD vector for each of the new seperate complex strings
    Dim oSurfaceBodyCon1 As IJSurfaceBody
    If TypeOf oSDConWrapper.object Is IJPlatePart Then
        Dim oPartSupp As IJPartSupport
        Dim oPlatePartSupp As IJPlatePartSupport
        Set oPartSupp = New PlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = oSDConWrapper.object

        If sMoldedSide = "Base" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oSurfaceBodyCon1
        Else
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oSurfaceBodyCon1
        End If

        oPlatePartSupp.GetThickness dThickness

        '*** For Debug ***'
'        Dim oModelBody As IJDModelBody
'        Set oModelBody = oSurfaceBodyCon1
'        Dim SBC1FileName As String
'        SBC1FileName = Environ("TEMP")
'        If SBC1FileName = "" Or SBC1FileName = vbNullString Then
'            SBC1FileName = "C:\temp" 'Only use C:\Temp if there is a %TEMP% failure
'        End If
'        SBC1FileName = SBC1FileName & "\oSurfaceBodyCon1.sat"
'        oModelBody.DebugToSATFile SBC1FileName
        '*****************'

        Set oPlatePartSupp = Nothing
        Set oPartSupp = Nothing
    ElseIf TypeOf oSDConWrapper.object Is IJProfilePart Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_RIGHT).Geometry
        Else
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry
        End If
        dThickness = oSDConWrapper.WebThickness
    ElseIf TypeOf oSDConWrapper.object Is ISPSMemberPartPrismatic Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPort(JXSEC_WEB_RIGHT).Geometry
        Else
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPort(JXSEC_WEB_LEFT).Geometry
        End If
        dThickness = oSDConWrapper.WebThickness
    End If

    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

    Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oInputPos, oSurfaceBodyCon1, Nothing)

    If Not oConnObjNormal Is Nothing Then
        Dim oNormalAtProjPoint As IJDVector
        Set oNormalAtProjPoint = oConnObjNormal.Clone
        oNormalAtProjPoint.Length = -1# * dThickness

        Set GetThicknessDirectionVectorAtAGivenPos = oNormalAtProjPoint
    Else
    Set GetThicknessDirectionVectorAtAGivenPos = oProjPos.Subtract(oInputPos)
    End If

CleanUp:
    Set oSurfaceBodyCon1 = Nothing
    'Set oMfgRuleHelper = Nothing
Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1001, , "RULES")
    GoTo CleanUp
End Function

'********************************************************************
' Routine: CreateThicknessDirectionMark
' Description: This function creates Thickness Direction "Bubble" Mark
'               .. and returns Geom2DCollection object.
' Inputs: Plate Part, UpSide, SelectiveRecompute flag, ReferenceObjColl and Geometry Type
' Outputs: Geom2DCollection object
' Notes:
'********************************************************************

Public Function CreateThicknessDirectionMark(ByVal Part As Object, _
                                                  ByVal UpSide As Long, _
                                                  ByVal bSelectiveRecompute As Boolean, _
                                                  ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                                                  ByVal sGeometryType As StrMfgGeometryType, _
                                                  ByVal sBubbleType As String) _
                                                  As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

    Const METHOD = "CreateThicknessDirectionMark"
    On Error GoTo ErrorHandler

    '*** Create the Plate Wrapper Object ***'
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr


    '*** Create the Profile Wrapper Object ***'
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    Dim oMfgPart As Object
    Dim oGeomCol2d As IJMfgGeomCol2d


    If TypeOf Part Is IJPlatePart Then

        'Set oMfgPart = New IJMfgPlatePart
        Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
        Set oPlateWrapper.object = Part

        '*** Get the Manufactured Plate Part ***'

        If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
            Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
            Set oMfgPlateWrapper.object = oMfgPart
        Else
            Exit Function
        End If
        Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries
    ElseIf TypeOf Part Is IJProfilePart Then

        'Set oMfgPart = New IJMfgProfilePart
        Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper.object = Part

        '*** Get the Manufactured Profile Part ***'
        Dim oMfgProfilePart As IJMfgProfilePart
        If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
            Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
            Set oMfgProfileWrapper.object = oMfgPart
            Set oMfgProfilePart = oMfgPart
            Set oGeomCol2d = oMfgProfilePart.FinalGeometriesAfterProcess2D
        Else
            Exit Function
        End If

    End If

    Dim oMfgMGHelper As IJMfgMGHelper

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If


    '*** Declaration of Variables ***'
    Dim oGeom2d As IJMfgGeom2d
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Dim oSGOWireBodyUtilities As IJSGOWireBodyUtilities
    Set oSGOWireBodyUtilities = New SGOWireBodyUtilities
    Dim oMfgGeomHelper As New MfgGeomHelper

    Dim i As Long

    Dim oCurve As IJCurve
    Dim oWB As IJWireBody, oWB1 As IJWireBody, oWB2 As IJWireBody, oWB3 As IJWireBody
    Dim oArcs As IArcs3d, oArcs1 As IArcs3d, oArcs2 As IArcs3d
    Dim oComplexStrings As IComplexStrings3d
    Dim oArc As New Arc3d
    Dim oArc1 As New Arc3d
    Dim oArc2 As New Arc3d
    Dim dNewPointOffset As Double
    Dim oCSColl As IJElements

    Dim oMidPos As New DPosition
    Dim oStartPos As New DPosition, oEndPos As New DPosition, oNewPos As New DPosition
    Dim oArcPos1 As New DPosition, oArcPos2 As New DPosition

    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    Dim dStartParam As Double, dEndParam As Double, dMidParam As Double, dMidPar As Double
    Dim dMidX As Double, dMidY As Double, dMidZ As Double

    Dim dTanX As Double, dTanY As Double, dTanZ As Double
    Dim dTan2x As Double, dTan2Y As Double, dTan2Z As Double

    Dim dTemp As Double
    Dim oNormalVec As New DVector
    Dim oNormalVecAtStart As New DVector, oTangentVecAtStart As New DVector, oZVector As New DVector
    Dim oTangentVecAtMid As New DVector
    Dim dRadius As Double

    Dim dLength As Double, dDistOnCurve As Double
    Dim dDotProd As Double
    Dim oCrossStart As New DVector, oCrossMid As New DVector

    Dim oCS1 As IJComplexString, oCS2 As IJComplexString, oCS3 As IJComplexString, oUpdatedCS As IJComplexString
    Dim oCS1Elem As IJElements, oCS2Elem As IJElements, oCS3Elem As IJElements, oUpdatedGeomColl As IJElements



    '=== Code to determine thickness direction ===

    Dim oRelataionHelper As IMSRelation.DRelationHelper
    Dim oCollectionHelper As IMSRelation.DCollectionHelper
    Dim oThicknessDir As IJDVector
    Dim o2DThicknessDirVec As IJDVector
    Dim oTransMat As IJDT4x4
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

    'Get the transformation matrix for thickness direction
    Set oTransMat = GetApproximate3DTo2DTransMatrix(oMfgPart, JXSEC_WEB_LEFT)
    '===

    Dim oDeleteMarkColl As New Collection
    Dim oDeletedFrame As IJMfgGeom2d

    For i = oGeomCol2d.GetCount To 1 Step -1
        Set oGeom2d = oGeomCol2d.GetGeometry(i)

        If oGeom2d.GetGeometryType = sGeometryType And oGeom2d.GetSubGeometryType <> STRMFG_FITTINGANGLE Then

            '*** Get the curve ***
            Set oCurve = oGeom2d.GetGeometry

            'Check if there is any frame mark overlapping with this location mark.
            Dim j As Long
            For j = oGeomCol2d.GetCount To 1 Step -1
                Dim oFrameGeom2d As IJMfgGeom2d

                Set oFrameGeom2d = oGeomCol2d.GetGeometry(j)

                If ((oFrameGeom2d.GetGeometryType = STRMFG_FRAMELINE_MARK) Or _
                    (oFrameGeom2d.GetGeometryType = STRMFG_BUTTOCKLINE_MARK) Or _
                    (oFrameGeom2d.GetGeometryType = STRMFG_WATERLINE_MARK)) Then

                    Dim oFrameCurve As IJCurve
                    Dim bOverlapExists As Boolean

                    Set oFrameCurve = oFrameGeom2d.GetGeometry
                    'TODO
                    bOverlapExists = oMfgGeomHelper.CheckOverlapBetweenTwoCurvesWithinTol(oFrameCurve, oCurve, 0.001)
                    If bOverlapExists Then
                        ' If this frame mark is already handled, skip the check
                        Dim bSkipThisFrame As Boolean
                        bSkipThisFrame = False
                        For Each oDeletedFrame In oDeleteMarkColl
                            If oDeletedFrame Is oFrameGeom2d Then
                                bSkipThisFrame = True
                            End If
                        Next

                        If bSkipThisFrame = False Then
                            ' Update the profile location geometry with the frame line geometry
                            Set oCurve = oFrameCurve
                            oDeleteMarkColl.Add oFrameGeom2d
                        Else
                            ' If the frame is handled, remove this location mark and goto the next geom2d
                            oDeleteMarkColl.Add oGeom2d
                            GoTo NextLocationMark
                        End If
                    End If
                    Set oFrameCurve = Nothing
                End If
                Set oFrameGeom2d = Nothing
            Next

            Set oRelataionHelper = oGeom2d
            Set oCollectionHelper = oRelataionHelper.CollectionRelations("{E6B9C8CA-4AC2-11D5-8151-0090276F4297}", "SystemMark2dParent")
            Set oMarkingInfo = oCollectionHelper.Item(1)
            Set oThicknessDir = oMarkingInfo.ThicknessDirection
            Set o2DThicknessDirVec = oTransMat.TransformVector(oThicknessDir)

            If sGeometryType = STRMFG_END_MARK Then
               ' Flange end connection mark, doesn't need bubble
               If o2DThicknessDirVec Is Nothing Then
                   GoTo NextLocationMark
               Else
                  If Abs(o2DThicknessDirVec.x) < 0.00001 And _
                     Abs(o2DThicknessDirVec.y) < 0.00001 And _
                     Abs(o2DThicknessDirVec.y) < 0.00001 Then
                      GoTo NextLocationMark
                  End If
               End If
            End If

            '*** Get the Start and End Points of the curve ***'
            oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

            oStartPos.x = dStartX
            oStartPos.y = dStartY
            oStartPos.z = dStartZ

            oEndPos.x = dEndX
            oEndPos.y = dEndY
            oEndPos.z = dEndZ

            '*** Get the parameters of the curve ***'
            oCurve.ParamRange dStartParam, dEndParam


            '*** Get the Normal Co-ordinates at the End Position ***'
            oCurve.Parameter oEndPos.x, oEndPos.y, oEndPos.z, dEndParam
            oCurve.Evaluate dEndParam, dTemp, dTemp, dTemp, dTanX, dTanY, dTanZ, dTan2x, dTan2Y, dTan2Z
            'oNormalVecAtStart.Set -dTanY, dTanx, dTanZ  '-ve dTanY places bubble in opposite direction

            '*** Get the Tangent Vector ***'
            oTangentVecAtStart.Set dTanX, dTanY, dTanZ

            '*** Set Z vector in order to get the Normal Vector given a Tangent Vector ***'
            oZVector.Set 0#, 0#, 1#

            '*** Get the Normal Vector ***'
            Set oNormalVecAtStart = oTangentVecAtStart.Cross(oZVector)

            dDotProd = oNormalVecAtStart.Dot(o2DThicknessDirVec)

            '... This dot product will be used to determine the Thickness direction (Left or Right)
            If dDotProd > 0 Then oNormalVecAtStart.Length = 1 Else oNormalVecAtStart.Length = -1

            '... This cross product will be used to determine the Thickness direction at the other points
            Set oCrossStart = oTangentVecAtStart.Cross(oNormalVecAtStart)

            '*** Get the length of the curve ***'
            dLength = oCurve.Length


            '*** Check for the type of the Bubble to be placed ***'

            If UCase(sBubbleType) = "CIRCLE" Then

                '=== Else if the location line length is < 5m, then place 1 bubble in the middle===

                If dLength < LOCATION_MARKING_LENGTH Then

                    '*** Determine Bubble Arc Diameter ***'
                    If dLength > LOCATION_MARKING_SEG_MAX Then
                        dNewPointOffset = Sqr((THICK_BUBBLE_RADIUS * THICK_BUBBLE_RADIUS) - (THICK_ARC_RADIUS * THICK_ARC_RADIUS))
                        dNewPointOffset = THICK_BUBBLE_RADIUS - dNewPointOffset
                        dDistOnCurve = THICK_ARC_RADIUS
                    ElseIf dLength > LOCATION_MARKING_SEG_MIN Then
                        dNewPointOffset = Sqr((THIN_BUBBLE_RADIUS * THIN_BUBBLE_RADIUS) - (THIN_ARC_RADIUS * THIN_ARC_RADIUS))
                        dNewPointOffset = THIN_BUBBLE_RADIUS - dNewPointOffset
                        dDistOnCurve = THIN_ARC_RADIUS
                    Else
                        'If the marking length < minimum i.e. < LOCATION_MARKING_SEG_MIN then skip the mark
                        GoTo NextLocationMark
                    End If


                    '*** Get the wirebody from complexstring ***'
                    Set oWB = oMfgRuleHelper.ComplexStringToWireBody(oCurve)

                    '*** Get the middle point of the curve ***'
                    Set oMidPos = oMfgRuleHelper.GetMiddlePoint(oWB)

                    '*** Get Normal co-ordinates to the curve ***'
                    oCurve.Parameter oMidPos.x, oMidPos.y, oMidPos.z, dMidPar
                    oCurve.Evaluate dMidPar, dTemp, dTemp, dTemp, dTanX, dTanY, dTanZ, dTan2x, dTan2Y, dTan2Z

                    '*** Get the Tangent vector for Midpoint ***'
                    oTangentVecAtMid.Set dTanX, dTanY, dTanZ

                    '*** Get the Cross Product of oCrossProd and this Tangent to get the Thickness Direction ***'
                    Set oCrossMid = oCrossStart.Cross(oTangentVecAtMid)

                    '*** Set the offset of through-point to 0.2232 ***'
                    oCrossMid.Length = dNewPointOffset

                    '*** Get the through-point of the arc of the bubble ***'
                    Set oNewPos = oMidPos.Offset(oCrossMid)


                    '*** Get the arc's left and right points (End Points of the arc) ***'
                    Set oArcPos1 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dDistOnCurve, oStartPos)
                    Set oArcPos2 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dDistOnCurve, oEndPos)


                    '*** Create the arc ***'
                    Set oArcs = New GeometryFactory
                    Set oComplexStrings = New GeometryFactory
                    Set oArc = oArcs.CreateBy3Points(Nothing, oArcPos1.x, oArcPos1.y, oArcPos1.z, oNewPos.x, oNewPos.y, oNewPos.z, oArcPos2.x, oArcPos2.y, oArcPos2.z)
                    dRadius = oArc.Radius



                    '*** Get the two parts of the Wired Bodies ***'
                    oSGOWireBodyUtilities.BoundWireByTwoPoints oWB, oStartPos, oArcPos1, oWB1
                    oSGOWireBodyUtilities.BoundWireByTwoPoints oWB, oArcPos2, oEndPos, oWB2

                    '*** Create the sets of the curves elements ***'
                    Set oCS1 = oMfgRuleHelper.WireBodyToComplexString(oWB1)
                    Set oCS2 = oMfgRuleHelper.WireBodyToComplexString(oWB2)

                    oCS1.GetCurves oCS1Elem
                    oCS2.GetCurves oCS2Elem

                    '*** Prepare the curves in order to create the "altered" location line ***'
                    Set oUpdatedGeomColl = New JObjectCollection

                    oUpdatedGeomColl.AddElements oCS1Elem
                    oUpdatedGeomColl.Add oArc
                    oUpdatedGeomColl.AddElements oCS2Elem

                    Set oUpdatedCS = oComplexStrings.CreateByCurves(Nothing, oUpdatedGeomColl)

                    '*** Add the curves in order ***'
                    oGeom2d.PutGeometry oUpdatedCS



                '=== If the location line length is > 5m, then place 2 bubbles one at each end ===
                Else

                    '*** Determine Bubble Arc Diameter ***'
                    dNewPointOffset = Sqr((THICK_BUBBLE_RADIUS * THICK_BUBBLE_RADIUS) - (THICK_ARC_RADIUS * THICK_ARC_RADIUS))
                    dNewPointOffset = THICK_BUBBLE_RADIUS - dNewPointOffset
                    dDistOnCurve = THICK_ARC_RADIUS


                    '*** Get the wirebody from complexstring ***'
                    Set oWB = oMfgRuleHelper.ComplexStringToWireBody(oCurve)


                    '*** Get the center point of the 1st Arc ***'
                    Dim oArc1Center As IJDPosition, oArc2Center As IJDPosition


                    Set oArc1Center = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oStartPos, OFFSET_DISTANCE, oEndPos)
                    Set oArc2Center = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oEndPos, OFFSET_DISTANCE, oStartPos)


                    '*** Get the arc's left and right points (End Points of the arc) ***'
                            '=== ARC 1 ==='
                    Dim oArc1Pos1 As IJDPosition, oArc1Pos2 As IJDPosition
                    Set oArc1Pos1 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oArc1Center, dDistOnCurve, oStartPos)
                    Set oArc1Pos2 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oArc1Center, dDistOnCurve, oEndPos)
                            '=== ARC 2 ==='
                    Dim oArc2Pos1 As IJDPosition, oArc2Pos2 As IJDPosition
                    Set oArc2Pos1 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oArc2Center, dDistOnCurve, oStartPos)
                    Set oArc2Pos2 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oArc2Center, dDistOnCurve, oEndPos)



                    '*** Get Normal co-ordinates to the curve for Arc 1 ***'
                    oCurve.Parameter oArc1Center.x, oArc1Center.y, oArc1Center.z, dMidPar
                    oCurve.Evaluate dMidPar, dTemp, dTemp, dTemp, dTanX, dTanY, dTanZ, dTan2x, dTan2Y, dTan2Z
    '                oNormalVec.Set dTanY, dTanx, dTanZ  '-ve dTanY places bubble in opposite direction

                    '*** Get the Tangent vector for Midpoint ***'
                    oTangentVecAtMid.Set dTanX, dTanY, dTanZ

                    '*** Get the Cross Product of oCrossProd and this Tangent to get the Thickness Direction ***'
                    Set oCrossMid = oCrossStart.Cross(oTangentVecAtMid)

                    '*** Set the offset of through-point to 0.2232 ***'
                    oCrossMid.Length = dNewPointOffset

                    '*** Get the through-point of the arc of the bubble ***'
                    Set oNewPos = oArc1Center.Offset(oCrossMid)

                    '*** Create the arc1 ***
                    Set oArcs1 = New GeometryFactory
                    Set oComplexStrings = New GeometryFactory
                    Set oArc1 = oArcs1.CreateBy3Points(Nothing, oArc1Pos1.x, oArc1Pos1.y, oArc1Pos1.z, oNewPos.x, oNewPos.y, oNewPos.z, oArc1Pos2.x, oArc1Pos2.y, oArc1Pos2.z)
                    dRadius = oArc1.Radius


                    '*** Set Temporary Variables to nothing ***'
                    Set oTangentVecAtMid = Nothing
                    Set oNormalVec = Nothing
                    Set oCrossMid = Nothing
                    '******************************************'


                    '*** Get Normal co-ordinates to the curve for Arc 2 ***'
                    oCurve.Parameter oArc2Center.x, oArc2Center.y, oArc2Center.z, dMidPar
                    oCurve.Evaluate dMidPar, dTemp, dTemp, dTemp, dTanX, dTanY, dTanZ, dTan2x, dTan2Y, dTan2Z

                    '*** Get the Tangent vector for Midpoint ***'
                    oTangentVecAtMid.Set dTanX, dTanY, dTanZ

                    '*** Get the Cross Product of oCrossProd and this Tangent to get the Thickness Direction ***'
                    Set oCrossMid = oCrossStart.Cross(oTangentVecAtMid)

                    '*** Set the offset of through-point to 0.2232 ***'
                    oCrossMid.Length = dNewPointOffset

                    '*** Get the through-point of the arc of the bubble ***'
                    Set oNewPos = oArc2Center.Offset(oCrossMid)

                    '*** Create the arc2 ***
                    Set oArcs2 = New GeometryFactory
                    Set oComplexStrings = New GeometryFactory
                    Set oArc2 = oArcs2.CreateBy3Points(Nothing, oArc2Pos1.x, oArc2Pos1.y, oArc2Pos1.z, oNewPos.x, oNewPos.y, oNewPos.z, oArc2Pos2.x, oArc2Pos2.y, oArc2Pos2.z)
                    dRadius = oArc2.Radius


                    '*** Get the three parts of the Wired Bodies ***'
                    oSGOWireBodyUtilities.BoundWireByTwoPoints oWB, oStartPos, oArc1Pos1, oWB1
                    oSGOWireBodyUtilities.BoundWireByTwoPoints oWB, oArc1Pos2, oArc2Pos1, oWB2
                    oSGOWireBodyUtilities.BoundWireByTwoPoints oWB, oArc2Pos2, oEndPos, oWB3


                    '*** Create the sets of the curves elements ***'
                    Set oCS1 = oMfgRuleHelper.WireBodyToComplexString(oWB1)
                    Set oCS2 = oMfgRuleHelper.WireBodyToComplexString(oWB2)
                    Set oCS3 = oMfgRuleHelper.WireBodyToComplexString(oWB3)

                    oCS1.GetCurves oCS1Elem
                    oCS2.GetCurves oCS2Elem
                    oCS3.GetCurves oCS3Elem


                     '*** Prepare the curves in order to create the "altered" location line ***'
                    Set oUpdatedGeomColl = New JObjectCollection

                    oUpdatedGeomColl.AddElements oCS1Elem
                    oUpdatedGeomColl.Add oArc1
                    oUpdatedGeomColl.AddElements oCS2Elem
                    oUpdatedGeomColl.Add oArc2
                    oUpdatedGeomColl.AddElements oCS3Elem


                    Set oUpdatedCS = oComplexStrings.CreateByCurves(Nothing, oUpdatedGeomColl)

                    '*** Add the curves in order ***'
                    oGeom2d.PutGeometry oUpdatedCS

                End If

            ElseIf UCase(sBubbleType) = "SQUARE" Then

                '*** Declaration of the variables ***'
                Dim oPos1 As New DPosition, oPos2 As New DPosition, oPos3 As New DPosition, oPos4 As New DPosition
                Dim oPos5 As New DPosition, oPos6 As New DPosition, oPosA As New DPosition, oPosB As New DPosition
                Dim oLine1 As New Line3d, oLine2 As New Line3d, oLine3 As New Line3d, oLine4 As New Line3d, oLine5 As New Line3d
                Dim oNewCS As IJComplexString
                Dim dThicknessLen As Double, dPointOffset As Double
                Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
                Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
                Dim oLandingCurve As IJWireBody
                Dim oThickDir As DVector, oStartDir As DVector, oEndDir As DVector, oXVector As DVector, oYVector As DVector
                Dim eTSide As GSCADSDPartSupport.ThicknessSide
                Dim oOrigin As New DPosition



                '*** CONSTRUCTION DETAILS OF THE BUBBLE ***'
                '         _________            .
                '        /3       4\          /|\
                '       /           \          |  Thickness
                '______/  .   .   .  \______  \|/
                '1    2   A  Mid   B  5     6  .
                '         <--->
                '      PointOffset
                '      <------------->
                ' FACE_PLATE_BUBBLE_LENGTH
                '******************************************'



                '*** Get the wirebody from complexstring ***'
                Set oWB = oMfgRuleHelper.ComplexStringToWireBody(oCurve)

                '*** Get the middle point of the curve ***'
                Set oMidPos = oMfgRuleHelper.GetMiddlePoint(oWB)

                '*** Get the ProfilePartSupport Object to get the Thickness direction & the Thickness direction side ***'
                Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
                Set oPartSupport.Part = Part
                Set oProfilePartSupport = oPartSupport

                '*** Get the Thickness Direction of the Profile/ER ***'
                oProfilePartSupport.GetProfilePartLandingCurve oLandingCurve, oThickDir, False, eTSide

                '*** Get the Web Thickness for the Height of the Thickness Bubble ***'
                oProfilePartSupport.GetThickness Web, dThicknessLen

                '*** Get the End Points and tangent vectors of the WireBody ***'
                oWB.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir

                Dim oNormal As New DVector
                Set oNormal = oEndDir.Cross(oZVector)

                '... This dot product will be used to determine the Thickness direction (Left or Right)
                dDotProd = oNormal.Dot(oThickDir)
                If dDotProd > 0 Then oNormal.Length = 1 Else oNormal.Length = -1

                '*** Set the Bubble Height to Web Thickness Length ***'
                oNormal.Length = dThicknessLen

                '*** Set the Bubble Length ***'
                '... If not mentioned through TSN_Constants, set it equal to Web Thickness
                If FACE_PLATE_BUBBLE_LENGTH <= 0# Then
                    dPointOffset = dThicknessLen
                Else
                    dPointOffset = FACE_PLATE_BUBBLE_LENGTH / 4
                End If

                '*** Get the 6 points of the bubble mark ***'
                '... Point A and Point B are temporary points which are used to create Point 3 and 4 respectively.
                Set oPos1 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset * 3, oStartPos)
                Set oPos2 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset * 2, oStartPos)
                Set oPosA = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset, oStartPos)
                Set oPosB = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset, oEndPos)
                Set oPos6 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset * 3, oEndPos)
                Set oPos5 = oMfgRuleHelper.GetPointAlongCurveAtDistance(oWB, oMidPos, dPointOffset * 2, oEndPos)
                Set oPos3 = oPosA.Offset(oNormal)
                Set oPos4 = oPosB.Offset(oNormal)

                '*** Construct the bubble using Lines ***'
                oLine1.DefineBy2Points oPos1.x, oPos1.y, oPos1.z, oPos2.x, oPos2.y, oPos2.z
                oLine2.DefineBy2Points oPos2.x, oPos2.y, oPos2.z, oPos3.x, oPos3.y, oPos3.z
                oLine3.DefineBy2Points oPos3.x, oPos3.y, oPos3.z, oPos4.x, oPos4.y, oPos4.z
                oLine4.DefineBy2Points oPos4.x, oPos4.y, oPos4.z, oPos5.x, oPos5.y, oPos5.z
                oLine5.DefineBy2Points oPos5.x, oPos5.y, oPos5.z, oPos6.x, oPos6.y, oPos6.z

                '*** Create the complexstring and add all the line segments in order ***'
                Set oNewCS = New ComplexString3d
                oNewCS.AddCurve oLine1, True
                oNewCS.AddCurve oLine2, True
                oNewCS.AddCurve oLine3, True
                oNewCS.AddCurve oLine4, True
                oNewCS.AddCurve oLine5, True


                '*** Finally, get the Geom3d object ***'
                '*** Add the curves in order ***'
                oGeom2d.PutGeometry oNewCS

            End If 'END Of IF condition to check they type of the bubble

        End If
NextLocationMark:
    Next

    Dim oGeom2dColFactory As New MfgGeomCol2dFactory

    'oGeomCol2d.AddGeometry oGeomCol2d.GetColcount + 1, oGeom2d
    Set CreateThicknessDirectionMark = oGeom2dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))


CleanUp:
    Set oGeom2d = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oSGOWireBodyUtilities = Nothing
    Set oCurve = Nothing
    Set oWB = Nothing
    Set oWB1 = Nothing
    Set oWB2 = Nothing
    Set oWB3 = Nothing
    Set oArcs = Nothing
    Set oArcs1 = Nothing
    Set oArcs2 = Nothing
    Set oComplexStrings = Nothing
    Set oArc = Nothing
    Set oArc1 = Nothing
    Set oArc2 = Nothing
    Set oCSColl = Nothing

    Set oMidPos = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oNewPos = Nothing
    Set oArcPos1 = Nothing
    Set oArcPos2 = Nothing

    Set oNormalVec = Nothing
    Set oNormalVecAtStart = Nothing
    Set oTangentVecAtStart = Nothing
    Set oZVector = Nothing
    Set oTangentVecAtMid = Nothing
    Set oCrossStart = Nothing
    Set oCrossMid = Nothing

    Set oCS1 = Nothing
    Set oCS2 = Nothing
    Set oCS3 = Nothing
    Set oUpdatedCS = Nothing
    Set oCS1Elem = Nothing
    Set oCS2Elem = Nothing
    Set oCS3Elem = Nothing
    Set oUpdatedGeomColl = Nothing



    Set oRelataionHelper = Nothing
    Set oCollectionHelper = Nothing
    Set oThicknessDir = Nothing
    Set o2DThicknessDirVec = Nothing
    Set oTransMat = Nothing


    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetApproximate3DTo2DTransMatrix(oMfgPart As Object, oFaceId As Long) As IJDT4x4
Const METHOD = "GetApproximate3DTo2DTransMatrix"
    On Error GoTo ErrorHandler
    Dim oMfgParent As IJMfgGeomParent
    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim oGeomCol3d As IJMfgGeomCol3d
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    Dim oChildren As IJDTargetObjectCol

    If TypeOf oMfgPart Is IJMfgProfilePart Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
        Dim oMfgProfilePart As IJMfgProfilePart
        Set oMfgProfilePart = oMfgPart
        Set oGeomCol2d = oMfgProfilePart.FinalGeometriesAfterProcess2D
    Else
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
        Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries
    End If

    Set oMfgParent = oMfgPart
    Set oChildren = oMfgParent.GetChildren

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked we can exit the function
        Exit Function
    End If

    ' Get the 3D and 2D UV Marks to compute the transformation matrix
    Dim oGeomColl As IJElements
    Dim oUV2DGeom As IJMfgGeom2d
    Dim oUV3DGeom As IJMfgGeom3D
    Dim oToolBox As New DGeomOpsToolBox
    Dim o2DGeom As IJMfgGeom2d
    Dim o3DGeom As IJMfgGeom3D
    Dim i As Long

    'Get the Geom2D collection which is After unfold
    Set oGeomColl = GetGeomColl(oGeomCol2d, STRMFG_UV_MARK)
    If Not oMfgPlateWrapper Is Nothing Then
        Set o2DGeom = oGeomColl.Item(1)
        Set oUV2DGeom = o2DGeom

        'Get the Geom3D collection which is Before unfold
        Set oGeomCol3d = oMfgPlateWrapper.GetGeometriesBeforeUnfold
        Set oGeomColl = GetGeomColl(oGeomCol3d, STRMFG_UV_MARK)

        Set o3DGeom = oGeomColl.Item(1)
        Set oUV3DGeom = o3DGeom

    ElseIf Not oMfgProfileWrapper Is Nothing Then
        For i = 1 To oGeomColl.Count
            Set o2DGeom = oGeomColl.Item(i)
            If o2DGeom.FaceId = oFaceId Then
                Set oUV2DGeom = o2DGeom
                Exit For
            End If
        Next

        'Get the Geom3D collection which is Before unfold
        Set oGeomCol3d = oChildren.Item(2)
        Set oGeomColl = GetGeomColl(oGeomCol3d, STRMFG_UV_MARK)


        For i = 1 To oGeomColl.Count
            Set o3DGeom = oGeomColl.Item(i)
            If o3DGeom.FaceId = oFaceId Then
                Set oUV3DGeom = o3DGeom
                Exit For
            End If
        Next
    End If

    If oUV2DGeom Is Nothing Then Exit Function
    If oUV3DGeom Is Nothing Then Exit Function

    ' Compute the transformation matrix
    Dim oMfgUtilMathGeom As IJMfgUtilMathGeom
    Dim oTransMat As IJDT4x4

    Set oMfgUtilMathGeom = New MfgMathGeom

    Set oTransMat = oMfgUtilMathGeom.ComputeFeatureTransf(oUV3DGeom, oUV2DGeom)

    Set GetApproximate3DTo2DTransMatrix = oTransMat

CleanUp:
    Set oGeomColl = Nothing
    Set oUV2DGeom = Nothing
    Set oUV3DGeom = Nothing
    Set oToolBox = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetGeomColl(oGeomCol As Object, eGeometryType As StrMfgGeometryType) As IJElements
Const METHOD = "GetGeomColl"
    On Error GoTo ErrorHandler

    Dim lGeomCount As Long
    Dim oMfgGeomCol3d As IJMfgGeomCol3d
    Dim oMfgGeomCol2d As IJMfgGeomCol2d
    Dim oOutputGeomColl As IJElements

    Set oOutputGeomColl = New JObjectCollection

    If TypeOf oGeomCol Is IJMfgGeomCol3d Then
        Set oMfgGeomCol3d = oGeomCol
        lGeomCount = oMfgGeomCol3d.GetCount
    ElseIf TypeOf oGeomCol Is IJMfgGeomCol2d Then
        Set oMfgGeomCol2d = oGeomCol
        lGeomCount = oMfgGeomCol2d.GetCount
    Else
        Exit Function
    End If

    Dim lIndex As Long

    For lIndex = 1 To lGeomCount
        Dim eThisGeometryType As StrMfgGeometryType
        If Not oMfgGeomCol3d Is Nothing Then
            Dim oMfgGeom3d As IJMfgGeom3D
            Set oMfgGeom3d = oMfgGeomCol3d.GetGeometry(lIndex)
            eThisGeometryType = oMfgGeom3d.GetGeometryType
            If eThisGeometryType = eGeometryType Then
                oOutputGeomColl.Add oMfgGeom3d
            End If
            Set oMfgGeom3d = Nothing
        Else
            Dim oMfgGeom2d As IJMfgGeom2d
            Set oMfgGeom2d = oMfgGeomCol2d.GetGeometry(lIndex)
            eThisGeometryType = oMfgGeom2d.GetGeometryType
            If eThisGeometryType = eGeometryType Then
                oOutputGeomColl.Add oMfgGeom2d
            End If
            Set oMfgGeom2d = Nothing
        End If
    Next lIndex

    Set GetGeomColl = oOutputGeomColl
CleanUp:
    Set oOutputGeomColl = Nothing
    Set oMfgGeomCol3d = Nothing
    Set oMfgGeomCol2d = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function EvaluateProfileCurvature(oProfileObj As Object) As ProfileCurvature

On Error GoTo CleanUp

    Dim strSectioName As String
    Dim dProfileStraightLength As Double, dBowStringDepth As Double, dMaxBowStringDepth As Double
    Dim bThicknessCentered As Boolean
    Dim oVector As IJDVector
    Dim oPartSupport As IJPartSupport
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oLandCurveWB As IJWireBody
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oQueryOutputValues() As Variant
    Dim strQuery As String
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Dim oSDProfilePart As New StructDetailObjects.ProfilePart

    Set oProfileWrapper.object = oProfileObj
    Set oSDProfilePart.object = oProfileObj

    If oSDProfilePart.ProfileType = sptEdgeReinforcement Then
        strSectioName = oSDProfilePart.SectionName

        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Set oPartSupport.Part = oProfileObj
        Set oProfilePartSupport = oPartSupport
        oProfilePartSupport.GetProfilePartLandingCurve oLandCurveWB, oVector, bThicknessCentered

        oLandCurveWB.GetEndPoints oStartPos, oEndPos
        dProfileStraightLength = oStartPos.DistPt(oEndPos)

        Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
        strQuery = "SELECT BowStringDepth From JUAMfgProfileBowStringLimit WHERE (SectionName = '" + strSectioName + "') AND (ProfileStraightLengthMin < " + CStr(dProfileStraightLength) + ") AND (ProfileStraightLengthMax >" + CStr(dProfileStraightLength) + ")"
        'MsgBox strQuery

        oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)

        If (UBound(oQueryOutputValues) >= LBound(oQueryOutputValues)) Then
            Dim dRadiusOfBow As Double
            dMaxBowStringDepth = oQueryOutputValues(0)
            'MsgBox "dMaxBowStringDepth = " & dMaxBowStringDepth
            dBowStringDepth = oProfileWrapper.GetDepthOfBowString(0, -1#, 0, Nothing, Nothing, CTX_INVALID, Nothing, Null, dRadiusOfBow)
            'MsgBox "dBowStringDepth = " & dBowStringDepth
            'MsgBox "dRadiusOfBow = " & dRadiusOfBow
            If dBowStringDepth < dMaxBowStringDepth Then
                EvaluateProfileCurvature = PROFILE_CURVATURE_Straight
            Else
                EvaluateProfileCurvature = PROFILE_CURVATURE_CurvedAlongWeb
            End If
        Else
            EvaluateProfileCurvature = oProfileWrapper.CurvatureType
        End If
    Else
        EvaluateProfileCurvature = oProfileWrapper.CurvatureType
    End If

CleanUp:
    Set oPartSupport = Nothing
    Set oProfilePartSupport = Nothing
    Set oLandCurveWB = Nothing
    Set oProfileWrapper = Nothing
    Set oVector = Nothing
    Set oMfgCatalogQueryHelper = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing

    Exit Function
End Function


Public Function IsProfileCurved(oThisPart As Object, oSDPlateWrapper As Object) As Boolean

    Dim oObj As Object, oConnectedObj As Object
    Dim oPort1 As IJPort
    Dim oNamed As IJNamedItem, oPlateName As IJNamedItem
    Dim pProfileCurvature As ProfileCurvature

    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport.Part = oThisPart

    Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection

    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol

    '*** Get the # of Physical Connections ***'
    Dim lNoOfObj As Long
    lNoOfObj = oConObjCol.Count

    '*** lNoOfObj > 1 means that there are other objects (other than base plate)
    '... having physical connection with the profile.
    '... So check which end of the profile is/are Free Ends, and place the marks accordingly.
    Dim i As Integer
    If lNoOfObj > 1 Then

        For i = 1 To lNoOfObj

            ' Get the object connected at the other port
            Set oPort1 = oOtherPortCol.Item(i)
            Set oObj = oPort1.Connectable

            ' This gives the name of the object connected at the other port
            Set oNamed = oObj

            ' This gives the name of the Base Plate
            Set oPlateName = oSDPlateWrapper.object

            'If the Connected Object is Base Plate, ignore it...
            '... and if not, then check if it a curved profile
            'If oNamed.Name <> oPlateName.Name Then
            If TypeOf oObj Is IJProfile Or TypeOf oObj Is IJProfilePart Then
                    If Not (oObj Is oSDPlateWrapper.object) Then

                        pProfileCurvature = EvaluateProfileCurvature(oObj)

                        If pProfileCurvature = PROFILE_CURVATURE_CurvedAlongFlange Or pProfileCurvature = PROFILE_CURVATURE_CurvedAlongWeb Then
                            IsProfileCurved = True
                            Exit Function
                        End If

                    End If
            End If
        Next

    Else
        'This means there are no other connections than the base plate
    End If
CleanUp:

    Set oPartSupport = Nothing
End Function

Public Function CreateRENDProfileMarks(Part As Object, UpSide As Long, oConnectionData As ConnectionData, oContourGeomCol3D As IJMfgGeomCol3d, oGeomCol3d As IJMfgGeomCol3d)
Const METHOD = "CreateRENDProfileMarks"

On Error GoTo ErrorHandler

    ' If the connected object is not profile part, exit the function
    If Not (TypeOf oConnectionData.ToConnectable Is IJProfilePart) Then
        Exit Function
    End If

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart
    Dim oNeutralSurface As IJSurfaceBody

    ' Initialize some helper classes
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = oConnectionData.ToConnectable
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oStiffenedPlate As Object
    Dim bIsSystem As Boolean
    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, False

    If Not oStiffenedPlate Is Nothing Then

        Dim oSys As IJSystem
        Dim oIJPartSupp As IJPartSupport
        Set oIJPartSupp = New PartSupport
        Set oIJPartSupp.Part = Part
        oIJPartSupp.IsSystemDerivedPart oSys, True

        If Not oSys Is Nothing Then
            'If given profile is ER and the plate part is not the stiffened plate, then exit the function
            'TR-CP-217868  SKDY-MFG-DEF-2012-0052_REmark wasPlaced
            If TypeOf oConnectionData.ToConnectable Is IJProfileER And Not oStiffenedPlate Is oSys Then
                Exit Function
            End If
        End If
    End If

    ' Get the plate's neutral surface
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
        Set oNeutralSurface = oMfgPlateCreation_AE.NeutralSurface
        Set oMfgPlateCreation_AE = Nothing
    Else
        Exit Function
    End If

    Dim oVector As IJDVector
    Dim bContourTee As Boolean
    Dim oTeeWire As IJWireBody
    Dim oCS As IJComplexString
    Dim oCSColl As IJElements
    Dim oMfgMathGeom As New MfgMathGeom
    Dim oOtherPort As IJPort

    ' Get the contour Tee geometry
    If TypeOf oConnectionData.ToConnectable Is IJProfileER Then
        bContourTee = oSDProfileWrapper.Connection_ContourTee(oConnectionData.AppConnection, SideA, oTeeWire, oVector)
        If Not bContourTee Then ' If ER's load point is its edge, it won't have a connection contour
            bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, SideA, oTeeWire, oVector)
        End If
    Else
        Exit Function
    End If

    'Convert the IJWireBody to ComplexStrings
    If bContourTee Then
        Set oCSColl = oMfgRuleHelper.WireBodyToComplexStrings(oTeeWire)
    End If
    If Not oCSColl Is Nothing Then
      If oCSColl.Count = 0 Then
          Set oCSColl = Nothing
      End If
    End If

    If (oCSColl Is Nothing) Then
      GoTo CleanUp
    End If

    ' For each complex string, convert it to line and arcs
    For Each oCS In oCSColl
        Dim oLineArcsCS As IJComplexString
        Dim oLineArcsColl As IJElements
        Dim oLineArcObj As Object, oLineArcObjPrev As Object, oLineArcObjNext As Object
        Dim lLineArcCnt As Long, indLineArc As Long

#If 0 Then ' Don't want Line-Arc conversion for now
        ' Convert this complex string to lines and arcs
        oMfgMathGeom.Convert3DCurvesToLineArc Nothing, oCS, 0.001, oLineArcsCS
#Else
        'Get the ComplexString with Line, Arc & Spline
        Set oLineArcsCS = oCS
#End If
        oLineArcsCS.GetCurves oLineArcsColl
        lLineArcCnt = oLineArcsColl.Count

        If lLineArcCnt > 10 Then
'            Dim oGH As New MfgGeomHelper
'            Dim sDGTFileName As String
'            sDGTFileName = Environ("TEMP")
'            If sDGTFileName = "" Or sDGTFileName = vbNullString Then
'                sDGTFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'            oGH.DumpGeometryToFile oCS, False, False, sDGTFileName, "OrigCS"
'            oGH.DumpGeometryToFile oLineArcsCS, False, False, sDGTFileName, "CSwith" & CStr(lLineArcCnt) & "LineArcs"
'            Set oGH = Nothing
        End If

        ' For each line/arc object, mark the ends
        For indLineArc = 1 To lLineArcCnt
            Dim lGeomCnt As Long, indGeom As Long
            Dim dMinDist As Double
            Dim oLineArcCurve As IJCurve, oClosestMfgGeom3dCrv As IJCurve
            Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
            Dim dDummy As Double, dTanX As Double, dTanY As Double, dTanZ As Double
            Dim oMarkPos As IJDPosition

            oLineArcsCS.GetCurve indLineArc, oLineArcObj
            Set oLineArcCurve = oLineArcObj

            ' Get the line/arc start and end points
            oLineArcCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
            If indLineArc > 1 Then
                Set oMarkPos = New DPosition
                oMarkPos.Set dStartX, dStartY, dStartZ
            Else
                GoTo NextMark
            End If

            '*** Check if the prev geometry is curve ***'
            oLineArcsCS.GetCurve indLineArc - 1, oLineArcObjPrev

            'If both arcs have same radius, SKIP
            Dim oCurrArc As Arc3d, oPrevArc As Arc3d

            'Put the mark at start and end of Spline as well
            If TypeOf oLineArcObjPrev Is IJArc Then
                Set oPrevArc = oLineArcObjPrev
            Else
                Set oPrevArc = Nothing
            End If

            If TypeOf oLineArcObj Is IJArc Then
                Set oCurrArc = oLineArcObj
            Else
                Set oCurrArc = Nothing
            End If

            'EXIT CONDITIONS
            ' If both radii are equal (Circular ER) or if Arc with large radius (>5m)
            If TypeOf oLineArcObjPrev Is IJBSplineCurve Then
                'Continue
            ElseIf Not oCurrArc Is Nothing And Not oPrevArc Is Nothing Then
                If Abs(oCurrArc.Radius - oPrevArc.Radius) < 0.001 Then
                    GoTo NextMark
                End If
            End If

            Dim dPar As Double, dStartParam As Double, dEndParam As Double, dPrevDotP As Double, dThisDotP As Double
            Dim oLineArcTan As New DVector, oMfgTan As New DVector
            Dim oProjPos As IJDPosition, oNewPos As IJDPosition, oNewPos2 As IJDPosition
            Dim oSurfaceNormal As IJDVector
            Dim oMatDir As IJDVector
            Dim o45Mark As IJDVector

            dMinDist = 1E+30
            lGeomCnt = oContourGeomCol3D.GetCount

            ' Get the tangent of line/arc at the mark position
            oLineArcCurve.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dPar
            oLineArcCurve.Evaluate dPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
            oLineArcTan.Set dTanX, dTanY, dTanZ

            Dim dT1 As Double, dT2 As Double, dT3 As Double
'            dim oLineArcCurvePrev as i
            Dim oLineArcCurvePrev As IJCurve
            Set oLineArcCurvePrev = oLineArcObjPrev
            oLineArcCurvePrev.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dPar
            oLineArcCurvePrev.Evaluate dPar, dDummy, dDummy, dDummy, dT1, dT2, dT3, dDummy, dDummy, dDummy

            Dim oLineArcPrevTan As New DVector
            oLineArcPrevTan.Set dT1, dT2, dT3
            'oLineArcTan.Set dTanX, dTanY, dTanZ

            'Check if tangent of both curves are parallel or not
            Dim oNorm As IJDVector
            Set oNorm = oLineArcPrevTan.Cross(oLineArcTan)

            Dim bKnuckled As Boolean
            'If parallel, its a roll (since next curve is tangent to previous)
            If Abs(oNorm.Length) < 0.00001 Then   'Parallel
                bKnuckled = False
            'If NOT parallel, its a knuckle
            Else
                bKnuckled = True
            End If

            For indGeom = 1 To lGeomCnt
                Dim oMfgGeom3d As IJMfgGeom3D
                Dim oMfgCurve As IJCurve
                Dim dDist As Double

                Set oMfgGeom3d = oContourGeomCol3D.GetGeometry(indGeom)
                Set oMfgCurve = oMfgGeom3d.GetGeometry
                oMfgCurve.DistanceBetween oMarkPos, dDist, dDummy, dDummy, dDummy, dDummy, dDummy, dDummy
                If Abs(dDist - dMinDist) < 0.001 Then
                    oMfgCurve.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dPar
                    oMfgCurve.Evaluate dPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
                    oMfgTan.Set dTanX, dTanY, dTanZ
                    dThisDotP = oMfgTan.Dot(oLineArcTan)
                    If (Abs(dThisDotP) > Abs(dPrevDotP)) Then
                        dMinDist = dDist
                        Set oClosestMfgGeom3dCrv = oMfgCurve
                        dPrevDotP = dThisDotP
                    End If
                ElseIf dDist < dMinDist Then
                    dMinDist = dDist
                    oMfgCurve.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dPar
                    oMfgCurve.Evaluate dPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
                    oMfgTan.Set dTanX, dTanY, dTanZ
                    dPrevDotP = oMfgTan.Dot(oLineArcTan)

                    Set oClosestMfgGeom3dCrv = oMfgCurve
                End If
            Next indGeom

            oClosestMfgGeom3dCrv.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dPar
            oClosestMfgGeom3dCrv.Evaluate dPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
            oMfgTan.Set dTanX, dTanY, dTanZ
            oMfgTan.Length = REND_FIT_MARKING_LINE_LENGTH

            '*** Get the base point of REND Mark ***'
            Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oMarkPos, oNeutralSurface, Nothing)

            oNeutralSurface.GetNormalFromPosition oProjPos, oSurfaceNormal
            Set oMatDir = oSurfaceNormal.Cross(oMfgTan)
            oMatDir.Length = REND_FIT_MARKING_LINE_LENGTH
            Set oNewPos = oProjPos.Offset(oMatDir)

            Dim oNorVec As IJDVector
            Set oNorVec = oMatDir

            Dim oLine As IJLine, oLine2 As IJLine
            Dim oLineCS As IJComplexString

            Set oLine = New Line3d
            Set oLine2 = New Line3d
            Set oLineCS = New ComplexString3d
'                Set oNewPos = New DPosition
            Set oNewPos2 = New DPosition

            '*** Construct the Mark ***'
'                oNewPos.x = oProjPos.x + oNorVec.x
'                oNewPos.y = oProjPos.y + oNorVec.y
'                oNewPos.z = oProjPos.z + oNorVec.z

            oLine.DefineBy2Points oProjPos.x, oProjPos.y, oProjPos.z, oNewPos.x, oNewPos.y, oNewPos.z
            oLineCS.AddCurve oLine, True

            Dim oCCurve As IJCurve
            Dim dCTanX As Double, dCTanY As Double, dCTanZ As Double
            Dim oCTanVec As New DVector

            If bKnuckled = False Then 'ONLY FOR RE's
            'If Previous is Arc, then create REND mark and orient towards previous arc
                If Not oPrevArc Is Nothing Or _
                    (TypeOf oLineArcObjPrev Is IJBSplineCurve And Not oPrevArc Is Nothing) Or _
                    (TypeOf oLineArcObjPrev Is IJBSplineCurve And TypeOf oLineArcObjPrev Is IJBSplineCurve) Then

                Set oCCurve = oLineArcObjPrev
                oCCurve.Parameter oProjPos.x, oProjPos.y, oProjPos.z, dPar
                oCCurve.Evaluate dPar, dDummy, dDummy, dDummy, dCTanX, dCTanY, dCTanZ, dDummy, dDummy, dDummy

                oCTanVec.Set dCTanX, dCTanY, dCTanZ

                'oMfgTan is tangent on entire curve, where as oCTanVec is tangent on current arc
                '... if both are in same direction, take oMfgTan, if in opposite direction, flip oMfgTan
                If oCTanVec.Dot(oMfgTan) < 0 Then
                    'Opposite
                    oMfgTan.Length = -1 * oMfgTan.Length
                End If

                Set o45Mark = oNorVec.Subtract(oMfgTan)
                o45Mark.Length = FittingMarkLength

                oNewPos2.x = oNewPos.x + o45Mark.x
                oNewPos2.y = oNewPos.y + o45Mark.y
                oNewPos2.z = oNewPos.z + o45Mark.z

                oLine2.DefineBy2Points oNewPos.x, oNewPos.y, oNewPos.z, oNewPos2.x, oNewPos2.y, oNewPos2.z
                oLineCS.AddCurve oLine2, True

            'If current is Arc and previous is not arc then create REND mark and orient towards current arc
            ElseIf Not oCurrArc Is Nothing Then ' => oPrevArc Is Nothing

                Set oCCurve = oLineArcObj
                oCCurve.Parameter oProjPos.x, oProjPos.y, oProjPos.z, dPar
                oCCurve.Evaluate dPar, dDummy, dDummy, dDummy, dCTanX, dCTanY, dCTanZ, dDummy, dDummy, dDummy

                oCTanVec.Set dCTanX, dCTanY, dCTanZ

                'oMfgTan is tangent on entire curve, where as oCTanVec is tangent on current arc
                '... if both are in same direction, take oMfgTan, if in opposite direction, flip oMfgTan
                If oCTanVec.Dot(oMfgTan) < 0 Then
                    'Opposite
                    oMfgTan.Length = -1 * oMfgTan.Length
                End If

                'oMfgTan.Length = -1#
                Set o45Mark = oNorVec.Add(oMfgTan)
                o45Mark.Length = 1 * FittingMarkLength

                oNewPos2.x = oNewPos.x + o45Mark.x
                oNewPos2.y = oNewPos.y + o45Mark.y
                oNewPos2.z = oNewPos.z + o45Mark.z

                oLine2.DefineBy2Points oNewPos.x, oNewPos.y, oNewPos.z, oNewPos2.x, oNewPos2.y, oNewPos2.z
                oLineCS.AddCurve oLine2, True

            Else ' => oCurrArc Is Nothing And oPrevArc Is Nothing
                GoTo NextMark
            End If
            End If

            '*** Create the Complex String ***'

            Dim oSystemMarkFactory As New GSCADMfgSystemMark.MfgSystemMarkFactory
            Dim oSystemMark As IJMfgSystemMark
            Dim oMarkingInfo As MarkingInfo
            Dim oGeom3d As IJMfgGeom3D
            Dim lGeomCount As Long
            Dim oMoniker As IMoniker

            Set oSystemMark = oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            '*** Set the marking side ***'
            oSystemMark.SetMarkingSide UpSide

            '*** QI for the MarkingInfo object on the SystemMark ***'
            Set oMarkingInfo = oSystemMark

            '*** Set the name and thickness for marking info ***'
            oMarkingInfo.Name = ""

            Dim oGeom3dFactory As New GSCADMfgGeometry.MfgGeom3dFactory

            Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            oGeom3d.PutGeometry oLineCS

            '*** Set the Type of the Mark ***'
            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
            oGeom3d.FaceId = UpSide
            oSystemMark.Set3dGeometry oGeom3d

            Set oMoniker = oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            lGeomCount = oGeomCol3d.GetCount
            oGeomCol3d.AddGeometry lGeomCount + 1, oGeom3d
            Set oLine = Nothing
            Set oLineCS = Nothing
NextMark:
            Set oMarkPos = Nothing
        Next ' each line/arc in line-arc coll
    Next ' each cs in cscoll

CleanUp:
    Set oSDProfileWrapper = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function GetDirAsFAUDIO(ByVal oVec As IJDVector, ByVal oPt As IJDPosition) As String
    Const METHOD = "GetDirAsFAUDIO"
    On Error GoTo ErrorHandler

    If oVec Is Nothing Then Exit Function
    If oVec.Length < 0.000001 Then Exit Function ' Null Vector has no direction!

    If Abs(oVec.x) > Abs(oVec.y) Then
        If Abs(oVec.x) > Abs(oVec.z) Then
            If oVec.x > 0 Then
                GetDirAsFAUDIO = "F"
            Else
                GetDirAsFAUDIO = "A"
            End If
        Else
            If oVec.z > 0 Then
                GetDirAsFAUDIO = "U"
            Else
                GetDirAsFAUDIO = "D"
            End If
        End If
    Else
        If Abs(oVec.z) > Abs(oVec.y) Then
            If oVec.z > 0 Then
                GetDirAsFAUDIO = "U"
            Else
                GetDirAsFAUDIO = "D"
            End If
        Else
            If oPt Is Nothing Then Exit Function

            If oVec.y > 0 And oPt.y >= 0 Or oVec.y < 0 And oPt.y <= 0 Then
                GetDirAsFAUDIO = "O"
            ElseIf oVec.y > 0 And oPt.y < 0 Or oVec.y < 0 And oPt.y > 0 Then
                GetDirAsFAUDIO = "I"
            End If
        End If
    End If

CleanUp:

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Sub RotateCSaboutPointNormal(oCS As IJComplexString, _
                                     oPivot As IJDPosition, _
                                     oNormal As IJDVector, _
                                     dAngle As Double)

    Const METHOD = "RotateCSaboutPointNormal"
    On Error GoTo ErrorHandler

    Dim oRotMat As New DT4x4
    oRotMat.LoadIdentity

    Dim oPivotPointAsVec As New DVector
    oPivotPointAsVec.Set oPivot.x, oPivot.y, oPivot.z
    oRotMat.Translate oPivotPointAsVec

    oRotMat.Rotate dAngle, oNormal

    oPivotPointAsVec.Set -oPivot.x, -oPivot.y, -oPivot.z
    oRotMat.Translate oPivotPointAsVec

    Dim oTransformCS As IJTransform
    Set oTransformCS = oCS

    oTransformCS.Transform oRotMat

    Set oPivotPointAsVec = Nothing
    Set oTransformCS = Nothing
    Set oRotMat = Nothing

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Private Function GetDeclivityAngleAtPosition(oInputPos As IJDPosition, _
                                             oConnectingPortSurface As IJSurfaceBody, _
                                             oProfilePart As IJProfilePart, _
                                             oPlatePart As IJPlatePart, _
                                             sMoldedSide As String, _
                                             oMfgRuleHelper As MfgRuleHelpers.Helper, _
                                             oTopologyLocate As IJTopologyLocate, _
                                             Optional oRetConObjNormal As IJDVector = Nothing) As Double
    Const METHOD = "GetDeclivityAngleAtPosition"
    On Error GoTo ErrorHandler

    ' Key for variable names in this Function.

    ' AB is the plate being manufactured, marked
    ' CD is the connected plate (oPlatePart) or profile (oProfilePart)
    ' oInputPos the point on CD at which the declivity angle is requested

    '                     D //
    '                      //
    '                     //
    '                    //
    '                   //
    ' A              C //                 B
    ' =====================================
    '

    ' oConnectingPortSurface is AB
    ' oProjPos is C -- oInputPos is projected on to AB to give C
    ' oThisPlateNormal is normal of AB at C
    ' oConnectedObjNormal is normal to CD at oInputPos
    ' oConnectedObjPlaneVec is a vector perpendicular to oConnectedObjNormal

    ' Get THIS plate's normal at the input frame position
    Dim oProjPos As IJDPosition
    Dim oThisPlateNormal As IJDVector
    Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oInputPos, oConnectingPortSurface, oThisPlateNormal)

    Dim oConnectedObjNormal As IJDVector
    Dim oConnectedObjPlaneVec As IJDVector

    If Not oProfilePart Is Nothing Then
        ' Get the profile cross section matrix at the input frame position
        Dim oCrossSectionMatrix As AutoMath.DT4x4
        Set oCrossSectionMatrix = oTopologyLocate.GetPenetratingCrossSectionMatrix(oProfilePart, oInputPos)

        Dim dMat() As Double
        dMat = oCrossSectionMatrix.GetMatrix

        Set oConnectedObjNormal = New DVector
        ' oConnectedObjNormal is the normal on WebRight surface
        oConnectedObjNormal.Set dMat(0), dMat(1), dMat(2)

        Set oConnectedObjPlaneVec = New DVector
        ' Get the plane on which the angle is measured
        oConnectedObjPlaneVec.Set dMat(8), dMat(9), dMat(10)

        Set oCrossSectionMatrix = Nothing
    ElseIf Not oPlatePart Is Nothing Then
        Dim oPartSupp As IJPartSupport
        Dim oPlatePartSupp As IJPlatePartSupport
        Set oPartSupp = New PlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = oPlatePart

        ' Get the normal of the connected surface at the input position
        Dim oMarkedSurfaceBody As IJSurfaceBody
        If sMoldedSide = "Base" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oMarkedSurfaceBody
        Else
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oMarkedSurfaceBody
        End If

        Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oInputPos, oMarkedSurfaceBody, Nothing)
        oMarkedSurfaceBody.GetNormalFromPosition oProjPos, oConnectedObjNormal
        Set oConnectedObjPlaneVec = oConnectedObjNormal.Cross(oThisPlateNormal)

        Set oMarkedSurfaceBody = Nothing
        Set oPartSupp = Nothing
        Set oPlatePartSupp = Nothing
    Else
        Exit Function
    End If

    If Not oRetConObjNormal Is Nothing Then
        Dim oTangVec As IJDVector
        Set oTangVec = oConnectedObjNormal.Cross(oThisPlateNormal)

        Set oRetConObjNormal = oTangVec.Cross(oThisPlateNormal)
        Set oTangVec = Nothing

        If oRetConObjNormal.Dot(oConnectedObjNormal) < 0 Then
            oRetConObjNormal.Length = -1
        Else
            oRetConObjNormal.Length = 1
        End If
    End If

    ' Calculate the declivity angle
    GetDeclivityAngleAtPosition = oConnectedObjNormal.Angle(oThisPlateNormal, oConnectedObjPlaneVec)

CleanUp:
    Set oCrossSectionMatrix = Nothing
    Set oMarkedSurfaceBody = Nothing
    Set oPartSupp = Nothing
    Set oPlatePartSupp = Nothing
    Set oThisPlateNormal = Nothing
    Set oConnectedObjNormal = Nothing
    Set oConnectedObjPlaneVec = Nothing
    Set oProjPos = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function CreateDeclivityMarks(Part As Object, oConnectionData As ConnectionData, oTeeCS As IJComplexString, _
                                     oMfgFrameSystem As IJDCoordinateSystem, oRefGeom3D As IJMfgGeom3D, _
                                     oGeomCol3d As IJMfgGeomCol3d, lGeomCount As Long, ByVal lMarkingSide As Long, _
                                     sMoldedSide As String, oRelatedPart As Object)
    Const METHOD = "CreateDeclivityMarks"

    Dim oGenericIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Dim iIndex As Long
    Dim nRefPlanes As Long
    Dim dPlateThickness As Double
    Dim oTeeCurve As IJCurve
    Dim oTeeWire As IJWireBody
    Dim dTeeLength As Double
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    Dim oToolBox As New DGeomOpsToolBox
    Dim oFramePosColl As IJElements
    Dim oMarkPosColl As IJElements
    Dim oFramePos As IJDPosition
    Dim oProfilePart As IJProfilePart
    Dim oPlatePart As IJPlatePart
    Dim oTopologyLocate As IJTopologyLocate
    Dim oConnectingPort As IJPort
    Dim oConnectingPortSurface As IJSurfaceBody
    Dim bThisPlateIsHull As Boolean
    Dim oSystemMarkFactory As New MfgSystemMarkFactory
    Dim oGeom3dFactory As New MfgGeom3dFactory
    Dim oSDObjectWrapper As Object
    Dim oFrameNormal As IJDVector
    Dim eSectionType As ProfileSectionType
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition, oOffsetPos As IJDPosition
    Dim oMarkTangentVec As New DVector
    Dim lConnMarkSide As Long
    Dim eProfileType As StructProfileType

    Set oTopologyLocate = New TopologyLocate

    eSectionType = UnknownProfile
    eProfileType = sptUnknown

    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport

    Dim oResMgr     As Object
    Set oResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oPartSupp As IJPartSupport
    Dim oPlatePartSupp As IJPlatePartSupport
    Set oPartSupp = New PlatePartSupport
    Set oPlatePartSupp = oPartSupp
    Set oPartSupp.Part = Part

    oPlatePartSupp.GetThickness dPlateThickness

    If oRelatedPart Is Nothing Then

        ' Check the connected part and initialize some helpers
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            ' If the connected part is a profile, get the Section type

            Set oProfilePart = oConnectionData.ToConnectable
            Set oSDObjectWrapper = New StructDetailObjects.ProfilePart
            Set oSDObjectWrapper.object = oConnectionData.ToConnectable
            Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport.Part = oProfilePart
            Set oProfilePartSupport = oPartSupport

            eProfileType = oSDObjectWrapper.ProfileType

            eSectionType = oProfilePartSupport.SectionType
            Set oPartSupport = Nothing
            Set oProfilePartSupport = Nothing
        Else
            Set oPlatePart = oConnectionData.ToConnectable
            Set oSDObjectWrapper = New StructDetailObjects.PlatePart
            Set oSDObjectWrapper.object = oConnectionData.ToConnectable
        End If

        ' Get the connecting face port ( for THIS plate part ) and its surface geometry
        Set oConnectingPort = oConnectionData.ConnectingPort
        Set oConnectingPortSurface = oConnectingPort.Geometry

        Dim oThisPlate As IJPlate
        Set oThisPlate = oConnectingPort.Connectable

        ' Check if this plate type is hull
        If oThisPlate.plateType = Hull Then
            bThisPlateIsHull = True
        Else
            bThisPlateIsHull = False
        End If

        Dim oStructPort As IJStructPort
        Set oStructPort = oConnectingPort

        Dim ConPortCID As eUSER_CTX_FLAGS
        ConPortCID = oStructPort.ContextID

            If (ConPortCID And CTX_BASE) > 0 Then
                lConnMarkSide = BaseSide
            ElseIf (ConPortCID And CTX_OFFSET) > 0 Then
                lConnMarkSide = OffsetSide
            ElseIf (ConPortCID And CTX_LATERAL) > 0 And _
                   TypeOf oProfilePart Is IJProfileER Then
                lConnMarkSide = -4

            Set oConnectingPort = oPlateWrapper.GetSurfacePort(lMarkingSide)
            Set oConnectingPortSurface = oConnectingPort.Geometry
            Else
                lConnMarkSide = UnDefinedSide
            End If

        Set oStructPort = Nothing

    Else ' APS Marks

        If TypeOf oRelatedPart Is IJProfilePart Then

            Set oProfilePart = oRelatedPart
            Set oSDObjectWrapper = New StructDetailObjects.ProfilePart
            Set oSDObjectWrapper.object = oRelatedPart
            Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport.Part = oProfilePart
            Set oProfilePartSupport = oPartSupport

            eSectionType = oProfilePartSupport.SectionType

            Dim eTSide As GSCADSDPartSupport.ThicknessSide
            eTSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint

            If eTSide = SideA Then
                sMoldedSide = "Base"
            ElseIf eTSide = SideB Then
                sMoldedSide = "Offset"
            End If

            Set oPartSupport = Nothing
            Set oProfilePartSupport = Nothing

        ElseIf TypeOf oRelatedPart Is IJPlatePart Then
            Set oPlatePart = oRelatedPart

            Set oSDObjectWrapper = New StructDetailObjects.PlatePart
            Set oSDObjectWrapper.object = oPlatePart
            sMoldedSide = oSDObjectWrapper.MoldedSide

        End If

        Dim oThisPort As IJPort
        GetMarkingSideAndPort oRefGeom3D, lConnMarkSide, oThisPort

        Dim oThisPlatePart As New StructDetailObjects.PlatePart
        Set oThisPlatePart.object = Part

        Set oConnectingPortSurface = oThisPort.Geometry

        ' Check if this plate type is hull
        If oThisPlatePart.plateType = Hull Then
            bThisPlateIsHull = True
        Else
            bThisPlateIsHull = False
        End If

    End If

    Set oFramePosColl = New JObjectCollection
    Set oMarkPosColl = New JObjectCollection

    ' Convert the wirebody to complex string and get its length
    Set oTeeCurve = oTeeCS
    dTeeLength = oTeeCurve.Length
    Set oTeeWire = oMfgRuleHelper.ComplexStringToWireBody(oTeeCS)

    If dTeeLength < DECLIVITY_MARK_MINIMUM_TEE_LENGTH Then
        Exit Function
    End If

    Dim ConnPartIsTwisted As Boolean
    ConnPartIsTwisted = False

    ' If the input plate is shell and the connected part is a profile, get the frame intersection points
    If ((bThisPlateIsHull = True) And (Not oProfilePart Is Nothing) And (eProfileType = sptLongitudinal) And dTeeLength > (DECLIVITY_CONNECTED_LENGTH)) Then
        Dim oProfileVec As IJDVector
        Dim eAxis As AxisType
        Dim oPointUtils As IJSGOPointsGraphUtilities
        Dim oGridsHelper As New SP3dGSMiddleHelper.SPGMiddleHelper
        Dim oRefPlaneCollection As IJElements

        ConnPartIsTwisted = True

        ' Check how the profile is oriented and get the axis type
        oTeeWire.GetEndPoints oStartPos, oEndPos
        Set oProfileVec = oStartPos.Subtract(oEndPos)
        oProfileVec.Length = 1#
        If (Abs(oProfileVec.x) > Abs(oProfileVec.y)) And (Abs(oProfileVec.x) > Abs(oProfileVec.z)) Then
            eAxis = x
        ElseIf (Abs(oProfileVec.y) > Abs(oProfileVec.x)) And (Abs(oProfileVec.y) > Abs(oProfileVec.z)) Then
            eAxis = y
        Else
            eAxis = z
        End If

        ' Need the Intersection Helper
        Set oGenericIntersector = New IMSModelGeomOps.DGeomOpsIntersect

        ' point utilities
        Set oPointUtils = New SGOPointsGraphUtilities

        ' Get all the frame planes within the profile range
        oGridsHelper.EnumPlanesInRange oResMgr, oTeeCS, oMfgFrameSystem, "", x, True, Primary, oRefPlaneCollection
        Set oGridsHelper = Nothing

        nRefPlanes = oRefPlaneCollection.Count

        ' loop thru each Frame, intersect it with the location mark
        For iIndex = 1 To nRefPlanes
            Dim oRefPlane As IJPlane
            Dim oIntersectedUnknown As IUnknown

            Set oRefPlane = oRefPlaneCollection.Item(iIndex)
            If Not oRefPlane Is Nothing Then
                Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double

                ' Intersect the frame plane with the location marking geometry
                oRefPlane.GetNormal dNormalX, dNormalY, dNormalZ
                Set oFrameNormal = New DVector
                oFrameNormal.Set dNormalX, dNormalY, dNormalZ

                On Error Resume Next
                oGenericIntersector.PlaceIntersectionObject Nothing, _
                                                            oTeeCS, oRefPlane, _
                                                            Nothing, _
                                                            oIntersectedUnknown
                On Error GoTo ErrorHandler

                If Not oIntersectedUnknown Is Nothing Then
                    ' Retrieve the IJPosition (Point) from the IntersectionObject
                    If TypeOf oIntersectedUnknown Is IJPointsGraphBody Then
                        Dim oPointsGraphBody As IJPointsGraphBody
                        Dim oPointsColl As Collection
                        Dim oIntersectionPos As IJDPosition

                        ' Collect the intersection points in an array
                        Set oPointsGraphBody = oIntersectedUnknown
                        Set oPointsColl = oPointUtils.GetPositionsFromPointsGraph(oPointsGraphBody)
                        For Each oIntersectionPos In oPointsColl
                            oFramePosColl.Add oIntersectionPos
                            oMarkPosColl.Add oIntersectionPos
                        Next
                    End If ' TypeOf oIntersectedUnknown Is IJPointsGraphBody
                End If ' Not oIntersectedUnknown Is Nothing
            End If ' Not oRefPlane Is Nothing
        Next
    Else ' Not a profile connected to Hull

        oTeeWire.GetEndPoints oStartPos, oEndPos

        If (Not oProfilePart Is Nothing) Then
            ' Update the start and end positions with length points.
            On Error GoTo Calculate_Angles
            Dim oLengthHelper As IJMfgLengthHelper
            Dim oBaseGeom3d As IJMfgGeom3D, oOffsetGeom3d As IJMfgGeom3D
            Dim oBaseGeomCS As IJComplexString, oOffsetGeomCS As IJComplexString
            Dim oBaseGeomCurve As IJCurve, oOffsetGeomCurve As IJCurve
            Dim dLocX1 As Double, dLocY1 As Double, dLocZ1 As Double, dLocX2 As Double, dLocY2 As Double, dLocZ2 As Double, dMinDist1 As Double, dMinDist2 As Double, dDummy As Double
            Dim dMinDistBetPlateAndLengthPt As Double
            Dim oVertices As Collection
            Dim oLengthPointPos As New DPosition, oClosestPos As IJDPosition

            Set oLengthHelper = CreateObject("MfgLengthHelper.CMfgLengthHelper")
            oLengthHelper.PROFILE = oProfilePart

            'Calculate length marks at the base cap
            Set oBaseGeom3d = oLengthHelper.GetLengthGeom(STRMFG_PROFILE_LENGTH, JXSEC_WEB_LEFT, True, False, oVertices)
            Set oBaseGeomCS = oBaseGeom3d.GetGeometry
            Set oBaseGeomCurve = oBaseGeomCS
            Set oVertices = Nothing

            'Calculate length marks at the offset cap
            Set oOffsetGeom3d = oLengthHelper.GetLengthGeom(STRMFG_PROFILE_LENGTH, JXSEC_WEB_LEFT, False, True, oVertices)
            Set oOffsetGeomCS = oOffsetGeom3d.GetGeometry
            Set oOffsetGeomCurve = oOffsetGeomCS
            Set oVertices = Nothing

            ' Check the closest length points for start position and update it
            oBaseGeomCurve.DistanceBetween oStartPos, dMinDist1, dLocX1, dLocY1, dLocZ1, dDummy, dDummy, dDummy
            oOffsetGeomCurve.DistanceBetween oStartPos, dMinDist2, dLocX2, dLocY2, dLocZ2, dDummy, dDummy, dDummy

            Dim oSurfaceBodyWithOutFeatures As IJDModelBody
            If sMoldedSide = "Base" Then
                oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oSurfaceBodyWithOutFeatures
            Else
                oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oSurfaceBodyWithOutFeatures
            End If

            If dMinDist1 < dMinDist2 Then
                GetDeclivityPoint oBaseGeomCurve, oStartPos, oStartPos.Subtract(oEndPos), oSurfaceBodyWithOutFeatures
            Else
                GetDeclivityPoint oOffsetGeomCurve, oStartPos, oStartPos.Subtract(oEndPos), oSurfaceBodyWithOutFeatures
            End If

            ' Check the closest length points for end position and update it
            oBaseGeomCurve.DistanceBetween oEndPos, dMinDist1, dLocX1, dLocY1, dLocZ1, dDummy, dDummy, dDummy
            oOffsetGeomCurve.DistanceBetween oEndPos, dMinDist2, dLocX2, dLocY2, dLocZ2, dDummy, dDummy, dDummy

            If dMinDist1 < dMinDist2 Then
                GetDeclivityPoint oBaseGeomCurve, oEndPos, oEndPos.Subtract(oStartPos), oSurfaceBodyWithOutFeatures
            Else
                GetDeclivityPoint oOffsetGeomCurve, oEndPos, oEndPos.Subtract(oStartPos), oSurfaceBodyWithOutFeatures
            End If
        End If

Calculate_Angles:
        On Error GoTo ErrorHandler
        Dim oGeom3dObj As IJDObject
        Set oGeom3dObj = oBaseGeom3d
        If Not oGeom3dObj Is Nothing Then
            oGeom3dObj.Remove
        End If

        Set oGeom3dObj = oOffsetGeom3d
        If Not oGeom3dObj Is Nothing Then
            oGeom3dObj.Remove
        End If

        Dim StartMountAngle As Double
        StartMountAngle = GetDeclivityAngleAtPosition(oStartPos, oConnectingPortSurface, _
                                                      oProfilePart, oPlatePart, sMoldedSide, _
                                                      oMfgRuleHelper, oTopologyLocate)

        Dim EndMountAngle As Double
        EndMountAngle = GetDeclivityAngleAtPosition(oEndPos, oConnectingPortSurface, _
                                                    oProfilePart, oPlatePart, sMoldedSide, _
                                                    oMfgRuleHelper, oTopologyLocate)

        If Abs(StartMountAngle - EndMountAngle) > (DECLIVITY_TWIST_TOLERANCE * PI / 180#) Then
            ' Mounting angle at start and end are not the same
            ConnPartIsTwisted = True

            '*** Handle case of one angle = 90 and other Obtuse ***'
            Dim bPutAcuteAngle As Boolean
            Dim dActualStartAngle As Double, dActualEndAngle As Double

            dActualStartAngle = StartMountAngle
            dActualEndAngle = EndMountAngle

            If StartMountAngle > PI Then
                dActualStartAngle = 2 * PI - StartMountAngle
            End If
            If EndMountAngle > PI Then
                dActualEndAngle = 2 * PI - EndMountAngle
            End If

            If (Abs(dActualStartAngle - PI / 2#) < 0.001 And Abs(dActualEndAngle) > PI / 2#) Or (Abs(dActualEndAngle - PI / 2#) < 0.001 And Abs(dActualStartAngle) > PI / 2#) Then
                bPutAcuteAngle = True
            End If
            '******************************************************'

            If dTeeLength > DECLIVITY_CONNECTED_LENGTH Then
                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oStartPos, DECLIVITY_CONN_END_OFFSET_ABS, oEndPos, oOffsetPos
                oFramePosColl.Add oStartPos
                oMarkPosColl.Add oOffsetPos
                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oEndPos, DECLIVITY_CONN_END_OFFSET_ABS, oStartPos, oOffsetPos
                oFramePosColl.Add oEndPos
                oMarkPosColl.Add oOffsetPos
            Else
                Dim DistToMeasure As Double
                DistToMeasure = dTeeLength * DECLIVITY_CONN_END_OFFSET_REL
                If DistToMeasure < DECLIVITY_MARKING_LINE_LENGTH Then
                    DistToMeasure = Abs(Sin(StartMountAngle / 2#)) * DECLIVITY_MARKING_LINE_LENGTH + DECLIVITY_OFFSET_FROM_FRAME
                End If

                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oStartPos, DistToMeasure, oEndPos, oOffsetPos
                oFramePosColl.Add oStartPos
                oMarkPosColl.Add oOffsetPos

                DistToMeasure = dTeeLength * DECLIVITY_CONN_END_OFFSET_REL
                If DistToMeasure < DECLIVITY_MARKING_LINE_LENGTH Then
                    DistToMeasure = Abs(Sin(EndMountAngle / 2#)) * DECLIVITY_MARKING_LINE_LENGTH + DECLIVITY_OFFSET_FROM_FRAME
                End If
                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oEndPos, DistToMeasure, oStartPos, oOffsetPos
                oFramePosColl.Add oEndPos
                oMarkPosColl.Add oOffsetPos
            End If
        Else
            ' Mounting angle at start and end are the same
            ConnPartIsTwisted = False

            If dTeeLength > DECLIVITY_CONNECTED_LENGTH Then
                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oStartPos, DECLIVITY_CONN_END_OFFSET_ABS, oEndPos, oOffsetPos
                oFramePosColl.Add oStartPos
                oMarkPosColl.Add oOffsetPos
                oToolBox.GetPointAlongCurveAtDistance oTeeWire, oEndPos, DECLIVITY_CONN_END_OFFSET_ABS, oStartPos, oOffsetPos
                oFramePosColl.Add oEndPos
                oMarkPosColl.Add oOffsetPos
            Else
                ' If the connected length is less than DECLIVITY_CONNECTED_LENGTH meters,
                ' calculate declivity angle at the mid point of location mark
                Dim oMidPos As IJDPosition
                oToolBox.GetMiddlePointOfCompositeCurve oTeeWire, oMidPos
                oFramePosColl.Add oMidPos
                oMarkPosColl.Add oMidPos
            End If ' end if length is greater thanDECLIVITY_CONNECTED_LENGTH
        End If ' end if mounting angles at ends are different
    End If ' End if profile connected to hull


    ' Get the MfgPlate, Neutral Surface and 3D Contour Collection
    Dim oOutrCntGeom3dColl As IJMfgGeomCol3d
    Dim oMfgPart As IJMfgPlatePart
    Dim oNeutralSurface As IJSurfaceBody
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
        Set oOutrCntGeom3dColl = oMfgPlateCreation_AE.GeometriesBeforeUnfold
        Set oNeutralSurface = oMfgPlateCreation_AE.NeutralSurface
        Set oMfgPlateCreation_AE = Nothing
    Else
        'Exit Function?
    End If
    Set oPlateWrapper = Nothing

    ' Calculate declivity angle at each position
    For Each oFramePos In oFramePosColl
        Dim dDeclivityAngle As Double
        Dim oThkDirVecAlongPlateSurf As New DVector
        dDeclivityAngle = GetDeclivityAngleAtPosition(oFramePos, oConnectingPortSurface, _
                                                      oProfilePart, oPlatePart, sMoldedSide, _
                                                      oMfgRuleHelper, oTopologyLocate, _
                                                      oThkDirVecAlongPlateSurf)

        ' If the angle is exactly 90 degrees ( 0.5 degree tolerance ),
        ' There is no need to create the mark. In this case, just goto the next position
        If Abs(Cos(dDeclivityAngle)) < Sin(DECLIVITY_SHOW_TOLERANCE * PI / 180#) And Not ConnPartIsTwisted Then
            GoTo NextFramePos
        End If

        If dDeclivityAngle > PI Then
            dDeclivityAngle = 2 * PI - dDeclivityAngle
        End If

        Dim bDeclivityIsOnBase As Boolean

        '1. For non-angles, if one angle is 90 and other is obtuse, flip the mark to other side with acute angle
        If bPutAcuteAngle And Not (eSectionType = Angle Or eSectionType = Bulb_Angle Or eSectionType = Bulb_Tee Or _
           eSectionType = Bulb_Type Or eSectionType = Fab_Angle) Then
                bDeclivityIsOnBase = False
                If (dDeclivityAngle - PI / 2#) > 0.001 Then
                    dDeclivityAngle = PI - dDeclivityAngle
                End If
        ElseIf eSectionType = Angle Or eSectionType = Bulb_Angle Or eSectionType = Bulb_Tee Or _
           eSectionType = Bulb_Type Or eSectionType = Fab_Angle Or ConnPartIsTwisted Then
            bDeclivityIsOnBase = True
        Else
        'Ref: TSN  Customization - mounting angle marked on the obtuse angle (Section 6.16)
        'Ref: SKDY Customization - mounting angle marked on the Acute side (Section 9.2.2)
            If dDeclivityAngle < PI / 2# And DECLIVITY_MARK_OBTUSE_ANGLE Or _
               dDeclivityAngle > PI / 2 And Not DECLIVITY_MARK_OBTUSE_ANGLE Then
                dDeclivityAngle = PI - dDeclivityAngle
                bDeclivityIsOnBase = False
            Else
                bDeclivityIsOnBase = True
            End If
        End If

        Dim oMarkPos As IJDPosition
        Set oMarkPos = oMarkPosColl.Item(oFramePosColl.GetIndex(oFramePos))

        Dim dFramePar As Double
        oTeeCurve.Parameter oMarkPos.x, oMarkPos.y, oMarkPos.z, dFramePar

        Dim dTanX As Double, dTanY As Double, dTanZ As Double
        oTeeCurve.Evaluate dFramePar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy

        Dim oDeclStartPos As IJDPosition
        Dim oDeclPosOnCS As IJDPosition
        Dim oDeclRootPos As IJDPosition

        Set oDeclPosOnCS = oMarkPos
        oMarkTangentVec.Set dTanX, dTanY, dTanZ

        If Not oFrameNormal Is Nothing Then
            If oMarkTangentVec.Dot(oFrameNormal) < 0# Then
                oMarkTangentVec.Set -1# * dTanX, -1# * dTanY, -1# * dTanZ
            End If
        End If
        oMarkTangentVec.Length = DECLIVITY_MARKING_LINE_LENGTH

        If Not oProfilePart Is Nothing Then
            If ((sMoldedSide = "Base" And bDeclivityIsOnBase = False) Or _
               (sMoldedSide = "Offset" And bDeclivityIsOnBase = True)) Then
                oThkDirVecAlongPlateSurf.Length = DECLIVITY_OFFSET_FROM_LOCATION
            Else
                oThkDirVecAlongPlateSurf.Length = -DECLIVITY_OFFSET_FROM_LOCATION
            End If
        Else
            If bDeclivityIsOnBase Then
                oThkDirVecAlongPlateSurf.Length = -DECLIVITY_OFFSET_FROM_LOCATION
            Else
                oThkDirVecAlongPlateSurf.Length = DECLIVITY_OFFSET_FROM_LOCATION
            End If
        End If

        Set oDeclRootPos = oDeclPosOnCS.Offset(oThkDirVecAlongPlateSurf)
        If Not oFrameNormal Is Nothing Then
            Set oDeclRootPos = oDeclRootPos.Offset(oMarkTangentVec)
        End If
        Set oDeclStartPos = oDeclRootPos.Offset(oMarkTangentVec)


        '*** Check if Declivity Mark is intersecting outer contour or outside outer contour ***'

        Dim oTCurve As IJCurve
        Dim dTemp As Double, dMinDist As Double
        Dim dPtX As Double, dPtY As Double, dPtZ As Double
        Dim dPtX2 As Double, dPtY2 As Double, dPtZ2 As Double
        Dim objGeom As IJMfgGeom3D
        Dim j As Integer
        Dim oRootContourVec As New DVector
        Dim oPosOnContour As New DPosition

        'Loop through all outer and inner contours to get the distance from root point
        '.. If the distance is less than (declivity mark length + offset), we will flip the direction and take complement angle.
        For j = 1 To oOutrCntGeom3dColl.GetCount

            Set objGeom = oOutrCntGeom3dColl.GetGeometry(j)

            'Check if the contour is outer / inner
            If objGeom.GetGeometryType = STRMFG_OUTER_CONTOUR Then
                Set oTCurve = objGeom.GetGeometry
                Dim dTanPar As Double
                Dim bFlipMark As Boolean
                Dim oOuterTan As New DVector

                bFlipMark = False

                'Get the distance of root of declivity mark from the contour
                oTCurve.DistanceBetween oDeclRootPos, dMinDist, dPtX, dPtY, dPtZ, dPtX2, dPtY2, dPtZ2

                Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double, d5 As Double, d6 As Double
                oTCurve.EndPoints d1, d2, d3, d4, d5, d6

                oTCurve.Parameter dPtX, dPtY, dPtZ, dTanPar

                oTCurve.Evaluate dTanPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
                oOuterTan.Set dTanX, dTanY, dTanZ

                'If distance is less than (declivity mark length + offset) then flip the direction and take complement angle.
                If dMinDist < DECLIVITY_MARKING_LINE_LENGTH + DECLIVITY_OFFSET_FROM_LOCATION Then

                    oPosOnContour.x = dPtX
                    oPosOnContour.y = dPtY
                    oPosOnContour.z = dPtZ

                    'oRootContourVec is a vector from root to the outer contour
                    Set oRootContourVec = oPosOnContour.Subtract(oDeclRootPos)

                    If sMoldedSide = "" Then
                        Dim oCrossP As IJDVector
                        Set oCrossP = oRootContourVec.Cross(oOuterTan)

                        ' Get THIS plate's normal at the input frame position
                        Dim oProjPos As IJDPosition
                        Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oMarkPos, oConnectingPortSurface, Nothing)

                        Dim oThisPlateNormal As IJDVector
                        oConnectingPortSurface.GetNormalFromPosition oProjPos, oThisPlateNormal

                        If (oCrossP.Dot(oThisPlateNormal) < 0) Then
                            bFlipMark = True
                        End If

                        Set oProjPos = Nothing
                        Set oThisPlateNormal = Nothing
                    Else
                        If (oRootContourVec.Dot(oThkDirVecAlongPlateSurf) > 0) Then
                            bFlipMark = True
                        End If
                    End If

                    ' Dot product > 0 means that declivity mark and contour are on same side
                    '.. so flip the position
                    If bFlipMark Then

                        oThkDirVecAlongPlateSurf.Length = oThkDirVecAlongPlateSurf.Length * -1#
                        dDeclivityAngle = PI - dDeclivityAngle
                        bDeclivityIsOnBase = Not bDeclivityIsOnBase

                        'Reset the Root Position and Start Position
                        Set oDeclRootPos = oDeclPosOnCS.Offset(oThkDirVecAlongPlateSurf)
                        If Not oFrameNormal Is Nothing Then
                            Set oDeclRootPos = oDeclRootPos.Offset(oMarkTangentVec)
                        End If
                        Set oDeclStartPos = oDeclRootPos.Offset(oMarkTangentVec)

                        'Quit the loop and continue marking
                        GoTo Marking
                    End If
                End If
            End If
        Next

Marking:

        Dim oLocalZVec As IJDVector
        Set oLocalZVec = oMarkTangentVec.Cross(oThkDirVecAlongPlateSurf)

        Dim oRotMat As New DT4x4
        oRotMat.LoadIdentity
        oRotMat.Rotate dDeclivityAngle, oLocalZVec

        Dim oDeclVec As IJDVector
        Set oDeclVec = oRotMat.TransformVector(oMarkTangentVec)
        oDeclVec.Length = DECLIVITY_MARKING_LINE_LENGTH

        Dim oDeclEndPos As IJDPosition
        Set oDeclEndPos = oDeclRootPos.Offset(oDeclVec)

        ' Construct a two segment geometry as declivity mark

        Dim oLines As ILines3d
        Set oLines = New GeometryFactory

        Dim oLine1 As IJLine
        Set oLine1 = oLines.CreateBy2Points(Nothing, oDeclStartPos.x, oDeclStartPos.y, oDeclStartPos.z, _
                                            oDeclRootPos.x, oDeclRootPos.y, oDeclRootPos.z)

        Dim oLine2 As IJLine
        Set oLine2 = oLines.CreateBy2Points(Nothing, oDeclRootPos.x, oDeclRootPos.y, oDeclRootPos.z, _
                                            oDeclEndPos.x, oDeclEndPos.y, oDeclEndPos.z)

        Dim oCurveColl As IJElements
        Set oCurveColl = New JObjectCollection
        oCurveColl.Add oLine1
        oCurveColl.Add oLine2

        Dim oComplexStrings As IComplexStrings3d
        Set oComplexStrings = oLines

        Dim oLineCS As IJComplexString
        Set oLineCS = oComplexStrings.CreateByCurves(Nothing, oCurveColl)

        If DECLIVITY_POINTS_TO_LOCATION Then
            ' Rotate the Declivity mark to point towards location mark
            RotateCSaboutPointNormal oLineCS, oDeclRootPos, oLocalZVec, (PI - dDeclivityAngle) / 2#

            If Not oProfilePart Is Nothing And bThisPlateIsHull = True Then
                Dim oMoveDclMrkVec As New DVector
                Set oMoveDclMrkVec = oMarkTangentVec.Clone
                oMoveDclMrkVec.Length = DECLIVITY_MARKING_LINE_LENGTH * (-1 + Sin(dDeclivityAngle / 2#)) _
                                        + DECLIVITY_OFFSET_FROM_FRAME

                oRotMat.LoadIdentity
                oRotMat.Translate oMoveDclMrkVec

                Dim oTransformDclMrk As IJTransform
                Set oTransformDclMrk = oLineCS
                oTransformDclMrk.Transform oRotMat
            End If
        End If

        'Create a SystemMark object to store additional information

        Dim oSystemMark As MfgSystemMark
        Set oSystemMark = oSystemMarkFactory.Create(oResMgr)

        'QI for the MarkingInfo object on the SystemMark
        Dim oMarkingInfo As MarkingInfo
        Set oMarkingInfo = oSystemMark

'       Commented because this is now controlled by annotation
'       (so this is what SHOULD be shown by Annotation)
'
'        oMarkingInfo.Name = ""
'        oMarkingInfo.Name = GetDirAsFAUDIO(oThkDirVecAlongPlateSurf, oDeclRootPos) & _
'                            CStr(Round(dDeclivityAngle * 180 / PI, 1))
'        If Not bDeclivityIsOnBase Then
'            oMarkingInfo.Name = oMarkingInfo.Name & vbLf & "C"
'        End If

        oMarkingInfo.FittingAngle = dDeclivityAngle
        oMarkingInfo.Direction = GetDirAsFAUDIO(oThkDirVecAlongPlateSurf, oDeclRootPos)
        oMarkingInfo.ThicknessDirection = oThkDirVecAlongPlateSurf

        oMarkingInfo.SetAttributeNameAndValue "DECLIVITY", CVar(dDeclivityAngle)
        oMarkingInfo.SetAttributeNameAndValue "MOUNTINGDIRECTION", CVar(oMarkingInfo.Direction)

        If lConnMarkSide = -4 Then ' Edge Reinforcement
            oMarkingInfo.SetAttributeNameAndValue "MEASURESIDE", CVar("base")
            If bDeclivityIsOnBase Then
                oMarkingInfo.SetAttributeNameAndValue "MARKSIDE", CVar("marking")
            Else
                oMarkingInfo.SetAttributeNameAndValue "MARKSIDE", CVar("anti_marking")
            End If
        Else
        If bDeclivityIsOnBase Then
            oMarkingInfo.SetAttributeNameAndValue "MEASURESIDE", CVar("base")
        Else
            oMarkingInfo.SetAttributeNameAndValue "MEASURESIDE", CVar("offset")
        End If
            If lMarkingSide = lConnMarkSide Then
            oMarkingInfo.SetAttributeNameAndValue "MARKSIDE", CVar("marking")
        Else
            oMarkingInfo.SetAttributeNameAndValue "MARKSIDE", CVar("anti_marking")
        End If
        End If

        oSystemMark.SetMarkingSide lMarkingSide

        Dim oGeom3d As IJMfgGeom3D
        Set oGeom3d = oGeom3dFactory.Create(oResMgr)
        oGeom3d.PutGeometry oLineCS
        oGeom3d.PutGeometrytype STRMFG_MOUNT_ANGLE_MARK

        Dim oMoniker As IMoniker
        If oRelatedPart Is Nothing Then
            Set oMoniker = oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
        Else
            Set oMoniker = oRefGeom3D.GetMoniker
        End If

        oGeom3d.PutMoniker oMoniker
        oGeom3d.TrimToBoundaries = False

        ' Some contract to decide if only label is required.
'        If Some_Condition Then
'            oGeom3d.IsSupportOnly = True
'        End If

        oSystemMark.Set3dGeometry oGeom3d

        oGeomCol3d.AddGeometry lGeomCount, oGeom3d
        lGeomCount = lGeomCount + 1

NextFramePos:
        Set oDeclStartPos = Nothing
        Set oDeclEndPos = Nothing
        Set oDeclPosOnCS = Nothing
        Set oDeclRootPos = Nothing
        Set oRotMat = Nothing
        Set oDeclVec = Nothing
        Set oCurveColl = Nothing
        Set oLine1 = Nothing
        Set oLine2 = Nothing
        Set oLineCS = Nothing
        Set oMoniker = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
    Next ' oFramePos In oFramePosColl

CleanUp:
    Set oGenericIntersector = Nothing
    Set oTeeCurve = Nothing
    Set oTeeWire = Nothing
    Set oMfgRuleHelper = Nothing
    Set oToolBox = Nothing
    Set oFramePosColl = Nothing
    Set oFramePos = Nothing
    Set oProfilePart = Nothing
    Set oTopologyLocate = Nothing
    Set oPlatePart = Nothing
    Set oConnectingPort = Nothing
    Set oConnectingPortSurface = Nothing
    Set oThisPlate = Nothing
    Set oSystemMarkFactory = Nothing
    Set oGeom3dFactory = Nothing
    Set oSDObjectWrapper = Nothing
    Set oFrameNormal = Nothing
    Set oLines = Nothing
    Set oComplexStrings = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oOffsetPos = Nothing
    Set oMarkTangentVec = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function CreateProfileLocationMark(ByVal Part As Object, ByVal UpSide As Long, _
                                          ByVal bSelectiveRecompute As Boolean, _
                                          ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                          Optional ByVal bAddERMarks As Boolean, _
                                          Optional ByVal bAddFittingMark As Boolean, _
                                          Optional ByVal bAddRENDMarks As Boolean, _
                                          Optional ByVal bAddDeclivityMark As Boolean, _
                                          Optional ByVal CONN_PART_CONDITION As Long = 0, _
                                          Optional ByVal dTrimValue As Double, _
                                          Optional ByVal bDeclMarkOnUpside As Boolean = False) As IJMfgGeomCol3d
    Const METHOD = "CreateProfileLocationMark"
    On Error Resume Next

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_PROFILELOCATION_MARK, ReferenceObjColl, oGeomCol3d
    Set CreateProfileLocationMark = oGeomCol3d

    Dim eSideOfInputPlateToBeMarked As enumPlateSide

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oGeomCol3Dextra As IJMfgGeomCol3d
    Dim pResMgr As IUnknown
    Dim oMarkingInfo As MarkingInfo
    Dim oMfgPart As IJMfgPlatePart
    Dim oPOM As IJDPOM
    Dim oCmnAppUtil As IJDCmnAppGenericUtil
    Dim oStructDetailHelper As New StructDetailHelper
    Dim validResultFlg As Boolean
    Dim ePlateType As StructPlateType
    Dim oVector1 As IJDVector
    Dim oStart As IJDPosition, oEnd As IJDPosition
    Dim oPlateNormal As IJDVector
    Dim oSurfaceBody As IJSurfaceBody

    ePlateType = oSDPlateWrapper.plateType

    Set oCmnAppUtil = New CmnAppGenericUtil

    oCmnAppUtil.GetPOMFromObject Part, oPOM

    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If

    Set oGeomCol3Dextra = oMfgPlateWrapper.GetGeometriesBeforeUnfold

    Dim oMoniker As IMoniker
    Dim oGeom3d As IJMfgGeom3D

    'Dim ogeomcol3dout As IJMfgGeomCol3d
    'Set ogeomcol3dout = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    Dim bContourTee As Boolean
    Dim oVector As IJDVector
    Dim oWB As IJWireBody, oTeeWire As IJWireBody
    Dim oCS As IJComplexString
    Dim oCSColl As IJElements
    Dim oSystemMark As IJMfgSystemMark
    Dim lGeomCount As Long
    lGeomCount = 1

    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    'Set oConObjsCol = oSDPlateWrapper.ConnectedObjects
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)

    If oConObjsCol Is Nothing Then
        'No connected objects so we can leave
        GoTo CleanUp
    End If

    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex As Long
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
            Exit For
        End If
    Next nIndex

    Dim Item As Object
    Dim oConnectionData As ConnectionData
    Dim dStartMargin As Double, dEndMargin As Double

    ' Loop thru each Physical Connections

    ' Get the moulded side from SD Wrapper
    Dim sMoldedSide As String
    sMoldedSide = oSDPlateWrapper.MoldedSide
    If sMoldedSide = "Base" Then
        eSideOfInputPlateToBeMarked = BaseSide
    ElseIf sMoldedSide = "Offset" Then
        eSideOfInputPlateToBeMarked = OffsetSide
    Else
        Dim lErrNumber As Long
        lErrNumber = LogMessage(Err, MODULE, METHOD, " Unexpected Molded side ")
    End If

     'Determine whether the Plate part is crossing the centre line
    Dim bCenterPart As Boolean
    Dim oHelperSupport As IJMfgRuleHelpersSupport

    Set oHelperSupport = New MfgRuleHelpersSupport
    bCenterPart = False

    If Not oHelperSupport Is Nothing Then
        Dim dMinX As Double, dMinY As Double, dMinZ As Double
        Dim dMaxX As Double, dMaxY As Double, dMaxZ As Double

        'Get the Part Range
        oHelperSupport.GetRange Part, dMinX, dMinY, dMinZ, dMaxX, dMaxY, dMaxZ
        'Check for Y value to determine whether part is crossing the Center line (y=0)
        If dMinY < 0 And dMaxY > 0 Then
            bCenterPart = True
        End If
    End If

    Set oHelperSupport = Nothing
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)

        'Check if the connected object is a Profile Part
        If TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Or _
           TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
            ' Check if this physical connection is a Root PC, which particpates in Split operation
            ' If so, Skip the marking line line creation for this Root PC.
            Dim oStructEntityOperation As IJDStructEntityOperation
            Dim opeartionProgID As String
            Dim opeartionID As StructOperation
            Dim oOperColl As New Collection

            Set oStructEntityOperation = oConnectionData.AppConnection
            oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl

            Set oStructEntityOperation = Nothing
            Set oOperColl = Nothing

            ' If the RootPC has Split operation in its graph, just goto the next pc.
            If opeartionID = ConnectionSplitOperation Then
                ' No need to create marking line.
                GoTo NextItem
            End If

            'Initialize the profile wrapper and the Physical Connection wrapper
            Dim oSDConWrapper As Object
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport

            If TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Then
                Set oSDConWrapper = New StructDetailObjects.ProfilePart
                Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Else
                Set oSDConWrapper = New StructDetailObjects.MemberPart
                Set oPartSupport = New GSCADSDPartSupport.MemberPartSupport
            End If

            Set oSDConWrapper.object = oConnectionData.ToConnectable
            Set oPartSupport.Part = oSDConWrapper.object

            Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oConnectionData.AppConnection

            'Fix for TR 54174
            'Pass the correct input argument (ThicknessSide) and use the two outputs of
            'the method, Connection_ContourTee() in order to get the correct desired results.
            'Using this approach fixes the TRs 54174, 50701, 53788 and 53675.
            'When LandingCurve() method was used, the landing curve was not getting trimmed
            'to the region of the plate part that supports the profile.
            'This fix takes care of the fact that the thickness direction of the detail profile
            'part should be determined based on the relative position of the web w.r.t. the load point
            'This fix also ensures that the landing curve should be used as the profile location line.


            Dim oConnPartSupport As GSCADSDPartSupport.IJProfilePartSupport
            Set oConnPartSupport = oPartSupport

            Dim eTSide As GSCADSDPartSupport.ThicknessSide
            eTSide = SideA ' Initialize to a safe default value

            If Not TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
                eTSide = oConnPartSupport.ThicknessSideAdjacentToLoadPoint
            End If

            'Connection_ContourTee() method fails if 'ThicknessSide' is "SideUnspecified". TR 55422
            'has been filed on StructDetail to resolve this
            '"SideUnspecified" means that the landing curve (load point) is exactly in the middle of the mounting face

            Set oConnPartSupport = Nothing
            Set oPartSupport = Nothing

            bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, _
                                                                eTSide, _
                                                                oTeeWire, _
                                                                oVector)
            If bContourTee = True Then
                ' Bound the wire based on split points, if there are any.
                Set oWB = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire)

                'Convert the IJWireBody to ComplexStrings
                Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWB)

                If Not oCSColl Is Nothing Then
                  If oCSColl.Count = 0 Then
                      Set oCSColl = Nothing
                  End If
                End If

                If (oCSColl Is Nothing) Then
                   GoTo NextItem
                End If

                For Each oCS In oCSColl

                    'Create a SystemMark object to store additional information
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                    'Set the marking side
                    Dim lMarkingSide As Long
                    lMarkingSide = oPlateWrapper.GetSide(oConnectionData.ConnectingPort)

                    oSystemMark.SetMarkingSide lMarkingSide

                    'QI for the MarkingInfo object on the SystemMark
                    Set oMarkingInfo = oSystemMark

                    Dim strProdName     As String
                    strProdName = GetManufacturingProductionName(Part, oConnectionData.ToConnectable, CONN_PART_CONDITION)
                    oMarkingInfo.Name = strProdName
                    'oMarkingInfo.Name = oSDConWrapper.Name

                    oMarkingInfo.Thickness = oSDConWrapper.WebThickness

                    Dim oPort As IJPort
                    Set oPort = oConnectionData.ToConnectedPort
                    Dim oIJSurface As IJSurface
                    Set oIJSurface = oPort.Geometry

                    'Should not use GetDirection because it won't consider the "out" and "in" cases
                    'Written new Method which will determine the direction string based on 1.Vector 2. Based on whether
                    'the Plate part is crossing centre line 3. And the Marking line curve position

                    'Fix for TR TR-CP·69190  MFG part monitor thickness symbol for plates and profiles are incorrect

                    'oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oVector)
                    'oMarkingInfo.Direction = m_oMfgRuleHelper.GetThicknessDirection(oCS, oVector, bCenterPart)

    '                oMarkingInfo.Direction = GetProfileThicknessDir(oConnectionData.ToConnectable)

                    oWB.GetEndPoints oStart, oEnd
                    Set oSurfaceBody = oSDPlateWrapper.BasePort(BPT_Base).Geometry
                    Set oStart = m_oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBody, oVector1)
                    oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal
                    Set oSurfaceBody = Nothing
                    Set oVector1 = Nothing
                    Set oStart = Nothing
                    Set oEnd = Nothing

                    If (Abs(oPlateNormal.x) > Abs(oPlateNormal.y) And Abs(oPlateNormal.x) > Abs(oPlateNormal.z)) Then
                        oVector.x = 0
                    ElseIf Abs(oPlateNormal.y) > Abs(oPlateNormal.z) Then
                        oVector.y = 0
                    Else
                        oVector.z = 0
                    End If

                    If eTSide = SideUnspecified Then
                        oMarkingInfo.Direction = "centered"
                    Else
                        oMarkingInfo.Direction = m_oMfgRuleHelper.GetThicknessDirection(oCS, oVector, bCenterPart)
                        If eTSide = SideA Then
                            sMoldedSide = "Base"
                        ElseIf eTSide = SideB Then
                            sMoldedSide = "Offset"
                        End If
                        oMarkingInfo.ThicknessDirection = GetThicknessDirectionVector(oCS, oSDConWrapper, sMoldedSide, oIJSurface, oVector)
                    End If


                    GetMarginExtensionDistances oGeomCol3Dextra, oCS, pResMgr, dStartMargin, dEndMargin

                    'If there is margin applied to the plate extend all profile location lines with that value
                    If dStartMargin <> 0 Then
                          m_oMfgRuleHelper.ExtendWire oCS, dStartMargin, 0
                    End If

                    If dEndMargin <> 0 Then
                          m_oMfgRuleHelper.ExtendWire oCS, dEndMargin, 1
                    End If

                    If dTrimValue > 0 Then
                        'Trim the curve in the ends
                        m_oMfgRuleHelper.TrimCurveEnds oCS, dTrimValue, Nothing
                    End If

                    Dim dFittingAngle As Double

                    dFittingAngle = oSDPhysicalConn.TeeMountingAngle

                    Dim oMeasuredDir As IJDVector
                    Dim dDotP As Double

                    If TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Or _
                       TypeOf oConnectionData.ToConnectable Is IJProfilePart _
                    Then
                        ' Get the declivity angle measurement direction
                        Set oMeasuredDir = GetDeclivityDirection(Part, oConnectionData.ToConnectable, oSDPhysicalConn, oWB)
                    End If

                    If (Abs(oPlateNormal.x) > Abs(oPlateNormal.y) And Abs(oPlateNormal.x) > Abs(oPlateNormal.z)) Then
                        oMeasuredDir.x = 0
                    ElseIf Abs(oPlateNormal.y) > Abs(oPlateNormal.z) Then
                        oMeasuredDir.y = 0
                    Else
                        oMeasuredDir.z = 0
                    End If

                    If eTSide = SideUnspecified Then
                        oMarkingInfo.ThicknessDirection = oMeasuredDir
                    Else
                        ' We need to complement the fitting angle, if part thickness direction is NOT in the same direction of declivity.
                        dDotP = oMeasuredDir.Dot(oVector)
                        If (dDotP < 0#) Then
                            dFittingAngle = PI - dFittingAngle
                        End If
                    End If

                    oMarkingInfo.FittingAngle = dFittingAngle
                    Dim strFlangeDir As String

                    Dim eProfileFlangeOrientation As StructMoldedOrientation
                    Dim eProfileType As StructProfileType

                    oSDConWrapper.Get_SecondOrientation eProfileFlangeOrientation, validResultFlg
                    eProfileType = oSDConWrapper.ProfileType
                    Dim bIsOpposite As Boolean
                    bIsOpposite = False
                    If eTSide = SideB Then
                        bIsOpposite = True
                    ElseIf eTSide = SideA Then
                        bIsOpposite = False
                    ElseIf eTSide = SideUnspecified Then

                    End If

                    'strFlangeDir = GetFlangeOrientation(ePlateType, eProfileType, eProfileFlangeOrientation, bCenterPart)
                    If eTSide = SideUnspecified Then
                        Dim oWB2 As IJWireBody, oTeeWire2 As IJWireBody
                        Dim oCS2 As IJComplexString

                        bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, _
                                                                 SideA, _
                                                                oTeeWire2, _
                                                                oVector)
                        If bContourTee = True Then
                            ' Bound the wire based on split points, if there are any.
                            Set oWB2 = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire2)

                            'Convert the IJWireBody to a IJComplexString
                            Set oCS2 = m_oMfgRuleHelper.WireBodyToComplexString(oWB2)
                            oWB2.GetEndPoints oStart, oEnd
                            Set oSurfaceBody = oSDPlateWrapper.BasePort(BPT_Base).Geometry
                            Set oStart = m_oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBody, oVector1)
                            oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal

                            If (Abs(oPlateNormal.x) > Abs(oPlateNormal.y) And Abs(oPlateNormal.x) > Abs(oPlateNormal.z)) Then
                                oVector.x = 0
                            ElseIf Abs(oPlateNormal.y) > Abs(oPlateNormal.z) Then
                                oVector.y = 0
                            Else
                                oVector.z = 0
                            End If

                            strFlangeDir = m_oMfgRuleHelper.GetThicknessDirection(oCS2, oVector, bCenterPart)
                            oMarkingInfo.FlangeDirection = strFlangeDir
                            Set oVector1 = Nothing
                            Set oStart = Nothing
                            Set oEnd = Nothing
                        End If
                    Else
                        If bIsOpposite Then
                            oMarkingInfo.FlangeDirection = "opposite"
                        Else
                            oMarkingInfo.FlangeDirection = "same"
                        End If
                    End If

                    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                    Dim oProfileSection As IJDProfileSection
                    Set oProfileSection = oConnectionData.ToConnectable

                    If oProfileSection Is Nothing Then
                        oGeom3d.PutGeometrytype STRMFG_PROFILELOCATION_MARK
                    Else
                        ' If the profile secondary orientation is changed such that webleft/webright become
                        ' the mounting face then treat the profile location mark as lap mark
                        If oProfileSection.mountingFace = LeftWeb Or oProfileSection.mountingFace = RightWeb Then
                            Dim oWebDirVector   As IJDVector
                            Set oWebDirVector = GetWebDirectionVector(oProfileSection)
                            oMarkingInfo.ThicknessDirection = oWebDirVector
                            oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oWebDirVector)

                            oGeom3d.PutGeometrytype STRMFG_LAP_MARK
                            bAddFittingMark = True ' For lap profile mark on Plate, show end fitting mark also
                        Else
                            oGeom3d.PutGeometrytype STRMFG_PROFILELOCATION_MARK
                        End If
                    End If

                    oGeom3d.PutGeometry oCS

                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker

                    oSystemMark.Set3dGeometry oGeom3d

                    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                    lGeomCount = lGeomCount + 1

                    '*** HARITSUKE Check ***'
                    'Create Indirect Bracket Location Marks for Each Profile location mark
                    CreateIndirectBracketLocationMark Part, oConnectionData.ToConnectable, ReferenceObjColl, oGeomCol3d, strProdName

                    '********************Start of Logic For Creating Fitting Marks at Bend Knuckle Position***********'
                    Dim bProfileKnuckle As Boolean
                    bProfileKnuckle = False

                    Dim oPrProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
                    Set oPrProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport

                    Dim oPrPartSupport As IJPartSupport
                    Set oPrPartSupport = oPrProfilePartSupport
                    Set oPrPartSupport.Part = oConnectionData.ToConnectable

                    Dim oBendColl As New Collection
                    Dim oKnkLineColl As New Collection
                    Dim oKnucAngColl As New Collection

                    '************** Getting the Bend Knuckle WireBody******************'
                    oPrProfilePartSupport.GetBendProfileKnuckleMfgData JXSEC_WEB_RIGHT, oBendColl, oKnkLineColl, oKnucAngColl
                    Set oBendColl = Nothing
                    Set oKnucAngColl = Nothing
                    Set oPrPartSupport = Nothing
                    '************** End ***********************'

                    If oKnkLineColl Is Nothing Then
                        bProfileKnuckle = False
                    Else
                        If oKnkLineColl.Count = 0 Then
                            bProfileKnuckle = False
                        Else
                            bProfileKnuckle = True
                        End If
                    End If

                    If bProfileKnuckle = True Then
                        CreateKnuckleFittingMarks oCS, oConnectionData, Part, sMoldedSide, UpSide, oKnkLineColl, Nothing, Nothing, oGeomCol3d
                    End If

                    '**********************End of the Logic***********************'

                    ' Check if declivity marks are needed
                    If bAddDeclivityMark = True Then
                        If bDeclMarkOnUpside Then
                            ' Call like so if you want declivity marks always on upside
                            CreateDeclivityMarks Part, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, UpSide, sMoldedSide, Nothing
                        Else
                            ' Call like so if you want declivity marks on same side as its location mark
                            CreateDeclivityMarks Part, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, lMarkingSide, sMoldedSide, Nothing
                        End If
                    End If

                    '*** CREATE END FITTING MARK ***'
                    'This function Creates the End Fitting Mark(s) at the Free End(s) of the
                    '... Profile Parts.
                    If bAddFittingMark Then
                        CreateEndFittingMark Part, oCS, oWB, oPartSupport, oSDPlateWrapper, _
                                             oSDConWrapper, sMoldedSide, lMarkingSide, oConnectionData, _
                                             oGeomCol3d, oPlateNormal, False
                    Else
                        '*******************************'


                        '' create the location of profile top location marking line and
                        '' add it to the oGeomCol3d collection.
                        Call CreateProfileTopLocationMarkingLine(eSideOfInputPlateToBeMarked, _
                                    oConnectionData.ConnectingPort, _
                                    m_oMfgRuleHelper.ComplexStringToWireBody(oCS), _
                                    Part, _
                                    oConnectionData.ToConnectable, _
                                    oGeomCol3d)
                    End If



                Next
            End If

            ' Create REND fitting mark for ERs.
            ' REND fitting marks should no be limited to Edge-reinforcements.
            If bAddRENDMarks Then
                CreateRENDProfileMarks Part, UpSide, oConnectionData, oGeomCol3Dextra, oGeomCol3d
            End If

NextItem:


            Set oWB = Nothing
            Set oTeeWire = Nothing
            Set oVector = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oGeom3d = Nothing
            Set oSDConWrapper = Nothing
            Set oSDPhysicalConn = Nothing
            If Not oCSColl Is Nothing Then
                oCSColl.Clear
                Set oCSColl = Nothing
            End If
        End If


    Next nIndex


    'Add Fitting marks for Edge Reinforcement Parts
    If bAddERMarks Then
        AddMarksForEdgeReinforcementParts Part, UpSide, oConObjsCol, ReferenceObjColl, oGeomCol3d
    End If

    'Create Indirect Mamber Marks for all indirectly connected members
    CreateIndirectMemberMarks Part, ReferenceObjColl, oGeomCol3d

    'Return the 3d collection
    Set CreateProfileLocationMark = oGeomCol3d

CleanUp:
    Set oPlateNormal = Nothing
    Set oMeasuredDir = Nothing
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oCS = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
    GoTo CleanUp
End Function


' ***********************************************************************************
' Private Sub CreateProfileTopLocationMarkingLine
'
' Description:  This creates the marking line at the end of the bounded profile on a
'               manufacturing plate. The profile could be bounded by another profile
'               or a plate.
'
' Inputs    :   [eSide As enumPlateSide] side of the plate on which marking to be placed
'               [oPlatePort As IJPort] port of the plate on which profile is connected
'               [oWireBody As IJWireBody] wire body of the landing curve of the profile
'               [oPart As Object] seleted plate part to be manufactured
'               [oProfilePart As IJProfilePart] profile on the plate
'               [oGeomCol3d As IJMfgGeomCol3d]  MfgGeom3d collection to which MfgGeom3d would be added
'
' ***********************************************************************************

Public Sub CreateProfileTopLocationMarkingLine(eSide As enumPlateSide, _
                    oPlatePort As IJPort, _
                    oWireBody As IJWireBody, _
                    oPart As Object, _
                    oProfilePart As IJProfilePart, _
                    oGeomCol3d As IJMfgGeomCol3d)

    Const METHOD = "CreateProfileTopLocationMarkingLine"
    On Error GoTo ErrorHandler

    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    Dim oStartDir           As IJDVector
    Dim oEndDir             As IJDVector

    '- get the start and end positions and their directions.
    oWireBody.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir

    '- find out the end in which a profile has to be bounded.
    Dim oSDProfileWrapper           As StructDetailObjects.ProfilePart
    Dim oPlatePartSuface            As Object 'IJSurface
    Dim oProfileBoundaries          As Collection
    Dim oProfileTopCS               As IJComplexString
    Dim oGeom3d                     As IJMfgGeom3D

    '- initialize the profile part wrapper
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = oProfilePart

    '- get the geometry of port of the plate to be passed to the method CreateComplexStringAtEndLocation
    Set oPlatePartSuface = oPlatePort.Geometry

    '- get the boundaries of the profile
    Set oProfileBoundaries = oSDProfileWrapper.ProfileBoundaries

    If Not oProfileBoundaries Is Nothing Then
        Dim oBoundary           As BoundaryData
        '- add the profile location mark at the start end
        If oProfileBoundaries.Count >= 1 Then
            oBoundary = oProfileBoundaries.Item(1)
            '- if the profile is bounded by stiffener or plate then proceed
            If oBoundary.ObjectType = "Stiffenerpart" Or _
                    oBoundary.ObjectType = "StiffenerSystem" Or _
                    oBoundary.ObjectType = "PlatePart" Or _
                    oBoundary.ObjectType = "PlateSystem" Then

                '- create the complex string and add it the geom3d collection
                Set oProfileTopCS = CreateComplexStringAtEndLocation(oStartPos, oStartDir, oPlatePartSuface)
                If Not oProfileTopCS Is Nothing Then
                    Set oGeom3d = CreateMfgGeom3d(eSide, oProfileTopCS, oSDProfileWrapper)
                    oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
                End If
            End If
        End If

        '- add the profile location line at the other end
        If oProfileBoundaries.Count >= 2 Then
            oBoundary = oProfileBoundaries.Item(2)
            '- if the profile is bounded by stiffener or plate then proceed
            If oBoundary.ObjectType = "Stiffenerpart" Or _
                    oBoundary.ObjectType = "StiffenerSystem" Or _
                    oBoundary.ObjectType = "PlatePart" Or _
                    oBoundary.ObjectType = "PlateSystem" Then

                '- create the complex string and add it the geom3d collection
                Set oProfileTopCS = CreateComplexStringAtEndLocation(oEndPos, oEndDir, oPlatePartSuface)
                If Not oProfileTopCS Is Nothing Then
                    Set oGeom3d = CreateMfgGeom3d(eSide, oProfileTopCS, oSDProfileWrapper)
                    oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
                End If
            End If
        End If

    End If

CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oSDProfileWrapper = Nothing
    Set oPlatePartSuface = Nothing
    Set oProfileBoundaries = Nothing
    Set oProfileTopCS = Nothing
    Set oGeom3d = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub


' ***********************************************************************************
' Private Function CreateComplexStringAtEndLocation
'
' Description:  This function returns the complex string for marking purpose at the
'               end of the profile bounded by a profile or plate. It creates a tangent
'               on the connecting plate surface and scales this vector by FittingMarkLength
'               and creates the line. The tangent on connecting surface has been created
'               by getting cross product of end direction and the normal at the end position
'               on plate surface.
'
' Inputs    :   [oProfileEndPosition As IJDPosition]  The marking postion on the plate, which is nothing
'                                               but end point of the landing curve of the profile.
'               [oEndDirection As IJDVector]  The direction of the landing curve at the end position .
'               [oPlateSurfaceBody As IJSurfaceBody] The geometry of the plate surface on which
'                                               profile is connected.
'
' ***********************************************************************************

Public Function CreateComplexStringAtEndLocation( _
            oProfileEndPosition As IJDPosition, _
            oEndDirection As IJDVector, _
            oPlateSurfaceBody As IJSurfaceBody) As IJComplexString

    Const METHOD = "CreateComplexStringAtEndLocation"
    On Error GoTo ErrorHandler

    Dim oPlateNormal            As IJDVector
    Dim oPlateTangent           As IJDVector
    Dim oVector1                As IJDVector

    '- project end position onto plate surfaces
    Set oProfileEndPosition = m_oMfgRuleHelper.ProjectPointOnSurface(oProfileEndPosition, oPlateSurfaceBody, oVector1)

    '- get the normal at end position of plate surface
    oPlateSurfaceBody.GetNormalFromPosition oProfileEndPosition, oPlateNormal

    '- get the tangent at the end postion. The tanget is nothing but direction
    '- for the marking line.
    Set oPlateTangent = oPlateNormal.Cross(oEndDirection)

    Dim xStart As Double
    Dim yStart As Double
    Dim zStart As Double
    Dim xEnd As Double
    Dim yEnd As Double
    Dim zEnd As Double

    ' scale the vector and find the start and end points for the marking line
    m_oMfgRuleHelper.ScaleVector oPlateTangent, FittingMarkLength / 2
    oProfileEndPosition.Get xStart, yStart, zStart
    oPlateTangent.Get xEnd, yEnd, zEnd

    ' generate a line
    Dim oLine As IngrGeom3D.Line3d
    Set oLine = New IngrGeom3D.Line3d
    oLine.DefineBy2Points xStart - xEnd, yStart - yEnd, zStart - zEnd, xStart + xEnd, yStart + yEnd, zStart + zEnd

    ' make the line into complex string
    Dim oLineCS As IJComplexString
    Set oLineCS = New ComplexString3d
    oLineCS.AddCurve oLine, False

    ' make the complex string into wire body
    Dim oWireBody As IJWireBody
    Set oWireBody = m_oMfgRuleHelper.ComplexStringToWireBody(oLineCS)

    ' project wire body to 'connecting surface' along projection vector
    Dim oCS As IJComplexString
    Set oCS = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oPlateSurfaceBody, oWireBody, oPlateNormal)

    'return complex string
    Set CreateComplexStringAtEndLocation = oCS

CleanUp:
    Set oCS = Nothing
    Set oWireBody = Nothing
    Set oLineCS = Nothing
    Set oLine = Nothing
    Set oVector1 = Nothing
    Set oPlateNormal = Nothing
    Set oPlateTangent = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Private Function CreateMfgGeom3d
'
' Description:  This function returns the MfgGeom3d for a given side, complex string and profile wrapper.
'
' Inputs    :   [lSide As Long] side of the plate on which marking to be placed
'               [oCS As IJComplexString] complex string of the landing curve of the profile
'               [oSDProfileWrapper As StructDetailObjects.ProfilePart] profile wrapper with profile on the plate
'
' ***********************************************************************************

Public Function CreateMfgGeom3d(lSide As Long, _
            oCS As IJComplexString, _
            oSDProfileWrapper As StructDetailObjects.ProfilePart) As IJMfgGeom3D

    Const METHOD = "CreateMfgGeom3d"
    On Error GoTo ErrorHandler

    Dim oSystemMark As IJMfgSystemMark
    Dim oUnkSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D

    ' create a SystemMark object to store additional information
    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    ' set the marking side
    oSystemMark.SetMarkingSide lSide

    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark

    ' set the name and thickness for marking info
    oMarkingInfo.Name = "FITTING MARK"
    oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness

    ' set the direction
    Dim oProfileLandingCurve As IJWireBody
    Dim oThickDir            As IJDVector
    Dim bIsThicknessCentered As Boolean

    'Get the profile part landing curve
    Call oSDProfileWrapper.LandingCurve(oProfileLandingCurve, oThickDir, bIsThicknessCentered)
    oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oThickDir)
    oMarkingInfo.ThicknessDirection = oThickDir

    Set oProfileLandingCurve = Nothing
    Set oThickDir = Nothing

    ' set the geometry for the marking line
    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
    Set oUnkSystemMark = oSystemMark

    oSystemMark.Set3dGeometry oGeom3d

    Set CreateMfgGeom3d = oGeom3d
CleanUp:
    Set oSystemMark = Nothing
    Set oUnkSystemMark = Nothing
    Set oMoniker = Nothing
    Set oMarkingInfo = Nothing
    Set oGeom3d = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function AddMarksForEdgeReinforcementParts(ByVal oPlatePart As Object, ByVal UpSide As Long, ByVal oConObjsCol As Collection, ByVal ReferenceObjColl As Collection, ByVal oGeom3dColl As Object)
    Const METHOD = "ProfileLocationMark: AddMarksForEdgeReinforcementParts"
    On Error GoTo ErrorHandler

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oPlatePart

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlatePart

    Dim oPlateSys As IJPlate
    Set oPlateSys = oPlateWrapper.GetRootSystem

    On Error Resume Next

    Dim oResourceManager As Object

    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim nConnections As Long
    nConnections = oConObjsCol.Count

    Dim iIndex As Long, nLines As Long, iLineIdx As Long
    Dim oIJProfileER As IJProfileER
    Dim oVector As IJDVector
    Dim oWB As IJWireBody, oTeeWire As IJWireBody
    Dim bContourTee As Boolean
    Dim oConnectionData As ConnectionData
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    Dim oGeom3d As IJMfgGeom3D
    Dim oMoniker As IMoniker
    Dim lGeomCount As Long
    Dim oPlateLateralSurfaceBody As IJSurfaceBody, oSurfaceBody As IJSurfaceBody
    'Dim oPosition As New DPosition
    'Dim oNormal As IJDVector
    Dim oGeomFac As New GeometryFactory
    Dim oLines3d As ILines3d
    Dim oComplexStrings3d As IComplexStrings3d
    Dim oLine As Line3d
    Dim oCrvElemets As IJElements
    Dim oOtherPort As IJPort
    Dim dStartPar As Double, dEndPar As Double, dOffsetPar As Double
    Dim vTanX As Double, vTanY As Double, vTanZ As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim bOffsetEdge As Boolean, bSlantER As Boolean
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    Dim oPlateNormal As IJDVector

    '*** Get Plate Molded Surface Body and its ModelBody ***'
    Set oSurfaceBody = oPlateWrapper.GetSurfacePort(UpSide).Geometry
    Dim oPlateModelbody As IJDModelBody
    Set oPlateModelbody = oSurfaceBody

    'Dim sPSBFileName As String
    'sPSBFileName = Environ("TEMP")
    'If sPSBFileName = "" Or sPSBFileName = vbNullString Then
    '    sPSBFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'sPSBFileName = sPSBFileName & "\oPlateSurfaceBody.sat"
    'oPlateModelbody.DebugToSATFile sPSBFileName

    '*** Get Plate Molded Surface Normal ***'
    Set oPlateNormal = GetPlateNeutralSurfaceNormal(oPlatePart)

    Set oCrvElemets = New JObjectCollection

    Set oLines3d = oGeomFac.Lines3d
    Set oComplexStrings3d = oGeomFac.ComplexStrings3d

    Dim oPosition As New DPosition
    Dim oNormal As IJDVector
    Dim oPCPortColl As IJElements
    Set oPCPortColl = New JObjectCollection

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart

    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If

    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex As Long
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
            Exit For
        End If
    Next nIndex

    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Set oPlateCreation_AE = oMfgPart.ActiveEntity

    '*** Collection of Outer Contours ***'
    Dim oGeomCol3dContour As IJMfgGeomCol3d
    Set oGeomCol3dContour = oPlateCreation_AE.GeometriesBeforeUnfold

    Dim oPos1 As New DPosition, oPos2 As New DPosition, dMinDist As Double
    Dim oBasePos As New DPosition
    Dim eGeomType As StrMfgGeometryType
    Dim dX As Double, dY As Double, dZ As Double
    Dim oCentroidPt As New DPosition
    Dim oRootSystemofER As Object, oRootSystemofPlate As Object

    lGeomCount = 1

    For iIndex = 1 To nConnections
        oConnectionData = oConObjsCol.Item(iIndex)
        Set oIJProfileER = Nothing
        Set oIJProfileER = oConnectionData.ToConnectable

        If oIJProfileER Is Nothing Then GoTo NextItem

        Set oSDProfileWrapper.object = oIJProfileER

        '*** Check if this plate system is parent for the ER. If NOT, DO NOT MARK ANYTHING ***'
        oSDProfileWrapper.GetStiffenedPlate oRootSystemofER, True
        Set oRootSystemofPlate = m_oMfgRuleHelper.GetTopMostParentSystem(oPlatePart)

        If Not oRootSystemofER Is oRootSystemofPlate Then
            GoTo NextItem
        End If

        '*************************************************************************************'

        'Set oParentPlate = oSDProfileWrapper.ParentSystem
        Dim oTopSurfaceCG As IJDPosition, oBottomSurfaceCG As IJDPosition
        Dim oProfTopSurface As IJSurfaceBody, oProfBottomSurface As IJSurfaceBody
        Dim oBottomModelbody As IJDModelBody, oTopModelbody As IJDModelBody

        'Get the Top & Bottom Sutfaces of the Profile ER
        Set oProfTopSurface = oSDProfileWrapper.SubPort(JXSEC_TOP).Geometry
        Set oProfBottomSurface = oSDProfileWrapper.SubPort(JXSEC_BOTTOM).Geometry
        Set oTopModelbody = oProfTopSurface
        Set oBottomModelbody = oProfBottomSurface

        '*** DEBUG ***'
'            Dim sPTSFileName As String
'            Dim sPBSFileName As String
'            sPTSFileName = Environ("TEMP")
'            sPBSFileName = Environ("TEMP")
'            If sPTSFileName = "" Or sPTSFileName = vbNullString Then
'                sPTSFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'            If sPBSFileName = "" Or sPBSFileName = vbNullString Then
'                sPBSFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'            sPTSFileName = sPTSFileName & "\oProfTopSurface.sat"
'            sPBSFileName = sPBSFileName & "\oProfBottomSurface.sat"
'            oTopModelbody.DebugToSATFile sPTSFileName
'            oBottomModelbody.DebugToSATFile sPBSFileName
        '*************'
        'Get the CGs of Surfaces for getting the Normals
        oProfTopSurface.GetCenterOfGravity oTopSurfaceCG
        oProfBottomSurface.GetCenterOfGravity oBottomSurfaceCG

        bOffsetEdge = False
        bSlantER = False

        Dim oSupp2 As IJPartSupport
        Dim oSys As IJSystem
        Dim oNamedItem As IJNamedItem

        Set oSupp2 = New PartSupport
        Set oSupp2.Part = oIJProfileER
        oSupp2.IsSystemDerivedPart oSys, True

        Set oNamedItem = oSys

        If Not oSys Is Nothing Then
            Dim oList As IJElements
            Dim objEROrientation As IJDProfileOrientation

            Set objEROrientation = oSys
            Set oList = objEROrientation.RegionCollection

            Dim objRegion As IJDProfileRegion
            Set objRegion = oList.Item(1)

            If Not objRegion Is Nothing Then
                Dim objPlanarRegion As IJDProfileRegionPlanar
                Set objPlanarRegion = objRegion

                'objPlanarRegion.PlanarAngle
                Dim dMountAngle As Double
                dMountAngle = objPlanarRegion.PlanarAngle * 180 / PI

                If Abs(dMountAngle - 90) > 0.001 And dMountAngle <> 0 Then
                    bSlantER = True
                End If

            End If
        End If

        '*** Get the Offset value ***'
        If oIJProfileER.Position = ER_OnEdgeOffset And oIJProfileER.EdgeReinforcmentOffset <> 0 Then ' We need to put Offset ER Mark
            'Get Global Direction of Primary Orientation
            Dim oIJProfileAttributes As IJProfileAttributes
            Set oIJProfileAttributes = New ProfileUtils

            Dim oTest As IJDProfileMoldedConventions
            Set oTest = oIJProfileER

            'Get ER's Primary orientation
            Dim eOrientation As StructMoldedOrientation
            eOrientation = oTest.Orientation

            Dim sOrient As String
            'Get the Global Orientation Value
            sOrient = GetFlangeOrientation(oPlateSys.plateType, sptEdgeReinforcement, eOrientation)
            'Get Single Letter Direction
            sOrient = GetDirectionValue(sOrient)

            'Set Offset Edge flag to TRUE
            bOffsetEdge = True

            'Get Distance between Molded Surface of the plate and nearby surface of ER
            ' ... this is Custom Attribute D value

            'Get Surface Normal Vector for Profile Top & Bottom Surface
            Dim oProfTopSurfaceNormalVec As IJDVector, oProfBottomSurfaceNormalVec As IJDVector
            oProfTopSurface.GetNormalFromPosition oTopSurfaceCG, oProfTopSurfaceNormalVec
            oProfBottomSurface.GetNormalFromPosition oBottomSurfaceCG, oProfBottomSurfaceNormalVec

            'Get Minimum distance between surfaces
            Dim dDistBetweenSurfaces As Double
            Dim dPos1 As New DPosition, dPos2 As New DPosition

            'Check which surface is parallel to Plate Normal Surface
            '(We need distance between Molded Plate Surface and Nearby Top/Bottom
            '.. Profile Surface (Custom Attribute C in case of ER)
            If oPlateNormal.Dot(oProfTopSurfaceNormalVec) = 1 Then
                oPlateModelbody.GetMinimumDistance oTopModelbody, dPos1, dPos2, dDistBetweenSurfaces
            ElseIf oPlateNormal.Dot(oProfBottomSurfaceNormalVec) = 1 Then
                oPlateModelbody.GetMinimumDistance oBottomModelbody, dPos1, dPos2, dDistBetweenSurfaces
            Else
                'No implementation
            End If

            'MsgBox "dDistBetweenSurfaces:" & dDistBetweenSurfaces
        End If
        '****************************'

        ' Check if this physical connection is a Root PC, which participates in Split operation
        ' If so, Skip the marking line line creation for this Root PC.
        Dim oStructEntityOperation As IJDStructEntityOperation
        Dim opeartionProgID As String
        Dim opeartionID As StructOperation
        Dim oOperColl As New Collection

        Set oStructEntityOperation = oConnectionData.AppConnection
        oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl

        Set oStructEntityOperation = Nothing
        Set oOperColl = Nothing

        ' If the RootPC has Split operation in its graph, just goto the next pc.
        If opeartionID = ConnectionSplitOperation Then
            ' No need to create marking line.
            GoTo NextItem
        End If

        Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
        Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport

        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Set oPartSupport.Part = oIJProfileER
        Set oProfilePartSupport = oPartSupport

        Dim eTSide As GSCADSDPartSupport.ThicknessSide
        eTSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint

        Dim bThisPartCrossOfTee As Boolean
        Dim eTestContourConnectionType As GSCADSDPartSupport.ContourConnectionType

        eTestContourConnectionType = PARTSUPPORT_CONNTYPE_UNKNOWN

        oPartSupport.GetConnectionTypeForContour oConnectionData.AppConnection, _
                                           eTestContourConnectionType, _
                                           bThisPartCrossOfTee

        bContourTee = False
        If eTestContourConnectionType = PARTSUPPORT_CONNTYPE_TEE And bThisPartCrossOfTee Then
            bContourTee = oSDProfileWrapper.Connection_ContourTee(oConnectionData.AppConnection, _
                                                                  eTSide, _
                                                                  oTeeWire, _
                                                                  oVector)
        End If
        If Not bContourTee Then GoTo NextItem

        ' Bound the wire based on split points, if there are any.
        Set oWB = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire)

        Dim nCrvCount As Long, nCrvIdx As Long
        Dim oCurve As IJCurve
        Dim oNextCurve As IJCurve, oPrevCurve As IJCurve
        Dim dStartX As Double, dStartY As Double, dStartZ As Double
        Dim dEndX As Double, dEndY As Double, dEndZ As Double
        Dim dStart2X As Double, dStart2Y As Double, dStart2Z As Double
        Dim dEnd2X As Double, dEnd2Y As Double, dEnd2Z As Double
        Dim oWireBodyCS As ComplexString3d, oCS As ComplexString3d, oProjCS As ComplexString3d
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        Set oOtherPort = oConnectionData.ToConnectedPort

        Set oPlateLateralSurfaceBody = oOtherPort.Geometry

        Set oWireBodyCS = m_oMfgRuleHelper.WireBodyToComplexString(oWB)

        Dim oPCMidPt As IJDPosition
        Set oPCMidPt = m_oMfgRuleHelper.GetMiddlePoint(oWB)

        '*** If The ER is Slant, place Declivity Mark ***'
        If bSlantER = True Then
            CreateDeclivityMarks oPlatePart, oConnectionData, oWireBodyCS, Nothing, _
                                 Nothing, oGeom3dColl, lGeomCount, UpSide, "", Nothing
        End If
        '************************************************'

        nCrvCount = oWireBodyCS.CurveCount

'            If nCrvCount = 1 Then 'And bOffsetEdge = False Then

            ' Get the PC port coll given a ref ER and connected ERs
            Dim oTempPCPortColl As IJElements, bCircularER As Boolean
            bCircularER = False
            Set oTempPCPortColl = GetPhysicalConnPortFromERParts(oIJProfileER, oConObjsCol, oPlatePart, bCircularER)

            Dim bSplitER    As Boolean
            Dim bNewPC      As Boolean

            bNewPC = False
            bSplitER = False

            If Not oTempPCPortColl Is Nothing Then

                ' if the collection is 2 - it means ER is having PC at ends so do not process this ER
                If oTempPCPortColl.Count = 2 And bCircularER = False Then
                   GoTo CheckKnuckleERMarks
                End If

                bSplitER = True

                If bCircularER = False Then

                    Dim oPCPort     As IJPort
                    Set oPCPort = oTempPCPortColl.Item(1)

                    ' if the PC port is not in collection then this PC is new and create toshi mark
                    If Not oPCPortColl.Contains(oPCPort) Then
                        oPCPortColl.Add oPCPort
                        bNewPC = True
                    End If
                Else

                    If oPCPortColl.Contains(oTempPCPortColl.Item(1)) Then
                        GoTo CheckKnuckleERMarks
                    Else
                        oPCPortColl.AddElements oTempPCPortColl
                        bNewPC = True
                    End If

                End If
            End If

            For nCrvIdx = 1 To 1
                'oWireBodyCS.GetCurve nCrvIdx, oCurve
                Set oCurve = oWireBodyCS
                oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

                Dim oStartPos As IJDPosition
                Set oStartPos = New DPosition

                oStartPos.Set dStartX, dStartY, dStartZ

                Dim oEndPos As IJDPosition
                Set oEndPos = New DPosition

                oEndPos.Set dEndX, dEndY, dEndZ

                Dim oModelBody As IJDModelBody
                Set oModelBody = oPCPort.Geometry

                Dim oClosePos As IJDPosition, oMarkPos As IJDPosition
                Set oMarkPos = New DPosition

                oModelBody.GetMinimumDistanceFromPosition oStartPos, oClosePos, dMinDist

                If dMinDist < 0.001 Then
                    Set oClosePos = oStartPos
                Else
                    Set oClosePos = oEndPos
                End If

'                    If nCrvIdx = 1 Then
                    nLines = 2
'                    Else
'                        nLines = 1
'                    End If

                For iLineIdx = 1 To nLines

                    If iLineIdx = 1 Then
                        Set oVector = oStartPos.Subtract(oEndPos)
                        oVector.Length = 1
                        oMarkPos.Set dEndX, dEndY, dEndZ
                        If nCrvIdx = nCrvCount Then
                            oPosition.Set dEndX, dEndY, dEndZ
                            oCurve.Parameter dEndX, dEndY, dEndZ, dEndPar
                            oCurve.Evaluate dEndPar, dEndX, dEndY, dEndZ, _
                                    vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
                            oCurve.Parameter (dEndX - vTanX * ER_OFFSET_DIST), _
                                (dEndY - vTanY * ER_OFFSET_DIST), _
                                (dEndZ - vTanZ * ER_OFFSET_DIST), dOffsetPar
                            oCurve.Evaluate dOffsetPar, dEndX, dEndY, dEndZ, _
                                    vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
                            oPosition.Set dEndX, dEndY, dEndZ
                        Else

                            oPosition.Set dEndX, dEndY, dEndZ
                        End If
                    Else
                        Set oVector = oEndPos.Subtract(oStartPos)
                        oVector.Length = 1

                        oMarkPos.Set dStartX, dStartY, dStartZ
                        oPosition.Set dStartX, dStartY, dStartZ
                        oCurve.Parameter dStartX, dStartY, dStartZ, dStartPar
                        oCurve.Evaluate dStartPar, dStartX, dStartY, dStartZ, _
                                vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
                        oCurve.Parameter (dStartX + vTanX * ER_OFFSET_DIST), _
                            (dStartY + vTanY * ER_OFFSET_DIST), _
                            (dStartZ + vTanZ * ER_OFFSET_DIST), dOffsetPar
                        oCurve.Evaluate dOffsetPar, dStartX, dStartY, dStartZ, _
                                vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
                        oPosition.Set dStartX, dStartY, dStartZ
                    End If

                    ' If below condition is true, it indicates that mark(for toshi annotation) is already created
                    If (Abs(oClosePos.DistPt(oMarkPos)) < 0.001) And bNewPC = False And bSplitER = True Then
                        GoTo NextMark
                    End If

                    ' Get the normal at the position on surface
                    oPlateLateralSurfaceBody.GetNormalFromPosition oPosition, oNormal

                    ' oNormal can be nothing when input point is not on the surface
                    If oNormal Is Nothing Then
                        Dim oSurfaceModelBody   As IJDModelBody
                        Set oSurfaceModelBody = oPlateLateralSurfaceBody

                        Dim oTempPos   As IJDPosition
                        oSurfaceModelBody.GetMinimumDistanceFromPosition oPosition, oTempPos, dMinDist

                        ' As per the G&T online help, a position is considered on the sheet body if it is within 1e-4 tolerance.
                        If dMinDist > 0.0001 Then
                            Set oPosition = oTempPos
                        End If

                        ' Get the normal at the position on surface
                        oPlateLateralSurfaceBody.GetNormalFromPosition oPosition, oNormal

                        ' If still oNormal is nothing then skip this mark
                        If oNormal Is Nothing Then
                            GoTo NextMark
                        End If
                    End If

                    oNormal.Length = ERFittingMarkLength

                    Set oLine = oLines3d.CreateBy2Points( _
                                Nothing, _
                                oPosition.x, oPosition.y, oPosition.z, _
                                oPosition.x + oNormal.x, _
                                oPosition.y + oNormal.y, _
                                oPosition.z + oNormal.z)

                    oCrvElemets.Add oLine

                    Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

                    oCrvElemets.Clear

                    'Create a SystemMark object to store additional information
                    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

                    'Set the marking side as upside always
                    oSystemMark.SetMarkingSide UpSide

                    'QI for the MarkingInfo object on the SystemMark
                    Set oMarkingInfo = oSystemMark

                    oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness

                    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                    oGeom3d.PutGeometry oCS
                    oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
                    ' If below condition is true, we need toshi annotation so mark direction as centered
                    If (Abs(oClosePos.DistPt(oMarkPos)) < 0.001) And bNewPC = True Then
                        oMarkingInfo.Direction = "centered"
                    ElseIf bCircularER = True Then
                        oMarkingInfo.Direction = "centered"
                    Else
                        'oMarkingInfo.Name = oSDProfileWrapper.Name

                        oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oVector)
                        oMarkingInfo.ThicknessDirection = oVector
                        oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "OverLap"
                        'oGeom3d.PutSubGeometryType STRMFG_FITTING_MARK
                    End If

                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker

                    oSystemMark.Set3dGeometry oGeom3d

                    oGeom3dColl.AddGeometry lGeomCount, oGeom3d
                    lGeomCount = lGeomCount + 1
NextMark:
                Next iLineIdx

            Next nCrvIdx

            'GoTo NextItem
        'End If

        Dim oKnucklePtCol As Collection
        'Set oKnucklePtCol = oSDProfileWrapper.KnucklePoints

        Dim oSGO As GSCADShipGeomOps.SGOWireBodyUtilities
        Set oSGO = New SGOWireBodyUtilities
        Set oKnucklePtCol = oSGO.GetKnucklePoints(oWB)

CheckKnuckleERMarks:
        For nCrvIdx = 1 To nCrvCount

            If oKnucklePtCol.Count < 1 Then GoTo OffSetEdge
            If nCrvIdx = nCrvCount Then GoTo OffSetEdge

            oWireBodyCS.GetCurve nCrvIdx, oCurve
            oWireBodyCS.GetCurve nCrvIdx + 1, oNextCurve

            'If this is a Curve, DO NOT place Knuckle Mark
            If TypeOf oCurve Is IJArc Or TypeOf oNextCurve Is IJArc Then
                GoTo NextERMark
            End If

            If Not TypeOf oCurve Is IJLine Or Not TypeOf oNextCurve Is IJLine Then
                GoTo NextERMark
            End If


            oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
            oNextCurve.EndPoints dStart2X, dStart2Y, dStart2Z, dEnd2X, dEnd2Y, dEnd2Z

            '*** New Implementation ***'
            Dim oEndPtOf1st As New DPosition
            Dim oStartPtOf2nd As New DPosition

            oEndPtOf1st.Set dEndX, dEndY, dEndZ
            oCurve.Parameter dEndX, dEndY, dEndZ, dEndPar
            oCurve.Evaluate dEndPar, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

            Dim oVecAtEnd As New DVector
            Dim oVecAtStart As New DVector
            Dim Markvector As New DVector

            Dim oMarkBasePos As New DPosition

            oVecAtEnd.Set vTanX, vTanY, vTanZ

            oStartPtOf2nd.Set dStart2X, dStart2Y, dStart2Z
            oNextCurve.Parameter dStart2X, dStart2Y, dStart2Z, dEndPar
            oNextCurve.Evaluate dEndPar, dStart2X, dStart2Y, dStart2Z, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

            oVecAtStart.Set -vTanX, -vTanY, -vTanZ

            Markvector.Set (oVecAtEnd.x + oVecAtStart.x) / 2, (oVecAtEnd.y + oVecAtStart.y) / 2, (oVecAtEnd.z + oVecAtStart.z) / 2

            oMarkBasePos.Set dEndX, dEndY, dEndZ

            Markvector.Length = ERFittingMarkLength


'                .--------.                  .--------.
'                |         \                /         |
'                |          \   Fitting    /          |
'                |           \   Marks    /           |
'                |            \/________\/            |
'                |            /          \            |
'                |                                    |
'                |                                    |




            'Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x, oMarkBasePos.y, oMarkBasePos.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)
            Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x - Markvector.x, oMarkBasePos.y - Markvector.y, oMarkBasePos.z - Markvector.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)

            '**************************'
            oCrvElemets.Add oLine

            Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)
            oCrvElemets.Clear

            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

            'Set the marking side
            'oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
            'Set the marking side as upside always
            oSystemMark.SetMarkingSide UpSide

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            'oMarkingInfo.Name = oSDProfileWrapper.Name
            oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness

            oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oVector)
            oMarkingInfo.ThicknessDirection = oVector

            Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK

            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            oSystemMark.Set3dGeometry oGeom3d

            oGeom3dColl.AddGeometry lGeomCount, oGeom3d
            lGeomCount = lGeomCount + 1
NextERMark:
        Next nCrvIdx
OffSetEdge:
        '*** This code is for Offset Edge Mark ***'
        If bOffsetEdge = True Then

            'Get the PC Curve
            Dim oPCPartCurve As IJCurve
            oWireBodyCS.GetCurve nCrvIdx, oCurve
            Set oPCPartCurve = oCurve
            oCurve.Evaluate 0.5, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
            oCurve.Centroid dX, dY, dZ
            oCentroidPt.Set dX, dY, dZ
'                    ____________________
'                    |                   |
'                    |                   |
'                    |     This is     |_|
'                    |     OffsetEdge  | |
'                    |     Mark          |
'                    |                   |
'                    |___________________|


            '*** Find contour close to PC ***'
            Dim i As Long
            Dim j As Long
            Dim dMinimum As Double
            Dim iMinIndex As Integer

            dMinimum = 999
            For i = 1 To oGeomCol3dContour.GetCount
                'Dim oGeom3d As IJMfgGeom3D
                Set oGeom3d = oGeomCol3dContour.GetGeometry(i)
                eGeomType = oGeom3d.GetGeometryType

                If eGeomType = STRMFG_OUTER_CONTOUR Then
                    Dim oMB1 As IJDModelBody
                    Dim oCString As IComplexStrings3d
                    Dim oWB123 As IJWireBody

                    Set oWB123 = m_oMfgRuleHelper.ComplexStringToWireBody(oGeom3d.GetGeometry)
                    Set oMB1 = oWB123

                    oMB1.GetMinimumDistanceFromPosition oCentroidPt, oPos1, dMinDist

                    If dMinDist < dMinimum Then
                        Set oBasePos = oPos1
                        dMinimum = dMinDist
                        iMinIndex = i
                    End If
                End If
            Next

            Set oGeom3d = oGeomCol3dContour.GetGeometry(iMinIndex)
            'This is PC curve/edge

            Set oCurve = oGeom3d.GetGeometry
            '*** Find contour close to PC ***'

            Dim dParam As Double, dPCParam As Double
            oCurve.Parameter oPCMidPt.x, oPCMidPt.y, oPCMidPt.z, dPCParam
            oCurve.Evaluate dPCParam, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z
            'oCurve.Centroid dEndX, dEndY, dEndZ

            Dim oTanVec As New DVector, oNormalVec As New DVector

            oTanVec.Set vTanX, vTanY, vTanZ
            oTanVec.Length = 0.025

            Set oNormalVec = oPlateNormal.Cross(oTanVec)
            oNormalVec.Length = 0.025

            'Set oCurve = oPCPartCurve
            'oPCPartCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
            'oPCPartCurve.Evaluate 0.5, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

            'oPCPartCurve.Centroid dEndX, dEndY, dEndZ
            ' This creates line between Curve-MidPoint and Point along with Vector on Surface (normal to tangent)
            'Set oLine = oLines3d.CreateBy2Points(Nothing, dEndX - oNormalVec.x, dEndY - oNormalVec.y, dEndZ - oNormalVec.z, dEndX + oNormalVec.x, dEndY + oNormalVec.y, dEndZ + oNormalVec.z)
            'oNormalVec.Length = -1#
            Set oLine = oLines3d.CreateBy2Points(Nothing, dEndX, dEndY, dEndZ, dEndX + oNormalVec.x, dEndY + oNormalVec.y, dEndZ + oNormalVec.z)

            ' This creates line parallel to tangent vector
'                    Set oLine = oLines3d.CreateBy2Points(Nothing, dEndX - oNormalVec.x - oTanVec.x, dEndY - oNormalVec.y - oTanVec.y, dEndZ - oNormalVec.z - oTanVec.z, dEndX - oNormalVec.x + oTanVec.x, dEndY - oNormalVec.y + oTanVec.y, dEndZ - oNormalVec.z + oTanVec.z)
'                    oCrvElemets.Add oLine
'                    Set oLine = oLines3d.CreateBy2Points(Nothing, dEndX + oNormalVec.x - oTanVec.x, dEndY + oNormalVec.y - oTanVec.y, dEndZ + oNormalVec.z - oTanVec.z, dEndX + oNormalVec.x + oTanVec.x, dEndY + oNormalVec.y + oTanVec.y, dEndZ + oNormalVec.z + oTanVec.z)
            oCrvElemets.Add oLine

            Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

            oMfgMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

            oProjCS.GetCurve 1, oCurve
            oCurve.Centroid dEndX, dEndY, dEndZ
            Set oLine = oLines3d.CreateBy2Points(Nothing, dEndX - oTanVec.x, dEndY - oTanVec.y, dEndZ - oTanVec.z, dEndX + oTanVec.x, dEndY + oTanVec.y, dEndZ + oTanVec.z)

            oCrvElemets.Clear
            oCrvElemets.Add oLine
            Set oProjCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

            oCrvElemets.Clear
            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

            'Set the marking side
            oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            oMarkingInfo.Name = oSDProfileWrapper.Name & "(" & OFFSET_MARK_PREFIX & sOrient & ")"
            oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness

            oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oNormalVec)
            oMarkingInfo.ThicknessDirection = oNormalVec

            Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
            oGeom3d.PutGeometry oProjCS
            oGeom3d.PutGeometrytype STRMFG_CONN_PART_MARK
            oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "OffsetEdge"
            'oGeom3d.PutSubGeometryType STRMFG_CONN_PART_MARK
            oGeom3d.IsSupportOnly = True


            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            oSystemMark.Set3dGeometry oGeom3d

            oGeom3dColl.AddGeometry lGeomCount, oGeom3d
            lGeomCount = lGeomCount + 1

        End If ' If offset edge

NextItem:
        Set oWB = Nothing
        Set oTeeWire = Nothing
        Set oProfilePartSupport = Nothing
        Set oPartSupport = Nothing

    Next iIndex


CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oSDPhysicalConn = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetMarginExtensionDistances(oGeomCol3Dextra As IJMfgGeomCol3d, oCS As IJComplexString, pResMgr As IUnknown, dStartMargin As Double, dEndMargin As Double)
Const METHOD = "ProfileLocationMark: GetMarginExtensionDistance"
    On Error GoTo ErrorHandler

    Dim dTotalMargin As Double

    Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
    Dim i As Long, j As Long
    Dim oGeom3d As IJMfgGeom3D
    Dim eGeomType As StrMfgGeometryType
    Dim oMoniker As IMoniker
    Dim oObject As IUnknown
    Dim oPort As IJPort
    Dim oMfgDefCol As Collection
    Dim oConstMargin As IJConstMargin
    Dim oObliqueMargin As IJObliqueMargin
    Dim oProfileLocationCurve As IJCurve
    Dim oOuterContourCurve As IJCurve
    Dim dMinDist As Double, profLocX As Double, profLocY As Double, profLocZ As Double
    Dim OuterCntX As Double, OuterCntY As Double, OuterCntZ As Double
    Dim dPar As Double, dStartParam As Double, dEndParam As Double

    dStartMargin = 0
    dEndMargin = 0

    lFabMargin = 0
    lAssyMargin = 0
    lCustomMargin = 0

    Set oProfileLocationCurve = oCS

    If Not oGeomCol3Dextra Is Nothing Then
    ' We found manufacturing information that could be margin values
        For i = 1 To oGeomCol3Dextra.GetCount
            Set oGeom3d = oGeomCol3Dextra.GetGeometry(i)
            eGeomType = oGeom3d.GetGeometryType

            If eGeomType = STRMFG_OUTER_CONTOUR Then
                Set oMoniker = oGeom3d.GetMoniker
                If oMoniker Is Nothing Then GoTo NextMoniker

                Set oObject = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)

                If Not TypeOf oObject Is IJPort Then GoTo NextMoniker

                Set oOuterContourCurve = oGeom3d.GetGeometry

                oProfileLocationCurve.DistanceBetween oOuterContourCurve, dMinDist, _
                        profLocX, profLocY, profLocZ, OuterCntX, OuterCntY, OuterCntZ

                If dMinDist > 0.01 Then GoTo NextMoniker

                oProfileLocationCurve.Parameter profLocX, profLocY, profLocZ, dPar
                oProfileLocationCurve.ParamRange dStartParam, dEndParam

                Set oPort = oObject
                Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)

                If oMfgDefCol.Count > 0 Then

                    For j = 1 To oMfgDefCol.Count
                        If TypeOf oMfgDefCol.Item(j) Is IJDFabMargin Then
                            Dim oFabMargin As IJDFabMargin
                            Set oFabMargin = oMfgDefCol.Item(j)

                            If oFabMargin.GeometryChange = AsMargin Then ''1 = As Margin, 2 = As Shrinkage, 3 = As Reference
                                If TypeOf oMfgDefCol.Item(j) Is IJAssyMarginChild Then
                                    Set oConstMargin = oMfgDefCol.Item(j)
                                    lAssyMargin = lAssyMargin + oConstMargin.Value
                                ElseIf TypeOf oMfgDefCol.Item(j) Is IJObliqueMargin Then
                                    Set oObliqueMargin = oMfgDefCol.Item(j)
                                    If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                                        lFabMargin = lFabMargin + oObliqueMargin.EndValue
                                    Else
                                        lFabMargin = lFabMargin + oObliqueMargin.StartValue
                                    End If
                                ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
                                    Set oConstMargin = oMfgDefCol.Item(j)
                                    lFabMargin = lFabMargin + oConstMargin.Value
                                End If
                            End If
                        End If
                    Next j

                    If ((dPar - dStartParam) < (dEndParam - dPar)) Then
                        dStartMargin = dStartMargin + lAssyMargin + lFabMargin + lCustomMargin
                    Else
                        dEndMargin = dEndMargin + lAssyMargin + lFabMargin + lCustomMargin
                    End If

                End If
            End If

NextMoniker:
            Set oOuterContourCurve = Nothing
            Set oMoniker = Nothing
            Set oObject = Nothing
            Set oPort = Nothing
            Set oMfgDefCol = Nothing
            Set oConstMargin = Nothing
            Set oObliqueMargin = Nothing
            Set oGeom3d = Nothing

        Next i
    End If

CleanUp:
    Set oOuterContourCurve = Nothing
    Set oProfileLocationCurve = Nothing
    Set oMoniker = Nothing
    Set oObject = Nothing
    Set oPort = Nothing
    Set oMfgDefCol = Nothing
    Set oConstMargin = Nothing
    Set oObliqueMargin = Nothing
    Set oGeom3d = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Public Function GetFlangeOrientation(enumStructPlateType As StructPlateType, enumStructProfileType As StructProfileType, enumStructMoldedOrientation As StructMoldedOrientation) As String
''Enum StructMoldedOrientation
''{
''    [helpstring("oriented to port")] PortOrient = 0,
''    [helpstring("starboard")] StarboardOrient = 1,
''
''    [helpstring("fore")] ForeOrient = 0,
''    [helpstring("aft")] AftOrient = 1,
''
''    [helpstring("above")] AboveOrient = 0,
''    [helpstring("below")] BelowOrient = 1,
''
''    [helpstring("In")]    InOrient = 0,
''    [helpstring("Out")]   OutOrient = 1,
''
''    [helpstring("Up")]    UpOrient = 0,
''    [helpstring("Down")]  DownOrient = 1,
''
''    [helpstring("Clockwise")]        ClockwiseOrient = 0,
''    [helpstring("CounterClockwise")]  CClockwiseOrient = 1,
''
''    [helpstring("inboard")] InboardOrient = 3,
''    [helpstring("outboard")] OutboardOrient = 4,
''
''    [helpstring("invalid profile type")] invalidOrientation = 0x2
''} StructMoldedOrientation;
''"fore"
''"port"
''"upper"
''"aft"
''"starboard"
''"lower"
    Const METHOD = ":GetFlangeOrientation"
    On Error GoTo ErrorHandler

    Select Case enumStructPlateType

        Case DeckPlate:
            If enumStructProfileType = sptLongitudinal Then
                If enumStructMoldedOrientation = PortOrient Then
                    GetFlangeOrientation = "port"
                ElseIf enumStructMoldedOrientation = StarboardOrient Then
                    GetFlangeOrientation = "starboard"
                ElseIf enumStructMoldedOrientation = InboardOrient Then
                    GetFlangeOrientation = "in"
                ElseIf enumStructMoldedOrientation = OutboardOrient Then
                    GetFlangeOrientation = "out"
                End If
            ElseIf enumStructProfileType = sptTransversal Then
                If enumStructMoldedOrientation = ForeOrient Then
                    GetFlangeOrientation = "fore"
                ElseIf enumStructMoldedOrientation = AftOrient Then
                    GetFlangeOrientation = "aft"
                End If
            ElseIf enumStructProfileType = sptEdgeReinforcement Then
                If enumStructMoldedOrientation = AboveOrient Then
                    GetFlangeOrientation = "above"
                ElseIf enumStructMoldedOrientation = BelowOrient Then
                    GetFlangeOrientation = "below"
                End If
            End If

        Case LBulkheadPlate:
            If enumStructProfileType = sptVertical Then
                If enumStructMoldedOrientation = ForeOrient Then
                    GetFlangeOrientation = "fore"
                ElseIf enumStructMoldedOrientation = AftOrient Then
                    GetFlangeOrientation = "aft"
                End If
            ElseIf enumStructProfileType = sptLongitudinal Then
                If enumStructMoldedOrientation = UpOrient Then
                    GetFlangeOrientation = "upper"
                ElseIf enumStructMoldedOrientation = DownOrient Then
                    GetFlangeOrientation = "lower"
                End If
            ElseIf enumStructProfileType = sptEdgeReinforcement Then
                If enumStructMoldedOrientation = PortOrient Then
                    GetFlangeOrientation = "port"
                ElseIf enumStructMoldedOrientation = StarboardOrient Then
                    GetFlangeOrientation = "starboard"
                ElseIf enumStructMoldedOrientation = InboardOrient Then
                    GetFlangeOrientation = "inboard"
                ElseIf enumStructMoldedOrientation = OutboardOrient Then
                    GetFlangeOrientation = "outboard"
                End If
            End If

        Case TBulkheadPlate:
            If enumStructProfileType = sptVertical Then
                If enumStructMoldedOrientation = PortOrient Then
                    GetFlangeOrientation = "port"
                ElseIf enumStructMoldedOrientation = StarboardOrient Then
                    GetFlangeOrientation = "starboard"
                ElseIf enumStructMoldedOrientation = InboardOrient Then
                    GetFlangeOrientation = "in"
                ElseIf enumStructMoldedOrientation = OutboardOrient Then
                    GetFlangeOrientation = "out"
                End If
            ElseIf enumStructProfileType = sptTransversal Then
                If enumStructMoldedOrientation = UpOrient Then
                    GetFlangeOrientation = "upper"
                ElseIf enumStructMoldedOrientation = DownOrient Then
                    GetFlangeOrientation = "lower"
                End If
            ElseIf enumStructProfileType = sptEdgeReinforcement Then
                If enumStructMoldedOrientation = ForeOrient Then
                    GetFlangeOrientation = "fore"
                ElseIf enumStructMoldedOrientation = AftOrient Then
                    GetFlangeOrientation = "aft"
                End If
            End If

        Case Hull:
          If enumStructProfileType = sptLongitudinal Or sptVertical Then
               If enumStructMoldedOrientation = PortOrient Then
                  GetFlangeOrientation = "port"
              ElseIf enumStructMoldedOrientation = StarboardOrient Then
                  GetFlangeOrientation = "starboard"
              ElseIf enumStructMoldedOrientation = InboardOrient Then
                  GetFlangeOrientation = "in"
              ElseIf enumStructMoldedOrientation = OutboardOrient Then
                  GetFlangeOrientation = "out"
              End If
          Else    'sptTransverse
              If enumStructMoldedOrientation = AboveOrient Then
                  GetFlangeOrientation = "upper"
              ElseIf enumStructMoldedOrientation = BelowOrient Then
                  GetFlangeOrientation = "lower"
              End If
          End If

        Case LongitudinalTube:
        Case TransverseTube:
        Case VerticalTube:

              If enumStructMoldedOrientation = PortOrient Then
                  GetFlangeOrientation = "port"
              ElseIf enumStructMoldedOrientation = StarboardOrient Then
                  GetFlangeOrientation = "starboard"
              ElseIf enumStructMoldedOrientation = AboveDir Then
                  GetFlangeOrientation = "upper"
              ElseIf enumStructMoldedOrientation = BelowDir Then
                  GetFlangeOrientation = "lower"
              ElseIf enumStructMoldedOrientation = ForeOrient Then
                  GetFlangeOrientation = "fore"
              ElseIf enumStructMoldedOrientation = AftOrient Then
                  GetFlangeOrientation = "aft"
              ElseIf enumStructMoldedOrientation = InOrient Or InboardDir Then
                  GetFlangeOrientation = "in"
              ElseIf enumStructMoldedOrientation = OutOrient Or OutboardDir Then
                  GetFlangeOrientation = "out"
              Else
                  GetFlangeOrientation = "Other"
              End If
    End Select

CleanUp:

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetFlangeDirection(strThickDir As String, bIsOpposite As Boolean) As String
    Const METHOD = ":GetFlangeDirection"
    On Error GoTo ErrorHandler

    If bIsOpposite Then
        GetFlangeDirection = GetOpposite(strThickDir)
    Else
        GetFlangeDirection = strThickDir
    End If

CleanUp:

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetOpposite(strThickDir As String) As String
    Const METHOD = ":GetOpposite"
    On Error GoTo ErrorHandler

    If strThickDir = "port" Then GetOpposite = "starboard"
    If strThickDir = "starboard" Then GetOpposite = "port"
    If strThickDir = "in" Then GetOpposite = "out"
    If strThickDir = "out" Then GetOpposite = "in"
    If strThickDir = "upper" Then GetOpposite = "lower"
    If strThickDir = "lower" Then GetOpposite = "upper"
    If strThickDir = "fore" Then GetOpposite = "aft"
    If strThickDir = "aft" Then GetOpposite = "fore"

CleanUp:

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetDeclivityDirection(oPlatePartObj As Object, oProfilePartObj As Object, oSDPhysicalConn As StructDetailObjects.PhysicalConn, oContour_TeeGeom As Object) As IJDVector
Const METHOD = "ProfileLocationMark: GetDeclivityDirection"
    On Error GoTo ErrorHandler

    Dim oLandingCrvWireBody As IJWireBody
    Dim oTopologyLocate As New TopologyLocate
    Dim oMarkingMidPos As DPosition
    Dim oProfilePoint As IJDPosition
    Dim oProfileVector As IJDVector
    Dim oCSMatrix As IJDT4x4
    Dim mountingFaceSectionID As JXSEC_CODE
    Dim iIndex As Integer
    Dim dCSUDirX As Double, dCSUDirY As Double, dCSUDirZ As Double, dDotP As Double
    Dim oCSUDir As New DVector
    Dim oDeclivityDir As New DVector

    ' Get the Profile System's landing curve
    Set oLandingCrvWireBody = oTopologyLocate.GetProfileParentWireBody(oProfilePartObj)

    ' Get the mid point of the contour_tee marking line
    Set oMarkingMidPos = m_oMfgRuleHelper.GetMiddlePoint(oContour_TeeGeom)

    ' Get the tangent of landing curve at the above mid position
    oTopologyLocate.GetProjectedPointOnModelBody oLandingCrvWireBody, oMarkingMidPos, oProfilePoint, oProfileVector

    ' Get the profile rientation matrix at the mid position
    Set oCSMatrix = oTopologyLocate.GetPenetratingCrossSectionMatrix(oProfilePartObj, oMarkingMidPos)

    ' Get the mounting face's section ID
    Dim oProfileSection As IJDProfileSection

    Set oProfileSection = oProfilePartObj

    mountingFaceSectionID = oProfileSection.mountingFaceSectionID

    ' if Bounded Profile Mounting Face ID is JXSEC_TOP_xxx or JXSEC_BOTTOM_xxx
    ' adjust the Bounded direction vector to be the Cross Section's U vector
    ' else use the Cross Section's V vector as the Bounded Direction vector
    If (mountingFaceSectionID = JXSEC_TOP Or _
         mountingFaceSectionID = JXSEC_TOP_FLANGE_LEFT_TOP Or _
         mountingFaceSectionID = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
         mountingFaceSectionID = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
         mountingFaceSectionID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
         mountingFaceSectionID = JXSEC_BOTTOM Or _
         mountingFaceSectionID = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
         mountingFaceSectionID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
         mountingFaceSectionID = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
         mountingFaceSectionID = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM) Then
        iIndex = 0
    Else
        iIndex = 4
    End If

    ' Get the U Vector
    oCSUDir.x = oCSMatrix.IndexValue(iIndex)
    oCSUDir.y = oCSMatrix.IndexValue(iIndex + 1)
    oCSUDir.z = oCSMatrix.IndexValue(iIndex + 2)

     ' Adjust the vector based on profile direction.
    dDotP = oCSUDir.Dot(oProfileVector)

    If (dDotP < -0.000001) Then
        oDeclivityDir.x = -1# * oCSUDir.x
        oDeclivityDir.y = -1# * oCSUDir.y
        oDeclivityDir.z = -1# * oCSUDir.z
    Else
        oDeclivityDir.x = oCSUDir.x
        oDeclivityDir.y = oCSUDir.y
        oDeclivityDir.z = oCSUDir.z
    End If

    oDeclivityDir.Length = -1#

    Set GetDeclivityDirection = oDeclivityDir

CleanUp:
    Set oLandingCrvWireBody = Nothing
    Set oTopologyLocate = Nothing
    Set oMarkingMidPos = Nothing
    Set oProfilePoint = Nothing
    Set oProfileVector = Nothing
    Set oCSMatrix = Nothing
    Set oCSUDir = Nothing
    Set oDeclivityDir = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function CreatePlateLocationMark(ByVal Part As Object, ByVal UpSide As Long, _
                                        ByVal bSelectiveRecompute As Boolean, _
                                        ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                        Optional ByVal bAddERMarks As Boolean, _
                                        Optional ByVal bAddFittingMark As Boolean, _
                                        Optional ByVal bAddRENDMarks As Boolean, _
                                        Optional ByVal bAddDeclivityMark As Boolean, _
                                        Optional ByVal CONN_PART_CONDITION As Long = 0, _
                                        Optional ByVal bDeclMarkOnUpside As Boolean = False, _
                                        Optional ByVal bMarkCenteredPlate As Boolean = True) As IJMfgGeomCol3d
    Const METHOD = "CreatePlateLocationMark"

    Dim eSideOfConnectedObjectToBeMarked As ThicknessSide

    On Error Resume Next

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oPartInfo As IJDPartInfo

    Set oPartInfo = New PartInfo

    'Get the Plate Part Physically Connected Objects
    Dim oCS As IJComplexString
    Dim oCSColl As IJElements
    Dim oConObjsCol As Collection
    'Set oConObjsCol = oSDPlateWrapper.ConnectedObjects
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_PLATELOCATION_MARK, ReferenceObjColl, oGeomCol3d
    CreateAPSMarkings STRMFG_MOUNT_ANGLE_MARK, ReferenceObjColl, oGeomCol3d

    Set CreatePlateLocationMark = oGeomCol3d

    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the function
        GoTo CleanUp
    End If

    Dim ArrayWithFlipBits() As Integer
    'ReDim'ed for Count+1 because 0th element is always set to false. The array is actually set from 1st element.
    ReDim ArrayWithFlipBits(oConObjsCol.Count + 1) As Integer
    ArrayWithFlipBits = CheckIfFlipMarkingSideNeeded(Part, oConObjsCol)

    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex As Long
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
        End If
    Next nIndex

    Dim Item As Object
    Dim oConnectionData As ConnectionData

    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn

    ' Loop thru each Physical Connections
    Dim bContourTee As Boolean
    Dim oVector As IJDVector
    Dim oWB As IJWireBody, oTeeWire As IJWireBody
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    lGeomCount = 1

    Dim oSupp2 As IJPartSupport
    Dim oSys As IJSystem
    Dim oPlateUtils As IJPlateAttributes
    Dim eSurfGeomType As SurGeoType

    Dim oNamedItem As IJNamedItem

    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)

        '*** GOTO NextItem if it is a BRACKET PLATE ***'
        Set oSupp2 = New PartSupport
        Set oSupp2.Part = oConnectionData.ToConnectable
        oSupp2.IsSystemDerivedPart oSys, True

        Set oNamedItem = oSys

        Set oPlateUtils = New PlateUtils
        If Not oSys Is Nothing Then
            If oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys) Then
                GoTo NextItem
            End If
        End If

        Set oPlateUtils = Nothing
        Set oSupp2 = Nothing
        '**********************************************'

        ' As both Bracket and Collar are implementing IJPlatePart we need to check if the
        ' conneected item is not either of those
        If TypeOf oConnectionData.ToConnectable Is IJSmartPlate Then
             GoTo NextItem
        End If
        If TypeOf oConnectionData.ToConnectable Is IJCollarPart Then
             GoTo NextItem
        End If
        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
             GoTo NextItem
        End If

        ' Check if this physical connection is a Root PC, which particpates in Split operation
        ' If so, Skip the marking line line creation for this Root PC.
        Dim oStructEntityOperation As IJDStructEntityOperation
        Dim opeartionProgID As String
        Dim opeartionID As StructOperation
        Dim oOperColl As New Collection

        Set oStructEntityOperation = oConnectionData.AppConnection
        oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl

        Set oStructEntityOperation = Nothing
        Set oOperColl = Nothing

        ' If the RootPC has Split operation in its graph, just goto the next pc.
        If opeartionID = ConnectionSplitOperation Then
            ' No need to create marking line.
            GoTo NextItem
        End If

        'Initialize the profile wrapper and the Physical Connection wrapper
        Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDConPlateWrapper.object = oConnectionData.ToConnectable

        Dim oConPlateSupp As IJPartSupport

        Set oConPlateSupp = New GSCADSDPartSupport.PlatePartSupport
        Set oConPlateSupp.Part = oSDConPlateWrapper.object

        Dim eMoldedDir As StructMoldedDirection

        eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(oConnectionData.ToConnectable)
        Dim sMoldedSide As String
        sMoldedSide = oSDConPlateWrapper.MoldedSide
        If eMoldedDir = Centered And bMarkCenteredPlate = True Then
            eSideOfConnectedObjectToBeMarked = SideUnspecified
        Else
            ' Fix for TR#39584
            ' Earlier we were filling the side to be marked on the
            ' connected plate based on the MoldedSide of the plate
            ' that is being manufactured which is wrong.

            ' Get the moulded side of the connected plate
            If sMoldedSide = "Base" Then
                eSideOfConnectedObjectToBeMarked = SideA
            ElseIf sMoldedSide = "Offset" Then
                eSideOfConnectedObjectToBeMarked = SideB
            Else
                Dim lErrNumber As Long
                lErrNumber = LogMessage(Err, MODULE, METHOD, " Unexpected Molded side ")
            End If
        End If


        '****** Flip the marking side as per ArrayWithFlipBits *******'
        If UBound(ArrayWithFlipBits) > 0 Then
            If eSideOfConnectedObjectToBeMarked = SideA And ArrayWithFlipBits(nIndex) = 2 Then
                eSideOfConnectedObjectToBeMarked = SideB
            ElseIf eSideOfConnectedObjectToBeMarked = SideB And ArrayWithFlipBits(nIndex) = 2 Then
                eSideOfConnectedObjectToBeMarked = SideA
            End If
        End If
        '*************************************************************'




        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)

        If bContourTee = True Then
            ' Bound the wire based on split points, if there are any.
            Set oWB = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire)

            Dim oVectorBase As IJDVector
            Dim oVector1 As IJDVector
            Dim oStart As IJDPosition, oEnd As IJDPosition, oBaseStart As IJDPosition
            Dim oDirVector As IJDVector
            Dim oCross As IJDVector, oPlateNormal As IJDVector
            Dim dDotProduct As Double
            Dim oSurfaceBody As IJSurfaceBody
            Dim oSurfaceBodyCon As IJSurfaceBody

            oWB.GetEndPoints oStart, oEnd

            Set oSurfaceBodyCon = oSDConPlateWrapper.BasePort(BPT_Base).Geometry
            Set oBaseStart = m_oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBodyCon, oVectorBase)
            Set oSurfaceBody = oSDPlateWrapper.BasePort(BPT_Base).Geometry

            Set oBaseStart = m_oMfgRuleHelper.ProjectPointOnSurface(oBaseStart, oSurfaceBody, oVector1)
            Set oStart = m_oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBody, oVector1)
            oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal

            'Convert the IJWireBody to ComplexStrings
            Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWB)

            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If

            If (oCSColl Is Nothing) Then
                GoTo NextItem
            End If

            For Each oCS In oCSColl
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                'Set the marking side
                Dim lMarkingSide As Long
                lMarkingSide = oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
                oSystemMark.SetMarkingSide lMarkingSide

                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

                ' Set the production name of connected part
                Dim strProdName     As String
                strProdName = GetManufacturingProductionName(Part, oConnectionData.ToConnectable, CONN_PART_CONDITION)
                oMarkingInfo.Name = strProdName

                oMarkingInfo.Thickness = oSDConPlateWrapper.PlateThickness

                If eMoldedDir = Centered And bMarkCenteredPlate = True Then
                    oMarkingInfo.Direction = "centered"
                Else
                    oMarkingInfo.Direction = GetPlateThicknessDir(oConnectionData.ToConnectable)

                    oMarkingInfo.ThicknessDirection = GetThicknessDirectionVector(oCS, oSDConPlateWrapper, sMoldedSide)

                    'If direction returned is Centered, reset it.
                    ' We only want center direction if <eMoldedDir = Centered And bMarkCenteredPlate = True>
                    ' ... which is the condition above
                    If eMoldedDir = Centered Then
                        oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oMarkingInfo.ThicknessDirection)
                    End If

                End If

                Dim dFittingAngle As Double

                dFittingAngle = oSDPhysicalConn.TeeMountingAngle
                If eMoldedDir = Centered And bMarkCenteredPlate = True Then
                    oVectorBase.Set oBaseStart.x - oStart.x, oBaseStart.y - oStart.y, oBaseStart.z - oStart.z
                    oMarkingInfo.ThicknessDirection = oVectorBase
                Else
                If oSDConPlateWrapper.MoldedSide = "Base" Then
                    dFittingAngle = PI - dFittingAngle
                End If
                End If

                oMarkingInfo.FittingAngle = dFittingAngle

                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_PLATELOCATION_MARK

                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
                oSystemMark.Set3dGeometry oGeom3d

                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1

                ' Check if declivity marks are needed
                If bAddDeclivityMark = True Then
                    If bDeclMarkOnUpside Then
                        ' Call like so if you want declivity marks always on upside
                        CreateDeclivityMarks Part, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, UpSide, sMoldedSide, Nothing
                    Else
                        ' Call like so if you want declivity marks on same side as its location mark
                        CreateDeclivityMarks Part, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, lMarkingSide, sMoldedSide, Nothing
                    End If
                End If

                '*** CREATE END FITTING MARK ***'
                'This function Creates the End Fitting Mark(s) at the Free End(s) of the
                '... Profile Parts.
                If bAddFittingMark Then

                    'Arg #4: oSDConPlateWrapper - Connected Object's GSCADSDPartSupport.ProfilePartSupport/PlatePartSupport
                    'Arg #5: oSDPlateWrapper - Base Part's StructDetailObjects.PlatePart/ProfilePart
                    'Arg #6: oSDConPlateWrapper - Connected Object's StructDetailObjects.ProfilePart/ PlatePart


                    CreateEndFittingMark Part, oCS, oWB, oConPlateSupp, oSDPlateWrapper, _
                             oSDConPlateWrapper, sMoldedSide, lMarkingSide, oConnectionData, _
                             oGeomCol3d, oPlateNormal, False
                End If
                '*******************************'

            Next

            Set oVectorBase = Nothing
            Set oVector1 = Nothing
            Set oBaseStart = Nothing
            Set oStart = Nothing
            Set oEnd = Nothing
            Set oDirVector = Nothing
            Set oCross = Nothing
            '                    Set oPlateNormal = Nothing
            Set oSurfaceBody = Nothing
            Set oSurfaceBodyCon = Nothing

            If Not oCSColl Is Nothing Then
                oCSColl.Clear
            End If
            Set oCSColl = Nothing
        End If
NextItem:

        Set oWB = Nothing
        Set oVector = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConPlateWrapper = Nothing
        Set oSDPhysicalConn = Nothing
        If Not oCSColl Is Nothing Then
            oCSColl.Clear
        End If
        Set oCSColl = Nothing
    Next nIndex

    'Return the 3d collection
    Set CreatePlateLocationMark = oGeomCol3d

CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oCS = Nothing
    Set oPartInfo = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES")
    GoTo CleanUp
End Function

' ***********************************************************************************
' Function: CheckIfFlipMarkingSideNeeded
'
' Description:  This function takes plate part and collection of geom2ds (Plate Location Marks)
'               and if the marked side is discountinous (due to different part thickness), checks the other side for a continous mark.
'               In case other side has a single continous mark, it marks the other side. If both sides have
'               discontinous marks, it marks the side with least difference

' Inputs    :   [Part] Input Plate Part, [oConObjsCol] Collection of Geom2ds (Plate Location Marks)
'
' ***********************************************************************************

Public Function CheckIfFlipMarkingSideNeeded(Part As Object, oConObjsCol As Collection) As Integer()

    Const METHOD = "CheckIfFlipMarkingSideNeeded"
    On Error GoTo ErrorHandler

    Dim dnt As Integer
    Dim ArrayWithFlipBits() As Integer
    Dim oSys As IJSystem, oConnSys As IJSystem
    Dim oIJPartSupp As IJPartSupport

    'Get SDPlate Wrapper for base plate part
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    ReDim ArrayWithFlipBits(oConObjsCol.Count + 1) As Integer
    '0 : Not Processed Yet
    '1 : Do not Flip
    '2 : Flip


    'Get Side to mark for this part
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oConnectionData As ConnectionData, oConnectionData2 As ConnectionData
    Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
    Dim bContourTee As Boolean

    'We dont process 0th element. So make it 1
    ArrayWithFlipBits(0) = 1

    Dim i As Integer, j As Integer
    For i = 1 To oConObjsCol.Count
        If ArrayWithFlipBits(i) = 0 Then
            oConnectionData = oConObjsCol.Item(i)
            If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
                GoTo NextI
            End If
            Set oSDConPlateWrapper.object = oConnectionData.ToConnectable

            Set oIJPartSupp = New PartSupport
            'oIJPartSupp.Part = oConnectionData

            'Get parent of current part
            Set oIJPartSupp.Part = oConnectionData.ToConnectable
            oIJPartSupp.IsSystemDerivedPart oSys, True

            If oSys Is Nothing Then GoTo NextI

            Dim oTeeWire As IJWireBody
            Dim oVector As IJDVector

            'Get current marked side
            eSideOfConnectedObjectToBeMarked = GetSideOfConnectedObjectToBeMarked(oConnectionData.ToConnectable)

            'Get Current TeeWire
            bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)

            If bContourTee = False Then
                GoTo NextI
            End If

            Dim TeeWireElemColl As IJElements
            Set TeeWireElemColl = New JObjectCollection
            TeeWireElemColl.Add oTeeWire

            'Get the teewire on other side
            If eSideOfConnectedObjectToBeMarked = SideA Then
                eSideOfConnectedObjectToBeMarked = SideB
            ElseIf eSideOfConnectedObjectToBeMarked = SideB Then
                eSideOfConnectedObjectToBeMarked = SideA
            End If

            bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)
            If bContourTee = False Then
                GoTo NextI
            End If

            Dim TeeWireElemColl2 As IJElements
            Set TeeWireElemColl2 = New JObjectCollection
            TeeWireElemColl2.Add oTeeWire

            Dim ArrayOfJ() As Integer
            ReDim ArrayOfJ(0 To oConObjsCol.Count) As Integer
            Dim k As Integer

            For j = 1 To oConObjsCol.Count

                'Just process one's which is not current and subsequent to the current part
                If j <> i And j > i Then
                    oConnectionData2 = oConObjsCol.Item(j)
                    If Not TypeOf oConnectionData2.ToConnectable Is IJPlatePart Then
                        GoTo NextJ
                    End If

                    Set oSDConPlateWrapper.object = oConnectionData2.ToConnectable
                    Set oIJPartSupp = New PartSupport
                    'oIJPartSupp.Part = oConnectionData2

                    'Get parent of next part
                    Set oIJPartSupp.Part = oConnectionData2.ToConnectable
                    oIJPartSupp.IsSystemDerivedPart oConnSys, True

                    If oConnSys Is Nothing Then
                        GoTo NextJ
                    End If

                    'If j's parent is i's parent then continue, else goto next j
                    If Not oSys Is oConnSys Then GoTo NextJ

                    Dim oTeeWire2 As IJWireBody
                    Dim oVector2 As IJDVector

                    'Get current marked side
                    eSideOfConnectedObjectToBeMarked = GetSideOfConnectedObjectToBeMarked(oConnectionData2.ToConnectable)

                    'Get Current TeeWire
                    bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData2.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire2, oVector2)
                    If bContourTee = False Then
                        GoTo NextJ
                    End If

                    'Add to current collection
                    TeeWireElemColl.Add oTeeWire2

                    'Change the marked side
                    If eSideOfConnectedObjectToBeMarked = SideA Then
                        eSideOfConnectedObjectToBeMarked = SideB
                    ElseIf eSideOfConnectedObjectToBeMarked = SideB Then
                        eSideOfConnectedObjectToBeMarked = SideA
                    End If

                    'Get TeeWire on other side
                    bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData2.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire2, oVector2)
                    If bContourTee = False Then
                        GoTo NextJ
                    End If

                    'Add to another collection
                    TeeWireElemColl2.Add oTeeWire2

                    'Get list of all parts belonging to parent of current part.
                    ArrayOfJ(k) = j
                    k = k + 1
                End If
NextJ:
            Next j

            Dim oMergedWB As IJWireBody
            Dim oMfgGeomHelper As New MfgGeomHelper

            Dim oMergedCSColl As IJElements
            Set oMergedCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(TeeWireElemColl)

            Dim oMergedCSColl2 As IJElements
            Dim x As Integer, icount As Integer, iCount2 As Integer
            If oMergedCSColl Is Nothing Or oMergedCSColl.Count < 1 Then
                GoTo CleanUp
            ElseIf oMergedCSColl.Count = 1 Then ' DO NOT Flip!
                ArrayWithFlipBits(i) = 1
                For x = 0 To UBound(ArrayOfJ)
                    ArrayWithFlipBits(ArrayOfJ(x)) = 1
                Next x
            Else
                Set oMergedCSColl2 = oMfgGeomHelper.OptimizedMergingOfInputCurves(TeeWireElemColl2)
                If oMergedCSColl2 Is Nothing Or oMergedCSColl2.Count < 1 Then GoTo CleanUp
                If oMergedCSColl2.Count = 1 Then ' Flip!
                    ArrayWithFlipBits(i) = 2
                    For x = 0 To UBound(ArrayOfJ)
                        If ArrayOfJ(x) <> 0 Then
                            ArrayWithFlipBits(ArrayOfJ(x)) = 2
                        End If
                    Next x
                Else 'oMergedCSColl.Count > 1 and oMergedCSColl2.Count > 1
                    'In this case, mark the side with least distance between parts
                    'Go through 1st collection to get the least distance
                    Dim dMin1 As Double, dMin2 As Double, dDist As Double
                    Dim oModelBody1 As IJDModelBody, oModelBody2 As IJDModelBody
                    Dim temp As IJDPosition
                    dMin1 = 1000
                    For icount = 1 To oMergedCSColl.Count
                        Set oModelBody1 = m_oMfgRuleHelper.ComplexStringToWireBody(oMergedCSColl.Item(icount))
                        If oModelBody1 Is Nothing Then GoTo NextiCount
                        For iCount2 = icount + 1 To oMergedCSColl.Count
                            Set oModelBody2 = m_oMfgRuleHelper.ComplexStringToWireBody(oMergedCSColl.Item(iCount2))
                            If oModelBody2 Is Nothing Then GoTo NextiCount2
                            oModelBody1.GetMinimumDistance oModelBody2, temp, temp, dDist
                            If dDist < dMin1 Then
                                dMin1 = dDist
                            End If
NextiCount2:
                        Next iCount2
NextiCount:
                    Next icount

                    dMin2 = 100
                    For icount = 1 To oMergedCSColl2.Count
                        Set oModelBody1 = m_oMfgRuleHelper.ComplexStringToWireBody(oMergedCSColl2.Item(icount))
                        If oModelBody1 Is Nothing Then GoTo GetNextiCount
                        For iCount2 = icount + 1 To oMergedCSColl2.Count
                            Set oModelBody2 = m_oMfgRuleHelper.ComplexStringToWireBody(oMergedCSColl2.Item(iCount2))
                            If oModelBody1 Is Nothing Then GoTo GetNextiCount2
                            oModelBody1.GetMinimumDistance oModelBody2, temp, temp, dDist
                            If dDist < dMin2 Then
                                dMin2 = dDist
                            End If
GetNextiCount2:
                        Next iCount2
GetNextiCount:
                    Next icount

                    'If the collection1 exhibits least distance, then continue else flip
                    If dMin1 <= dMin2 Then
                        ArrayWithFlipBits(i) = 1
                        For x = 0 To UBound(ArrayOfJ)
                            ArrayWithFlipBits(ArrayOfJ(x)) = 1
                        Next x
                    Else
                        ArrayWithFlipBits(i) = 2
                        For x = 0 To UBound(ArrayOfJ)
                            If ArrayOfJ(x) <> 0 Then
                                ArrayWithFlipBits(ArrayOfJ(x)) = 2
                            End If
                        Next x
                    End If

                End If
            End If


'                If oMergedCSColl.Count = 1 Then  ' DO NOT Flip!
'                    ArrayWithFlipBits(i) = 1
'                    For x = 0 To UBound(ArrayOfJ)
'                        'j-1 is true because "Next j" above increments j
'                        ArrayWithFlipBits(ArrayOfJ(x)) = 1
'                    Next x
'                ElseIf oMergedCSColl2.Count = 1 Then ' Flip!
'                    ArrayWithFlipBits(i) = 2
'                    For x = 0 To UBound(ArrayOfJ)
'                        If ArrayOfJ(x) <> 0 Then
'                        'j-1 is true because "Next j" above increments j
'                            ArrayWithFlipBits(ArrayOfJ(x)) = 2
'                        End If
'                        'ArrayWithFlipBits(j - 1) = 2
'                    Next x
'
'                End If
'            End If

        End If ' End of ArrayWithFlipBits(i) = 0
NextI:
    Next i

    CheckIfFlipMarkingSideNeeded = ArrayWithFlipBits

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES")
    GoTo CleanUp

End Function


Public Function GetSideOfConnectedObjectToBeMarked(oObject As Object) As ThicknessSide

    Const METHOD = "GetSideOfConnectedObjectToBeMarked"
    On Error GoTo ErrorHandler

    Dim oPartInfo As IJDPartInfo
    Set oPartInfo = New PartInfo

    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Set oSDConPlateWrapper.object = oObject
    Dim sMoldedSide As String
    sMoldedSide = oSDConPlateWrapper.MoldedSide

    Dim eMoldedDir As StructMoldedDirection
    eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(oObject)

        ' Get the moulded side of the connected plate
    If sMoldedSide = "Base" Then
        GetSideOfConnectedObjectToBeMarked = SideA
    ElseIf sMoldedSide = "Offset" Then
        GetSideOfConnectedObjectToBeMarked = SideB
    Else
        Dim lErrNumber As Long
        lErrNumber = LogMessage(Err, MODULE, METHOD, " Unexpected Molded side ")
    End If

CleanUp:

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES")
    GoTo CleanUp

End Function


Public Function GetPlateThicknessDir(oPlatePart As Object) As String

    Const METHOD = "GetPlateThicknessDir"
    On Error GoTo GetPlateThicknessDir_Error

    Dim oPlate As IJPlate
    Dim oPlateMC As IJDPlateMoldedConventions
    Dim Dir As StructMoldedDirection
    Dim strDir As String
    Dim ePlateType As StructPlateType
    Dim oStandAlonPart As IJDStandAlonePlatePart
    If TypeOf oPlatePart Is IJDStandAlonePlatePart Then
        Set oStandAlonPart = oPlatePart
        Set oPlate = oPlatePart
        ePlateType = oPlate.plateType
        Dir = oStandAlonPart.plateThicknessDirection
    Else
        ' Sastry    11-Mar-2003
        ' Getting Parent Plate System recursively
        Set oPlate = GetParentPlateSystem(oPlatePart)

        Set oPlateMC = oPlate

        ePlateType = oPlate.plateType

        Dir = oPlateMC.plateThicknessDirection
    End If
    Select Case ePlateType
        Case DeckPlate:
            Select Case Dir
                Case AboveDir:
                    strDir = "upper"
                Case BelowDir:
                    strDir = "lower"
            End Select

        Case LBulkheadPlate
            Select Case Dir
                Case StarDir:
                    strDir = "starboard"
                Case PortDir:
                    strDir = "port"
                Case InboardDir:
                    strDir = "in"
                Case OutboardDir:
                    strDir = "out"
                Case Centered:
                    strDir = "centered"
            End Select
        Case TBulkheadPlate
            Select Case Dir
                Case ForeDir:
                    strDir = "fore"
                Case AftDir:
                    strDir = "aft"
            End Select
        Case TransverseTube, VerticalTube, LongitudinalTube
            Select Case Dir
                Case InDir:
                    strDir = "in"
                Case OutDir:
                    strDir = "out"
                Case Centered:
                    strDir = "centered"
            End Select
    End Select

    GetPlateThicknessDir = strDir


    Exit Function

GetPlateThicknessDir_Error:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' Sastry    11-Mar-2003
' This method gets the parent plate system recursively
Public Function GetParentPlateSystem(oPlatePart As Object) As Object

    Const METHOD = "GetParentPlateSystem"
    On Error GoTo ErrHandler

    Dim oChild  As IJSystemChild
    Dim oPlate As IJPlate

    Set oChild = oPlatePart

    Set oPlate = oChild.GetParent()
    If TypeOf oPlate Is IJPlate Then
        If TypeOf oPlate Is IJPlateSystem Then
            Set GetParentPlateSystem = oPlate
        Else
            Set GetParentPlateSystem = GetParentPlateSystem(oPlate)
        End If
    Else
        GoTo ErrHandler
    End If

CleanUp:
    Set oPlate = Nothing
    Set oChild = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Function

Public Function CreateBracketLocationMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                                          Optional ByVal bAddERMarks As Boolean, Optional ByVal bAddFittingMark As Boolean, Optional ByVal bAddRENDMarks As Boolean, Optional ByVal bAddDeclivityMark As Boolean, _
                                          Optional ByVal CONN_PART_CONDITION As Long = 0, Optional ByVal bAddCenterFittingMark As Boolean) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "CreateBracketLocationMark"
    On Error GoTo ErrorHandler

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_BRACKETLOCATION_MARK, ReferenceObjColl, oGeomCol3d

    Set CreateBracketLocationMark = oGeomCol3d

    Dim eSideOfConnectedObjectToBeMarked As enumPlateSide

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)

    If oConObjsCol Is Nothing Then
        'Since there is no connecting strcuture we can leave the marking rule
        GoTo CleanUp
    End If

    Dim Item As Object
    Dim oConnectionData As ConnectionData

    Dim nIndex As Long
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn

    ' Loop thru each Physical Connections
    Dim bContourTee As Boolean
    Dim oVector As IJDVector
    Dim oWB As IJWireBody, oTeeWire As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oSurfaceBody As IJSurfaceBody
    Dim oGeom3d As IJMfgGeom3D
    Dim oStart As IJDPosition, oEnd As IJDPosition
    Dim oPlateNormal As IJDVector
    Dim lGeomCount As Long
    Dim oSupp2 As IJPartSupport
    Dim oSys As IJSystem
    Dim oNamedItem As IJNamedItem
    Dim oPlateUtils As IJPlateAttributes
    Dim eSurfGeomType As SurGeoType

    lGeomCount = 1

    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oPlatePartSupport As GSCADSDPartSupport.IJPlatePartSupport

    Set oPartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport.Part = oSDPlateWrapper.object
    Set oPlatePartSupport = oPartSupport

    'For the connected parts
    Dim oPartSupport2 As GSCADSDPartSupport.IJPartSupport
    Set oPartSupport2 = New PartSupport


    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
        End If
    Next nIndex

    Dim oSDPlateWrapper2 As StructDetailObjects.PlatePart
    Set oSDPlateWrapper2 = New StructDetailObjects.PlatePart

    Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart

    Dim oPort As IJPort
    Dim oBracketPort As IJPort
    Dim oStructPort As IJStructPort
    Dim pProfileCurvature As ProfileCurvature
    Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
    Dim bIsProfileCurved As Boolean
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition, oMidPos As New DPosition
    Dim oStartDir As New DVector, oEndDir As New DVector
    Dim lMarkingSide As Long
    Dim bFlip As Boolean
    Set oPlateUtils = New PlateUtils

    If bAddCenterFittingMark = True Then
        Set oPartSupport2.Part = Part

        'Getting All the Connected Objects to the
        oPartSupport2.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol

        Dim oMfgMGHelper As New GSCADMathGeom.MfgMGHelper
        Dim oAppConnection As IJAppConnection

        Dim oPlateAtt As IJPlateAttributes
        Set oPlateAtt = New PlateUtils

        Dim oPlane                          As IJPlane
        Dim oPoint1                         As IJPoint
        Dim oPoint2                         As IJPoint
        Dim strRootSelector                 As String
        Dim oObject1                        As Object
        Dim oObject2                        As Object
        Dim oObject3                        As Object
        Dim oObject4                        As Object
        Dim oObject5                        As Object
        Dim oRootPartSystem                As IJSystem
        Dim oStructDetailHelper             As GSCADStructDetailUtil.StructDetailHelper
        Dim oBracketSupportCol              As New Collection

        Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        oStructDetailHelper.IsPartDerivedFromSystem Part, oRootPartSystem, True

        If Not oRootPartSystem Is Nothing Then
            If oPlateUtils.IsBracketByPlane(oRootPartSystem) Then


                oPlateAtt.GetInput_BracketByPlane_IndividualSupports oRootPartSystem, oPlane, oPoint1, oPoint2, strRootSelector, oObject1, oObject2, oObject3, oObject4, oObject5
                Set oPlane = Nothing
                Set oPoint1 = Nothing
                Set oPoint2 = Nothing
                Dim bIs2SBracket As Boolean

                If oObject3 Is Nothing Then
                    bIs2SBracket = True
                    oBracketSupportCol.Add oObject1
                    oBracketSupportCol.Add oObject2
                End If

                If bIs2SBracket = True Then        'A Case of 2S Brackets
                    Dim oPartSystem             As IJSystem
                    Dim oRPSChildColl           As IJDTargetObjectCol
                    Dim oSystem                 As IJSystem
                    Dim oProfilePart            As IJProfilePart
                    Dim oStiffenedPlate         As IJPlate
                    Dim oConnectable            As IJConnectable
                    Dim oStructConnectable      As IJStructConnectable
                    Dim oRootBoundingConnection As IJStructConnection
                    Dim icount                  As Integer

                    Set oStructConnectable = oRootPartSystem
                    Set oRootPartSystem = Nothing

                    'Checking for ShellProfile Type and Getting the OtherPart
                    For icount = 1 To 2
                        If TypeOf oBracketSupportCol.Item(icount) Is IJStiffenerSystem Then
                            Set oPartSystem = oBracketSupportCol.Item(icount)
                            Set oRPSChildColl = oPartSystem.GetChildren
                            For nIndex = 1 To oRPSChildColl.Count
                                If TypeOf oRPSChildColl.Item(nIndex) Is IJProfile Then
                                    Set oSystem = oRPSChildColl.Item(nIndex)
                                    Set oProfilePart = oSystem.GetChildren.Item(1)
                                    Set oSystem = Nothing
                                    Set oSDProfileWrapper.object = oProfilePart
                                    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, True
                                    Set oSDPlateWrapper2.object = oStiffenedPlate
                                    If oSDPlateWrapper2.plateType = Hull Then
                                        If icount = 1 Then
                                            Set oConnectable = oBracketSupportCol.Item(icount + 1)
                                        Else
                                            Set oConnectable = oBracketSupportCol.Item(icount - 1)
                                        End If
                                    End If
                                End If
                            Next nIndex
                        End If
                    Next icount

                    Set oPartSystem = Nothing
                    Set oRPSChildColl = Nothing
                    Set oSystem = Nothing

                    oStructConnectable.GetRootConnection oConnectable, oRootBoundingConnection
                    Set oRootPartSystem = Nothing

                    For nIndex = 1 To oConObjCol.Count

                        oStructDetailHelper.IsPartDerivedFromSystem oConObjCol.Item(nIndex), oRootPartSystem, True

                        If oRootPartSystem Is oConnectable Then
                            Set oPort = oThisPortCol.Item(nIndex)
                            Set oAppConnection = oConnCol.Item(nIndex)
                            Dim oPCWireBody As IJWireBody
                            Set oPCWireBody = oRootBoundingConnection.ModelBody

                            'Getting the Mid point of PC
                            Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oPCWireBody)
                            Set oSurfaceBody = oPort.Geometry
                            oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceBody, oMidPos, oStart, oPlateNormal
                            oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal

                            Set oBracketPort = oPlateWrapper.GetSurfacePort(UpSide)

                            Dim oBracketSurfaceBody  As IJSurfaceBody
                            Set oBracketSurfaceBody = oBracketPort.Geometry

                            Set oCS = CreateMarkAtPosition(oBracketSurfaceBody, oStart, oPlateNormal, END_FITTING_MARK_LENGTH)

                            'Create a SystemMark object to store additional information
                            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                            'QI for the MarkingInfo object on the SystemMark
                            Set oMarkingInfo = oSystemMark

                            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                            oGeom3d.PutGeometry oCS
                            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
                            'oGeom3d.PutSubGeometryType STRMFG_BRACKETLOCATION_MARK
                            oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "BracketLocation"
                            oSystemMark.Set3dGeometry oGeom3d

                            'Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oAppConnection)
                            oGeom3d.PutMoniker oMoniker

                            oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                            lGeomCount = lGeomCount + 1
                        End If
                    Next nIndex
                    Set oCS = Nothing
                    Set oAppConnection = Nothing
                    Set oMidPos = Nothing
                End If
            End If
        End If
    End If


    For nIndex = 1 To oConObjsCol.Count

        oConnectionData = oConObjsCol.Item(nIndex)

        'pProfileCurvature = EvaluateProfileCurvature(oConnectionData.ToConnectable)

        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
             GoTo NextItem
        End If

        Set oPartSupport2.Part = oConnectionData.ToConnectable
        oPartSupport2.IsSystemDerivedPart oSys, True
        Set oNamedItem = oSys
        Set oNamedItem = oPartSupport2.Part
        Set oPlateUtils = New PlateUtils

        If Not oSys Is Nothing Then
            If Not (oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys)) Then
                GoTo NextItem
            End If
        Else
            GoTo NextItem
        End If


        Set oPort = oConnectionData.ToConnectedPort
        Set oStructPort = oPort
        'oStructPort.ContextID

        Set oPartSupport.Part = oConnectionData.ToConnectable
        Set oPlatePartSupport = oPartSupport

        Set oSDPlateWrapper2.object = oConnectionData.ToConnectable

        '*** Check if the bracket has connection with a shell profile ***'
        bIsProfileCurved = IsProfileCurved(oConnectionData.ToConnectable, oSDPlateWrapper)

        ' Check if this physical connection is a Root PC, which particpates in Split operation
        ' If so, Skip the marking line line creation for this Root PC.
        Dim oStructEntityOperation As IJDStructEntityOperation
        Dim opeartionProgID As String
        Dim opeartionID As StructOperation
        Dim oOperColl As New Collection

        Set oStructEntityOperation = oConnectionData.AppConnection
        oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl

        Set oStructEntityOperation = Nothing
        Set oOperColl = Nothing

        ' If the RootPC has Split operation in its graph, just goto the next pc.
        If opeartionID = ConnectionSplitOperation Then
            ' No need to create marking line.
            GoTo NextItem
        End If

        Dim oSDConPlateWrapper As StructDetailObjects.PlatePart
        Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDConPlateWrapper.object = oConnectionData.ToConnectable

        Dim sMoldedSide As String
        sMoldedSide = oSDConPlateWrapper.MoldedSide
        If sMoldedSide = "Base" Then
            eSideOfConnectedObjectToBeMarked = BaseSide
        ElseIf sMoldedSide = "Offset" Then
            eSideOfConnectedObjectToBeMarked = OffsetSide
        Else
            Dim lErrNumber As Long
            lErrNumber = LogMessage(Err, MODULE, METHOD, " Unexpected Molded side ")
        End If

        'Initialize the profile wrapper and the Physical Connection wrapper
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        bContourTee = oSDPlateWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)

        If bContourTee = True Then
            ' Bound the wire based on split points, if there are any.
            Set oWB = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire)

            Set oSurfaceBody = oSDPlateWrapper.BasePort(BPT_Base).Geometry

            Dim oVector1 As IJDVector

            oWB.GetEndPoints oStart, oEnd

            Set oStart = m_oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBody, oVector1)

            oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal
            Set oVector1 = Nothing

            'Convert the IJWireBody to a IJComplexString
            Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWB)

            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            'Set the marking side
            oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            Dim strProdName     As String
            strProdName = GetManufacturingProductionName(Part, oConnectionData.ToConnectable, CONN_PART_CONDITION)
            If Not strProdName = vbNullString Then
                oMarkingInfo.Name = strProdName
            End If

            oMarkingInfo.Direction = m_oMfgRuleHelper.GetDirection(oVector)
            oMarkingInfo.ThicknessDirection = GetThicknessDirectionVector(oCS, oSDConPlateWrapper, sMoldedSide)
            oMarkingInfo.FittingAngle = oSDPhysicalConn.TeeMountingAngle

            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_BRACKETLOCATION_MARK
            oSystemMark.Set3dGeometry oGeom3d

            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            oGeomCol3d.AddGeometry lGeomCount, oGeom3d
            lGeomCount = lGeomCount + 1

            ' Check if declivity marks are needed
            If bAddDeclivityMark = True Then
                CreateDeclivityMarks Part, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, UpSide, sMoldedSide, Nothing
            End If

            'Set oGeom3d = GetEndFittingMarkAtPoint(oStartPos, oSDPlateWrapper2, sMoldedSide, UpSide, oConnectionData, oStartDir, oPlateNormal, oTDir)

            '*** CREATE END FITTING MARK ***'
            'This function Creates the End Fitting Mark(s) at the Free End(s) of the
            '... Profile Parts.

            If bIsProfileCurved Then

                If bAddCenterFittingMark = True Then
                    oWB.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir

                    '*** Get the middle point of the curve ***'
                    Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oWB)
                    Set oGeom3d = GetEndFittingMarkAtPoint(oMidPos, oSDPlateWrapper2, sMoldedSide, UpSide, oConnectionData, oEndDir, oPlateNormal, oStartDir, BRACKETSTIFF_FITTING_MARK_LENGTH)

                    Set oMarkingInfo = oGeom3d.SystemMark
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "BracketLocation"
                    '*** Add the Geom3d object to collection ***'
                    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                    lGeomCount = lGeomCount + 1
                End If

            Else

                If oPlateWrapper.GetSide(oConnectionData.ConnectingPort) <> UpSide Then
                    bFlip = True
                End If

                If bAddFittingMark Then
                    CreateEndFittingMark Part, oCS, oWB, oPartSupport, oSDPlateWrapper, _
                             oSDPlateWrapper2, sMoldedSide, oPlateWrapper.GetSide(oConnectionData.ConnectingPort), oConnectionData, _
                             oGeomCol3d, oPlateNormal, bFlip
                End If

            End If

            'GetMarkingSideAndPort oGeom3d, lMarkingSide

            '*******************************'


        End If
NextItem:

        Set oWB = Nothing
        Set oTeeWire = Nothing
        Set oVector = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConPlateWrapper = Nothing
        Set oSDPhysicalConn = Nothing
        'Set oSmartPlate = Nothing
    Next nIndex

    'Return the 3d collection
    Set CreateBracketLocationMark = oGeomCol3d

    Set Item = Nothing
    Set oCS = Nothing

CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing

Exit Function

ErrorHandler:
    'MsgBox "Err:" & Err.Description
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1004, , "RULES")
    GoTo CleanUp
End Function

Public Function CreatePenetrationMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, bFullMark As Boolean, bWebPenMark As Boolean, bMarkOnlyTripElem As Boolean) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "CreatePenetrationMark"
    On Error GoTo ErrorHandler

    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_PROFILE_TO_PLATE_PENETRATION_MARK, ReferenceObjColl, oGeomCol3d
    Set CreatePenetrationMark = oGeomCol3d

    If bSelectiveRecompute Then
        Exit Function
    End If

    'wrap plate part
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    'Hull plates cannot be processed
    Dim ePlateType As StructPlateType
    ePlateType = oSDPlateWrapper.plateType
    If ePlateType = Hull Then
        Set oPlateWrapper = Nothing
        Set oSDPlateWrapper = Nothing
        Exit Function
    End If

    'Non-planar plates cannot be processed
    If Not oPlateWrapper.CurvatureType = PLATE_CURVATURE_Flat Then
        'GoTo CleanUp
    End If

    'Get plate's marking surface
    Dim oSurface As IJSurfaceBody
    Dim oThiknessSide As PlateThicknessSide
    'Convert between enums
    If UpSide = BaseSide Then
        oThiknessSide = PlateBaseSide
    Else
        oThiknessSide = PlateOffsetSide
    End If

    Dim oThisPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Dim oUpSideGeometry As IUnknown

    Set oThisPlateWrapper.object = Part

    Set oSurface = oThisPlateWrapper.GetSurfaceWithoutFeatures(oThiknessSide)

    Set oUpSideGeometry = oSurface

    Dim oUpSideSurfaceNormal As New DVector
    Dim oPlane As IJPlane
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double

    'Use GetPlateNeutralSurfaceNormal to get oUpSideSurfaceNormal
    Set oUpSideSurfaceNormal = GetPlateNeutralSurfaceNormal(Part)

    'Retrieve collection of Feature objects for this plate
    Dim oFeatureObjectsColl As Collection
    Set oFeatureObjectsColl = oSDPlateWrapper.PlateFeatures

    If oFeatureObjectsColl Is Nothing Then
        'Since there are not features on this plate there will be no penetration marks to be made
        GoTo CleanUp
    End If

    Dim oPlateParentSystem As Object
    Set oPlateParentSystem = m_oMfgRuleHelper.GetTopMostParentSystem(Part)

    'Dim oConnectionData As StructDetailObjects.ConnectionData
    Dim iFeatureIndex As Integer 'iteration object
    Dim oPt As New DPosition
    Dim oCol1 As Collection, oPortCol As Collection
    Dim bWebLeftMark As Boolean

    Dim oWebLeftPortColl As IJElements
    Set oWebLeftPortColl = New JObjectCollection
        
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    'iterate through the collection of features
    For iFeatureIndex = 1 To oFeatureObjectsColl.Count

        Set oPortCol = New Collection
        ' Sastry
        ' TR#40585
        If Not TypeOf oFeatureObjectsColl.Item(iFeatureIndex) Is IJStructFeature Then
            GoTo NextFeatureIndex
        End If

        'Drop if not Profile/Plate penetration (Slot).
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oFeatureObjectsColl.Item(iFeatureIndex)
        If Not oStructFeature.get_StructFeatureType = SF_Slot Then GoTo NextFeatureIndex

        'Wrap the slot
        Dim oSlotWrapper As New StructDetailObjects.Slot
        Set oSlotWrapper.object = oFeatureObjectsColl.Item(iFeatureIndex)

        If Not (TypeOf oSlotWrapper.Penetrating Is IJProfilePart) Then GoTo NextFeatureIndex

        'Get penetrating profile object
        Dim oPenetratingProfile As IJProfilePart
        Set oPenetratingProfile = oSlotWrapper.Penetrating

        Dim bIsThereAnyTrippingElement As Boolean
        Dim bIsCrossSectionAllowed As Boolean
        Dim oProfileParentSystem As Object

        Set oProfileParentSystem = m_oMfgRuleHelper.GetTopMostParentSystem(oPenetratingProfile)

        '*** Check for tripping elements ***'
        bIsThereAnyTrippingElement = m_oMfgRuleHelper.CheckIfThereIsAnyTrippingElem(oPlateParentSystem, oProfileParentSystem)
        
        Dim nProfileIndex As Long
        Dim oConnectionData As ConnectionData
        Dim oProfileAtt As IJProfileAttributes
        
        Set oProfileAtt = New GSCADCreateModifyUtilities.ProfileUtils
    
        For nProfileIndex = 1 To oConObjsCol.Count
            Dim oStifParentSystem As Object
            oConnectionData = oConObjsCol.Item(nProfileIndex)
            ' This is the stiffener ON the plate
            Set oStifParentSystem = m_oMfgRuleHelper.GetTopMostParentSystem(oConnectionData.ToConnectable)
            If Not oStifParentSystem Is Nothing Then
                If TypeOf oStifParentSystem Is IJProfile Then
                    Dim oProfileBoundaries As Collection
                    ' Get the boundaries of the stiffener
                    Set oProfileBoundaries = oProfileAtt.GetProfileSystemBoundary(oStifParentSystem)
                    
                    Dim nBoundaryIndex As Long
                    For nBoundaryIndex = 1 To oProfileBoundaries.Count
                        ' If any of the profile boundaries matches with the penetrating profile, this is candidate for marking
                        If (oProfileBoundaries.Item(nBoundaryIndex) Is oProfileParentSystem) Then
                            bIsThereAnyTrippingElement = True
                            GoTo GotTrippingElem
                        End If
                    Next
                    Set oProfileBoundaries = Nothing
                End If
            End If
            Set oStifParentSystem = Nothing
        Next

GotTrippingElem:
        If bMarkOnlyTripElem <> True Then
            bIsThereAnyTrippingElement = False
        End If

        If bMarkOnlyTripElem = True And bIsThereAnyTrippingElement = False Then
            GoTo NextFeatureIndex
        End If

        bIsCrossSectionAllowed = CheckIfCrossSectionAllowed(oPenetratingProfile)

        Set oProfileParentSystem = Nothing

        'wrap penetrating profile
        Dim oPenetratingProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
        Set oPenetratingProfileWrapper.object = oPenetratingProfile

        'Get profile port opposite mounting face
        Dim oPortOppositMount As IJPort
        Set oPortOppositMount = oPenetratingProfileWrapper.GetPortOppositeMountingFace

        oPortCol.Add oPortOppositMount

        '*** Case - If Web Penetration Mark is needed ***'
'                               |
'                              _|___
'                           _ /     \_
'                            |   __ |
'                            |  |
'                     _______|  |__________

        If bWebPenMark = True Then
            Set oPortOppositMount = oPenetratingProfileWrapper.GetSurfacePort(JXSEC_WEB_LEFT)
            oPortCol.Add oPortOppositMount
            oWebLeftPortColl.Add oPortOppositMount
        End If
        '**************************************************'

        For Each oPortOppositMount In oPortCol

            Set oCol1 = New Collection

            'Get intersection product between two surfaces
            Dim oIntersectionWB As IJWireBody
            Dim oPortGeometry As IUnknown

            If Not oPortOppositMount Is Nothing Then
                Set oPortGeometry = oPortOppositMount.Geometry
            End If

            Set oIntersectionWB = m_oMfgRuleHelper.GetCommonGeometry(oUpSideGeometry, oPortGeometry)
            If oIntersectionWB Is Nothing Then
                GoTo NextFeatureIndex
                'Exit Function
            End If

            '*** Offset the penetration mark by 1mm  as per customer request ***'
            'Get SurfaceNormal at Flange Top Surface - We need to shift it by 0.001m
            ' ... As per Customer request.

            Dim oFlangeNormal As IJDVector
            Dim oDummy As IJDPosition
            Dim oTL As GSCADStructGeomUtilities.TopologyLocate
            Set oTL = New TopologyLocate
            oTL.FindApproxCenterAndNormal oPortGeometry, oDummy, oFlangeNormal

            'Get two end positions
            Dim oStartPos  As IJDPosition
            Dim oEndPos   As IJDPosition
            Dim oStartPosDir As IJDVector
            Dim oEndPosDir As IJDVector
            oIntersectionWB.GetEndPoints oStartPos, oEndPos, oStartPosDir, oEndPosDir

            'Create  line through the positions

            Dim xStart As Double
            Dim yStart As Double
            Dim zStart As Double
            Dim xEnd As Double
            Dim yEnd As Double
            Dim zEnd As Double
            Dim endX As Double
            Dim endY As Double
            Dim endZ As Double

            Dim oMarkCS As IJComplexString
            Dim oMarkWB As IJWireBody, oOffsetMarkWB As IJWireBody

            Dim oEndPosColl As Collection
            Set oEndPosColl = New Collection

            oEndPosColl.Add oStartPos
            oEndPosColl.Add oEndPos

            Dim oTempPos As IJDPosition
            Dim iMarkCount As Integer
            iMarkCount = 0

            For Each oTempPos In oEndPosColl

                iMarkCount = iMarkCount + 1
                Dim oLine As New Line3d
                oStartPos.Get xStart, yStart, zStart
                oEndPos.Get xEnd, yEnd, zEnd

                'Create line
                If oTempPos Is oStartPos Then
                    oLine.DefineBy2Points xStart, yStart, zStart, (2# * xStart) - xEnd, (2# * yStart) - yEnd, (2# * zStart) - zEnd
                ElseIf oTempPos Is oEndPos Then
                    oLine.DefineBy2Points xEnd, yEnd, zEnd, (2# * xEnd) - xStart, (2# * yEnd) - yStart, (2# * zEnd) - zStart
                End If

                oLine.Length = 5

                Dim oCS As New ComplexString3d
                Dim oLineWB As IJWireBody
                oCS.AddCurve oLine, False
                Set oLineWB = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)

                'Retrieve symbol contour
                oStructFeature.CacheSymbolOutput
                Dim oAttribWire As IJWireBody
                Dim dMaximumExtrusionDistance As Double
                Dim dMinimumExtrusionDistance As Double
                Dim bUseExtrusionDistance As Boolean
                Dim oProjectionVecors As IJElements
                'oStructFeature.GetCachedSymbolOutput oAttribWire, dMaximumExtrusionDistance, _
                    dMinimumExtrusionDistance, bUseExtrusionDistance, oProjectionVecors

                '--------------------------------------------------------------------------------
                ' The above method reutrns the feature geometry on neutral surface of the plate
                ' Instead, use the below method which returns the feature geometry on upside
                Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
                Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
                Set oSDPartSupport.Part = Part

                Dim oContourColl As Collection
                Dim oPortMonikerColl As Collection

                ' Get the feature contour collection
                oSDPartSupport.GetFeatureInfo oStructFeature, UpSide, oContourColl, oPortMonikerColl

                Dim oMfgMGHelper As IJMfgMGHelper
                Set oMfgMGHelper = New MfgMGHelper
                Dim oCurveElems As IJElements
                Set oCurveElems = New JObjectCollection
                Dim oCurveElems1 As IJElements
                Dim icount      As Long

                If Not oContourColl Is Nothing Then
                    For icount = 1 To oContourColl.Count
                        Dim oTempElems As IJElements
                        oMfgMGHelper.WireBodyToComplexStrings oContourColl.Item(icount), oTempElems

                        Set oCurveElems1 = New JObjectCollection
                        Dim iCount2 As Long
                        For iCount2 = 1 To oTempElems.Count
                            Dim oTempCS As New ComplexString3d
                            Set oTempCS = oTempElems.Item(iCount2)

                            oTempCS.GetCurves oCurveElems1

                        Next iCount2

                        oCurveElems.AddElements oCurveElems1
                    Next
                Else
                    Exit Function
                End If

                Dim oFeatureCS As New ComplexString3d
                oFeatureCS.SetCurves oCurveElems
                Set oAttribWire = m_oMfgRuleHelper.ComplexStringToWireBody(oFeatureCS)
                '--------------------------------------------------------------------------------

                Dim oMDBody As IJDModelBody
                Dim oPosOnFeature As New DPosition
                Dim oPosOnLine As New DPosition
                Dim dDist As Double

                Set oMDBody = oAttribWire
                oMDBody.GetMinimumDistance oLineWB, oPosOnFeature, oPosOnLine, dDist

                'Prepare mark point collection
                Dim oMarkPosColl As New Collection
                Set oMarkPosColl = New Collection

                'retrieve intersection positions
                'Set oMarkPosColl = m_oMfgRuleHelper.GetPositionsFromPointsGraph(oCommonGeom)
                oMarkPosColl.Add oPosOnFeature

                If oMarkPosColl.Count = 0 Then Exit Function

                'Iterate through the mark points collection and create ML
                Dim oMarkPointPos As IJDPosition
                Dim iIndex As Integer

                'For Each oMarkPointPosCount In oMarkPosColl (supposed to be only one)
                For iIndex = 1 To oMarkPosColl.Count

                    'Project oLine on plate's surface if nessesary
                    If TypeOf oMarkPosColl.Item(iIndex) Is IJDPosition Then
                        Set oMarkPointPos = oMarkPosColl.Item(iIndex)
                    Else
                        GoTo NextMarkPointPos
                    End If

                    'Create mark line from the part of oLine
                    Dim oMarkLine As New Line3d

                    oMarkPointPos.Get xStart, yStart, zStart
                    oLine.GetDirection xEnd, yEnd, zEnd
                    endX = xStart + (xEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    endY = yStart + (yEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    endZ = zStart + (zEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    oMarkLine.DefineBy2Points xStart, yStart, zStart, endX, endY, endZ

                    '*** Case I - Partial Marks ***'
'                              _|___
'                           _ /     \_
'                            |   __ |
'                            |  |
'                     _______|  |__________

                    If bFullMark = False And bIsThereAnyTrippingElement = False Then

                        ' Offset the line created here as there is an issue with slot type 'SlotT-T1'
                        ' when we offset the intersection wire(PortOppMountingFace and Plate surface) before
                        If Not oWebLeftPortColl.Contains(oPortOppositMount) Then
                            Set oMarkCS = New ComplexString3d
                            oMarkCS.AddCurve oMarkLine, False
                            Set oMarkWB = m_oMfgRuleHelper.ComplexStringToWireBody(oMarkCS)

                            ' Offset the mark by PROFILE_TO_PLATE_PEN_STRETCH_LENGTH
                            If Not oFlangeNormal Is Nothing Then
                                Set oOffsetMarkWB = m_oMfgRuleHelper.OffsetCurve(oUpSideGeometry, oMarkWB, oFlangeNormal, PROFILE_TO_PLATE_PEN_STRETCH_LENGTH, False)
                            End If

                            Set oMarkCS = m_oMfgRuleHelper.WireBodyToComplexString(oOffsetMarkWB)
                            Set oMarkLine = Nothing
                            'oMarkCS.GetCurve 1, oMarkLine

                        End If

                        If Not oMarkLine Is Nothing Then
                            Set oGeomCol3d = GetPenMarkGeom3D(oGeomCol3d, oMarkCS, UpSide, oStructFeature)
                        End If

                        Set oMarkCS = Nothing
                        Set oMarkWB = Nothing
                        Set oOffsetMarkWB = Nothing
                    End If
                    '*******************************'

                    Set oPt = New DPosition

                    oPt.x = endX
                    oPt.y = endY
                    oPt.z = endZ

                    oCol1.Add oPt

                    Set oMarkLine = Nothing

                Next iIndex

                'CleanUp
                Set oLine = Nothing
                Set oCS = Nothing
                Set oLineWB = Nothing

            Next oTempPos

            Dim oPt1 As New DPosition, oPt2 As New DPosition

            Set oPt1 = oCol1.Item(1)
            Set oPt2 = oCol1.Item(2)

            '*** Case II - Full Marks ***'
'                              _|___
'                           __/_|___\_
'                            |  | __|
'                            |  ||
'                     _______|  ||__________

            If bFullMark = True Or (bIsThereAnyTrippingElement = True And bIsCrossSectionAllowed = True) Then
                oMarkLine.DefineBy2Points oPt1.x, oPt1.y, oPt1.z, oPt2.x, oPt2.y, oPt2.z

                If Not oWebLeftPortColl.Contains(oPortOppositMount) Then
                    Set oMarkCS = New ComplexString3d
                    oMarkCS.AddCurve oMarkLine, False
                    Set oMarkWB = m_oMfgRuleHelper.ComplexStringToWireBody(oMarkCS)

                    ' Offset the mark by PROFILE_TO_PLATE_PEN_STRETCH_LENGTH
                    If Not oFlangeNormal Is Nothing Then
                        Set oOffsetMarkWB = m_oMfgRuleHelper.OffsetCurve(oUpSideGeometry, oMarkWB, oFlangeNormal, PROFILE_TO_PLATE_PEN_STRETCH_LENGTH, False)
                    End If

                    Set oMarkCS = m_oMfgRuleHelper.WireBodyToComplexString(oOffsetMarkWB)
                    Set oMarkLine = Nothing

                End If

                If Not oMarkLine Is Nothing Then
                    Set oGeomCol3d = GetPenMarkGeom3D(oGeomCol3d, oMarkCS, UpSide, oStructFeature)
                End If

                Set oMarkCS = Nothing
                Set oMarkWB = Nothing
                Set oOffsetMarkWB = Nothing
            End If
            '*******************************'

NextMarkPointPos:
            'CleanUp
            Set oMarkPointPos = Nothing
            Set oCol1 = Nothing
            bWebLeftMark = False

            'CleanUp
            Set oEndPosColl = Nothing
            'Set oPortCol = Nothing
        Next oPortOppositMount
NextFeatureIndex:
    Next iFeatureIndex

    'Return collection of ML's geometry
    Set CreatePenetrationMark = oGeomCol3d

CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oPlateParentSystem = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1024, , "RULES")
    GoTo CleanUp
End Function

Private Function CheckIfCrossSectionAllowed(oProfilePart As Object) As Boolean

    Dim oProfileSection As IJDProfileSection
    Set oProfileSection = oProfilePart

    Dim oCrossSection As IJCrossSection
    Set oCrossSection = oProfileSection.crossSection

    Dim oCSPartClass As IJDCrossSectionPartClass
    Set oCSPartClass = oCrossSection.GetPartClass

    Dim strProfileCrossSection As String
    strProfileCrossSection = oCSPartClass.CrossSectionTypeName

'    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
'    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    CheckIfCrossSectionAllowed = False

    Select Case UCase(strProfileCrossSection)

        Case "EQUALANGLE", "UNEQUALANGLE", "BULBFLAT", "FLATBAR", "IBAR", "CHANNEL", "TEEBAR"
            CheckIfCrossSectionAllowed = True
        Case Else

    End Select

CleanUp:

    Set oProfileSection = Nothing
    Set oCrossSection = Nothing
    Set oCSPartClass = Nothing

End Function

Private Function GetPenMarkGeom3D(oGeomCol3d As IJMfgGeomCol3d, oCS As IJComplexString, ByVal UpSide As Long, oStructFeature As IJStructFeature) As IJMfgGeomCol3d

    'Dim oCS_ML As New ComplexString3d
    'oCS_ML.AddCurve oMarkLine, False

    Dim oSystemMark As IJMfgSystemMark
    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oSystemMark.SetMarkingSide UpSide

    Dim oMarkingInfo As MarkingInfo
    Dim oMoniker As IMoniker

    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark
    oMarkingInfo.Name = "PROFILE PEN FITTING MARK"

    'Add ML's geometry to collection
    Dim oGeom3d As IJMfgGeom3D
    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype STRMFG_PROFILE_TO_PLATE_PENETRATION_MARK

    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oStructFeature)
    oGeom3d.PutMoniker oMoniker

    oSystemMark.Set3dGeometry oGeom3d
    oGeomCol3d.AddGeometry 1, oGeom3d

    Set GetPenMarkGeom3D = oGeomCol3d

    Set oCS = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oMoniker = Nothing
    Set oGeom3d = Nothing

Exit Function
ErrorHandler:

End Function

'***********************************************************************
' METHOD:  GetProjectedDirectionOnSurface
'
' DESCRIPTION:  Used to get a direction on a surface that is projected by another direction...

'***********************************************************************
Public Function GetProjectedDirectionOnSurface(ByVal oOriginalDir As IJDVector, ByVal oSurfaceNormal As IJDVector, Optional ByVal oProjectedAlong As IJDVector)
    Const METHOD As String = "GetProjectedDirectionOnSurface"
    On Error GoTo ErrorHandler

    Dim dPI As Double
    dPI = Atn(1) * 4

    Dim oTemVector As IJDVector
    Set oTemVector = New DVector

    Dim dAngelBetNomarlAndDir As Double

    '//Get a Temp dir to calculate the angel......

    If oProjectedAlong Is Nothing Then
        Set oProjectedAlong = oSurfaceNormal.Clone
    End If

    Set oTemVector = oOriginalDir.Cross(oProjectedAlong)
    dAngelBetNomarlAndDir = oOriginalDir.Angle(oProjectedAlong, oTemVector)

    If dAngelBetNomarlAndDir > dPI Then
        dAngelBetNomarlAndDir = 2 * dPI - dAngelBetNomarlAndDir
    ElseIf Abs(dAngelBetNomarlAndDir - 0) < 0.01 _
            Or Abs(dAngelBetNomarlAndDir - dPI) < 0.01 Then
        Exit Function '// we will not project the Dir when the angel is very small
    Else
        '// Nothing
    End If

    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper

    '//Do a double cross to get the projected vector
    Set oTemVector = oOriginalDir.Cross(oProjectedAlong)
    Set oTemVector = oTemVector.Cross(oSurfaceNormal)

    Dim dDot As Double
    dDot = oOriginalDir.Dot(oTemVector)  '//If the Dir is oppsite to the original one, then inverse it...

    If VBA.Round(dDot, 7) < 0 Then
        oMfgRuleHelper.ScaleVector oTemVector, -1
    End If

    Dim dLength As Double
    dLength = oTemVector.Length

    oMfgRuleHelper.ScaleVector oTemVector, 1 / dLength

    Set GetProjectedDirectionOnSurface = New DVector

    Set GetProjectedDirectionOnSurface = oTemVector.Clone

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Public Function Create3DLineComplexString
'
' Description:  Method will create a line as a ComplexString,
'               and the direction information will be in the MarkingInfo.Name ans Direction.
'
' ***********************************************************************************
Public Function Create3DLineComplexString(oPos As IJDPosition, oSurfacePort As IJPort, oDirection As IJDVector, UpSide As Long, ByRef oMark As IJComplexString, ByVal dLineLength As Double, Optional ByVal oProjectedDir As IJDVector) As Boolean

    Const METHOD = "Create3DLineComplexString"
    On Error GoTo ErrorHandler

    Dim endX As Double
    Dim endY As Double
    Dim endZ As Double

    Dim oPlate As IJPlatePart
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper

    Set oPlate = oSurfacePort.Connectable
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr ' .MfgPlatePartHlpr
    Set oPlateWrapper.object = oPlate

    oMfgRuleHelper.ScaleVector oDirection, 1 ' For TSN do not need to inverse the direction

    Dim oLine As IJLine
    Set oLine = New Line3d
    endX = oPos.x + (oDirection.x * dLineLength)
    endY = oPos.y + (oDirection.y * dLineLength)
    endZ = oPos.z + (oDirection.z * dLineLength)
    oLine.DefineBy2Points oPos.x, oPos.y, oPos.z, endX, endY, endZ

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True

    Dim oUpsidePort As IJPort
    Dim oUnkSurface As IUnknown
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(UpSide)
    Set oUnkSurface = oUpsidePort.Geometry

    ' Make sure the line is on the surface
    ' by taking the endpoint of the newly created line
    ' and projecting it into the surface.
    ' Output direction vector will indicate a succeeded projection
    ' and will demand the hole line to be projected aswell.
    Dim x As Double, y As Double, z As Double
    Dim oLineEndPoint As IJDPosition
    Set oLineEndPoint = New DPosition
    oLine.GetEndPoint x, y, z
    oLineEndPoint.Set x, y, z

    Dim oProjectionVector As IJDVector
    Dim oProjectedPosition As IJDPosition
    Set oProjectedPosition = oMfgRuleHelper.ProjectPointOnSurface(oLineEndPoint, oUnkSurface, oProjectionVector)

    'Dim oWireBody As IJWireBody
    'Set oWireBody = oMfgRuleHelper.ComplexStringToWireBody(oCS)

    'in case when mark is offset from the surface by margin value ,the mark
    'may not project on surface,so we need infinite geometry
    'Dim oUnboundedSurface As IJSurfaceBody
    'Set oUnboundedSurface = oPlateWrapper.GetUnboundedSurface(Upside)

    'Set oCS = Nothing
    If oProjectedDir Is Nothing Then
        Set oProjectedDir = oProjectionVector.Clone
    End If

    'Set oCS = oMfgRuleHelper.CurveAlongVectorOnToSurface(oUnkSurface, oWireBody, oProjectedDir)
    If oCS Is Nothing Then GoTo ErrorHandler
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    Dim oRevCS As IJComplexString
    oMfgMGHelper.ProjectComplexStringToSurface oCS, oUnkSurface, oProjectionVector, oRevCS
    Dim bRev As Boolean
    oMfgMGHelper.IsPositionedAtStart oRevCS, oPos, bRev
    Set oCS = Nothing
    If Not bRev Then
        oMfgMGHelper.ReverseComplexString oRevCS, oCS
    Else
        Set oCS = oRevCS
    End If
    Set oMark = oCS

    Create3DLineComplexString = True

    Set oMfgRuleHelper = Nothing

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Create3DLineComplexString = False
End Function


' ***********************************************************************************
' Public Function GetWebDirectionVector
'
' Description:  Method will get the web direction
' ***********************************************************************************
Public Function GetWebDirectionVector(oProfilePart As Object) As IJDVector
Const METHOD = "GetWebDirectionVector"
On Error GoTo ErrorHandler

    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart

    Dim oWBLandingCurve As IJWireBody
    Set oWBLandingCurve = oProfileWrapper.GetLandingCurve

    Dim oLCStartPos As IJDPosition, oLCEndPos As IJDPosition
    oWBLandingCurve.GetEndPoints oLCStartPos, oLCEndPos

    Dim oProfileSupport As IJProfilePartSupport
    Set oProfileSupport = New ProfilePartSupport

    Dim oPartSupport2    As IJPartSupport
    Set oPartSupport2 = oProfileSupport
    Set oPartSupport2.Part = oProfilePart

    ' From the start point, get the CrossSection orientation
    Dim oStartFlangeDirVec As IJDVector
    Dim oStartWebDirVec As IJDVector
    Dim oOrigin As IJDPosition
    oProfileSupport.GetOrientation oLCStartPos, oStartFlangeDirVec, oStartWebDirVec, oOrigin

    Set GetWebDirectionVector = oStartWebDirVec

    Set oPartSupport2 = Nothing
    Set oProfileSupport = Nothing

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetDirectionValue(sDirection As String) As String

    Select Case UCase(sDirection)

    Case "INBOARD", "IN"
        GetDirectionValue = "I"
    Case "OUTBOARD", "OUT"
        GetDirectionValue = "O"
    Case "UP", "UPPER", "ABOVE"
        GetDirectionValue = "U"
    Case "DOWN", "LOWER", "BELOW"
        GetDirectionValue = "D"
    Case "FORWARD", "FORE"
        GetDirectionValue = "F"
    Case "AFT"
        GetDirectionValue = "A"
    Case "PORT"
        GetDirectionValue = "P"
    Case "STARBOARD"
        GetDirectionValue = "S"
    Case Else
        GetDirectionValue = sDirection
    End Select


End Function

' ***********************************************************************************
' Method: IsALine
'
' Description:This Function is written to check whether given geometry is curved or
'             linear
' ***********************************************************************************
Public Function IsALine(oCS As Object) As Boolean
Const METHOD = "IsALine"
On Error GoTo ErrorHandler

    Dim oCurveElements As IJElements
    Set oCurveElements = New JObjectCollection

    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    oCurveElements.Add oCS

    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElements)

    Dim Length1 As Double, length2 As Double, length3 As Double
    Dim Points(1 To 4) As IJDPosition

    Set Points(1) = oBoxPoints.Item(1)
    Set Points(2) = oBoxPoints.Item(2)
    Set Points(3) = oBoxPoints.Item(3)
    Set Points(4) = oBoxPoints.Item(4)

    Length1 = Points(1).DistPt(Points(2))
    length2 = Points(2).DistPt(Points(3))
    length3 = Points(1).DistPt(Points(4))

    If (Length1 < 0.01 And length2 < 0.01) Or (length2 < 0.01 And length3 < 0.01) Or (Length1 < 0.01 And length3 < 0.01) Then
        IsALine = True
    End If

CleanUp:
    Set Points(1) = Nothing
    Set Points(2) = Nothing
    Set Points(3) = Nothing
    Set Points(4) = Nothing
    Set oBoxPoints = Nothing
    Set oCurveElements = Nothing
    Set oMfgGeomHelper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Create and plane given the orrt point and the normal.
'            Use the geometry factory to create a transient plane 3d object and return the plane
'--------------------------------------------------------------------------------------------------
Public Function CreatePlane(oRootPoint As IJDPosition, oNormalVec As IJDVector) As IJPlane
On Error GoTo ErrorHandler
Const METHOD = "CreatePlane"

    Dim oGeometryFactory    As IngrGeom3D.GeometryFactory
    Dim oPlane3D            As IngrGeom3D.IPlanes3d

    ' create persistent point
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oPlane3D = oGeometryFactory.Planes3d
    Set CreatePlane = oPlane3D.CreateByPointNormal(Nothing, oRootPoint.x, oRootPoint.y, oRootPoint.z, oNormalVec.x, oNormalVec.y, oNormalVec.z)

    Set oGeometryFactory = Nothing
    Set oPlane3D = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetButtConnectionData(ByVal oThisPart As Object, ByVal oReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As Collection

Const METHOD = "GetButtConnectionData"
On Error GoTo ErrorHandler

    Dim index As Long
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oConnection As IJAppConnection
    Dim aConnectionData As ConnectionData

    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = oThisPart

    Set GetButtConnectionData = New Collection

    For index = 1 To oReferenceObjColl.Count
        If TypeOf oReferenceObjColl.Item(index) Is IJStructPhysicalConnection Then
            Dim bIsCrossOfTee As Boolean
            Dim oConnType As ContourConnectionType

            Set oConnection = oReferenceObjColl.Item(index)
            oSDPartSupport.GetConnectionTypeForContour oConnection, _
                                                       oConnType, _
                                                       bIsCrossOfTee

            If oConnType = PARTSUPPORT_CONNTYPE_BUTT Then

               Dim oPortElements As IJElements
               oConnection.enumPorts oPortElements

               Dim oPort1 As IJPort
               Dim oport2 As IJPort

               Set oPort1 = oPortElements.Item(1)
               Set oport2 = oPortElements.Item(2)

               If (oPort1.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oPort1
                    Set aConnectionData.ToConnectable = oport2.Connectable
                    Set aConnectionData.ToConnectedPort = oport2
                    GetButtConnectionData.Add aConnectionData
               ElseIf (oport2.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oport2
                    Set aConnectionData.ToConnectable = oPort1.Connectable
                    Set aConnectionData.ToConnectedPort = oPort1
                    GetButtConnectionData.Add aConnectionData
               End If
            End If
        End If

    Next index

    Set oSDPartSupport = Nothing
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' ***********************************************************************************
' Private Sub CreateHoleCrossMarks
'
' Description:  Helper function to create hole mark for the input part.
'               Input arguments: this part, Wire Body of the marking line geometry, marking side.
'
' ***********************************************************************************
Public Sub CreateHoleCrossMarks(oThisPart As Object, _
                                    oCurve As IJWireBody, _
                                    UpSide As Long, _
                                    oFeature As IUnknown, _
                                    oGeom3dCustom As IJMfgGeom3D, _
                                    ByRef oGeomCol3d As MfgGeomCol3d)

    Const sMETHOD = "CreateHoleCrossMarks"
    On Error GoTo ErrorHandler

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oThisPart

    Dim oFeatureWireBody As IJWireBody
    Set oFeatureWireBody = oCurve

    Dim oWBUtil As IJSGOWireBodyUtilities
    Set oWBUtil = New SGOWireBodyUtilities

    Dim oMfgMGHelper As IJMfgMGHelper

    If oWBUtil.IsWireBodyClosed(oFeatureWireBody) = True Then

        Dim oCurveElems As IJElements
        Set oCurveElems = New JObjectCollection

        Dim oTempElems As IJElements
        Set oMfgMGHelper = New MfgMGHelper

        oMfgMGHelper.WireBodyToComplexStrings oFeatureWireBody, oTempElems
        oCurveElems.AddElements oTempElems

        Dim oCOG As IJDPosition
        Dim oProjPoint As IJDPosition
        Dim oProjVector As IJDVector

        Dim oSurfaceBody As IJSurfaceBody
        Set oSurfaceBody = oPlateWrapper.GetSurfaceWithoutFeatures(UpSide)

        oSurfaceBody.GetCenterOfGravity oCOG
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceBody, oCOG, oProjPoint, oProjVector

        Dim oSurfaceNormal As IJDVector
        oSurfaceBody.GetNormalFromPosition oProjPoint, oSurfaceNormal

        Dim oMfgGeomHelper As New MfgGeomHelper

        Dim oBoxPoints As IJElements
        Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElems, oSurfaceNormal)

        Dim oPoints(1 To 4) As IJDPosition

        Set oPoints(1) = oBoxPoints.Item(1)
        Set oPoints(2) = oBoxPoints.Item(2)
        Set oPoints(3) = oBoxPoints.Item(3)
        Set oPoints(4) = oBoxPoints.Item(4)

        Dim oPoints4 As IJDPosition
        Set oPoints4 = New DPosition

        oPoints4.x = oPoints(1).x + oPoints(3).x - oPoints(2).x
        oPoints4.y = oPoints(1).y + oPoints(3).y - oPoints(2).y
        oPoints4.z = oPoints(1).z + oPoints(3).z - oPoints(2).z

        Dim dLength(1 To 2) As Double
        dLength(1) = oPoints(1).DistPt(oPoints(2))
        dLength(2) = oPoints(2).DistPt(oPoints(3))

        If Abs(dLength(1) - dLength(2) < 0.0001) Then
            If IsALine(oCurveElems.Item(1)) = False Then

                Dim oBoxVector As IJDVector
                Set oBoxVector = oPoints(1).Subtract(oPoints(2))
                oBoxVector.Length = 1

                Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
                Set oXVector = New DVector
                Set oYVector = New DVector
                Set oZVector = New DVector

                oXVector.Set 1, 0, 0
                oYVector.Set 0, 1, 0
                oZVector.Set 0, 0, 1

                If Not (Abs(oBoxVector.Dot(oXVector)) > 0.999999 Or Abs(oBoxVector.Dot(oYVector)) > 0.999999 Or Abs(oBoxVector.Dot(oZVector)) > 0.999999) Then
                    Dim oVecElems As IJElements
                    Set oVecElems = New JObjectCollection

                    Dim dXDotP As Double, dYDotP As Double, dZDotP As Double

                    ' Take absolute value of dot product as surface normal can be either positive or negative
                    dXDotP = Abs(oSurfaceNormal.Dot(oXVector))
                    dYDotP = Abs(oSurfaceNormal.Dot(oYVector))
                    dZDotP = Abs(oSurfaceNormal.Dot(oZVector))

                    'The order in which vectors are added to the collection is important.
                    ' The vector aligning with the surface normal need to be added at the end of collection
                    If ((dXDotP > dYDotP) And (dXDotP > dZDotP)) Then
                        oVecElems.Add oYVector
                        oVecElems.Add oZVector
                        oVecElems.Add oXVector
                    ElseIf ((dYDotP > dXDotP) And (dYDotP > dZDotP)) Then
                        oVecElems.Add oXVector
                        oVecElems.Add oZVector
                        oVecElems.Add oYVector
                    Else
                        oVecElems.Add oXVector
                        oVecElems.Add oYVector
                        oVecElems.Add oZVector
                    End If

                    ' Get Min bounding box oiented in global axes only when surface normal aligned to one of global axes
                    If dXDotP > 0.99 Or dYDotP > 0.99 Or dZDotP > 0.99 Then
                        Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVectors(oCurveElems, oVecElems)
                    Else
                        Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElems, oSurfaceNormal)
                    End If

                    Set oPoints(1) = oBoxPoints.Item(1)
                    Set oPoints(2) = oBoxPoints.Item(2)
                    Set oPoints(3) = oBoxPoints.Item(3)
                    Set oPoints(4) = oBoxPoints.Item(4)

                    'Checking for getting the three BoxPoints which lies on the same plane
                    'Check the distance if it is same use the 4th point
                    If Abs(oPoints(1).DistPt(oPoints(2))) < 0.001 Then
                        Set oPoints(1) = oPoints(4)
                    ElseIf Abs(oPoints(2).DistPt(oPoints(3))) < 0.001 Then
                        Set oPoints(2) = oPoints(4)
                    ElseIf Abs(oPoints(3).DistPt(oPoints(1))) < 0.001 Then
                        Set oPoints(3) = oPoints(4)
                    End If

                    oPoints4.x = oPoints(1).x + oPoints(3).x - oPoints(2).x
                    oPoints4.y = oPoints(1).y + oPoints(3).y - oPoints(2).y
                    oPoints4.z = oPoints(1).z + oPoints(3).z - oPoints(2).z

                End If
            End If
        End If

        Dim oCrossPoint1 As IJDPosition
        Set oCrossPoint1 = New DPosition

        oCrossPoint1.x = (oPoints(1).x + oPoints(2).x) / 2
        oCrossPoint1.y = (oPoints(1).y + oPoints(2).y) / 2
        oCrossPoint1.z = (oPoints(1).z + oPoints(2).z) / 2

        Dim oCrossPoint2 As IJDPosition
        Set oCrossPoint2 = New DPosition

        oCrossPoint2.x = (oPoints(2).x + oPoints(3).x) / 2
        oCrossPoint2.y = (oPoints(2).y + oPoints(3).y) / 2
        oCrossPoint2.z = (oPoints(2).z + oPoints(3).z) / 2

        Dim oCrossPoint3 As IJDPosition
        Set oCrossPoint3 = New DPosition

        oCrossPoint3.x = (oPoints(3).x + oPoints4.x) / 2
        oCrossPoint3.y = (oPoints(3).y + oPoints4.y) / 2
        oCrossPoint3.z = (oPoints(3).z + oPoints4.z) / 2

        Dim oCrossPoint4 As IJDPosition
        Set oCrossPoint4 = New DPosition

        oCrossPoint4.x = (oPoints(1).x + oPoints4.x) / 2
        oCrossPoint4.y = (oPoints(1).y + oPoints4.y) / 2
        oCrossPoint4.z = (oPoints(1).z + oPoints4.z) / 2

        Dim oCrossLine1 As IJLine
        Dim oCrossLine2 As IJLine
        Set oCrossLine1 = New Line3d
        Set oCrossLine2 = New Line3d

        oCrossLine1.DefineBy2Points oCrossPoint1.x, oCrossPoint1.y, oCrossPoint1.z, oCrossPoint3.x, oCrossPoint3.y, oCrossPoint3.z
        oCrossLine2.DefineBy2Points oCrossPoint2.x, oCrossPoint2.y, oCrossPoint2.z, oCrossPoint4.x, oCrossPoint4.y, oCrossPoint4.z

        Dim dVecX   As Double, dVecY   As Double, dVecZ   As Double
        Dim bNameFirstCS    As Boolean
        Dim strCrossSecName As String
        If oCrossLine1.Length < HOLE_MINIMUM_DIAMETER Or oCrossLine2.Length < HOLE_MINIMUM_DIAMETER Then
            If Abs(oCrossLine1.Length - oCrossLine2.Length) < 0.001 Then
                strCrossSecName = CInt(oCrossLine1.Length * 1000)

                oCrossLine1.GetDirection dVecX, dVecY, dVecZ

                If Abs(dVecX) > 0.5 Then
                    bNameFirstCS = True
                End If
            Else
                strCrossSecName = CInt(oCrossLine1.Length * 1000) & "X" & CInt(oCrossLine2.Length * 1000)
                If Abs(oCrossLine1.Length) > Abs(oCrossLine2.Length) Then
                    bNameFirstCS = True
                End If
            End If
        Else
            strCrossSecName = " "
        End If

        Dim oCS1 As IJComplexString
        Dim oCS2 As IJComplexString
        Set oCS1 = New ComplexString3d
        Set oCS2 = New ComplexString3d

        oCS1.AddCurve oCrossLine1, True
        oCS2.AddCurve oCrossLine2, True

        Dim oTempCS1    As IJComplexString
        Dim oTempCS2    As IJComplexString

        oMfgMGHelper.ProjectComplexStringToSurface oCS1, oSurfaceBody, Nothing, oTempCS1
        oMfgMGHelper.ProjectComplexStringToSurface oCS2, oSurfaceBody, Nothing, oTempCS2

        Set oCS1 = Nothing
        Set oCS1 = oTempCS1

        Set oCS2 = Nothing
        Set oCS2 = oTempCS2

        Dim oMfgGeomUtilwrapper1 As New MfgGeomUtilWrapper
        oMfgGeomUtilwrapper1.ExtendWire oCS1, 0.05
        Dim oMfgGeomUtilwrapper2 As New MfgGeomUtilWrapper
        oMfgGeomUtilwrapper2.ExtendWire oCS2, 0.05

        Dim oCS As IJComplexString
        Dim oSystemMark As IJMfgSystemMark
        Dim oMoniker As IMoniker
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3d As IJMfgGeom3D

        Dim jCount As Integer
        For jCount = 1 To 2
            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            'Set the marking side
            oSystemMark.SetMarkingSide UpSide

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            If jCount = 1 Then
                oGeom3d.PutGeometry oCS1
                If bNameFirstCS = True Then
                    oMarkingInfo.Name = strCrossSecName
                End If
            Else
                oGeom3d.PutGeometry oCS2
                If bNameFirstCS = False Then
                    oMarkingInfo.Name = strCrossSecName
                End If
            End If

            If oCrossLine1.Length < HOLE_MINIMUM_DIAMETER Or oCrossLine2.Length < HOLE_MINIMUM_DIAMETER Then
                oGeom3d.PutGeometrytype STRMFG_HOLE_TRACE_MARK  'Hole Trace Mark
            Else
               oGeom3d.PutGeometrytype STRMFG_HOLE_REF_MARK  'Hole Ref Mark
            End If

            oGeomCol3d.AddGeometry 1, oGeom3d
            oSystemMark.Set3dGeometry oGeom3d

            If oGeom3dCustom Is Nothing Then
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oFeature)
            Else
                Set oMoniker = oGeom3dCustom.GetMoniker
            End If

            oGeom3d.PutMoniker oMoniker

        Next jCount

        Set oTempCS1 = Nothing
        Set oTempCS2 = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oMoniker = Nothing

    End If


    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ***********************************************************************************
' Private Sub GetPhysicalConnPortFromERParts
'
' Description:  Helper function to get the PC port for the input ER part and related connecton data collection.
'               Input arguments: reference ER part, connecton data collection and connected Platepart of reference ER.
'
' ***********************************************************************************
Private Function GetPhysicalConnPortFromERParts(ByVal oThisPart As Object, ByVal oConObjsCol As Collection, ByVal oPlatePart As Object, ByRef bCircularER As Boolean) As IJElements
    Const METHOD = "GetPhysicalConnPortFromERParts"
    On Error GoTo ErrorHandler

    Dim lConnections As Long, lIndex As Long
    lConnections = oConObjsCol.Count

    Dim oOtherPart As Object
    Dim oConnectionData As ConnectionData

    Dim oTempColl As IJElements
    Set oTempColl = New JObjectCollection

    For lIndex = 1 To lConnections

        oConnectionData = oConObjsCol.Item(lIndex)
        Set oOtherPart = Nothing
        Set oOtherPart = oConnectionData.ToConnectable

        ' check if the input part is ER
        If TypeOf oOtherPart Is IJProfileER And Not oOtherPart Is oThisPart Then

            Dim icount As Integer
            Dim oConnectable As IJConnectable
            Set oConnectable = oThisPart

            Dim bIsConnected As Boolean
            Dim oConnections As IJElements

            ' check if the other part(in collection) is connected to reference ER
            oConnectable.isConnectedTo oOtherPart, bIsConnected, oConnections

            If bIsConnected = True Then
                ' Get the PC and return
                For icount = 1 To oConnections.Count
                    If TypeOf oConnections.Item(icount) Is IJStructPhysicalConnection Then
                        Dim oAppConn As IJAppConnection
                        Set oAppConn = oConnections.Item(icount)

                        Dim oPorts As IJElements
                        oAppConn.enumPorts oPorts

                        oTempColl.Add oPorts.Item(1)

                        If oConnections.Count > 2 Then
                            bCircularER = True
                                                Else
                            GoTo NextConn
                        End If

                    End If
                Next

            End If
        End If
NextConn:
    Next

    If oTempColl.Count > 0 Then
        Set GetPhysicalConnPortFromERParts = oTempColl
    End If

CleanUp:
    Set oConnectable = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function




Public Function CreateConnPartInfoMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, Optional IGNORE_CHAMFER As Boolean = False) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ConnPartMark: CreateConnPartInfoMark"

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart

    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If

    '*** Get base surface geometry of this plate ***'
    Dim oPlateBasePort As IJPort
    Dim oPartBaseSurface As IUnknown
    Set oPlateBasePort = oSDPlateWrapper.BasePort(BPT_Base)
    Set oPartBaseSurface = oPlateBasePort.Geometry

    If IGNORE_CHAMFER = True Then
    'Get Surface without Features
        Set oPartBaseSurface = oPlateWrapper.GetSurfaceWithoutFeatures(PlateBaseSide)
    End If

    '*** Get ModelBody from geometry ***'
    Dim oPlateBaseModelBody As IJDModelBody
    Set oPlateBaseModelBody = oPartBaseSurface
    'Dim PBFileName As String
    'PBFileName = Environ("TEMP")
    'If PBFileName = "" Or PBFileName = vbNullString Then
    '    PBFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'PBFileName = PBFileName & "\oPlateBase.sat"
    'oPlateBaseModelBody.DebugToSATFile PBFileName
    '***********************************************'

    '*** Get base surface geometry of this plate ***'
    Dim oPlateOffsetPort As IJPort
    Dim oPartOffsetSurface As IUnknown
    Set oPlateOffsetPort = oSDPlateWrapper.BasePort(BPT_Offset)
    Set oPartOffsetSurface = oPlateOffsetPort.Geometry

    If IGNORE_CHAMFER = True Then
    'Get Surface without Features
        Set oPartOffsetSurface = oPlateWrapper.GetSurfaceWithoutFeatures(PlateOffsetSide)
    End If

    '*** Get ModelBody from geometry ***'
    Dim oPlateOffsetModelBody As IJDModelBody
    Set oPlateOffsetModelBody = oPartOffsetSurface
    'Dim POFileName As String
    'POFileName = Environ("TEMP")
    'If POFileName = "" Or POFileName = vbNullString Then
    '    POFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'POFileName = POFileName & "\oPlateOffset.sat"
    'oPlateOffsetModelBody.DebugToSATFile POFileName
    '***********************************************'

    '*** Get Plate Surface Normal ***'
    Dim oPlateNormal As IJDVector
    'Set oPlateNormal = GetPlateNeutralSurfaceNormal(Part)
    '********************************'

    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetButtConnectionData(Part, ReferenceObjColl)

    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If

    Dim Item As Object
    Dim oConnectionData As ConnectionData

    Dim nIndex As Long, nWBIndex As Long
    Dim oSDConnPlateWrapper As New StructDetailObjects.PlatePart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn

    Dim dThickness As Double, dConnThickness As Double
    Dim oSupp As IJPartSupport
    Dim oSys As IJSystem
    Dim oNamedItem As IJNamedItem

    Set oSupp = New PartSupport
    Set oSupp.Part = Part
    oSupp.IsSystemDerivedPart oSys, True

    Set oNamedItem = oSys

    '*** Get this plate's thickness ***'
    dThickness = oSDPlateWrapper.PlateThickness

    '*** Loop thru each Physical Connections ***'
    Dim bContourEnd As Boolean
    Dim oWBColl As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim dBaseBaseDist As Double, dOffsetOffsetDist As Double
    Dim dOffsetBaseDist As Double, dBaseOffsetDist As Double, dDistDifference As Double
    Dim oPos1 As New DPosition, oPos2 As New DPosition, dMinDist As Double
    Dim oBasePos As New DPosition
    Dim eGeomType As StrMfgGeometryType

    '*** Create an instance of the StrMfg math geom helper ***'
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    lGeomCount = 1

    '*** Collection of Connected Part Marks ***'
    Dim oConnGeomCol3d As IJMfgGeomCol3d
    Set oConnGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    '*** Collection of APS Custom Marks ***'
    Dim oGeomCol3DCustom As IJMfgGeomCol3d
    Set oGeomCol3DCustom = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_CONN_PART_MARK, ReferenceObjColl, oGeomCol3DCustom

    Dim oResourceManager As Object

    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Set oPlateCreation_AE = oMfgPart.ActiveEntity

    '*** Collection of Outer Contours ***'
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = oPlateCreation_AE.GeometriesBeforeUnfold

    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)

        Dim oPort As IJPort
        Set oPort = oConnectionData.ConnectingPort
        Dim oMB As IJDModelBody
        Set oMB = oPort.Geometry
        'Dim PC2FileName As String
        'PC2FileName = Environ("TEMP")
        'If PC2FileName = "" Or PC2FileName = vbNullString Then
        '    PC2FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'PC2FileName = PC2FileName & "\PhyConn2.sat"
        'oMB.DebugToSATFile PC2FileName

        '*** Check if the connected object is a plate ***'
        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
            Exit Function
        End If

        '*** Initialize the plate wrapper and the Physical Connection wrapper ***'
        Set oSDConnPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDConnPlateWrapper.object = oConnectionData.ToConnectable

        '*** Get System Name of connected plate part ***'
        Set oSupp = New PartSupport
        Set oSupp.Part = oConnectionData.ToConnectable
        oSupp.IsSystemDerivedPart oSys, True

        Set oNamedItem = oSys
        Set oNamedItem = oConnectionData.ToConnectable
        '***********************************************'

        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        Dim bChamfer As Boolean
        Dim oChamfer As Object

        '*** Get connected plate's thickness ***'
        dConnThickness = oSDConnPlateWrapper.PlateThickness

        '*** Check if base surfaces are aligned ***'

        Dim oConnPlateWrapper As MfgRuleHelpers.PlatePartHlpr
        Set oConnPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
        Set oConnPlateWrapper.object = oConnectionData.ToConnectable


        'Get base surface geometry of connected plate
        Dim oConnPlateBasePort As IJPort
        Dim oConnPartBaseSurface As IUnknown, oConnPartOffsetSurface As IUnknown
        Dim oConnPlateModelBody As IJDModelBody
        Dim oConnPlateOffsetPort As IJPort
        '***********************************************'

        '*** Get BaseBaseDist ***'
        Set oConnPlateBasePort = oSDConnPlateWrapper.BasePort(BPT_Base)
        Set oConnPartBaseSurface = oConnPlateBasePort.Geometry
        If IGNORE_CHAMFER = True Then
        'Get Surface without features
            Set oConnPartBaseSurface = oConnPlateWrapper.GetSurfaceWithoutFeatures(PlateBaseSide)
        End If

        Set oConnPlateModelBody = oConnPartBaseSurface

        'Dim CPBSFileName As String
        'CPBSFileName = Environ("TEMP")
        'If CPBSFileName = "" Or CPBSFileName = vbNullString Then
        '    CPBSFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'CPBSFileName = CPBSFileName & "\oConnPartBaseSurface.sat"
        'oConnPlateModelBody.DebugToSATFile CPBSFileName
        oPlateBaseModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dBaseBaseDist

        '*** Get OffsetBaseDist ***'
        oPlateOffsetModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dOffsetBaseDist

        '*** Get OffsetOffsetDist ***'
        Set oConnPlateOffsetPort = oSDConnPlateWrapper.BasePort(BPT_Offset)
        Set oConnPartOffsetSurface = oConnPlateOffsetPort.Geometry

        If IGNORE_CHAMFER = True Then
        'Get Surface without features
            Set oConnPartOffsetSurface = oConnPlateWrapper.GetSurfaceWithoutFeatures(PlateOffsetSide)
        End If

        Set oConnPlateModelBody = oConnPartOffsetSurface

        'Dim CPOSFileName As String
        'If CPOSFileName = "" Or CPOSFileName = vbNullString Then
        '    CPOSFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'CPOSFileName = CPOSFileName & "\oConnPartOffsetSurface.sat"
        'oConnPlateModelBody.DebugToSATFile CPOSFileName
        oPlateOffsetModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dOffsetOffsetDist

        '*** Get BaseOffsetDist ***'
        oPlateBaseModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dBaseOffsetDist

        '**********************************************'

        '*** Check if Molded surfaces have difference ***'
        Dim sSide As String
        Dim sConnPlateSide As String
        sSide = oSDPlateWrapper.AlternateMoldedSide
        sConnPlateSide = oSDConnPlateWrapper.AlternateMoldedSide

        Dim oMfgGeomHelper As MfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper

        Dim bIsBaseMolded1 As Boolean, bIsBaseMolded2 As Boolean
        Dim oPlateMoldedSurface As IJSurfaceBody, oConnPlateMoldedSurface As IJSurfaceBody

        If sSide = "Base" Then
            Set oPlateMoldedSurface = oPartBaseSurface
        ElseIf sSide = "Offset" Then
            Set oPlateMoldedSurface = oPartOffsetSurface
        End If

        If sConnPlateSide = "Base" Then
            Set oConnPlateMoldedSurface = oConnPartBaseSurface
        ElseIf sConnPlateSide = "Offset" Then
            Set oConnPlateMoldedSurface = oConnPartOffsetSurface
        End If


'        Set oPlateMoldedSurface = oMfgGeomHelper.GetSurfaceFromPlatePart(Part, bIsBaseMolded1)
'        Set oConnPlateMoldedSurface = oMfgGeomHelper.GetSurfaceFromPlatePart(oConnectionData.ToConnectable, bIsBaseMolded2)
'
        Dim oPlateMoldedSurfaceMB As IJDModelBody, oConnPlateMoldedSurfaceMB As IJDModelBody
        Set oPlateMoldedSurfaceMB = oPlateMoldedSurface
        Set oConnPlateMoldedSurfaceMB = oConnPlateMoldedSurface

        Dim dMoldedSurfaceDist As Double
        oPlateMoldedSurfaceMB.GetMinimumDistance oConnPlateMoldedSurfaceMB, oPos1, oPos2, dMoldedSurfaceDist

        'Dim PMSFileName As String
        'Dim CPMSFileName As String
        'PMSFileName = Environ("TEMP")
        'CPMSFileName = Environ("TEMP")
        'If PMSFileName = "" Or PMSFileName = vbNullString Then
        '    PMSFileName = "C:\temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'If CPMSFileName = "" Or CPMSFileName = vbNullString Then
        '    CPMSFileName = "C:\temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'PMSFileName = PMSFileName & "\oPlateMoldedSurfaceMB.sat"
        'CPMSFileName = CPMSFileName & "\oConnPlateMoldedSurfaceMB.sat"
'        oPlateMoldedSurfaceMB.DebugToSATFile PMSFileName
'        oConnPlateMoldedSurfaceMB.DebugToSATFile CPMSFileName

        '************************************************'

        'Removed the chamfer condition
        If dThickness > dConnThickness Then  'And Not oSDPhysicalConn.FromChamfer(oChamfer) Then

            '*** Get difference between Upside ***'
            If UpSide = BaseSide Then
                dDistDifference = IIf(dBaseBaseDist < dBaseOffsetDist, dBaseBaseDist, dBaseOffsetDist)
            ElseIf UpSide = OffsetSide Then
                dDistDifference = IIf(dOffsetOffsetDist < dOffsetBaseDist, dOffsetOffsetDist, dOffsetBaseDist)
            End If

            '*** Get MidPoint/Centroid of PC ***'
            Dim oMidPt As New DPosition

            Dim oStructConnection As IJStructConnection
            Set oStructConnection = oSDPhysicalConn.object
            Dim oPCWB As IJWireBody
            Set oPCWB = oStructConnection.ModelBody

            Dim ppCurves As IJElements
            oMfgMGHelper.WireBodyToComplexStrings oPCWB, ppCurves

            Dim oCurve As IJCurve
            Set oCurve = ppCurves.Item(1)

            Dim dX As Double, dY As Double, dZ As Double
            Dim oCentroidPt As New DPosition

            oCurve.Centroid dX, dY, dZ
            oCentroidPt.Set dX, dY, dZ
            '***************************'

            '*** Find contour close to PC ***'
            Dim i As Long
            Dim j As Long
            Dim dMinimum As Double
            Dim iMinIndex As Integer

            dMinimum = 999
            For i = 1 To oGeomCol3d.GetCount
                'Dim oGeom3d As IJMfgGeom3D
                Set oGeom3d = oGeomCol3d.GetGeometry(i)
                eGeomType = oGeom3d.GetGeometryType

                If eGeomType = STRMFG_OUTER_CONTOUR Then
                    Dim oMB1 As IJDModelBody
                    Dim oCString As IComplexStrings3d
                    Dim oWB123 As IJWireBody

                    Set oWB123 = m_oMfgRuleHelper.ComplexStringToWireBody(oGeom3d.GetGeometry)
                    Set oMB1 = oWB123

                    oMB1.GetMinimumDistanceFromPosition oCentroidPt, oPos1, dMinDist

                    If dMinDist < dMinimum Then
                        Set oBasePos = oPos1
                        dMinimum = dMinDist
                        iMinIndex = i
                    End If
                End If
            Next

            Set oGeom3d = oGeomCol3d.GetGeometry(iMinIndex)
            'This is PC curve/edge
            Set oCurve = oGeom3d.GetGeometry

            '*** Get Tangent to PC MidPoint ***'
'            ________________________
'                                    .
'                                    .
'                                    .
'                                    ^
'                      This is    | /|\
'                      the Mark   |  |   PC Centroid & Tangent
'                                    .
'                                    .
'            ________________________.

            Dim dStartPar As Double, dEndPar As Double, dOffsetPar As Double
            Dim vTanX As Double, vTanY As Double, vTanZ As Double
            Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
            Dim dStartX As Double, dStartY As Double, dStartZ As Double
            Dim dParam As Double
            oCurve.Parameter oBasePos.x, oBasePos.y, oBasePos.z, dParam
            oCurve.Evaluate dParam, dStartX, dStartY, dStartZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

            Dim oTanVec As New DVector, oVecOnSurface As New DVector
            oTanVec.Set vTanX, vTanY, vTanZ
'            oTanVec.Length = 0.05
            '***********************************'

            'Get Normal at the base point
            Dim oSolidBody As IJSurfaceBody
            Dim oToolBox As New DGeomOpsToolBox
            Dim oProjBasePos As IJDPosition
            Dim oProjBaseVector As IJDVector
            If UpSide = BaseSide Then
                Set oSolidBody = oPartBaseSurface
            Else
                Set oSolidBody = oPartOffsetSurface
            End If
            oToolBox.ProjectPointOnSurfaceBody oSolidBody, oBasePos, oProjBasePos, oProjBaseVector
            oSolidBody.GetNormalFromPosition oProjBasePos, oPlateNormal

            If Not oPlateNormal Is Nothing Then
                Set oVecOnSurface = oPlateNormal.Cross(oTanVec)

                'Offseting the PC Centrioid
                oVecOnSurface.Length = 0.025

                Set oTanVec = oVecOnSurface.Cross(oPlateNormal)

                oTanVec.Length = 0.025
            End If


            Dim oGeomFac As New GeometryFactory
            Dim oCrvElemets As IJElements

            Set oCrvElemets = New JObjectCollection

            Dim oLines3d As ILines3d
            Dim oLine As Line3d
            Dim oComplexStrings3d As IComplexStrings3d

            Set oLines3d = oGeomFac.Lines3d
            Set oLine = oLines3d.CreateBy2Points(Nothing, oProjBasePos.x + oVecOnSurface.x, oProjBasePos.y + oVecOnSurface.y, oProjBasePos.z + oVecOnSurface.z, oProjBasePos.x + oVecOnSurface.x + oTanVec.x, oProjBasePos.y + oVecOnSurface.y + oTanVec.y, oProjBasePos.z + oVecOnSurface.z + oTanVec.z)

            oCrvElemets.Add oLine
            Set oComplexStrings3d = oGeomFac.ComplexStrings3d

            Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

            'Set the marking side
            oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            oMarkingInfo.SetAttributeNameAndValue "MoldedSideDifference", dMoldedSurfaceDist
            oMarkingInfo.SetAttributeNameAndValue "BaseBaseDifference", dBaseBaseDist
            oMarkingInfo.SetAttributeNameAndValue "BaseOffsetDifference", dBaseOffsetDist
            oMarkingInfo.SetAttributeNameAndValue "UpSideDifference", dDistDifference
            oMarkingInfo.SetAttributeNameAndValue "UpSide", UpSide
            oMarkingInfo.SetAttributeNameAndValue "OffsetBaseDifference", dOffsetBaseDist
            oMarkingInfo.SetAttributeNameAndValue "OffsetOffsetDifference", dOffsetOffsetDist
            oMarkingInfo.ThicknessDirection = oVecOnSurface

            Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_CONN_PART_MARK

            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker
            oGeom3d.IsSupportOnly = True

            oSystemMark.Set3dGeometry oGeom3d

            oConnGeomCol3d.AddGeometry lGeomCount, oGeom3d
            lGeomCount = lGeomCount + 1

        End If

NextItem:
        Set oBasePos = Nothing
        Set oWB = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        'Set oSDPlateWrapper = Nothing
        'Set oSDPhysicalConn = Nothing

    Next nIndex

    Dim lCustomMarkCounter  As Long
    Dim lGeomCountCustom    As Long
    Dim oGeom3dCustom       As IJMfgGeom3D

    If Not oGeomCol3DCustom Is Nothing Then
        lGeomCountCustom = oGeomCol3DCustom.GetCount
        lGeomCount = oConnGeomCol3d.GetCount + 1

        For lCustomMarkCounter = 1 To lGeomCountCustom
            'Get the Geom3d object from Geom3dCollection
            Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(lCustomMarkCounter)
            oConnGeomCol3d.AddGeometry lGeomCount, oGeom3dCustom
            lGeomCount = lGeomCount + 1
        Next lCustomMarkCounter

        ' Get rid of oGeomCol3DCustom and oGeom3dCustom
        Dim oGeomObject As IJDObject
        While oGeomCol3DCustom.GetCount > 0
            Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(1)
            oGeomCol3DCustom.RemoveGeometry oGeom3dCustom
        Wend
    End If

    'Return the 3d collection
    Set CreateConnPartInfoMark = oConnGeomCol3d

    Set oCS = Nothing

CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oMfgMGHelper = Nothing
    Set oSDPhysicalConn = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1009, , "RULES")
    GoTo CleanUp
End Function

' ***********************************************************************************
' CreateSeamControlMark3D
'
' Description:  This function creates the Seam control marks based on given inputs.
'
' ***********************************************************************************
Public Function CreateSeamControlMark3D(ByVal Part As Object, ByVal UpSide As Long, _
                                        ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                        Optional ByVal bProcessAllBUTTConns As Boolean) As IJMfgGeomCol3d

    Const METHOD = "CreateSeamControlMark3D"
    On Error GoTo ErrorHandler

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim pResMgr As IUnknown
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(pResMgr)

    CreateAPSMarkings STRMFG_SEAM_MARK, ReferenceObjColl, oGeomCol3d
    Set CreateSeamControlMark3D = oGeomCol3d

    Dim Item As Object
    Dim oConnectionData As ConnectionData

    On Error Resume Next
    Dim nConnectionData As Long
    Dim oConObjsCol As Collection
    Dim oGeomOpsToolBox As IJDTopologyToolBox

    Set oGeomOpsToolBox = New DGeomOpsToolBox

    nConnectionData = 0

    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)

    nConnectionData = oConObjsCol.Count
    If nConnectionData < 1 Then
        Exit Function
    End If

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If

    Dim oUpsidePort As IJPort
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(UpSide)

    Dim bConnWithProfile As Boolean
    Dim nIndex As Variant
    Dim projInd As Integer

    Dim oSurfaceWithOutFeatures     As IJSurfaceBody
    Set oSurfaceWithOutFeatures = oPlateWrapper.GetSurfaceWithoutFeatures(UpSide)

    'Can't use "For Each" when user defined types in collection
    For nIndex = 1 To nConnectionData

        oConnectionData = oConObjsCol.Item(nIndex)

        'Check if the connected object is a PlatePart or bProcessAllBUTTConns = True else goto next item
        If TypeOf oConnectionData.ToConnectable Is IJPlatePart Or bProcessAllBUTTConns = True Then
            ' Check if it is a butt-to-butt connection (both are connected with lateral ports)

            bConnWithProfile = False
            ' If bAllBUTTConns = True then process Plate-Profile connections also
            If bProcessAllBUTTConns = True Then
                If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
                    bConnWithProfile = True
                End If
            End If

            Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oConnectionData.AppConnection

            Dim rootgapvalue As Double

            oSDPhysicalConn.GetBevelParameterValue "RootGap", rootgapvalue, DoubleType

            Dim res1 As Integer
            Dim res2 As Integer

            Dim port1Flag As IMSStructConnection.eUSER_CTX_FLAGS
            Dim port2Flag As IMSStructConnection.eUSER_CTX_FLAGS

            Dim oStructPort1 As IJStructPort, oStructPort2 As IJStructPort
            Set oStructPort1 = oConnectionData.ConnectingPort
            Set oStructPort2 = oConnectionData.ToConnectedPort

            port1Flag = oStructPort1.ContextID
            port2Flag = oStructPort2.ContextID

            res1 = port1Flag And IMSStructConnection.CTX_LATERAL

            ' Check If the Plate is connected with profile
            If bConnWithProfile = True Then
                res2 = port2Flag And IMSStructConnection.CTX_BASE
                If Not res2 > 0 Then
                    res2 = port2Flag And IMSStructConnection.CTX_OFFSET
                End If
            Else
                res2 = port2Flag And IMSStructConnection.CTX_LATERAL
            End If

            If res1 > 0 And res2 > 0 Then
                ' Get connecting side port
                Dim oConnectingSidePort As IJPort
                Set oConnectingSidePort = oConnectionData.ConnectingPort

                Dim oCommonGeometry As IUnknown
                Dim oPhyConGeom As Object
                Dim oPhyConGeomColl As IJDObjectCollection
                Dim oPhyConCurve As IJCurve
                Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
                Dim oPhyConPos As New DPosition
                Dim oProjectedPhyConPos As IJDPosition
                Dim dDist As Double
                Dim pNormal As IJDVector
                Dim oMfgGeomHelper As New MfgGeomHelper
                Dim oTouchingComplexStringColl As IJElements

                Set oTouchingComplexStringColl = New JObjectCollection

                Set oPhyConGeomColl = oSDPhysicalConn.GetConnectionGeometries

                For Each oPhyConGeom In oPhyConGeomColl
                    Set oPhyConCurve = oPhyConGeom
                    oPhyConCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

                    oPhyConPos.Set dStartX, dStartY, dStartZ

                    oGeomOpsToolBox.ProjectPointOnSurfaceBody oUpsidePort.Geometry, oPhyConPos, _
                                                              oProjectedPhyConPos, pNormal
                    Err.Clear

                    If oProjectedPhyConPos Is Nothing Then
                        dDist = 1000
                    Else
                        dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
                    End If

                    If dDist < 0.001 Then
                        Set oProjectedPhyConPos = Nothing
                        Set pNormal = Nothing

                        oPhyConPos.Set dEndX, dEndY, dEndZ
                        oGeomOpsToolBox.ProjectPointOnSurfaceBody oUpsidePort.Geometry, _
                                oPhyConPos, oProjectedPhyConPos, pNormal
                        Err.Clear

                        If oProjectedPhyConPos Is Nothing Then
                            dDist = 1000
                        Else
                            dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
                        End If

                        If dDist < 0.001 Then
                            oTouchingComplexStringColl.Add m_oMfgRuleHelper.ComplexStringToWireBody(oPhyConGeom)
                        End If
                    End If
                Next

                If oTouchingComplexStringColl Is Nothing Then GoTo NextItem

                If oTouchingComplexStringColl.Count = 0 Then GoTo NextItem

                Dim oMergedCSColl As IJElements
                Set oMergedCSColl = oMfgGeomHelper.MergeCollectionToComplexStrings(oTouchingComplexStringColl)

                If Not oMergedCSColl Is Nothing Then
                    If oMergedCSColl.Count = 0 Then
                        Set oMergedCSColl = Nothing
                    End If
                End If

                If oMergedCSColl Is Nothing Then
                    GoTo NextItem
                End If

                Dim oMergedCS As IJComplexString
                For Each oMergedCS In oMergedCSColl

                    Set oCommonGeometry = m_oMfgRuleHelper.ComplexStringToWireBody(oMergedCS)

                    If oCommonGeometry Is Nothing Then GoTo NextMergedCS

                    Dim oEdgeWire As IJWireBody
                    Set oEdgeWire = oCommonGeometry
                    Set oCommonGeometry = Nothing

                    ' OffSet the wirebody from the edge into the platesurface as a fixed distance
                    ' Get input parameters
                    Dim oOffSetDirection            As IJDVector
                    Dim oConnectingSideSurface      As IJSurfaceBody

                    Set oConnectingSideSurface = oConnectingSidePort.Geometry

                    ' get a direction from somewhere on the edge
                    Dim oMidPosition As IJDPosition
                    Dim oTangent1 As IJDVector, oNormalDirection As IJDVector, oCrossVec As IJDVector

                    Set oMidPosition = m_oMfgRuleHelper.GetMiddlePoint(oEdgeWire)
                    Set oTangent1 = oMfgGeomHelper.GetTangentByPointOnCurve(oEdgeWire, oMidPosition)

                    Dim oProjPoint As IJDPosition
                    Set oProjPoint = m_oMfgRuleHelper.ProjectPointOnSurface(oMidPosition, oSurfaceWithOutFeatures, oNormalDirection)

                    Set oCrossVec = oTangent1.Cross(oNormalDirection)
                    m_oMfgRuleHelper.ScaleVector oNormalDirection, -1

                    ' this gets the direction pointing outwards
                    oConnectingSideSurface.GetNormalFromPosition oMidPosition, oOffSetDirection

                    Dim dOffSet As Double
                    Dim oUnkTransientSheet As IUnknown
                    dOffSet = GetSeamDistance

                    ' Note that often the rootgapvalue is a negative value and therefore
                    ' I need to add the value to the offset in order to substract
                    If rootgapvalue > 0 Then
                       dOffSet = dOffSet + rootgapvalue
                    Else
                       If rootgapvalue < 0 Then
                          dOffSet = dOffSet - rootgapvalue
                       End If
                    End If

                    ' In case the offset value is less or equal to zero then reset it
                    ' to the intial offset values.
                    If dOffSet <= 0 Then
                        dOffSet = GetSeamDistance
                    End If

                    Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
                    lFabMargin = 0
                    lAssyMargin = 0
                    lCustomMargin = 0

                    Dim i As Integer
                    Dim oConstMargin As IJConstMargin
                    Dim oObliqueMargin As IJObliqueMargin

                    Dim oMfgDefport As IJPort
                    Dim pMfgDef As IJMfgDefinition
                    For i = 1 To ReferenceObjColl.Count
                        If TypeOf ReferenceObjColl.Item(i) Is IJMfgDefinition Then
                            Set pMfgDef = ReferenceObjColl.Item(i)
                            Set oMfgDefport = pMfgDef.GetPort
                            'check if Margin exist on the same port where SeamControl mark is being created
                            If oConnectingSidePort Is oMfgDefport Then
                                If TypeOf pMfgDef Is IJDFabMargin Then
                                    Dim oFabMargin As IJDFabMargin
                                    Set oFabMargin = pMfgDef

                                    If oFabMargin.GeometryChange = AsMargin Then ''1 = As Margin, 2 = As Shrinkage, 3 = As Reference
                                        If TypeOf pMfgDef Is IJAssyMarginChild Then
                                            Set oConstMargin = pMfgDef
                                            lAssyMargin = lAssyMargin + oConstMargin.Value
                                        ElseIf TypeOf pMfgDef Is IJObliqueMargin Then
                                            Set oObliqueMargin = pMfgDef
                                            If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                                                lFabMargin = lFabMargin + oObliqueMargin.EndValue
                                            Else
                                                lFabMargin = lFabMargin + oObliqueMargin.StartValue
                                            End If
                                        ElseIf TypeOf pMfgDef Is IJConstMargin Then
                                            Set oConstMargin = pMfgDef
                                            lFabMargin = lFabMargin + oConstMargin.Value
                                        'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                    If lAssyMargin <> 0 Or lFabMargin <> 0 Or lCustomMargin <> 0 Then

                         Dim TotMargin As Double
                         TotMargin = lAssyMargin + lFabMargin + lCustomMargin
                         If TotMargin < 0.08 Then
                            dOffSet = dOffSet - TotMargin
                         End If
                    End If

                    Dim oCSColl As IJElements
                    Dim oCS As IJComplexString
                    Dim oProject As New IMSModelGeomOps.Project
                    Dim oSolutionUnk As IUnknown

                    If oCrossVec.Dot(oOffSetDirection) < 0# Then
                        dOffSet = -1# * dOffSet
                    End If

                    For projInd = 1 To 2
                        If Not oCSColl Is Nothing Then
                            If oCSColl.Count = 0 Then
                                Set oCSColl = Nothing
                            End If
                        End If

                        If (oCSColl Is Nothing) Then
                            dOffSet = -1# * dOffSet

                            Set oUnkTransientSheet = m_oMfgRuleHelper.OffsetCurve(oSurfaceWithOutFeatures, oEdgeWire, Nothing, dOffSet, False)

                            If oUnkTransientSheet Is Nothing Then
                                 'GoTo NextItem
                                GoTo NextProjInd
                            End If

                            oProject.CurveAlongVectorOnToSurface Nothing, oSurfaceWithOutFeatures, oUnkTransientSheet, oNormalDirection, Nothing, oSolutionUnk

                            If Not oSolutionUnk Is Nothing Then
                                Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oSolutionUnk)
                                Set oSolutionUnk = Nothing
                            End If
                        End If
NextProjInd:
                    Next projInd

                    Set oProject = Nothing

                    If Not oCSColl Is Nothing Then
                        If oCSColl.Count = 0 Then
                            Set oCSColl = Nothing
                        End If
                    End If

                    If (oCSColl Is Nothing) Then
                        GoTo NextMergedCS
                    End If

                    For Each oCS In oCSColl
                        Dim oPCCurve As IJCurve
                        Set oPCCurve = oCS

                        ' If the length of the CS is less than 1 mm then skip that curve
                        If oPCCurve.Length < 0.001 Then
                            GoTo NextCS
                        End If

                        ' Project the CS Start position onto the wirebody. If it is too far away, ignore it.
                        Dim oCSStartPos As IJDPosition, oProjPos As IJDPosition
                        oPCCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

                        Set oCSStartPos = New DPosition
                        oCSStartPos.Set dStartX, dStartY, dStartZ
                        dDist = 1000
                        oGeomOpsToolBox.GetNearestPointOnWireBodyFromPoint oUnkTransientSheet, oCSStartPos, oCSStartPos, oProjPos
                        dDist = oCSStartPos.DistPt(oProjPos)
                        If dDist > Abs(dOffSet) Then
                            GoTo NextCS
                        End If

                        'Create in systemmark object and add to output-collection
                        Dim oSystemMark As IJMfgSystemMark
                        Dim oMarkingInfo As MarkingInfo
                        Dim oGeom3d As IJMfgGeom3D
                        Set oGeom3d = m_oGeom3dFactory.Create(pResMgr)

                        'Create a SystemMark object to store additional information
                        Set oSystemMark = m_oSystemMarkFactory.Create(pResMgr)
                        'Set the marking side
                        oSystemMark.SetMarkingSide UpSide '(should be connecting side)

                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark

                        oMarkingInfo.Name = "SEAM CONTROL"

                        oSystemMark.Set3dGeometry oGeom3d
                        oGeom3d.PutGeometry oCS

                        Set oCS = Nothing
                        oGeom3d.PutGeometrytype STRMFG_SEAM_MARK

                        oGeom3d.PutMoniker m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                        'ApplyDirection oSystemMark, oConnectionData.ConnectingPort
                        oGeomCol3d.AddGeometry 1, oGeom3d
NextCS:
                    Next
NextMergedCS:
                    If Not oCSColl Is Nothing Then
                        oCSColl.Clear
                        Set oCSColl = Nothing
                    End If
                Next

                If Not oCSColl Is Nothing Then
                    oCSColl.Clear
                    Set oCSColl = Nothing
                End If

            End If
        End If
NextItem:
            Set oStructPort1 = Nothing
            Set oStructPort2 = Nothing
            Set oConnectingSidePort = Nothing
            Set oCommonGeometry = Nothing
            Set oEdgeWire = Nothing
            Set oOffSetDirection = Nothing
            Set oConnectingSideSurface = Nothing
            Set oMidPosition = Nothing
            Set oTangent1 = Nothing
            Set oNormalDirection = Nothing
            Set oCrossVec = Nothing
            Set oUnkTransientSheet = Nothing
            Set oCS = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oGeom3d = Nothing
    Next nIndex

    'Return the 3d collection
    Set CreateSeamControlMark3D = oGeomCol3d

CleanUp:
    Set oUpsidePort = Nothing
    Set oSurfaceWithOutFeatures = Nothing
    Set oSDPlateWrapper = Nothing
    Set oSDPhysicalConn = Nothing
    Set oPlateWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateMarkAtPosition()
'
' Description:  It create A mark of fixed length at the given position along given vector
'
' ***********************************************************************************
Public Function CreateMarkAtPosition(ByVal oSurfaceBody As IJSurfaceBody, ByVal oMarkPos As IJDPosition, ByVal oMarkVector As IJDVector, ByVal MarkLength As Double) As IJComplexString
    Const METHOD = "CreateMarkAtPosition"
    On Error GoTo ErrorHandler

    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d

    oMarkLine.DefineBy2Points oMarkPos.x - oMarkVector.x, oMarkPos.y - oMarkVector.y, oMarkPos.z - oMarkVector.z, oMarkPos.x + oMarkVector.x, oMarkPos.y + oMarkVector.y, oMarkPos.z + oMarkVector.z

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False

    Dim oProjCS As IJComplexString
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    ' Project the line created on the plate part surface
    oMfgMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

    If oProjCS Is Nothing Then
        GoTo CleanUp
    End If

    Dim oCurve As IJCurve
    Set oCurve = oProjCS

    Dim dCurveLength     As Double
    dCurveLength = oCurve.Length

    ' If the length of projects CS is more then 60mm then trim it.
    If dCurveLength > (MarkLength - 0.001) Then

        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        Dim oStartPos As IJDPosition
        Set oStartPos = New DPosition

        Dim oEndPos As IJDPosition
        Set oEndPos = New DPosition

        oStartPos.Set dStartX, dStartY, dStartZ
        oEndPos.Set dEndX, dEndY, dEndZ

        ' we need Fitting marks with length 15 mm
        If oStartPos.DistPt(oMarkPos) > oEndPos.DistPt(oMarkPos) Then
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - MarkLength, oStartPos
        Else
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - MarkLength, oEndPos
        End If
    End If

    Set CreateMarkAtPosition = oProjCS

CleanUp:
    Set oMarkLine = Nothing
    Set oCS = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


'---------------------------------------------------------------------------------------
' Purpose : Process Corrugated Marks
'---------------------------------------------------------------------------------------
Public Sub ProcessCorrugatedMarks(Part As IJPlatePart, ByVal UpSide As Long, CorrugateUpIsWaveUp As Boolean, RemoveKnuckles As Boolean, bRemoveRollLines As Boolean, bRemoveRollBoundary As Boolean)
    Const METHOD = "ProcessCorrugatedMarks"
    On Error GoTo ErrorHandler

    Const OFFSET_FROM_CORRUGATE As Double = 0.2    ' 200mm
    Const LENGTH_FITTING_MARK As Double = 0.4    ' 400mm

    If Not IsCorrugated(Part) And Not IsSwage(Part) And Not IsSwadge(Part) Then
        Exit Sub
    End If

    '*** Create the Plate Wrapper Object ***'
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oMfgPart As IJMfgPlatePart
    If Not oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Exit Sub
    End If

    Dim oMfgParent As IJMfgGeomParent
    Set oMfgParent = oMfgPart

    Dim oChildColl As IJDTargetObjectCol
    Set oChildColl = oMfgParent.GetChildren

    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oChildColl.Item(1)

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If

    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    Dim oMfgMiscHelper As New MfgRuleHelpers.Helper

    Dim LargeXYplane As IJSurface
    Set LargeXYplane = oMfgMiscHelper.CreateSurfaceByPointNormal(0#, 0#, 0#, 0#, 0#, 1#, True)

    Dim oNormalVec As New DVector
    oNormalVec.Set 0#, 0#, 1#

    Dim oCSColl As IJElements
    Set oCSColl = New JObjectCollection

    Dim i As Long
    For i = oGeomCol2d.GetCount To 1 Step -1
        Dim oGeom2d As IJMfgGeom2d
        Set oGeom2d = oGeomCol2d.GetGeometry(i)

        If RemoveKnuckles And oGeom2d.GetGeometryType = STRMFG_KNUCKLE_MARK Then
            oGeomCol2d.RemoveGeometry oGeom2d

            Dim oObjToDel As IJDObject
            Set oObjToDel = oGeom2d
            oObjToDel.Remove
            Set oObjToDel = Nothing

            GoTo ContinueNextGeom
        End If
        'Removing Roll_Lines Mark When option is OnlyCorrugateUp and OnlyCorrugateDown
        If bRemoveRollLines And oGeom2d.GetGeometryType = STRMFG_ROLL_LINES_MARK Then
            oGeomCol2d.RemoveGeometry oGeom2d

            Set oObjToDel = oGeom2d
            oObjToDel.Remove
            Set oObjToDel = Nothing

            GoTo ContinueNextGeom
        End If
        'Removing Roll_Boundary Mark When option is OnlyCorrugateUp and OnlyCorrugateDown
        If bRemoveRollBoundary And oGeom2d.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
            oGeomCol2d.RemoveGeometry oGeom2d

            Set oObjToDel = oGeom2d
            oObjToDel.Remove
            Set oObjToDel = Nothing

            GoTo ContinueNextGeom
        End If

        If oGeom2d.GetGeometryType <> STRMFG_CORRUGATE_MARK Then GoTo ContinueNextGeom

        Dim oMarkInfo As MarkingInfo
        Dim oExistingSystemMark As IJMfgSystemMark
        Set oExistingSystemMark = oGeom2d.SystemMark
        Set oMarkInfo = oGeom2d.SystemMark

        'Declaring the Custom Attributes
        Dim vLabel_Value As Variant, vUsage_Value As Variant, vBump_Value As Variant, vModel_Dim_Value As Variant
        oExistingSystemMark.SetMarkingSide UpSide

        Dim sCorDir As String
        sCorDir = oMarkInfo.Direction

        If (sCorDir = "up" And CorrugateUpIsWaveUp) Or _
           (sCorDir = "down" And Not CorrugateUpIsWaveUp) _
        Then

            'Setting the Values for Custom Attributes
            vLabel_Value = ""
            vUsage_Value = "MAIN"
            vModel_Dim_Value = oMarkInfo.Radius

            If sCorDir = "up" Then
                vBump_Value = "UP"
            ElseIf sCorDir = "down" Then
                vBump_Value = "DOWN"
            End If

            oMarkInfo.SetAttributeNameAndValue "LABEL", vLabel_Value
            oMarkInfo.SetAttributeNameAndValue "USAGE", vUsage_Value
            oMarkInfo.SetAttributeNameAndValue "BUMP", vBump_Value
            oMarkInfo.SetAttributeNameAndValue "MODEL_DIM", vModel_Dim_Value


            Dim oCorrCS As IJComplexString
            Set oCorrCS = oGeom2d.GetGeometry

            Dim oCorrCrv As IJCurve
            Set oCorrCrv = oCorrCS

            Dim sX As Double, sY As Double, sZ As Double
            Dim eX As Double, eY As Double, eZ As Double
            oCorrCrv.EndPoints sX, sY, sZ, eX, eY, eZ

            Dim oStartPos As New DPosition
            oStartPos.Set sX, sY, sZ

            Dim oEndPos As New DPosition
            oEndPos.Set eX, eY, eZ

            Dim CorLen As Double
            CorLen = oCorrCrv.Length

            If CorLen < LENGTH_FITTING_MARK Then GoTo ContinueNextGeom

            Dim j As Integer
            For j = 1 To 4
                Dim oCS As IJComplexString
                oMfgMGHelper.CloneComplexString oCorrCS, oCS

                Dim oTrimPos As IJDPosition
                Dim oAlongVec As IJDVector
                If j <= 2 Then
                    Set oTrimPos = oMfgGeomHelper.GetPointAtDistAlongCurve(oCS, oStartPos, LENGTH_FITTING_MARK)
                    oMfgMGHelper.TrimCurveByPoints oCS, oStartPos, oTrimPos
                    Set oAlongVec = oTrimPos.Subtract(oStartPos)
                Else
                    Set oTrimPos = oMfgGeomHelper.GetPointAtDistAlongCurve(oCS, oStartPos, CorLen - LENGTH_FITTING_MARK)
                    oMfgMGHelper.TrimCurveByPoints oCS, oTrimPos, oEndPos
                    Set oAlongVec = oEndPos.Subtract(oTrimPos)
                End If

                oAlongVec.Length = 1#

                Dim oOffsetVec As IJDVector
                Set oOffsetVec = oAlongVec.Cross(oNormalVec)

                Dim oWBi As IJWireBody
                oMfgMGHelper.ComplexStringToWireBody oCS, oWBi

                Dim oWBo As IJWireBody
                If j Mod 2 Then
                    oOffsetVec.Length = -1# * oOffsetVec.Length
                End If

                Set oWBo = oMfgMiscHelper.OffsetCurve(LargeXYplane, oWBi, oOffsetVec, OFFSET_FROM_CORRUGATE, False)

                Dim oOffCS As IJComplexString
                oMfgMGHelper.WireBodyToComplexString oWBo, oOffCS

                oCSColl.Add oOffCS
            Next
        Else
            'Setting the Values for Custom Attributes
            vLabel_Value = ""
            vUsage_Value = "SUB"
            vModel_Dim_Value = oMarkInfo.Radius

            If sCorDir = "up" Then
                vBump_Value = "UP"
            ElseIf sCorDir = "down" Then
                vBump_Value = "DOWN"
            End If

            oMarkInfo.SetAttributeNameAndValue "LABEL", vLabel_Value
            oMarkInfo.SetAttributeNameAndValue "USAGE", vUsage_Value
            oMarkInfo.SetAttributeNameAndValue "BUMP", vBump_Value
            oMarkInfo.SetAttributeNameAndValue "MODEL_DIM", vModel_Dim_Value

        End If

ContinueNextGeom:
        Set oGeom2d = Nothing
    Next

    Dim oGeom2dFactory As GSCADMfgGeometry.MfgGeom2dFactory
    Set oGeom2dFactory = New GSCADMfgGeometry.MfgGeom2dFactory

    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo

    For i = 1 To oCSColl.Count

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark

        'Setting Upside as Marking Side
        oSystemMark.SetMarkingSide UpSide

        Dim vRef_Label_Value As Variant, vRef_Usage_Value As Variant

        vRef_Label_Value = ""
        vRef_Usage_Value = "REF"

        oMarkingInfo.SetAttributeNameAndValue "LABEL", vRef_Label_Value
        oMarkingInfo.SetAttributeNameAndValue "USAGE", vRef_Usage_Value

        ' Create a Geom2d for each complex string
        Dim oNewGeom2d As IJMfgGeom2d
        Set oNewGeom2d = oGeom2dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        oNewGeom2d.PutGeometrytype STRMFG_CORRUGATE_MARK
        oNewGeom2d.PutGeometry oCSColl.Item(i)
        oNewGeom2d.TrimToBoundaries = True

        oSystemMark.Set2dGeometry oNewGeom2d
        oGeomCol2d.AddGeometry oGeomCol2d.GetCount, oNewGeom2d
        Set oNewGeom2d = Nothing
        Set oSystemMark = Nothing
    Next


CleanUp:
    Set oPlateWrapper = Nothing
    Set oMfgPart = Nothing
    Set oMfgParent = Nothing
    Set oChildColl = Nothing
    Set oGeomCol2d = Nothing
    Set oGeom2dFactory = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1011, , "RULES")
    GoTo CleanUp
End Sub

Public Sub GetDeclivityPoint(ByVal oLengthGeomCurve As IJComplexString, _
                                ByRef oInputPosToReset As IJDPosition, _
                                ByRef oMaterialVec As IJDVector, _
                                ByVal oInputSurface As IJDModelBody)

    Dim oLineStr As IJLineString
    Dim Points() As Double
    Dim lPtCount As Long, ind As Long
    Dim oLengthPointPos As New DPosition
    Dim dFarthestDist As Double, dMinDist As Double, dX As Double, dY As Double, dZ As Double
    Dim dTotalDist As Double, dAverageDist As Double

    oLengthGeomCurve.GetCurve 1, oLineStr

    oMaterialVec.Length = 1#

    oLineStr.GetPoints lPtCount, Points

    For ind = 0 To (lPtCount - 1)
        oLengthPointPos.x = Points(3 * ind + 0)
        oLengthPointPos.y = Points(3 * ind + 1)
        oLengthPointPos.z = Points(3 * ind + 2)
        dTotalDist = dTotalDist + oInputPosToReset.DistPt(oLengthPointPos)
    Next

    dAverageDist = dTotalDist / lPtCount

    dFarthestDist = -1#
    For ind = 0 To (lPtCount - 1)
        Dim oClosestPos As IJDPosition
        oLengthPointPos.x = Points(3 * ind + 0)
        oLengthPointPos.y = Points(3 * ind + 1)
        oLengthPointPos.z = Points(3 * ind + 2)

        If ind = 0 Then
            oInputPosToReset.Get dX, dY, dZ
        End If

        If oLengthPointPos.Subtract(oInputPosToReset).Dot(oMaterialVec) > 0# Then
            If oInputPosToReset.DistPt(oLengthPointPos) < dAverageDist Then
                oInputSurface.GetMinimumDistanceFromPosition oLengthPointPos, oClosestPos, dMinDist
                dMinDist = oInputPosToReset.DistPt(oClosestPos)
                If dMinDist > dFarthestDist Then
                    dFarthestDist = dMinDist
                    oClosestPos.Get dX, dY, dZ
                End If
            End If
        End If
        Set oClosestPos = Nothing
    Next

    oInputPosToReset.Set dX, dY, dZ

CleanUp:
    Set oLengthPointPos = Nothing
    Set oLineStr = Nothing
    Exit Sub

ErrorHandler:
    GoTo CleanUp
End Sub

' ***********************************************************************************
' Public Function CreateDimensionMarksForCorrugate
'
' Description:  It create marks Dimension marks for corrugate marks
'
' ***********************************************************************************
Public Sub CreateDimensionMarksForCorrugate(oPlatePart As IJPlatePart, bCorrugateUp As Boolean, UpSide As Long)
    Const METHOD = "CreateDimensionMarksForCorrugate"
    On Error GoTo ErrorHandler

    '*** Create the Plate Wrapper Object ***'
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlatePart

    Dim oMfgPart As IJMfgPlatePart
    If Not oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Exit Sub
    End If

    Dim oMfgParent As IJMfgGeomParent
    Set oMfgParent = oMfgPart

    Dim oChildColl As IJDTargetObjectCol
    Set oChildColl = oMfgParent.GetChildren

    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oChildColl.Item(1)

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If

    Dim i As Long
    Dim oGeom2d As IJMfgGeom2d
    Dim oMarkingInfo As MarkingInfo
    Dim dDistanceFromEdge As Double
    Dim bPlateWidthMoreThan3000mm As Boolean

    Dim oCorrugateMarkColl As Collection
    Set oCorrugateMarkColl = New Collection

    'Getting all the Corrugate Marks
    For i = 1 To oGeomCol2d.GetCount
        Set oGeom2d = oGeomCol2d.GetGeometry(i)
        If oGeom2d.GetGeometryType = STRMFG_CORRUGATE_MARK Then

            If Not oGeom2d.SystemMark Is Nothing Then
                oCorrugateMarkColl.Add oGeom2d

                Set oMarkingInfo = oGeom2d.SystemMark
                dDistanceFromEdge = oMarkingInfo.Radius

                'Checking the Width Condition
                If dDistanceFromEdge > 3 Then
                    bPlateWidthMoreThan3000mm = True
                End If
            End If
        End If
    Next

    Dim dMarkDistnceFromEdge As Double
    Dim dMarkDistnceFromEdgePrevious As Double

    Dim oGeom2dNext As IJMfgGeom2d
    Dim oMarkingInfoNext As MarkingInfo

    Dim oValidGeom2d As IJMfgGeom2d
    Dim oMarkingInfoValidGeom2d As MarkingInfo

    Dim oValidGeom2dColl As Collection
    Set oValidGeom2dColl = New Collection

    Dim oMarkPostionColl As Collection
    Set oMarkPostionColl = New Collection

    Dim oMarkingNameColl As New Collection

    Dim oCorrugateUpMarkColl As IJElements
    Set oCorrugateUpMarkColl = New JObjectCollection
    Dim oCorrugateDownMarkColl As IJElements
    Set oCorrugateDownMarkColl = New JObjectCollection

    'Looping Through the Corrugate Marks for Checking the Distance Condition
    For i = 1 To oCorrugateMarkColl.Count

        Dim sCorrugateDir As String
        Set oGeom2d = oCorrugateMarkColl.Item(i)
        Set oMarkingInfo = oGeom2d.SystemMark
        sCorrugateDir = oMarkingInfo.Direction

        'Adding the first item and Last item(Plate Edges) into both the collections
        If i = 1 Or i = oCorrugateMarkColl.Count Then
            oCorrugateDownMarkColl.Add oCorrugateMarkColl.Item(i)
            oCorrugateUpMarkColl.Add oCorrugateMarkColl.Item(i)
        End If

        'Separating the Up and Down Marks into the respective collections
        If sCorrugateDir = "down" Then
            oCorrugateDownMarkColl.Add oCorrugateMarkColl.Item(i)
        ElseIf sCorrugateDir = "up" Then
            oCorrugateUpMarkColl.Add oCorrugateMarkColl.Item(i)

        End If
    Next


    Dim lSplitedCount As Long
    Dim oTempColl As IJElements
    Set oTempColl = New JObjectCollection

    'Getting the counts and collection as per CorrugateUp or CorrugateDown
    If bCorrugateUp = True Then
        lSplitedCount = oCorrugateDownMarkColl.Count
        oTempColl.AddElements oCorrugateDownMarkColl
    Else
        lSplitedCount = oCorrugateUpMarkColl.Count
        oTempColl.AddElements oCorrugateUpMarkColl
    End If
    'Looping through the collection and adding valid marks into seperate collection
    For i = 1 To lSplitedCount
        If i = 1 Then
            oValidGeom2dColl.Add oTempColl.Item(i)
            GoTo NextItem
        End If

        Set oGeom2d = oTempColl.Item(i)
        Set oMarkingInfo = oGeom2d.SystemMark
        dMarkDistnceFromEdge = oMarkingInfo.Radius

        If Abs(dMarkDistnceFromEdge - dMarkDistnceFromEdgePrevious) > 3 Then
            If i = 2 Then
                oValidGeom2dColl.Add oTempColl.Item(i)
            Else
                oValidGeom2dColl.Add oTempColl.Item(i - 1)
            End If

            Set oValidGeom2d = oValidGeom2dColl.Item(oValidGeom2dColl.Count)
            Set oMarkingInfoValidGeom2d = oValidGeom2d.SystemMark
            dMarkDistnceFromEdgePrevious = oMarkingInfoValidGeom2d.Radius
        Else
            GoTo NextItem
        End If

NextItem:
    Next

    'if the last item is not added into the valid collection - adding it
    If Not oValidGeom2dColl.Item(oValidGeom2dColl.Count) Is oTempColl.Item(lSplitedCount) Then
        oValidGeom2dColl.Add oTempColl.Item(lSplitedCount)
    End If


    Dim oComplexString As IJComplexString
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    'Looping through the ValidGeom2d Coll to get the all Dimension Mark Positions
    For i = 1 To oValidGeom2dColl.Count

        If i <> oValidGeom2dColl.Count Then
            Set oGeom2d = oValidGeom2dColl.Item(i)
            Set oMarkingInfo = oGeom2d.SystemMark

            Set oGeom2dNext = oValidGeom2dColl.Item(i + 1)
            Set oMarkingInfoNext = oGeom2dNext.SystemMark

            oMarkingNameColl.Add Abs(oMarkingInfo.Radius - oMarkingInfoNext.Radius)

            Set oGeom2d = Nothing
            Set oGeom2dNext = Nothing
        End If

        Set oGeom2d = oValidGeom2dColl.Item(i)
        Set oComplexString = oGeom2d.GetGeometry

        Dim oCurve As IJCurve
        Set oCurve = oComplexString

        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double

        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        Dim oStartPosition As IJDPosition
        Set oStartPosition = New DPosition
        oStartPosition.Set dStartX, dStartY, dStartZ

        Dim oMarkPosition As IJDPosition
        'Finding the position on the mark Either at 3/5 or 11/20 of Mark Height based on the width of the plate
        If bPlateWidthMoreThan3000mm = True Then

            Dim dInterpolatedValue As Double
            If dInterpolatedValue = 0 Then
                dInterpolatedValue = 0.55 ' Intially Set the Value to 0.55 (11/20)
            End If

            Set oMarkPosition = oMfgGeomHelper.GetPointAtDistAlongCurve(oComplexString, oStartPosition, (dInterpolatedValue * oCurve.Length))
            oMarkPostionColl.Add oMarkPosition
            'For Interpoling
            'Other than the First item and last item everything should be interpolated
            If Not (i = 1) And Not (i = oValidGeom2dColl.Count) Then
                dInterpolatedValue = dInterpolatedValue + (0.05 / (oValidGeom2dColl.Count - 1))
                Set oMarkPosition = oMfgGeomHelper.GetPointAtDistAlongCurve(oComplexString, oStartPosition, (dInterpolatedValue * oCurve.Length))
                oMarkPostionColl.Add oMarkPosition
            End If

        Else
            Set oMarkPosition = oMfgGeomHelper.GetPointAtDistAlongCurve(oComplexString, oStartPosition, (0.6 * oCurve.Length))
            oMarkPostionColl.Add oMarkPosition
        End If
    Next

    Dim oGeom2dFactory As GSCADMfgGeometry.MfgGeom2dFactory
    Set oGeom2dFactory = New GSCADMfgGeometry.MfgGeom2dFactory

    Dim dMarkPositionCount As Double
    'Setting the MarkingPosition Count based on the count
    If (oMarkPostionColl.Count) Mod 2 = 0 Then
        dMarkPositionCount = oMarkPostionColl.Count
    Else
        dMarkPositionCount = oMarkPostionColl.Count - 1
    End If

    'Looping through the Positions Coll to put the Mark
    For i = 1 To dMarkPositionCount

        Dim oMarkPosition1 As IJDPosition
        Dim oMarkPosition2 As IJDPosition

        'Only For Odd Numbers Should Create Mark
        If i Mod 2 <> 0 Then
            Set oMarkPosition1 = oMarkPostionColl.Item(i)
            Set oMarkPosition2 = oMarkPostionColl.Item(i + 1)

            Dim oMarkLine As IJLine
            Set oMarkLine = New Line3d
            oMarkLine.DefineBy2Points oMarkPosition1.x, oMarkPosition1.y, oMarkPosition1.z, oMarkPosition2.x, oMarkPosition2.y, oMarkPosition2.z

            Dim oCS As IJComplexString
            Set oCS = New ComplexString3d
            oCS.AddCurve oMarkLine, False

            Dim oSystemMark As IJMfgSystemMark
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            'Set the marking side
            oSystemMark.SetMarkingSide UpSide

            Dim oNewGeom2d As IJMfgGeom2d
            Set oNewGeom2d = oGeom2dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            oNewGeom2d.PutGeometrytype STRMFG_REF_MARK
            oNewGeom2d.PutGeometry oCS

            Dim oCorrugateMarkingInfo As MarkingInfo
            Set oCorrugateMarkingInfo = oSystemMark

            Dim vDim_Label_Value As Variant, vDim_Usage_Value As Variant
            vDim_Usage_Value = "DIM"
            vDim_Label_Value = Format((oMarkingNameColl.Item((i + 1) / 2) * 1000), 0) 'Converting to mm and rounding it off to zero decimal places

            oCorrugateMarkingInfo.SetAttributeNameAndValue "USAGE", vDim_Usage_Value
            oCorrugateMarkingInfo.SetAttributeNameAndValue "LABEL", vDim_Label_Value
            oCorrugateMarkingInfo.Name = Format((oMarkingNameColl.Item((i + 1) / 2) * 1000), 0)

            oSystemMark.Set2dGeometry oNewGeom2d

            oGeomCol2d.AddGeometry oGeomCol2d.GetCount, oNewGeom2d
            Set oNewGeom2d = Nothing

        End If
    Next

CleanUp:
    Set oPlateWrapper = Nothing
    Set oMfgPart = Nothing
    Set oMfgParent = Nothing
    Set oChildColl = Nothing
    Set oGeomCol2d = Nothing
    Set oNewGeom2d = Nothing
    Set oGeom2dFactory = Nothing


    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

' ***********************************************************************************
' Public Function CreateRollLineMarkAndCorrugateMarks
'
' Description:  It create Roll Lines Marks, Corrugate Marks and Dimension Marks
'
' ***********************************************************************************
Public Function CreateRollLineMarkAndCorrugateMarks(ByVal Part As Object, ByVal UpSide As Long, _
                                        ByVal bSelectiveRecompute As Boolean, _
                                        ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                        Optional ByVal bRollMarksOnly As Boolean, _
                                        Optional ByVal bCorrugateUp As Boolean, _
                                        Optional ByVal bCorrugateDown As Boolean, Optional ByVal bCreateDimension As Boolean, Optional ByVal bRollLineTrimmed As Boolean) As IJMfgGeomCol2d

Const METHOD = "CreateRollLineMarkAndCorrugateMarks"

    'For Only Roll Lines
    If bRollMarksOnly = True Then

        Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
        Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
        Set oPlateWrapper.object = Part

        Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
        Dim oMfgPart As IJMfgPlatePart
        If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
            Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
            Set oMfgPlateWrapper.object = oMfgPart
        Else
            Exit Function
        End If

        Dim oGeomCol2d As IJMfgGeomCol2d
        Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries

        If oGeomCol2d Is Nothing Then
            'You can exit the function after cleanup since there is nothing to be marked
            GoTo CleanUp
        End If

        Dim lSwagedPlate As Long
        ' Default is not Swaged
        lSwagedPlate = 0

        Dim oIJDAttributes As IJDAttributes
        Dim oIJDAttr As IJDAttribute
        Dim oAttribCol As IJDAttributesCol
        Dim varIJUASwedgePlate As Variant
        Dim oAttribMetaData As IJDAttributeMetaData
        Dim oObject         As IJDObject

        Set oObject = Part
        Set oAttribMetaData = oObject.ResourceManager
        Set oObject = Nothing

        If ROLL_LINES_DISPLAY = ShowAllGeneratedBySystem Then
            On Error Resume Next
            varIJUASwedgePlate = oAttribMetaData.IID("IJUASwagePlate")
            Err.Clear
            If Trim(CStr(varIJUASwedgePlate)) <> vbNullString Then
                Set oIJDAttributes = Part

                Set oAttribCol = oIJDAttributes.CollectionOfAttributes(varIJUASwedgePlate)

                If Not (oAttribCol Is Nothing) Then
                    Set oIJDAttr = oAttribCol.Item("SwagePlate")
                    lSwagedPlate = CLng(oIJDAttr.Value)
                End If
            End If
            On Error GoTo ErrorHandler
        End If

        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlatePart = oMfgPart

        Dim oMfgPlateProcessSettings As IJMfgPlateProcessSettings
        Set oMfgPlateProcessSettings = oMfgPlatePart.GetProcessSettings

        Dim bstrUnfoldAlg   As String
        bstrUnfoldAlg = oMfgPlateProcessSettings.UnwrapAlgorithm

        Dim oRollLinesToDel As IJElements
        Set oRollLinesToDel = New JObjectCollection

        Dim oGeom2d As IJMfgGeom2d, oKnuGeom2D As IJMfgGeom2d
        Dim oSystemMark As IJMfgSystemMark
        Dim oMarkingInfo As MarkingInfo

        Dim i As Long, kk As Long

        For i = oGeomCol2d.GetCount To 1 Step -1
            Set oGeom2d = oGeomCol2d.GetGeometry(i)
            If oGeom2d.GetGeometryType = STRMFG_ROLL_LINES_MARK Then
                If lSwagedPlate > 0 Or ROLL_LINES_DISPLAY = DoNotShowRollLines Then
                    ' This is a Swaged plate
                    oGeomCol2d.RemoveGeometry oGeom2d
                    ' remove persistent object
                    Dim oGeomObject As IJDObject
                    Set oGeomObject = oGeom2d
                    oGeomObject.Remove
                    Set oGeomObject = Nothing
                ElseIf ROLL_LINES_DISPLAY = ShowAllGeneratedBySystem Then
                    ' This is not a Swaged plate
                    Set oSystemMark = oGeom2d.SystemMark

                    If oSystemMark Is Nothing Then
                        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        oSystemMark.Set2dGeometry oGeom2d
                    End If

                    oSystemMark.SetMarkingSide UpSide

                    'QI for the MarkingInfo object on the SystemMark
                    Set oMarkingInfo = oSystemMark
                    oMarkingInfo.Name = "ROLL LINE"
                ElseIf ROLL_LINES_DISPLAY = GenerateRollAnnotationOnly And bstrUnfoldAlg = "Developable" Then
                    If oRollLinesToDel.Contains(oGeom2d) Then GoTo ContinueNextI

                    ' Pick middle Roll Line for this roll region
                    Dim MidRollLine As IJMfgGeom2d
                    Dim oRB(1 To 2) As IJMfgGeom2d
                    Dim OtherRollLines As IJElements
                    ProcessRollRegion oGeom2d, oGeomCol2d, MidRollLine, oRB, OtherRollLines

                    If MidRollLine Is Nothing Then GoTo ContinueNextI
                    Dim bKnuckleMark As Boolean
                '*** Check if Roll Line comes from Knuckle or rolled plate ***'
                    For kk = oGeomCol2d.GetCount To 1 Step -1
                        Set oKnuGeom2D = oGeomCol2d.GetGeometry(kk)
                        bKnuckleMark = False
                        If oKnuGeom2D.GetGeometryType = STRMFG_KNUCKLE_MARK Then
                            Dim oKnuCenterPt As New DPosition
                            Dim oRB1MidPt As New DPosition
                            Dim oRB2MidPt As New DPosition
                            Dim oRBCurve As IJCurve
                            Dim oMidRollLineCurve As IJCurve
                            Dim dStartX As Double
                            Dim dStartY As Double
                            Dim dStartZ As Double
                            Dim dEndX As Double
                            Dim dEndY As Double
                            Dim dEndZ As Double


                            Set oRBCurve = oRB(1).GetGeometry
                            'oRBCurve.Centroid oRB1MidPt.x, oRB1MidPt.y, oRB1MidPt.z
                            oRBCurve.Centroid dStartX, dStartY, dStartZ
                            oRB1MidPt.Set dStartX, dStartY, dStartZ

                            Set oRBCurve = oRB(2).GetGeometry
                            'oRBCurve.Position 0.5, oRB2MidPt.x, oRB2MidPt.y, oRB2MidPt.z
                            oRBCurve.Centroid dStartX, dStartY, dStartZ
                            oRB2MidPt.Set dStartX, dStartY, dStartZ

                            Set oMidRollLineCurve = oKnuGeom2D.GetGeometry
                            'oMidRollLineCurve.Position 0.5, oKnuCenterPt.x, oKnuCenterPt.y, oKnuCenterPt.z
                            oMidRollLineCurve.Centroid dStartX, dStartY, dStartZ
                            oKnuCenterPt.Set dStartX, dStartY, dStartZ

                            'Get 1st Vector
                            Dim oVec1 As IJDVector, oVec2 As IJDVector
                            Set oVec1 = oRB1MidPt.Subtract(oKnuCenterPt)

                            'Get 1st Vector
                            Set oVec2 = oRB2MidPt.Subtract(oKnuCenterPt)

                            'Check Dot Product
                            If oVec1.Dot(oVec2) > 0 Then
                                'This means its Roll (No Knuckle)
                                'Do Nothing
                            Else
                                'This means its Roll coming from Knuckle
                                'Set Subgeometry type so that appropriate Annotation Can be placed
                                bKnuckleMark = True
                                oKnuGeom2D.IsSupportOnly = False
                                GoTo SystemMark
                            End If
                        End If
                    Next kk
                '*************************************************************'
SystemMark:
                    Set oSystemMark = MidRollLine.SystemMark
                    Set oMarkingInfo = oSystemMark
                    If bKnuckleMark = True Then
                        oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Knuckle"
                    End If

                    Dim DistBetRollBoundaries As Double
                    DistBetRollBoundaries = oMarkingInfo.Radius * oMarkingInfo.FittingAngle

                    Dim Prefix As String
'                    If oMarkingInfo.Direction = "up" Then
'                        Prefix = "%"
'                    ElseIf oMarkingInfo.Direction = "down" Then
'                        Prefix = "$"
'                    End If
                    oMarkingInfo.Name = Prefix & CStr(Round(oMarkingInfo.Radius * 1000, 0)) & "R"

                    Dim IsLine(1 To 2) As Boolean
                    Dim oCS(1 To 2) As IJComplexString
                    Dim oSpt(1 To 2) As IJDPosition
                    Dim oMpt(1 To 2) As IJDPosition
                    Dim oEpt(1 To 2) As IJDPosition

                    Dim jj As Integer
                    For jj = 1 To 2
                        IsLine(jj) = GetRBgeomProps(oRB(jj), oCS(jj), oSpt(jj), oMpt(jj), oEpt(jj))
                    Next jj

                    If IsLine(1) And IsLine(2) Then
                        If oMarkingInfo.Radius > LARGEvsSMALL_CRITERION_1_ROLLRADIUS Or _
                           DistBetRollBoundaries > LARGEvsSMALL_CRITERION_2_DISTBETROLLB Then
                           ' Region with "large" roll radius
                           MidRollLine.IsSupportOnly = True ' Don't show roll line for large radius.
                        End If ' end if large radius
                    End If ' end if both Roll boundaries are lines

                    oRollLinesToDel.AddElements OtherRollLines
                End If ' end if GenerateRollAnnotationOnly
            End If ' end if geom type is roll line mark

ContinueNextI:
            Set oGeom2d = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
        Next i

        Dim RedundantRollLine As IJMfgGeom2d
        For Each RedundantRollLine In oRollLinesToDel
            oGeomCol2d.RemoveGeometry RedundantRollLine
            Set oGeomObject = RedundantRollLine
            oGeomObject.Remove
            Set oGeomObject = Nothing
            Set RedundantRollLine = Nothing
        Next

    End If


    If bCorrugateUp = True Then
        If bCreateDimension = True Then
            CreateCorrugateMarksAndDimensionMarks Part, UpSide, True, False, False, False, True
        Else
            CreateCorrugateMarksAndDimensionMarks Part, UpSide, True, False, False, False, False
        End If
    End If

    If bCorrugateDown = True Then
        If bCreateDimension = True Then
            CreateCorrugateMarksAndDimensionMarks Part, UpSide, False, True, False, False, True
        Else
            CreateCorrugateMarksAndDimensionMarks Part, UpSide, False, True, False, False, False
        End If
    End If

    If bRollLineTrimmed = True Then

        Dim lCount As Long

        For i = oGeomCol2d.GetCount To 1 Step -1
            Set oGeom2d = oGeomCol2d.GetGeometry(i)

            'only for Roll Lines this Implementation is applicable
            If oGeom2d.GetGeometryType = STRMFG_ROLL_LINES_MARK Then
                Dim oRollLineCS As IJComplexString
                Set oRollLineCS = oGeom2d.GetGeometry
                Dim oRollLineCurve As IJCurve
                Set oRollLineCurve = oRollLineCS
                Dim dCurveLength As Double
                dCurveLength = oRollLineCurve.Length

                'Checking for the Length conditions
                If dCurveLength > SHELL_PLATE_ROLL_LINE_LENGTH Then
                    Dim oRollLineWireBody As IJWireBody
                    Set oRollLineWireBody = m_oMfgRuleHelper.ComplexStringToWireBody(oRollLineCS)

                    Dim oMfgMGHelper As IJMfgMGHelper
                    Set oMfgMGHelper = New MfgMGHelper

                    Dim oOffsetStartPoint As New DPosition, oOffsetEndPoint As New DPosition, oStartPoint As New DPosition, oEndPoint As New DPosition

                    oRollLineWireBody.GetEndPoints oStartPoint, oEndPoint

                    'Triiming the Roll Lines for 1 meter or 1000 mm
                    Set oOffsetStartPoint = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oRollLineWireBody, oStartPoint, SHELL_PLATE_ROLL_LINE_SHORTLENGTH, oEndPoint)

                    Set oOffsetEndPoint = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oRollLineWireBody, oEndPoint, SHELL_PLATE_ROLL_LINE_SHORTLENGTH, oStartPoint)

                    Dim oNewGeom2d As IJMfgGeom2d
                    Set oNewGeom2d = m_oGeom2dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                    Set oRollLineCS = m_oMfgRuleHelper.WireBodyToComplexString(oRollLineWireBody)

                    oMfgMGHelper.TrimCurveByPoints oRollLineCS, oStartPoint, oOffsetStartPoint

                    lCount = oGeomCol2d.GetCount + 1
                    oNewGeom2d.PutGeometry oRollLineCS

                    'Adding the new Trimmed Roll Lines to the existing Geom2d Collection
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    oSystemMark.Set2dGeometry oNewGeom2d
                    oSystemMark.SetMarkingSide UpSide

                    oNewGeom2d.PutGeometrytype STRMFG_ROLL_LINES_MARK
                    'oNewGeom2d.PutSubGeometryType STRMFG_ROLL_LINES_MARK

                    Set oMarkingInfo = oSystemMark
                    oMarkingInfo.Name = "LR"
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Trimmed"
                    oGeomCol2d.AddGeometry lCount, oNewGeom2d

                    Set oNewGeom2d = Nothing
                    Set oRollLineCS = Nothing

                    Set oNewGeom2d = m_oGeom2dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                    Set oRollLineCS = m_oMfgRuleHelper.WireBodyToComplexString(oRollLineWireBody)

                    oMfgMGHelper.TrimCurveByPoints oRollLineCS, oEndPoint, oOffsetEndPoint

                    lCount = lCount + 1

                    oNewGeom2d.PutGeometry oRollLineCS
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    oSystemMark.Set2dGeometry oNewGeom2d
                    oSystemMark.SetMarkingSide UpSide

                    oNewGeom2d.PutGeometrytype STRMFG_ROLL_LINES_MARK
                    Set oMarkingInfo = oSystemMark
                    oMarkingInfo.Name = "LR"
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Trimmed"
                    oGeomCol2d.AddGeometry lCount, oNewGeom2d

                    'Removing the unTrimmed Roll line from the Geom2d Collection
                    oGeomCol2d.RemoveGeometry oGeom2d
                    Set oGeomObject = oGeom2d
                    oGeomObject.Remove

                    Set oNewGeom2d = Nothing
                    Set oOffsetStartPoint = Nothing
                    Set oOffsetEndPoint = Nothing
                    Set oEndPoint = Nothing
                    Set oStartPoint = Nothing
                    Set oMfgMGHelper = Nothing
                    Set oRollLineWireBody = Nothing
                Else
                    GoTo CleanUp
                End If
                Set oRollLineCS = Nothing
                Set oRollLineCurve = Nothing
            End If
        Next i

        Set oGeom2d = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeomObject = Nothing
    End If

CleanUp:
    Set oPlateWrapper = Nothing
    Set oMfgPlateWrapper = Nothing
    Set oMfgPart = Nothing
    Set oMfgPlatePart = Nothing
    Set oMfgPlateProcessSettings = Nothing

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateCorrugateMarksAndDimensionMarks
'
' Description:  It create Corrugate Marks and Dimension Marks
'
' ***********************************************************************************
Public Sub CreateCorrugateMarksAndDimensionMarks(ByVal Part As Object, ByVal UpSide As Long, ByVal bCorrugateUp As Boolean, ByVal bCorrugateDown As Boolean, ByVal bRemoveRollLines As Boolean, ByVal bRemoveRollBoundary As Boolean, ByVal bCreateDimensionMarks As Boolean)
    Const METHOD = "CreateCorrugateMarksAndDimensionMarks"
    On Error GoTo ErrorHandler


    'Checking for Corrugate Plate
    If Not IsCorrugated(Part) And Not IsSwage(Part) Then
        Exit Sub
    End If

    'For CorrugateUp and Down
    'For Creating Corrugate Marks
    ProcessCorrugatedMarks Part, UpSide, bCorrugateUp, True, bRemoveRollLines, bRemoveRollBoundary

    If bCreateDimensionMarks = True Then
        'For Creating Dimension Marks
        CreateDimensionMarksForCorrugate Part, bCorrugateUp, UpSide
    End If


CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function GetRollRegionIndex(RollGeom As IJMfgGeom2d) As Long
    Const METHOD As String = "GetRollRegionIndex"
    On Error GoTo ErrorHandler

    GetRollRegionIndex = -1

    If RollGeom Is Nothing Then Exit Function

    Dim SystemMark As IJMfgSystemMark
    Set SystemMark = RollGeom.SystemMark

    Dim oMarkingInfo As MarkingInfo
    Set oMarkingInfo = SystemMark

    If oMarkingInfo Is Nothing Then Exit Function

    Dim varValue As Variant
    oMarkingInfo.GetAttributeValue "ROLL_REGION_INDEX", varValue

    GetRollRegionIndex = CLng(varValue)

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Sub ProcessRollRegion(SomeRollLine As IJMfgGeom2d, oGeomCol2d As IJMfgGeomCol2d, _
                              MidRollLine As IJMfgGeom2d, _
                              RollBoundary() As IJMfgGeom2d, _
                              OtherRollLines As IJElements)
    Const METHOD As String = "ProcessRollRegion"
    On Error GoTo ErrorHandler

    If SomeRollLine Is Nothing Or oGeomCol2d Is Nothing Or _
       LBound(RollBoundary) <> 1 Or UBound(RollBoundary) <> 2 _
    Then
       GoTo CleanUp
    End If

    Set OtherRollLines = New JObjectCollection
    Set RollBoundary(1) = Nothing
    Set RollBoundary(2) = Nothing

    Dim CurRollRegion As Long
    CurRollRegion = GetRollRegionIndex(SomeRollLine)

    If CurRollRegion <= 0 Then GoTo CleanUp

    Dim i As Long
    For i = 1 To oGeomCol2d.GetCount
        Dim oGeom2d As IJMfgGeom2d
        Set oGeom2d = oGeomCol2d.GetGeometry(i)
        If oGeom2d.GetGeometryType = STRMFG_ROLL_LINES_MARK Then
            If GetRollRegionIndex(oGeom2d) = CurRollRegion Then
                OtherRollLines.Add oGeom2d
            End If
        ElseIf oGeom2d.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
            If GetRollRegionIndex(oGeom2d) = CurRollRegion Then
                If RollBoundary(1) Is Nothing Then
                    Set RollBoundary(1) = oGeom2d
                Else
                    Set RollBoundary(2) = oGeom2d
                End If
            End If
        End If
    Next

    Dim MidIdx As Long
    MidIdx = 1 + OtherRollLines.Count \ 2

    Set MidRollLine = OtherRollLines.Item(MidIdx)
    OtherRollLines.Remove (MidIdx)

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Private Function GetRBgeomProps(oMfgGeom2d As IJMfgGeom2d, _
                                oCS As IJComplexString, _
                                oSpt As IJDPosition, _
                                oMpt As IJDPosition, _
                                oEpt As IJDPosition) As Boolean
    Const METHOD As String = "GetRBgeomProps"
    On Error GoTo ErrorHandler

    GetRBgeomProps = False
    If oMfgGeom2d Is Nothing Then Exit Function

    Set oCS = oMfgGeom2d.GetGeometry

    Dim oCrv As IJCurve
    Set oCrv = oCS

    Dim CrvScope As Geom3dCurveScopeConstants
    Dim x1 As Double, y1 As Double, z1 As Double
    oCrv.Normal CrvScope, x1, y1, z1

    If CrvScope <> CURVE_SCOPE_COLINEAR Then
        GetRBgeomProps = False
        Exit Function
    End If

    Dim x2 As Double, y2 As Double, z2 As Double
    oCrv.EndPoints x1, y1, z1, x2, y2, z2

    Set oSpt = New DPosition
    oSpt.Set x1, y1, z1

    Set oEpt = New DPosition
    oEpt.Set x2, y2, z2

    oCrv.PositionFRatio 0.5, x1, y1, z1

    Set oMpt = New DPosition
    oMpt.Set x1, y1, z1

    GetRBgeomProps = True

CleanUp:
    Exit Function

ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateKnuckleFittingLine
'
' Description:  It creates A Knuckle Fitting Markmark of fixed length at the given
'               position along given vector
'
' ***********************************************************************************
Public Function CreateKnuckleFittingLine(ByVal oSurfaceBody As IJSurfaceBody, ByVal oMarkPos As IJDPosition, ByVal oMarkVector As IJDVector, ByVal MarkLength As Double) As IJComplexString
    Const METHOD = "CreateKnuckleFittingLine"
    On Error GoTo ErrorHandler

    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d

    oMarkLine.DefineBy2Points oMarkPos.x, oMarkPos.y, oMarkPos.z, oMarkPos.x + oMarkVector.x, oMarkPos.y + oMarkVector.y, oMarkPos.z + oMarkVector.z

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False

    Dim oProjCS As IJComplexString
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    ' Project the line created on the plate part surface
    oMfgMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

    If oProjCS Is Nothing Then
        GoTo CleanUp
    End If

    Dim oCurve As IJCurve
    Set oCurve = oProjCS

    Dim dCurveLength     As Double
    dCurveLength = oCurve.Length

    ' If the length of projects CS is more then 60mm then trim it.
    If dCurveLength > (MarkLength - 0.001) Then

        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        Dim oStartPos As IJDPosition
        Set oStartPos = New DPosition

        Dim oEndPos As IJDPosition
        Set oEndPos = New DPosition

        oStartPos.Set dStartX, dStartY, dStartZ
        oEndPos.Set dEndX, dEndY, dEndZ

        ' we need Fitting marks with length 15 mm
        If oStartPos.DistPt(oMarkPos) > oEndPos.DistPt(oMarkPos) Then
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - MarkLength, oStartPos
        Else
            m_oMfgRuleHelper.TrimCurveEnds oProjCS, dCurveLength - MarkLength, oEndPos
        End If
    End If

    Set CreateKnuckleFittingLine = oProjCS

CleanUp:
    Set oMarkLine = Nothing
    Set oCS = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


' ************************************************************************************************************
' Public Sub CreateKnuckleFittingMarks
'
' Description:  It creates A Knuckle Fitting Marks on a given location Mark if the
'               Location Mark has Knuckles
'
'Inputs to this Function are:
'1) Location Mark Curve As ComplexString
'2) ConnectionData Information of the Location Mark
'3) Plate Part on which the Location marks are Placed
'4) Moldedside of the Profilepart/Platepart whose location mark will be placed
'5) Plate Part Upside
'6) Bend Knuckles Collection - Collection of wirebodies of Knuckles
'7) Ignore Knuckles Collection - Collection of wirebodies of Knuckles
'8) Split Knuckles Collection - Collection of wirebodies of Knuckles
'
'Output of this Function is: GeomCol3D of the Knuckle Fitting Marks
' ************************************************************************************************************
Public Sub CreateKnuckleFittingMarks(ByVal oLocationMarkCS As IJComplexString, oConnectionData As ConnectionData, ByVal oPart As Object, ByVal sMoldedSide As String, ByVal UpSide As Long, ByVal oBendKnuckleWBColl As Collection, ByVal oIgnoreKnuckleWBColl As Collection, ByVal oSplitKnuckleWBColl As Collection, oGeomCol3d As IJMfgGeomCol3d)
    Const METHOD = "CreateKnuckleFittingMarks"
    On Error GoTo ErrorHandler

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPart

    'Initialize the profile wrapper and the Physical Connection wrapper
    Dim oSDConWrapper As Object
    If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
        Set oSDConWrapper = New StructDetailObjects.PlatePart
    ElseIf TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Then
        Set oSDConWrapper = New StructDetailObjects.ProfilePart
    Else
        Set oSDConWrapper = New StructDetailObjects.MemberPart
    End If
    Set oSDConWrapper.object = oConnectionData.ToConnectable

    Dim oKnucklePtCol As New Collection

    Dim jCount As Integer
    Dim oKnuckleCollection As New Collection

    Dim oLocationMarkWB As IJWireBody
    Set oLocationMarkWB = m_oMfgRuleHelper.ComplexStringToWireBody(oLocationMarkCS)

    Dim oLocationMarkMB As IJDModelBody
    Set oLocationMarkMB = oLocationMarkWB

    If Not oBendKnuckleWBColl Is Nothing Then
        For jCount = 1 To oBendKnuckleWBColl.Count
            oKnuckleCollection.Add oBendKnuckleWBColl.Item(jCount)
        Next jCount
    End If

    Dim oPlateNormal As IJDVector
    Dim oPrVector As IJDVector

    Dim oMoniker As IMoniker
    Dim oGeom3d As IJMfgGeom3D
    Dim oMarkingInfo As MarkingInfo
    Dim oSystemMark As IJMfgSystemMark

    Dim oThicknessDirVec As New DVector
    'Set oThicknessDirVec = GetThicknessDirectionVector(oLocationMarkCS, oSDConWrapper, sMoldedSide)
    'oThicknessDirVec.Length = 1

    Dim oCurve As IJCurve

    Dim oFittingMarkCS As IJComplexString
    Dim oKnuStartPos As IJDPosition, oKnuEndPos As IJDPosition

    Dim oGeomHelper As New MfgGeomHelper

    Dim oPoint As IJDPosition
    Dim dValue As Double
    Dim dMinDist As Double

    Dim oKnuckleWireBody As IJWireBody
    Dim oKnuckleMB As IJDModelBody

    Dim dStartX As Double, dStartY As Double, dEndY As Double, dStartZ As Double, dEndX As Double, dEndZ As Double, dStart2X As Double, dStart2Y As Double, dStart2Z As Double
    Dim dEndPar As Double, vTanX As Double, vTanY As Double, vTanZ As Double, vTan2X As Double, vTan2Y As Double, vTan2Z As Double

    Dim oTanVector As New DVector
    Dim oNormalVector As New DVector

    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    Dim oPlateSurfaceBody As IJSurfaceBody
    Set oPlateSurfaceBody = oPlateWrapper.GetSurfacePort(UpSide).Geometry

    Dim oProjStartPos As IJDPosition

    Set oCurve = oLocationMarkCS


    '******************************Code For Bend Knuckles Only **********************************

    '********Projecting the End Points of the Bend Knuckle WireBody on the Profile Location Mark (PC)******'
    If oKnuckleCollection.Count <> 0 Then
        For jCount = 1 To oKnuckleCollection.Count

            Dim oProjEndPos As IJDPosition
            Set oKnuckleWireBody = oKnuckleCollection.Item(jCount)
            Set oKnuckleMB = oKnuckleWireBody
            oLocationMarkMB.GetMinimumDistance oKnuckleMB, oProjStartPos, oProjEndPos, dMinDist
            oKnucklePtCol.Add oProjStartPos
            Set oProjStartPos = Nothing
            Set oProjEndPos = Nothing
        Next jCount

        '******************************************End*********************'

        '****************Creating the Fitting Mark on the PC at the Bend Knuckle Position****************'


        For jCount = 1 To oKnucklePtCol.Count

            Set oPoint = oKnucklePtCol.Item(jCount)

            'Getting the Tangent at the Knuckle Position along the
            Set oTanVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPoint)
            oTanVector.Length = 1

            oMfgMGHelper.ProjectPointOnSurfaceBody oPlateWrapper.GetSurfacePort(UpSide).Geometry, oPoint, oProjStartPos, oPrVector
            oPlateSurfaceBody.GetNormalFromPosition oProjStartPos, oPlateNormal

            oPlateNormal.Length = 1


            'Gettting the normal to the Location Mark Curve At the Knuckle Point
            Set oNormalVector = oPlateNormal.Cross(oTanVector)
            oNormalVector.Length = 1

            Set oThicknessDirVec = GetThicknessDirectionVectorAtAGivenPos(oPoint, oSDConWrapper, Nothing, sMoldedSide)
            oThicknessDirVec.Length = 1



            '**********The Below code Ensures that the Knuckle Fitting Mark respects the Thickness Direction of the Connected Part*******
            dValue = oNormalVector.Dot(oThicknessDirVec)
            If dValue < 0 Then
                oNormalVector.Length = -1
            End If
            '************End of the code************

            Set oFittingMarkCS = CreateKnuckleFittingLine(oPlateWrapper.GetSurfacePort(UpSide).Geometry, oPoint, oNormalVector, END_FITTING_MARK_LENGTH)

            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oFittingMarkCS
            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
            oSystemMark.Set3dGeometry oGeom3d

            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            oGeomCol3d.AddGeometry 1, oGeom3d
        Next jCount

        Set oKnuckleCollection = Nothing
        Set oKnuckleWireBody = Nothing
        Set oKnucklePtCol = Nothing
    End If

    '***************************************End of Code for Bend Knuckle Collection***********************************

    Set oKnucklePtCol = New Collection
    Set oKnuckleCollection = New Collection


    '*************************Code for Knuckle Fitting Marks if the Type of Knuckle is Split/ Ignore******************


    If Not oIgnoreKnuckleWBColl Is Nothing Then
        For jCount = 1 To oIgnoreKnuckleWBColl.Count
            oKnuckleCollection.Add oIgnoreKnuckleWBColl.Item(jCount)
        Next jCount
    End If

    If Not oSplitKnuckleWBColl Is Nothing Then
        For jCount = 1 To oSplitKnuckleWBColl.Count
            oKnuckleCollection.Add oSplitKnuckleWBColl.Item(jCount)
        Next jCount
    End If

    If oKnuckleCollection.Count > 0 Then
        For jCount = 1 To oKnuckleCollection.Count

            Dim oKnuckle As IJKnuckle
            Set oKnuckle = oKnuckleCollection.Item(jCount)

            Set oKnuckleWireBody = oKnuckle.GetKnuckleCurve
            oKnuckleWireBody.GetEndPoints oKnuStartPos, oKnuEndPos

            Set oKnuckleMB = oKnuckleWireBody
            oLocationMarkMB.GetMinimumDistance oKnuckleMB, oProjStartPos, oProjEndPos, dMinDist

            oKnucklePtCol.Add oProjStartPos

        Next jCount


        For jCount = 1 To oKnucklePtCol.Count
            Set oPoint = oKnucklePtCol.Item(jCount)

            '*** New Implementation ***'
            Dim oEndPtOf1st As New DPosition
            Dim oStartPtOf2nd As New DPosition

            Set oStartPtOf2nd = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oPoint, 0.03)
            Set oEndPtOf1st = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oPoint, -0.03)

            oTanVector.Set oStartPtOf2nd.x - oEndPtOf1st.x, oStartPtOf2nd.y - oEndPtOf1st.y, oStartPtOf2nd.z - oEndPtOf1st.z
            oTanVector.Length = 1


            oMfgMGHelper.ProjectPointOnSurfaceBody oPlateWrapper.GetSurfacePort(UpSide).Geometry, oPoint, oProjStartPos, oPrVector
            oPlateSurfaceBody.GetNormalFromPosition oProjStartPos, oPlateNormal

            oPlateNormal.Length = 1

            Set oNormalVector = oPlateNormal.Cross(oTanVector)
            oNormalVector.Length = 1

            Set oThicknessDirVec = GetThicknessDirectionVectorAtAGivenPos(oPoint, oSDConWrapper, Nothing, sMoldedSide)

            '**********The Below code Ensures that the Knuckle Fitting Mark respects the Thickness Direction of the Connected Part*******
            dValue = oNormalVector.Dot(oThicknessDirVec)
            If dValue < 0 Then
                oNormalVector.Length = -1
            End If
            '************End of the code************

            Set oFittingMarkCS = CreateKnuckleFittingLine(oPlateWrapper.GetSurfacePort(UpSide).Geometry, oPoint, oNormalVector, END_FITTING_MARK_LENGTH)
            'Create a SystemMark object to store additional information

            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark

            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oFittingMarkCS
            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
            oSystemMark.Set3dGeometry oGeom3d

            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker

            oGeomCol3d.AddGeometry 1, oGeom3d
        Next jCount
    End If


CleanUp:

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Sub


' ***********************************************************************************
' Public Function CreateIndirectBracketLocationMark()
'
' Description:  Creates Indirect Bracket Location Mark when brackets are attached to profile on base plate
'               and wither the bracket part is in HARITSUKE stage code (in Ogumi) or in same assembly of base plate
' ***********************************************************************************
Public Sub CreateIndirectBracketLocationMark(Part As Object, oProfilePart As Object, ByVal ReferenceObjColl As JCmnShp_CollectionAlias, oGeomCol3d As IJMfgGeomCol3d, _
                                        sProdName As String)

    Const METHOD = "CreateIndirectBracketLocationMark"
    On Error GoTo ErrorHandler

    Dim oAttribute As IJDAttribute
    Dim oAttrCol As IJDAttributesCol


    Dim oSDProfileWrapper As Object
    'Create the SD profile Wrapper and initialize it
    If TypeOf oProfilePart Is IJStiffenerPart Then
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = oProfilePart
    ElseIf TypeOf Part Is IJBeamPart Then
        Set oSDProfileWrapper = New StructDetailObjects.BeamPart
        Set oSDProfileWrapper.object = oProfilePart
    End If

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart

    Dim oPartInfo As IJDPartInfo
    Set oPartInfo = New PartInfo

    'Get the Profile Part Physically Connected Objects
    Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    Set oPartSupport.Part = oProfilePart
    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol

    Dim oMfgMGHelper    As MfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper

    'Set oConObjCol = GetPhysicalConnectionData(oProfilePart, ReferenceObjColl, True)

    Dim oFirstConObj As IJConnectable, oSecondConObj As IJConnectable
    Dim oConnections As IJElements
    Dim bIsConnected As Boolean

    Set oFirstConObj = Part

    Dim oConnectionData As ConnectionData
    Dim lGeomCount As Long
    lGeomCount = 1

    Dim nIndex As Integer
    For nIndex = 1 To oConObjCol.Count
        Dim oMfgGeom3d As IJMfgGeom3D

        Set oConnectionData.ToConnectable = oConObjCol.Item(nIndex)
        Set oConnectionData.AppConnection = oConnCol.Item(nIndex)

        Dim oNamed As IJNamedItem
        Set oNamed = oConnectionData.ToConnectable

        Dim sConnectedPartName As String
        sConnectedPartName = oNamed.Name

        On Error GoTo NextMark
        If Not TypeOf oConnectionData.ToConnectable Is IJProfilePart Or Not TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Then

            'Check if the bracket is already connected to this PLATE.
            ' ... if already connected, then dont need to mark
            Set oSecondConObj = oConnectionData.ToConnectable
            oFirstConObj.isConnectedTo oSecondConObj, bIsConnected, oConnections

            If bIsConnected Then GoTo NextMark

            'Check if it is bracket
            If CheckIfPartIsBracket(oConnectionData.ToConnectable) = True Then

                If (GetAssemblyType(oConnectionData.ToConnectable) = 1 And _
                    UCase(GetProductionRoutingStageCode(oConnectionData.ToConnectable)) = "StageA") Or _
                    ArePartsInSameAssembly(Part, oConnectionData.ToConnectable) Then

                        'Create support only mark for Indirect Location Mark for bracket
                        'Get the point on profile location mark to put support only mark
                        Dim oPort As IJPort
                        Set oPort = oThisPortCol.Item(nIndex)

                        Dim oAppConnection As IJAppConnection
                        Set oAppConnection = oConnCol.Item(nIndex)

                        Dim oRootPartSystem As IJSystem
                        Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
                        Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
                        oStructDetailHelper.IsPartDerivedFromSystem oConnectionData.ToConnectable, oRootPartSystem, True

                        'Method #1 to get PC
                        Dim oConnectable            As IJConnectable
                        Dim oStructConnectable      As IJStructConnectable
                        Dim oRootBoundingConnection As IJStructConnection
                        Set oStructConnectable = oRootPartSystem
                        Set oNamed = oRootPartSystem
                        Set oConnectable = oProfilePart
                        'oStructConnectable.GetRootConnection oConnectable, oRootBoundingConnection

                        Dim oPCWireBody As IJWireBody
                        'Set oPCWireBody = oRootBoundingConnection.ModelBody

                        Dim oStartPt As IJDPosition, oEndPt As IJDPosition
                        'oPCWireBody.GetEndPoints oStartPt, oEndPt

                        Dim oPCWireMB As IJDModelBody
                        'Set oPCWireMB = oPCWireBody

                        Dim oPtOnProfLocMark As New DPosition
                        Dim dMinDist As Double
                        'oPCWireMB.GetMinimumDistanceFromPosition oEndPt, oPtOnProfLocMark, dMinDist

                        'Method #2 to get PC
                        Dim oProfilePartSupport As IJProfilePartSupport
                        Dim oPartSupp As IJPartSupport

                        Set oPartSupp = New GSCADSDPartSupport.ProfilePartSupport
                        'MsgBox TypeName(oProfilePart)
                        Dim oIJProfilePart As IJProfilePart
                        Set oIJProfilePart = oProfilePart
                        Set oPartSupp.Part = oIJProfilePart

                        Set oProfilePartSupport = oPartSupp
                        Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
                        eSideOfConnectedObjectToBeMarked = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint

                        Dim oTeeWire As IJWireBody
                        Dim oVector As IJDVector
                        Dim bContourTee As Boolean
                        bContourTee = oSDProfileWrapper.Connection_ContourTee(oConnectionData.AppConnection, SideA, oTeeWire, oVector)
                        Set oPCWireMB = oTeeWire

                        Dim oPCCS As IJComplexString, oProjCS As IJComplexString
                        oMfgMGHelper.WireBodyToComplexString oTeeWire, oPCCS

                        'oPCWireMB.GetMinimumDistanceFromPosition oEndPt, oPtOnProfLocMark, dMinDist

                        Dim oSDPlateWrapper As StructDetailObjects.PlatePart
                        Set oSDPlateWrapper = New StructDetailObjects.PlatePart
                        Set oSDPlateWrapper.object = Part

                        Dim oPlateSurface As IJSurfaceBody
                        Set oPlateSurface = oSDPlateWrapper.BasePort(BPT_Base).Geometry

                        oMfgMGHelper.ProjectComplexStringToSurface oPCCS, oPlateSurface, oVector, oProjCS

                        Dim sMarkingName As String
                        sMarkingName = "&" & sConnectedPartName

                        If Not oPCCS Is Nothing Then
                            Set oMfgGeom3d = CreateIndirectMarkFromCS(oPCCS, STRMFG_ANNOTATION_MARK, _
                                                                    True, oAppConnection, sMarkingName)
                        End If

                        If Not oMfgGeom3d Is Nothing Then
                            oGeomCol3d.AddGeometry oGeomCol3d.GetCount + 1, oMfgGeom3d
                        End If

                End If
            End If
        End If
NextMark:
    Next nIndex

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
    GoTo CleanUp

End Sub



Public Function CreateIndirectMarkFromCS(oCS As IJComplexString, enumGeomType As StrMfgGeometryType, bSupportOnly As Boolean, _
                                    oAppConnection As IJAppConnection, sMarkingName As String) As IJMfgGeom3D

    Const METHOD = "CreateIndirectMarkFromCS"
    On Error GoTo ErrorHandler

    'Create Geom3D
    Dim oResourceManager As Object
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oSystemMark As IJMfgSystemMark
    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

    'QI for the MarkingInfo object on the SystemMark
    Dim oMarkingInfo As MarkingInfo
    Set oMarkingInfo = oSystemMark
    oMarkingInfo.Name = sMarkingName

    'oSystemMark.SetMarkingSide 0
    Dim oGeom3d As IJMfgGeom3D
    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)

    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype enumGeomType
    oGeom3d.IsSupportOnly = bSupportOnly

    Dim oMoniker As IMoniker
    'Set oMoniker = m_oMfgRuleHelper.GetMoniker(oAppConnection)
    'oGeom3d.PutMoniker oMoniker
    oGeom3d.PutGeometrytype enumGeomType

    oSystemMark.Set3dGeometry oGeom3d

    Set CreateIndirectMarkFromCS = oGeom3d

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
    GoTo CleanUp

End Function

Public Sub CreateIndirectMemberMarks(Part As Object, ByVal ReferenceObjColl As JCmnShp_CollectionAlias, oGeomCol3d As IJMfgGeomCol3d)

    Const METHOD = "CreateIndirectMemberMarks"
    On Error GoTo ErrorHandler

    'Get the Plate Part Physically Connected Objects
    Dim oConMembersCol As IJElements
    'Set oConObjsCol = oSDPlateWrapper.ConnectedObjects
    Set oConMembersCol = GetMembersInSameAssembly(Part)
    If oConMembersCol Is Nothing Then
        'No connected objects so we can leave
        GoTo CleanUp
    End If

    If oConMembersCol.Count < 1 Then
        GoTo CleanUp
    End If

    Dim oCurrMember As ISPSMemberPartPrismatic

    Dim ii As Integer
    Dim oNamed As IJNamedItem
    For ii = 1 To oConMembersCol.Count
        Set oNamed = oConMembersCol.Item(ii)
        Set oCurrMember = oConMembersCol.Item(ii)
        Dim dMinStartX As Double, dMinStartY As Double, dMinStartZ As Double
        Dim dMaxStartX As Double, dMaxStartY As Double, dMaxStartZ As Double
        Dim dMinEndX As Double, dMinEndY As Double, dMinEndZ As Double
        Dim dMaxEndX As Double, dMaxEndY As Double, dMaxEndZ As Double

        oCurrMember.GetEndFacePositions dMinStartX, dMinStartY, dMinStartZ, dMaxStartX, dMaxStartY, dMaxStartZ, _
                                        dMinEndX, dMinEndY, dMinEndZ, dMaxEndX, dMaxEndY, dMaxEndZ

        Dim oAxisCurve As IJCurve
        Set oAxisCurve = oCurrMember.Axis

        Dim oMemberUtil As IJSPSMemberUtils
        Set oMemberUtil = New SPSUtils

        Dim oPrimaryVec As IJDVector
        Dim oSecondaryVec As IJDVector
        oMemberUtil.GetMemberOrientationVectors oCurrMember.MemberSystem, oPrimaryVec, oSecondaryVec

        'Create Geom3d
        Dim oMarkLine As IJLine
        Set oMarkLine = New Line3d

        oMarkLine.DefineBy2Points dMinEndX, dMinEndY, dMinEndZ, dMinEndX + oPrimaryVec.x, dMinEndY + oPrimaryVec.y, dMinEndZ + oPrimaryVec.z

        Dim oCS As IJComplexString
        Set oCS = New ComplexString3d
        oCS.AddCurve oMarkLine, False

        Dim oSDPlateWrapper As StructDetailObjects.PlatePart
        Set oSDPlateWrapper = New StructDetailObjects.PlatePart
                        Set oSDPlateWrapper.object = Part

        Dim oPlateSurface As IJSurfaceBody
        Set oPlateSurface = oSDPlateWrapper.BasePort(BPT_Base).Geometry

        Dim oVector As IJDVector
        Dim oMfgMGHelper    As MfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper

        Dim oProjCS As IJComplexString

        oMfgMGHelper.ProjectComplexStringToSurface oCS, oPlateSurface, oVector, oProjCS

        Dim oDummy As IJAppConnection

        Set oNamed = oConMembersCol.Item(ii)
        Dim sMemberName As String
        sMemberName = oNamed.Name

        Dim oMfgGeom3d As IJMfgGeom3D
        Set oMfgGeom3d = CreateIndirectMarkFromCS(oProjCS, STRMFG_MEMBER_LOCATION_MARK, False, oDummy, sMemberName)

        oGeomCol3d.AddGeometry oGeomCol3d.GetCount + 1, oMfgGeom3d

    Next ii

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
    GoTo CleanUp

End Sub

Public Function GetMembersInSameAssembly(Part As Object) As IJElements

    Const METHOD = "GetMembersInSameAssembly"
    On Error GoTo ErrorHandler

    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = Part

    'Check if ThisPartAssembly is Nothing
    If oAssemblyChild.Parent Is Nothing Then
        'Set GetMembersInSameAssembly = vbNullString
        Exit Function
    End If

    If Not TypeOf oAssemblyChild.Parent Is IJPlanningAssembly Then
        'Set GetMembersInSameAssembly = vbNullString
        Exit Function
    End If

    Dim oThisPartAssy   As IJPlanningAssembly
    Set oThisPartAssy = oAssemblyChild.Parent

    Dim oPlnIntHelper As GSCADPlnIntHelper.IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper

    'Get first hierarchy children assemblies of ThisAssembly
    Set GetMembersInSameAssembly = oPlnIntHelper.GetAssemblyChildren(oThisPartAssy, "ISPSMemberPartPrismatic", False)

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
    GoTo CleanUp
End Function

Public Function ArePartsInSameAssembly(oPart1 As Object, oPart2 As Object) As Boolean

Const METHOD = "ArePartsInSameAssembly"
On Error GoTo ErrorHandler

    ArePartsInSameAssembly = False

    Dim oAssemblyChild1 As IJAssemblyChild, oAssemblyChild2 As IJAssemblyChild
    Set oAssemblyChild1 = oPart1
    Set oAssemblyChild2 = oPart2

    'Check if ThisPartAssembly is Nothing
    If Not oAssemblyChild1.Parent Is Nothing And Not oAssemblyChild2.Parent Is Nothing Then
        If oAssemblyChild1.Parent Is oAssemblyChild2.Parent Then
            ArePartsInSameAssembly = True
        End If
    End If

CleanUp:

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
End Function



Public Function GetAttributeColl(pObject As Object, strInterfaceName As String) As IJDAttributesCol
Const METHOD = "GetAttribute"
On Error GoTo ErrorHandler

    Dim oAttrMetaData   As IJDAttributeMetaData
    Set oAttrMetaData = pObject

    Dim varOldAttribInt As Variant
    varOldAttribInt = oAttrMetaData.IID(strInterfaceName)
    Set oAttrMetaData = Nothing

    Dim oAttributes     As IJDAttributes
    Set oAttributes = pObject

    Dim oAttributesCol  As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(varOldAttribInt)
    Set GetAttributeColl = oAttributesCol

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


' ***********************************************************************************
' Public Function GetAssemblyType
'
' Description:  Method will give the Type of Assembly
'
' ***********************************************************************************
Private Function GetAssemblyType(oPart As Object) As Long
    Const METHOD = "GetAssemblyType"
    On Error GoTo ErrorHandler

    Dim oAssemblyChild As IJAssemblyChild
    Dim oUAFolder As IJAssemblyChild
    Dim oAssemblybase As IJAssemblyBase
    Dim lAssemblyType As Long

    Set oAssemblyChild = oPart

    If oAssemblyChild.Parent Is Nothing Then 'This will Exist when the Part is under Config Root
        Exit Function
    ElseIf TypeOf oAssemblyChild.Parent Is IJAssemblyBase Then
        Set oAssemblybase = oAssemblyChild.Parent
    'Below Condition will occurs when the Part is under UnProcessedParts or FailedParts
    ElseIf TypeOf oAssemblyChild.Parent Is IJPlnUnprocessedParts Or TypeOf oAssemblyChild.Parent Is IJPlnFailedParts Then
        Set oUAFolder = oAssemblyChild.Parent
        Set oAssemblybase = oUAFolder.Parent
    End If

    If Not oAssemblybase Is Nothing Then
        lAssemblyType = oAssemblybase.Type
    Else
        Exit Function
    End If

    GetAssemblyType = lAssemblyType

    Exit Function

ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function GetMarkingLinesForThisPart(ByVal Part As Object, ByVal oReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
Const METHOD = "GetMarkingLinesForThisPart"
On Error GoTo ErrorHandler

    Dim index As Long
    Set GetMarkingLinesForThisPart = New Collection

    For index = 1 To oReferenceObjColl.Count
        If TypeOf oReferenceObjColl.Item(index) Is IJMfgMarkingLinesData Then
            Dim oMfgMarkingLinesData As IJMfgMarkingLinesData
            Dim oMfgMarkingLines_AE As IJMfgMarkingLines_AE
            Dim oMarkingPart As Object
            
            Set oMfgMarkingLinesData = oReferenceObjColl.Item(index)
            Set oMfgMarkingLines_AE = oMfgMarkingLinesData.GetMfgMarkingLines_AE
            Set oMarkingPart = oMfgMarkingLines_AE.GetMfgMarkingPart
            
            If oMarkingPart Is Part Then
                GetMarkingLinesForThisPart.Add oReferenceObjColl.Item(index)
            End If
            
            Set oMarkingPart = Nothing
            Set oMfgMarkingLines_AE = Nothing
            Set oMfgMarkingLinesData = Nothing
        End If
    Next

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
