Attribute VB_Name = "LadderBSpanAndOrigin"

Option Explicit
'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  File:  LadderBSpanAndOrigin.bas
'
'  Description: The file contains the functions used for finding and setting the Origin and the Height of the ladder.
'
'  History:
'   03/08/2000 - JM - Creation
'   04/18/2000 - JM - Get and Set of Attributes has changed.  Given the collection
'                                        now, so no need to have the IID to find it.
'   06/02/2000 - JM - Changed the Matrix and Offset Directions so that the positive
'                     offset direction is correct and the lighting is correct.
'   Jan 2002    AP    V3 Additions new Functionality
'   09-May-06   JMS DI#97751 - All lines are constructed through the
'                   geometry factory instead of being "new"ed
'  June 12, 2006 AS TR#88968: Added code to normalize and make the vectors in the xform matrix orthogonal
'  May 05, 2007  AS CR#116177 Added new functions used for mirorr copy support
'  09/12/2007 C.C.P. Created for CR113595 Ladder on Corrugated Bulkhead. Copied from GenericOriginAndSpan.bas.
'  01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
'  Jul-17-2008  WR  TR-CP-144430 - Changing vertical offset of ladder results in incorrect geometry.
'                                  Span was not taking into account the top extension thus Span was computed incorrectly.
'  BS Mar-06-2009   CR-154170    - removed DetailPhysicalView and added CenterlineView mode, implemented CenterlineView to include side frames, steps, hoops and support leg pitch
'
'******************************************************************
    
Private Const MODULE = "SPSLadderBMacros.LadderBSetSpanAndOrigin"
Private m_oErrors As IJEditErrors
Private Const TOL = 0.0000001
Public Const E_FAIL = -2147467259
Private oLocalizer As IJLocalizer

Public Enum eRepresentationType
    SimpleRep = 1
    centerlinerep
End Enum



'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
Public Function GetDistTolerance() As Double
On Error GoTo ErrorHandler

    Dim oGeomSrvs As IGeometryServices

    Set oGeomSrvs = New IngrGeom3D.GeometryFactory
    GetDistTolerance = oGeomSrvs.DistTolerance
    
    Set oGeomSrvs = Nothing
    
Exit Function
ErrorHandler:
    GetDistTolerance = 0.000001
End Function

Public Function GetGeomPI() As Double
On Error GoTo ErrorHandler

    Dim oGeomSrvs As IGeometryServices

    Set oGeomSrvs = New IngrGeom3D.GeometryFactory
    GetGeomPI = oGeomSrvs.GeomPI
    
    Set oGeomSrvs = Nothing
    
Exit Function
ErrorHandler:
    GetGeomPI = 4 * Atn(1)
End Function

Public Function GetAngTolerance() As Double
On Error GoTo ErrorHandler

    Dim oGeomSrvs As IGeometryServices

    Set oGeomSrvs = New IngrGeom3D.GeometryFactory
    GetAngTolerance = oGeomSrvs.AngTolerance
    
    Set oGeomSrvs = Nothing
    
Exit Function
ErrorHandler:
    GetAngTolerance = 0.0000000001
End Function
'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.


'This function will be used to set the Position and Matrix of the Ladder and
'determine the height of the ladder.
'This is just a generic function that can be rewriten or changed for
'each symbol. (LadderBO, Top, Bottom, Ref, OccAttributeCol)
Public Function SetOriginAndSpan(LadderBo As ISPSLadder, ByRef Top As Object, _
                                 ByRef bottom As Object, ByRef Ref As Object, _
                                 PartOccInfoCol As IJDInfosCol) As Variant
Set oLocalizer = New IMSLocalizer.Localizer
oLocalizer.Initialize App.Path & "\" & "SPSLadderBMacros"
On Error GoTo ErrorHandler
Const METHOD = "SetOriginAndSpan"
    
    Dim Flip As Boolean
    Dim oBottomSurf As IJSurface
    Dim oBottomPlane As IJPlane
    Dim BottomNormal As DVector
    Set BottomNormal = New DVector
    Dim dHoriz As Double
    Dim dVert As Double
    Dim TopExt As Double
    Dim BottomExt As Double
    Dim gPosOrig As DPosition
    Dim oMountingWall As IJPlane
    Dim oTopSurf As IJSurface
    Dim oMountingNormal As DVector
    Set oMountingNormal = New DVector
    Dim oTopArc As IJArc
    Dim oTopEdge As IJLine
    Dim oTopCurve As IJCurve
    Dim oRefEdge As IJLine
    Dim oRefCurve As IJCurve
    Dim oRefPlane As IJPlane
    Dim Matrix As DT4x4
    Set Matrix = New DT4x4
    Dim startpt As DPosition
    Set startpt = New DPosition
    Dim endpt As DPosition
    Set endpt = New DPosition
    Dim SParamU As Double, SParamV As Double, EParamU As Double, EParamV As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim x As Double, y As Double, z As Double
    Dim Rootx As Double, Rooty As Double, Rootz As Double
    
    Dim BoolArc As Boolean
    BoolArc = False
    Dim OccAttrs As IJDAttributes
    Set OccAttrs = LadderBo
    
    
    'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    'Boolean flags to determine how we should proceed.
    Dim bRotationOffsets As Boolean
    bRotationOffsets = True 'Assume we can do RotationOffsets.
    Dim bRotation As Boolean
    bRotation = True 'Assume we can do Rotation.

    'Attributes to define positions of ladder left side, ladder center, ladder right side.
    Dim lJustify As Long
    lJustify = 0
    Dim dWidth As Double
    dWidth = 0#
    Dim dSideFrameThick As Double
    dSideFrameThick = 0#
    Dim dSideFrameWidth As Double
    dSideFrameWidth = 0#
    Dim dSupportLegThick As Double
    dSupportLegThick = 0#
    Dim dSupportLegWidth As Double
    dSupportLegWidth = 0#
    Dim dWallOffset As Double
    dWallOffset = 0#
    
    'Attributes specific to doing rotation and offsets.
    Dim dRotation As Double
    dRotation = 0#
    Dim dLeftSupportLegLength As Double
    dLeftSupportLegLength = 0#
    Dim dRightSupportLegLength As Double
    dRightSupportLegLength = 0#
    Dim dExtraOffset As Double
    dExtraOffset = 0#
    Dim dLeftSupportLegOffset As Double 'Not an attribute, now derived from corresponding length attribute.
    dLeftSupportLegOffset = 0#
    Dim dRightSupportLegOffset As Double 'Not an attribute, now derived from corresponding length attribute.
    dRightSupportLegOffset = 0#
                       
    'Temporary working variables and objects.
    Dim dXtmp As Double
    Dim dYtmp As Double
    Dim dZtmp As Double
    dXtmp = 0#
    dYtmp = 0#
    dZtmp = 0#
    Dim dXtmp1 As Double
    Dim dYtmp1 As Double
    Dim dZtmp1 As Double
    dXtmp1 = 0#
    dYtmp1 = 0#
    dZtmp1 = 0#
    Dim dXtmp2 As Double
    Dim dYtmp2 As Double
    Dim dZtmp2 As Double
    dXtmp2 = 0#
    dYtmp2 = 0#
    dZtmp2 = 0#
    Dim oVecTmp2 As DVector
    Set oVecTmp2 = New DVector
    Dim oVecTemp As DVector
    Set oVecTemp = New DVector
    Dim oPosTemp As DPosition
    Set oPosTemp = New DPosition
    oPosTemp.Set 0, 0, 0
    Dim dTemp As Double
    dTemp = 0#
    Dim oTestObject As Object
    Set oTestObject = Nothing

    'Ladder Factory and services, and Geometry Factory and Services.
    Dim oSPSLadderFactory As SPSLadderFactory
    Dim oISpsLadderUtil As ISPSLadderUtil
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Dim oGeometryServices As IGeometryServices
    
    'To get the oTopDir, oTopAway, gPosOrig from transform matrix of LadderBO with Flip/Direc already applied.
    Dim oVecXRotTopDir As DVector
    Set oVecXRotTopDir = New DVector
    Dim oVecYRotTopAway As DVector
    Set oVecYRotTopAway = New DVector
    Dim oVecZRotVertUp As DVector
    Set oVecZRotVertUp = New DVector
    Dim oPosOriginPivot As DPosition
    Set oPosOriginPivot = New DPosition
    
    'To get center origin and left and right plane root point positions and normal vectors.
    Dim dPos As Double  'Necessary adjustment to ladder position based on Justification.
    dPos = 0#
    Dim dCenterX As Double
    Dim dCenterY As Double
    Dim dCenterZ As Double
    Dim dLeftXroot As Double
    Dim dLeftYroot As Double
    Dim dLeftZroot As Double
    Dim dRightXroot As Double
    Dim dRightYroot As Double
    Dim dRightZroot As Double
    Dim dLeftXnorm As Double
    Dim dLeftYnorm As Double
    Dim dLeftZnorm As Double
    Dim dRightXnorm As Double
    Dim dRightYnorm As Double
    Dim dRightZnorm As Double
    Dim oVecLeftNormal As DVector
    Set oVecLeftNormal = New DVector
    Dim oVecRightNormal As DVector
    Set oVecRightNormal = New DVector
    Dim oPosLeftRoot As DPosition
    Set oPosLeftRoot = New DPosition
    oPosLeftRoot.Set 0, 0, 0
    Dim oPosRightRoot As DPosition
    Set oPosRightRoot = New DPosition
    oPosRightRoot.Set 0, 0, 0
    Dim oPosCenterOrigin As DPosition
    Set oPosCenterOrigin = New DPosition
    oPosCenterOrigin.Set 0, 0, 0
    
    'To create left and right planes and planar surfaces.
    Dim oLeftPlane As IJPlane
    Dim oRightPlane As IJPlane
    Dim oLeftSurface As IJSurface
    Dim oRightSurface As IJSurface
    
    'To intersect left planar surface with top curve.
    Dim oLeftIntersectElements As IJElements 'Don't preallocate (New) the collection, Intersect will allocate and return it.
    'Set oLeftIntersectElements = New JObjectCollection 'Elements
    Dim LeftIntersectCode As Geom3dIntersectConstants
    Dim oLeftIntersectPoint As IJPoint
    Dim dLeftIntersectX As Double
    Dim dLeftIntersectY As Double
    Dim dLeftIntersectZ As Double
    Dim oPosLeftIntersect As DPosition
    Set oPosLeftIntersect = New DPosition
    
    'To intersect right planar surface with top curve.
    Dim oRightIntersectElements As IJElements 'Don't preallocate (New) the collection, Intersect will allocate and return it.
    'Set oRightIntersectElements = New JObjectCollection 'Elements
    Dim RightIntersectCode As Geom3dIntersectConstants
    Dim oRightIntersectPoint As IJPoint
    Dim dRightIntersectX As Double
    Dim dRightIntersectY As Double
    Dim dRightIntersectZ As Double
    Dim oPosRightIntersect As DPosition
    Set oPosRightIntersect = New DPosition

    'To intersect top curve with left and right planar surfaces.
    Dim lLeftNumIntersects As Long
    Dim lLeftNumOverlaps As Long
    Dim dLeftIntersectPoints() As Double
    Dim lRightNumIntersects As Long
    Dim lRightNumOverlaps As Long
    Dim dRightIntersectPoints() As Double
    
    'To check intersection points and choose closest to center.
    Dim lPointIndex As Long
    Dim lPointJndex As Long
    Dim bIntersect As Boolean
    Dim dDistIntSqr As Double
    Dim dMinDistIntSqr As Double
    
    'To extract the partial top curve (oSubCurve) from the whole top curve (oTopCurve).
    Dim oSubCurve As IJCurve
    Dim dLeftIntersectPar As Double
    Dim dRightIntersectPar As Double
    Dim dStartPar As Double
    Dim dDirecPar As Double
    Dim dEndPar As Double
    
    'To get MinMaxTypedCurveDir for the partial top curve (oSubCurve).
    Dim dDistanceTolerance As Double
    Dim dDirX As Double
    Dim dDirY As Double
    Dim dDirZ As Double
    Dim dMinX As Double
    Dim dMinY As Double
    Dim dMinZ As Double
    Dim dMaxX As Double
    Dim dMaxY As Double
    Dim dMaxZ As Double
    dMinX = 0
    dMinY = 0
    dMinZ = 0
    dMaxX = 0
    dMaxY = 0
    dMaxZ = 0
    
    'To get dot products and projected lengths to get extra offset.
    Dim dFrCtrToMinX As Double
    Dim dFrCtrToMinY As Double
    Dim dFrCtrToMinZ As Double
    Dim oVecFrCtrToMin As DVector
    Set oVecFrCtrToMin = New DVector
    Dim dDotProductMin As Double
    Dim dProjLengthMin As Double
    Dim dFrCtrToMaxX As Double
    Dim dFrCtrToMaxY As Double
    Dim dFrCtrToMaxZ As Double
    Dim oVecFrCtrToMax As DVector
    Set oVecFrCtrToMax = New DVector
    Dim dDotProductMax As Double
    Dim dProjLengthMax As Double
    
    'Total offset vector to use to translate all ladder positions away from curve.
    Dim dTotalOffset As Double
    Dim dTransX As Double
    Dim dTransY As Double
    Dim dTransZ As Double
    Dim oVecTotalOffset As DVector
    Set oVecTotalOffset = New DVector
    
    'To create infinite vertical lines and typed curves through translated ladder left, right side positions.
    Dim oLeftLine As IJLine
    Dim oRightLine As IJLine
    Dim oLeftCurve As IJCurve
    Dim oRightCurve As IJCurve

    'To get minimum distance between translated ladder left side and left end of sub curve.
    Dim dLeftMinDist As Double
    Dim dLeftMinPosX As Double
    Dim dLeftMinPosY As Double
    Dim dLeftMinPosZ As Double
    dLeftMinDist = 0
    dLeftMinPosX = 0
    dLeftMinPosY = 0
    dLeftMinPosZ = 0
    
    'To get minimum distance between translated ladder right side and right end of sub curve.
    Dim dRightMinDist As Double
    Dim dRightMinPosX As Double
    Dim dRightMinPosY As Double
    Dim dRightMinPosZ As Double
    dRightMinDist = 0
    dRightMinPosX = 0
    dRightMinPosY = 0
    dRightMinPosZ = 0
    'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    

    'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    On Error GoTo ErrorHandler  'Make sure any previous error and On Error Resume Next is reset/cleared.
    
    
    'Get a new LadderFactory and get a new LadderServices interface on it.
    'Get the LadderFactory.
    Set oSPSLadderFactory = New SPSLadderFactory
    If (oSPSLadderFactory Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Get the LadderServices.
    Set oISpsLadderUtil = oSPSLadderFactory.CreateLadderServices
    If (oISpsLadderUtil Is Nothing) Then
        GoTo ErrorHandler
    End If
    
    
    'Get a new GeometryFactory and get a new GeometryServices interface on it.
    'Get the GeometryFactory.
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    If (oGeometryFactory Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Get the GeometryServices.
    Set oGeometryServices = oGeometryFactory.GeometryServices
    If (oGeometryServices Is Nothing) Then
        GoTo ErrorHandler
    End If
        
        
    'Get attributes to define positions of ladder left side, ladder center, ladder right side.
    'Get the Justification attribute.
    lJustify = GetAttribute(OccAttrs, "Justification", PartOccInfoCol)
    'Get the Width attribute.
    dWidth = GetAttribute(OccAttrs, "Width", PartOccInfoCol)
    'Get the SideFrameThickness attribute.
    dSideFrameThick = GetAttribute(OccAttrs, "SideFrameThickness", PartOccInfoCol)
    'Get the SideFrameWidth attribute.
    dSideFrameWidth = GetAttribute(OccAttrs, "SideFrameWidth", PartOccInfoCol)
    'Get the SupportLegThickness attribute.
    dSupportLegThick = GetAttribute(OccAttrs, "SupportLegThickness", PartOccInfoCol)
    'Get the SupportLegWidth attribute.
    dSupportLegWidth = GetAttribute(OccAttrs, "SupportLegWidth", PartOccInfoCol)
    'Get the WallOffset attribute.
    dWallOffset = GetAttribute(OccAttrs, "WallOffset", PartOccInfoCol)
    
    'Get attributes specific doing rotation and offsets.
    'Get the Rotation attribute.
    dRotation = GetAttribute(OccAttrs, "Rotation", PartOccInfoCol)
    dLeftSupportLegLength = GetAttribute(OccAttrs, "LeftSupportLegLength", PartOccInfoCol)
    'Get the RightSupportLegLength attribute.
    dRightSupportLegLength = GetAttribute(OccAttrs, "RightSupportLegLength", PartOccInfoCol)
    'Get the ExtraOffset attribute.
    dExtraOffset = GetAttribute(OccAttrs, "ExtraOffset", PartOccInfoCol)
    dLeftSupportLegOffset = dLeftSupportLegLength - dSideFrameWidth / 2#
    dRightSupportLegOffset = dRightSupportLegLength - dSideFrameWidth / 2#
    'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    
    
    'Get the Normal of the Bottom Plane.
    '09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    'Removed a series of redundantly repeated On Error Resume Next statements from legacy code.
    'Left the first On Error Resume Next in the series in effect to avoid breaking legacy code.
    'Added IsTypeOf checks to try to avoid triggering errors by Set statements expected to fail.
    On Error Resume Next
    
    If (TypeOf bottom Is IJPlane) Then
        Set oBottomPlane = bottom
    End If
    
    LadderBo.GetMatrix Matrix

    If (TypeOf bottom Is IJSurface) Then
        Set oBottomSurf = bottom
    End If

    '09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    'Deleted former useless if-then path and replaced former else path with following if not then path.
    If Not (oBottomPlane Is Nothing) Then
        oBottomPlane.GetNormal x, y, z

        If (TypeOf oBottomPlane Is IJSurface) Then
            Set oBottomSurf = oBottomPlane
        End If
    End If
     
    BottomNormal.Set 0, 0, 1 ' x, y, z
    BottomNormal.Length = 1
    Matrix.IndexValue(8) = 0
    Matrix.IndexValue(9) = 0
    Matrix.IndexValue(10) = 1
    
    'set Origin of Ladder to 0,0,0
    Set gPosOrig = New DPosition
    gPosOrig.Set 0, 0, 0


    If (TypeOf Top Is IJLine) Then
        Set oTopEdge = Top
    End If

    If (TypeOf Top Is IJCurve) Then
        Set oTopCurve = Top
    End If
    If oTopCurve Is Nothing Then
        If (TypeOf oTopEdge Is IJCurve) Then
            Set oTopCurve = oTopEdge
        End If
    End If

    'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    If (oTopCurve Is Nothing) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        bRotation = False 'We cannot do Rotation.
    End If
    'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.

    If (TypeOf Top Is IJArc) Then
        Set oTopArc = Top
    End If

    If oTopEdge Is Nothing And Not oTopCurve Is Nothing Then
        BoolArc = True
    End If
    
    If Not oTopCurve Is Nothing Then
        If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Then
            BoolArc = True
        End If
    End If

    If (TypeOf Top Is IJSurface) Then
        Set oTopSurf = Top
    End If

    If (TypeOf Top Is IJPlane) Then
        Set oMountingWall = Top
    End If
    If oTopSurf Is Nothing Then
        If (TypeOf oMountingWall Is IJSurface) Then
            Set oTopSurf = oMountingWall
        End If
    End If
    
    ' The following check is to make sure that we have proper reference object
    ' if the top edge is a surface or a plane. Other wise, raise error.
    If Not oTopSurf Is Nothing Or Not oMountingWall Is Nothing Then
        If Ref Is Nothing Then
            GoTo RefEdgeMissingHandler
        End If
    End If

    If (TypeOf Ref Is IJLine) Then
        Set oRefEdge = Ref
    End If

    If (TypeOf Ref Is IJCurve) Then
        Set oRefCurve = Ref
    End If

    If oRefCurve Is Nothing Then
        If (TypeOf oRefEdge Is IJCurve) Then
            Set oRefCurve = oRefEdge
        End If
    End If

    If (TypeOf Ref Is IJPlane) Then
        Set oRefPlane = Ref
    End If

    Dim oTopDir As DVector
    Dim oTopAway As DVector
    Set oTopDir = New DVector
    Set oTopAway = New DVector
    
    If oTopCurve Is Nothing And oTopEdge Is Nothing Then
        If oTopSurf Is Nothing Then
            Err.Clear
            GoTo ErrorHandler
        End If
    End If
    
    If oBottomSurf Is Nothing And oBottomPlane Is Nothing Then
        Err.Clear
        GoTo ErrorHandler
    End If
    
    Dim Direc As Integer
    Direc = 1
    Flip = GetAttribute(OccAttrs, "TopSupportSide", PartOccInfoCol)
      
    If Flip Then
        Direc = -1
    End If
    
    Dim proj As IJProjection
    If (TypeOf Top Is IJProjection) Then
        Set proj = Top
    End If

    '01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
    'If Not BoolArc Then    'First step toward a fix for TR130780, TR124571, TR133869, TR133944 is to retrieve horizontal offset for curved cases.
        dHoriz = LadderBo.HorizVertOffset.HorizontalOffset
    'End If
    dVert = LadderBo.HorizVertOffset.VerticalOffset
    TopExt = GetAttribute(OccAttrs, "TopExtension", PartOccInfoCol)
    Dim Maxdist As Double, Mindist As Double, position As DPosition
    Set position = New DPosition
    
    If Not proj Is Nothing Or oTopSurf Is Nothing Then
        'Begin 01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
        'Call MaxOffset(dHoriz, Maxdist, Mindist, position, Top, Ref)
        Call MaxOffset(BoolArc, dHoriz, Maxdist, Mindist, position, Top, Ref)
        
        'If Not proj Is Nothing Then
        If Not proj Is Nothing Or BoolArc Then
        'End 01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
            LadderBo.GetPosition x, y, z
            If x <> position.x Or y <> position.y Or z <> position.z Then
                LadderBo.SetPosition position.x, position.y, position.z
            End If
        End If
        
        If dHoriz <= Maxdist And dHoriz >= Mindist Then
        Else
            dHoriz = Mindist
            LadderBo.HorizVertOffset.HorizontalOffset = dHoriz
            LadderBo.SetPosition position.x, position.y, position.z
        End If
    End If
      
    If Not oTopCurve Is Nothing And Not BoolArc Then
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oTopDir.Set xEn - xSt, yEn - ySt, 0 ' zEn - zSt
        oTopDir.Length = 1
        Set oTopAway = oTopDir.Cross(BottomNormal)
                
        'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
        If (True = bRotation) Then 'We can do Rotation.
            Set oVecTmp2 = ApplyLadderRotationAboutVertical(dRotation, oTopAway)
            oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2
            oVecTmp2.Length = 1
            oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2
            oTopAway.Set dXtmp2, dYtmp2, dZtmp2
            'Re-obtain the topDir so that it is orthogonal to TopAway vector.
            'Otherwise GTransf can not be obtained for the matrix in SetMatrix.
            oTopAway.Length = 1
            'Get the oTopAway again using the cross so that it is orthogonal.
            Set oTopDir = BottomNormal.Cross(oTopAway)
            oTopDir.Length = 1
        End If '(True = bRotation)
        'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.

        
        Matrix.IndexValue(4) = oTopAway.x * Direc
        Matrix.IndexValue(5) = oTopAway.y * Direc
        Matrix.IndexValue(6) = oTopAway.z * Direc
    ElseIf Not oTopSurf Is Nothing And Not oRefCurve Is Nothing Then
        LadderBo.GetPosition x, y, z
        oTopSurf.Parameter x, y, z, SParamU, SParamV
        oTopSurf.Normal SParamU, SParamV, x, y, z

'************************************************************************************************
        
        Dim x1 As Double, y1 As Double, z1 As Double  'Temp Vars
        If Not proj Is Nothing Then
            If proj.IsOrthogonal Then
'                x = x * -1
'                y = y * -1
'                z = z * -1
            End If
        End If
'************************************************************************************************
        oMountingNormal.Set x, y, z
        Matrix.IndexValue(4) = x
        Matrix.IndexValue(5) = y
        Matrix.IndexValue(6) = z
    ElseIf BoolArc Then  ' for ARC
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        LadderBo.GetPosition x, y, z
        Dim param As Double
        Dim xTan1 As Double, yTan1 As Double, zTan1 As Double, xTan2 As Double, yTan2 As Double, zTan2 As Double
        oTopCurve.Parameter x, y, z, param
        oTopCurve.Evaluate param, x, y, z, xTan1, yTan1, zTan1, xTan2, yTan2, zTan2
        oTopDir.Set xTan1, yTan1, zTan1
        oTopDir.Length = 1
        Set oTopAway = oTopDir.Cross(BottomNormal)
        
        'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
        If (True = bRotation) Then 'We can do Rotation.
            Set oVecTmp2 = ApplyLadderRotationAboutVertical(dRotation, oTopAway)
            oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2
            oVecTmp2.Length = 1
            oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2
            oTopAway.Set dXtmp2, dYtmp2, dZtmp2
        End If '(True = bRotation)
        'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.

        'Re-obtain the topDir so that it is orthogonal to TopAway vector. Otherwise GTransf can not be obtained for the matrix
        ' in SetMatrix
        oTopAway.Length = 1
        'get the oTopAway again using the cross so that it is orhtogonal
        Set oTopDir = BottomNormal.Cross(oTopAway)
        oTopDir.Length = 1
        
        Matrix.IndexValue(4) = oTopAway.x * Direc
        Matrix.IndexValue(5) = oTopAway.y * Direc
        Matrix.IndexValue(6) = oTopAway.z * Direc
    End If
    
    startpt.Set xSt, ySt, zSt
    endpt.Set xEn, yEn, zEn
       
    Dim IntersectPt As DPosition
    Set IntersectPt = New DPosition
    Set IntersectPt = GetIntersectPoint(Top, Ref)
    If oRefCurve Is Nothing And BoolArc Or Not proj Is Nothing Then
        LadderBo.GetPosition x, y, z
        gPosOrig.Set x, y, IntersectPt.z
    Else
        'TR-198568
        gPosOrig.Set IntersectPt.x, IntersectPt.y, IntersectPt.z
        'Set gPosOrig = IntersectPt.Clone 'GetIntersectPoint(Top, Ref) 'IntersectPt
    End If
  
    'Apply the Vertical Offset to the origin.
    gPosOrig.x = gPosOrig.x + ((dVert + TopExt) * BottomNormal.x)
    gPosOrig.y = gPosOrig.y + ((dVert + TopExt) * BottomNormal.y)
    gPosOrig.z = gPosOrig.z + ((dVert + TopExt) * BottomNormal.z)

    'Horizontal Offset Direction
    Dim OffsetDir As DVector
    Set OffsetDir = New DVector
    Dim iDirMod As Integer
    iDirMod = 1
    Dim dot As Double
    
    If Not oTopCurve Is Nothing And Not BoolArc Then
        Dim Direc1 As DVector
        Set Direc1 = New DVector
        Dim Direc2 As DVector
        Set Direc2 = New DVector
        
        Direc1.Set endpt.x - IntersectPt.x, endpt.y - IntersectPt.y, endpt.z - IntersectPt.z
        Direc2.Set startpt.x - IntersectPt.x, startpt.y - IntersectPt.y, startpt.z - IntersectPt.z
        Direc1.Length = 1
        Direc2.Length = 1
        OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
        dot = Direc1.dot(Direc2)
            If (dot < 0.000001) And (dot > -0.0000001) Then
                dot = 0#
            End If

        If dot >= 0# Then
            Set OffsetDir = Direc2
            If IntersectPt.DistPt(startpt) <= TOL Then
                Set OffsetDir = Direc1
            ElseIf IntersectPt.DistPt(endpt) <= TOL Then
                Set OffsetDir = Direc2
            End If
            iDirMod = 1

        Else
            OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
        End If
    
    ElseIf Not oTopSurf Is Nothing And Not BoolArc Then
        Set OffsetDir = oMountingNormal.Cross(BottomNormal)
        OffsetDir.Length = 1
        
        
    ElseIf BoolArc Then

    End If
    
    OffsetDir.Length = 1 'Unit Directional vector
    If Not BoolArc And proj Is Nothing Then
        gPosOrig.x = gPosOrig.x + (dHoriz * OffsetDir.x) * iDirMod
        gPosOrig.y = gPosOrig.y + (dHoriz * OffsetDir.y) * iDirMod
        gPosOrig.z = gPosOrig.z + (dHoriz * OffsetDir.z) * iDirMod
    End If

    'The offset direction is not the same as the matrices x-direction (Matrix 0 1 2).  Incorrect values in here
    'result in the lighting for the object to be incorrect.
 
   If Not oTopCurve Is Nothing Or BoolArc Then
        Matrix.IndexValue(0) = oTopDir.x * -1 * Direc
        Matrix.IndexValue(1) = oTopDir.y * -1 * Direc
        Matrix.IndexValue(2) = oTopDir.z * -1 * Direc
    ElseIf Not oTopSurf Is Nothing Then
        Dim oTemp As DVector
        Set oTemp = New DVector
        Set oTemp = oMountingNormal.Cross(BottomNormal)
        Matrix.IndexValue(0) = oTemp.x
        Matrix.IndexValue(1) = oTemp.y
        Matrix.IndexValue(2) = oTemp.z
    End If
    
    Matrix.IndexValue(12) = gPosOrig.x
    Matrix.IndexValue(13) = gPosOrig.y
    Matrix.IndexValue(14) = gPosOrig.z
    Matrix.IndexValue(3) = 0
    Matrix.IndexValue(7) = 0
    Matrix.IndexValue(11) = 0
    Matrix.IndexValue(15) = 1

    Dim Span As Double
    Dim Height As Double

    LadderBo.SetMatrix Matrix

     'Code for sloped surface as bottom input
    Dim Angle As Double
    Angle = GetAttribute(OccAttrs, "Angle", PartOccInfoCol)

    Dim IdMatrix As DT4x4
    Set IdMatrix = New DT4x4
    IdMatrix.LoadIdentity
    
    Dim Vec1 As DVector
    Dim Vec2 As DVector
    Dim Vec3 As DVector
    Dim Vec4 As DVector
    Set Vec1 = New DVector
    Set Vec2 = New DVector
    Set Vec3 = New DVector
    Set Vec4 = New DVector
    Vec1.Set Matrix.IndexValue(4), Matrix.IndexValue(5), Matrix.IndexValue(6)
    Vec2.Set 0, 0, 1  'gPosOrig.x, gPosOrig.y, gPosOrig.z - 0.1
'    Vec2.Length = 1#
    Set Vec3 = Vec1.Cross(Vec2) ' Normal Vector to V1 and V2
    IdMatrix.Rotate Angle, Vec3
    Set Vec4 = IdMatrix.TransformVector(Vec1)
    Vec4.Get x, y, z
    
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    Dim line As Line3d
    Set line = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, gPosOrig.x, gPosOrig.y, gPosOrig.z, x, y, z, 1)
    line.Infinite = True
    '0, 1 * Cos(Angle), -1 * Sin(Angle)
    'line.SetDirection 0, 0, -1
    Dim temp As IJElements
    Set temp = New JObjectCollection ' elements

    'oBottomSurf should be made infinite so that we calculate the height
    'always even though the top intersect pt is not exactly above the bottom
    oBottomPlane.GetRootPoint x, y, z
    oBottomPlane.GetNormal Rootx, Rooty, Rootz

    Dim DummyFace As New Plane3d
    Dim oNewBottomSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, Rootx, Rooty, Rootz)
    Set oNewBottomSurf = DummyFace

    oNewBottomSurf.Intersect line, temp, code
      
     If temp.Count <> 0 Then
          Dim pt1 As Double, pt2 As Double, pt3 As Double
          Dim point As IJPoint
          Set point = New Point3d
          Set point = temp.Item(1)
          Dim dist As Double
          point.GetPoint pt1, pt2, pt3
          Dim NewPos As DPosition
          Set NewPos = New DPosition
          NewPos.Set pt1, pt2, pt3
         
          dist = NewPos.DistPt(gPosOrig)
     End If
     
    If pt3 >= gPosOrig.z Then
          SetOriginAndSpan = E_FAIL
    End If
    
    Height = dist * Sin(Angle)
    Span = Height - dVert - TopExt '[TR-CP-144430]
    Call SetAttribute(OccAttrs, Height, "Height", PartOccInfoCol)
    Call SetAttribute(OccAttrs, Span, "Span", PartOccInfoCol)

    'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    On Error GoTo ErrorHandler  'Make sure any previous error and On Error Resume Next is reset/cleared.
   
   
   'If we cannot do Rotation, and Rotation is not zero, reset Rotation attribute to zero.
    If (False = bRotation) Then 'If we cannot do Rotation.
        If (dRotation <> 0#) Then 'If Rotation is not zero
            dRotation = 0#  'Reset Rotation attribute to zero.
            'Try to set the Rotation attribute.
            Call SetAttribute(OccAttrs, dRotation, "Rotation", PartOccInfoCol)
        End If ' Rotation is not zero
    End If '(False = bRotation)


    'If we cannot do RotationOffsets, skip remaining offset calculations and set offsets so that we revert to old logic.
    If (False = bRotationOffsets) Then   'If we cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets  'Skip remaining offset calculations and set offsets so that we revert to old logic.
    End If


    'Try to do offset calculations. If we run into trouble skip remaining offset calculations and set offsets so that we revert to old logic.

    'Get oTopDir, oTopAway, gPosOrig from transform matrix of LadderBO with Flip/Direc already applied.
    'Get XrotTopDir vector from transform matrix of LadderBO with Flip/Direc already applied.
    oVecXRotTopDir.x = Matrix.IndexValue(0)
    oVecXRotTopDir.y = Matrix.IndexValue(1)
    oVecXRotTopDir.z = Matrix.IndexValue(2)
    oVecXRotTopDir.Length = 1
    'Get YRotTopAway vector from transform matrix of LadderBO with Flip/Direc already applied.
    oVecYRotTopAway.x = Matrix.IndexValue(4)
    oVecYRotTopAway.y = Matrix.IndexValue(5)
    oVecYRotTopAway.z = Matrix.IndexValue(6)
    oVecYRotTopAway.Length = 1
    'Get ZRotVertUp vector from transform matrix of LadderBO with Flip/Direc already applied.
    oVecZRotVertUp.x = Matrix.IndexValue(8)     'Should be (x=0, y=0, z=1) vertical up vector
    oVecZRotVertUp.y = Matrix.IndexValue(9)
    oVecZRotVertUp.z = Matrix.IndexValue(10)
    oVecZRotVertUp.Length = 1
    'Get OriginPivot position from transform matrix of LadderBO with Flip/Direc already applied.
    oPosOriginPivot.x = Matrix.IndexValue(12)
    oPosOriginPivot.y = Matrix.IndexValue(13)
    oPosOriginPivot.z = Matrix.IndexValue(14)


    'Calc untranslated ladder center, left, right positions based on above attributes, transform vectors, origin position.
    
    'Calc untranslated ladder center origin position as (0,0,0) in local symbol coords and same as oPosOriginPivot in global coords.
    dCenterX = 0    'In local symbol coords
    dCenterY = 0
    dCenterZ = 0
    oPosTemp.Set dCenterX, dCenterY, dCenterZ
    Set oPosCenterOrigin = Matrix.TransformPosition(oPosTemp)
    oPosCenterOrigin.Get dCenterX, dCenterY, dCenterZ
    'Confirm the consistency of the local (0,0,0) origin and the global oPosOriginPivot position.
    oPosOriginPivot.Get dXtmp, dYtmp, dZtmp
    
    
    'Calc left and right plane normal vectors same as +X dir in local symbol coords and same as XRotTopDir in global coords.
    'Calc left plane normal vector direction.
    oVecTemp.Set 1, 0, 0    '+X dir in local symbol coords.
    Set oVecLeftNormal = Matrix.TransformVector(oVecTemp)
    oVecLeftNormal.Get dLeftXnorm, dLeftYnorm, dLeftZnorm  'In global world coords. Same as oVecXRotTopDir from Matrix.
    'Calc right plane normal vector direction. Make it same as left plane normal direction.
    oVecRightNormal.Set dLeftXnorm, dLeftYnorm, dLeftZnorm     'The above test is over, set right plane normal same as left.
    oVecRightNormal.Get dRightXnorm, dRightYnorm, dRightZnorm  'We want normal vectors of both vectors to be same direction.
    'The oVecLeftPlaneNormal will be same as oVecXRotTopDir from LadderBO matrix with Flip/Direc already applied.
    'Confirm the consistency of the local +X vector and the global oVecXRotTopDir vector.
    oVecXRotTopDir.Get dXtmp, dYtmp, dZtmp
    'Confirm the consistency of the local +Y vector and the global oVecYRotTopAway vector.
    oVecYRotTopAway.Get dXtmp1, dYtmp1, dZtmp1
    oVecTemp.Set 0, 1, 0    '+Y dir in local symbol coords.
    Set oVecTmp2 = Matrix.TransformVector(oVecTemp)
    oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2  'In global world coords. Same as oVecYRotTopAway from Matrix.


    'Calc left and right plane root point positions as untranslated ladder left and right support leg positions.
    'Note: Here the Left and Right sides are truly left and right as shown in the graphic view.
    'Similar legacy code calculations elsewhere such as Physical have left and right reversed.

    'Calc necessary adjustment for ladder left and right side positions based on justification.
    dPos = 0#   'Zero adjustment for Center Justify.
    If lJustify = 2 Then            'Left Justify
        dPos = ((dWidth / 2#) + dSideFrameThick)
    'Elseif lJustify = 1 then       'Center Justify
    '   dPos = 0                'for Center Justify
    ElseIf lJustify = 3 Then        'Right Justify
        dPos = -((dWidth / 2#) + dSideFrameThick)
    End If


    'Calc the left plane root point as the untranslated ladder left support leg position.
    dLeftXroot = 0 + (dWidth / 2#) + (dSideFrameThick / 2#) - dPos     'Left side frame position in local symbol coords.
    dLeftXroot = dLeftXroot + dSideFrameThick / 2# + dSupportLegThick / 2#    'Left support leg position in local symbol coords.
    dLeftYroot = 0
    dLeftZroot = 0
    oPosTemp.Set dLeftXroot, dLeftYroot, dLeftZroot
    Set oPosLeftRoot = Matrix.TransformPosition(oPosTemp)
    oPosLeftRoot.Get dLeftXroot, dLeftYroot, dLeftZroot
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'DO NOT DELETE THIS COMMENT BLOCK! IT DOCUMENTS WHAT WE ARE REALLY CALCULATING.
    'dLeftX is not immediately obvious just from inspection of the above calculation.
    'This comment block shows the basic algebra to simplify what dLeftX really is.
    'These are given in local symbol coords and then transformed to global coords.
    '
    'For left justify   dPos = ((dWidth / 2#) + dSideFrameThick)
    'so dLeftX = 0 + w/2 + sft/2 - w/2 - sft
    'or dLeftX = 0 + 0 - (1/2 * sft)
    'or dLeftX = 0 - (dSideFrameThick / 2#)
    'dLeftX + sft/2 + slt / 2 ==> 0 - sft/2 + sft/2 + slt/2
    ' ==> dLeftX = 0 + dSupportLegThick / 2#                                    for left justify
    '
    'For center justify dPos = 0 for center justify
    'so dLeftX = 0 + w/2 + sft/2
    'or dLeftX = 0 + (dWidth / 2#) + (dSideFrameThick / 2#)
    'dLeftX + sft/2 + slt / 2 ==> 0 + w/2 + sft/2 + sft/2 + slt/2
    ' ==> dLeftX = 0 + dWidth / 2# + dSideFrameThick + dSupportLegThick / 2#    for center justify
    '
    'For right justify  dPos = -((dWidth / 2#) + dSideFrameThick)
    'so dLeftX = 0 + w/2 + sft/2 -( -((w/2) + sft) )
    'or dLeftX = 0 + w/2 + sft/2 + w/2 + sft
    'so dLeftX = 0 + dWidth + (3/2 * dSideFrameThick)
    'dLeftX + sft/2 + slt / 2 ==> 0 + w + 3/2 * sft + sft/2 + slt/2
    ' ==> dLeftX = 0 + dWidth + 2# * dSideFrameThick + dSupportLegThick / 2#    for right justify
    '
    'dLeftY = 0 for untranslated ladder
    'dLeftZ = 0 for untranslated ladder
    'DO NOT DELETE THIS COMMENT BLOCK! IT DOCUMENTS WHAT WE ARE REALLY CALCULATING.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  
    'Calc the right plane root point as the untranslated ladder right support leg position.
    dRightXroot = 0 - (dWidth / 2#) - dSideFrameThick / 2# - dPos    'Right side frame position in local symbol coords.
    dRightXroot = dRightXroot - dSideFrameThick / 2# - dSupportLegThick / 2#  'Right support leg position in local symbol coords.
    dRightYroot = 0
    dRightZroot = 0
    oPosTemp.Set dRightXroot, dRightYroot, dRightZroot
    Set oPosRightRoot = Matrix.TransformPosition(oPosTemp)
    oPosRightRoot.Get dRightXroot, dRightYroot, dRightZroot

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'DO NOT DELETE THIS COMMENT BLOCK! IT DOCUMENTS WHAT WE ARE REALLY CALCULATING.
    'dRightX is not immediately obvious just from inspection of the above calculation.
    'This comment block shows the basic algebra to simplify what dRightX really is.
    'These are given in local symbol coords and then transformed to global coords.
    '
    'For left justify   dPos = ((dWidth / 2#) + dSideFrameThick)
    'so dRightX = 0 - w/2 - sft/2 - w/2 - sft
    'or dRightX = 0 - dWidth - (3/2 * dSideFrameThick)
    'dRightX - sft/2 - slt / 2 ==> 0 - w - 3/2*sft - sft/2 - slt/2
    ' ==> dRightX = 0 - dWidth - 2# * dSideFrameThick - dSupportLegThick / 2#   for left justify
    '
    'For center justify dPos = 0
    'so dRightX = 0 - w/2 - sft/2
    'or dRightX = 0 - (dWidth / 2#) - (dSideFrameThick / 2#)
    'dRightX - sft/2 - slt / 2 ==> 0 - w/2 - sft/2 - sft/2 - slt/2
    ' ==> dRightX = 0 - dWidth / 2# - dSideFrameThick - dSupportLegThick / 2#   for center justify
    '
    'For right justify  dPos = -((dWidth / 2#) + dSideFrameThick)
    'so dRightX = 0 - dw/2 - sft/2 + dw/2 + sft
    'or dRightX = 0 - 0 + (1/2 * sft)
    'so dRightX = 0 + dSideFrameThick / 2#
    'dRightX - sft/2 - slt / 2 ==> 0 + sft/2 - sft/2 - slt/2
    ' ==> dRightX = 0 - dSupportLegThick / 2#                                   for right justify
    '
    'dRightY = 0 for untranslated ladder
    'dRightZ = 0 for untranslated ladder
    'DO NOT DELETE THIS COMMENT BLOCK! IT DOCUMENTS WHAT WE ARE REALLY CALCULATING.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    'Create the infinite left and right planes and planar surfaces on each side of untranslated ladder.
    'Create infinite left plane.
    Set oLeftPlane = oGeometryFactory.Planes3d.CreateByPointNormal(Nothing, dLeftXroot, dLeftYroot, dLeftZroot, dLeftXnorm, dLeftYnorm, dLeftZnorm)
    If (oLeftPlane Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Get infinite left planar surface from infinite left plane.
    Set oLeftSurface = oLeftPlane
    If (oLeftSurface Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Create infinite right plane.
    Set oRightPlane = oGeometryFactory.Planes3d.CreateByPointNormal(Nothing, dRightXroot, dRightYroot, dRightZroot, dRightXnorm, dRightYnorm, dRightZnorm)
    If (oRightPlane Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Get infinite right planar surface from infinite right plane.
    Set oRightSurface = oRightPlane
    If (oRightSurface Is Nothing) Then
        GoTo ErrorHandler
    End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DO NOT DELETE THIS COMMENT BLOCK! WE MAY STILL DECIDE WE WANT TO USE THIS.
'   I considered doing the intersection as below, but chose to do it differently instead.
'   For now I will still leave it here, but commented out, in case I change my mind.
'
'    'Try to intersect left plane surface with top curve.
'    'TBD - If multiple intersection points returned, choose closest one to center origin point.
'    On Error Resume Next
'    oLeftSurface.Intersect oTopCurve, oLeftIntersectElements, LeftIntersectCode
'    If ((Err.Number <> 0) Or (oLeftIntersectElements Is Nothing)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    ElseIf ((oLeftIntersectElements.Count <> 1) Or (LeftIntersectCode <> ISECT_UNKNOWN)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    ElseIf Not (TypeOf oLeftIntersectElements.Item(1) Is IJPoint) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    End If
'    On Error GoTo ErrorHandler
'    'TBD check to make sure we get this successfully.
'    Set oLeftIntersectPoint = oLeftIntersectElements.Item(1)  'This should succeed since our previous checks all passed.
'    oLeftIntersectPoint.GetPoint dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ
'    oPosLeftIntersect.Set dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ
'    If (Not oTopCurve.IsPointOn(dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    End If
'
'
'    'Try to intersect right plane surface with top curve to get right intersection point.
'    'TBD - If multiple intersection points returned, choose closest one to center origin point.
'    On Error Resume Next
'    oRightSurface.Intersect oTopCurve, oRightIntersectElements, RightIntersectCode
'    If ((Err.Number <> 0) Or (oRightIntersectElements Is Nothing)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    ElseIf ((oRightIntersectElements.Count <> 1) Or (RightIntersectCode <> ISECT_UNKNOWN)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    ElseIf Not (TypeOf oRightIntersectElements.Item(1) Is IJPoint) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    End If
'    On Error GoTo ErrorHandler
'    'TBD check to make sure we get this successfully.
'    Set oRightIntersectPoint = oRightIntersectElements.Item(1)  'This should succeed since our previous checks all passed.
'    oRightIntersectPoint.GetPoint dRightIntersectX, dRightIntersectY, dRightIntersectZ
'    oPosRightIntersect.Set dRightIntersectX, dRightIntersectY, dRightIntersectZ
'    If (Not oTopCurve.IsPointOn(dRightIntersectX, dRightIntersectY, dRightIntersectZ)) Then
'        bRotationOffsets = False 'We cannot do RotationOffsets.
'        GoTo RevertToOldLogicIfCannotDoRotationOffsets
'    End If
'
'   I considered doing the intersection as above, but chose to do it as below instead.
'   For now I will still leave it here, but commented out, in case I change my mind.
'   DO NOT DELETE THIS COMMENT BLOCK! WE MAY STILL DECIDE WE WANT TO USE THIS.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    'Try to intersect top curve with left plane surface seeking at least one intersection point.
    On Error Resume Next
    oTopCurve.Intersect oLeftSurface, lLeftNumIntersects, dLeftIntersectPoints, lLeftNumOverlaps, LeftIntersectCode
    If ((Err.Number <> 0)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    ElseIf ((lLeftNumIntersects < 1) Or (LeftIntersectCode <> ISECT_UNKNOWN)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    'Check resulting intersection points. Choose left intersect point closest to center origin point.
    bIntersect = False
    lPointJndex = 0
    dMinDistIntSqr = 0#
    For lPointIndex = 1 To lLeftNumIntersects
        dXtmp = dLeftIntersectPoints((lPointIndex - 1) * 3)
        dYtmp = dLeftIntersectPoints((lPointIndex - 1) * 3 + 1)
        dZtmp = dLeftIntersectPoints((lPointIndex - 1) * 3 + 2)
        If oTopCurve.IsPointOn(dXtmp, dYtmp, dZtmp) And oLeftSurface.IsPointOn(dXtmp, dYtmp, dZtmp) Then
        'Confirmed that point is on both the curve and the surface.
            dDistIntSqr = (dCenterX - dXtmp) * (dCenterX - dXtmp) + (dCenterY - dYtmp) * (dCenterY - dYtmp) + (dCenterZ - dZtmp) * (dCenterZ - dZtmp)
            
            If ((False = bIntersect) Or (dDistIntSqr < dMinDistIntSqr)) Then
                bIntersect = True
                lPointJndex = lPointIndex
                dLeftIntersectX = dXtmp
                dLeftIntersectY = dYtmp
                dLeftIntersectZ = dZtmp
                dMinDistIntSqr = dDistIntSqr
            End If
        End If
    Next lPointIndex
    If (False = bIntersect) Then
        Err.Clear
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    oPosLeftIntersect.Set dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ


    'Try to intersect the top curve with the right plane surface seeking at least one intersection point.
    On Error Resume Next
    oTopCurve.Intersect oRightSurface, lRightNumIntersects, dRightIntersectPoints, lRightNumOverlaps, RightIntersectCode
    If ((Err.Number <> 0)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    ElseIf ((lRightNumIntersects < 1) Or (RightIntersectCode <> ISECT_UNKNOWN)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    'Check resulting intersection points. Choose right intersect point closest to center origin point.
    bIntersect = False
    lPointJndex = 0
    dMinDistIntSqr = 0#
    For lPointIndex = 1 To lRightNumIntersects
        dXtmp = dRightIntersectPoints((lPointIndex - 1) * 3)
        dYtmp = dRightIntersectPoints((lPointIndex - 1) * 3 + 1)
        dZtmp = dRightIntersectPoints((lPointIndex - 1) * 3 + 2)
        If oTopCurve.IsPointOn(dXtmp, dYtmp, dZtmp) And oRightSurface.IsPointOn(dXtmp, dYtmp, dZtmp) Then
        'Confirmed that point is on both the curve and the surface.
            dDistIntSqr = (dCenterX - dXtmp) * (dCenterX - dXtmp) + (dCenterY - dYtmp) * (dCenterY - dYtmp) + (dCenterZ - dZtmp) * (dCenterZ - dZtmp)
            
            If ((False = bIntersect) Or (dDistIntSqr < dMinDistIntSqr)) Then
                bIntersect = True
                lPointJndex = lPointIndex
                dRightIntersectX = dXtmp
                dRightIntersectY = dYtmp
                dRightIntersectZ = dZtmp
                dMinDistIntSqr = dDistIntSqr
            End If
        End If
    Next lPointIndex
    If (False = bIntersect) Then
        Err.Clear
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    oPosRightIntersect.Set dRightIntersectX, dRightIntersectY, dRightIntersectZ

    
    'Get the curve parameters to extract the partial top curve (oSubCurve) from the whole top curve (oTopCurve).
    oTopCurve.Parameter dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ, dLeftIntersectPar
    oTopCurve.Parameter dRightIntersectX, dRightIntersectY, dRightIntersectZ, dRightIntersectPar
    If (dLeftIntersectPar < dRightIntersectPar) Then
        dStartPar = dLeftIntersectPar
        dEndPar = dRightIntersectPar
    ElseIf (dLeftIntersectPar > dRightIntersectPar) Then
        dStartPar = dRightIntersectPar
        dEndPar = dLeftIntersectPar
    Else
        Err.Clear
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    dDirecPar = (dStartPar + dEndPar) / 2#
    'Try to extract the partial top curve (oSubCurve) from the whole top curve (oTopCurve).
    'This will create a new curve matching part of original curve between given parameters.
    On Error Resume Next
    Set oTestObject = oGeometryServices.CreateByPartOfCurve(Nothing, oTopCurve, dStartPar, dDirecPar, dEndPar)
    If ((Err.Number <> 0) Or (oTestObject Is Nothing)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    'Put the resulting generic dispatch Object into the IJCurve oSubCurve.
    On Error Resume Next
    Set oSubCurve = oTestObject
    If ((Err.Number <> 0) Or (oSubCurve Is Nothing)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    
    
    'Try to get MinMaxTypedCurveDir for the partial top curve (oSubCurve).
    'This routine finds the minimum and maximum points of a given GTypedCurve along a given vector.
    dDistanceTolerance = GetDistTolerance()
    oVecYRotTopAway.Get dDirX, dDirY, dDirZ
    On Error Resume Next
    oISpsLadderUtil.MinMaxTypedCurveDir oSubCurve, dDistanceTolerance, dDirX, dDirY, dDirZ, dMinX, dMinY, dMinZ, dMaxX, dMaxY, dMaxZ
    If ((Err.Number <> 0)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    
    
    'Get extra offset as dot product of ovecXRotTopAway with vectors from center origin to subcurve min and max.
    'Note: The Zvalues of subcurve min and max points differ from Zvalue of center origin point by TopExtension.
    'Get length of vector from center origin to sub curve min point projected onto oVecYRotTopAway.
    oPosCenterOrigin.Get dCenterX, dCenterY, dCenterZ 'These coordinates should be already defined.
    dFrCtrToMinX = dMinX - dCenterX
    dFrCtrToMinY = dMinY - dCenterY
    dFrCtrToMinZ = dMinZ - dCenterZ
    oVecFrCtrToMin.Set dFrCtrToMinX, dFrCtrToMinY, dFrCtrToMinZ
    dDotProductMin = oVecYRotTopAway.dot(oVecFrCtrToMin)    'oVecYRotTopAway is already a unit vector.
    dProjLengthMin = dDotProductMin     'Dot product gives length of oVecFrCtrToMin projected onto oVecYRotTopAway.
    'Get length of vector from center origin to sub curve max point projected onto oVecYRotTopAway.
    dFrCtrToMaxX = dMaxX - dCenterX
    dFrCtrToMaxY = dMaxY - dCenterY
    dFrCtrToMaxZ = dMaxZ - dCenterZ
    oVecFrCtrToMax.Set dFrCtrToMaxX, dFrCtrToMaxY, dFrCtrToMaxZ
    dDotProductMax = oVecYRotTopAway.dot(oVecFrCtrToMax)    'oVecYRotTopAway is already a unit vector.
    dProjLengthMax = dDotProductMax     'Dot product gives length of oVecFrCtrToMax projected onto oVecYRotTopAway.
    'Choose the larger of the two dot products / projected lengths. The larger should be >= zero, but make sure.
    dExtraOffset = 0#    'The extra offset should never be negative.
    If (dProjLengthMin > dExtraOffset) Then
        dExtraOffset = dProjLengthMin
    End If
    If (dProjLengthMax > dExtraOffset) Then
        dExtraOffset = dProjLengthMax
    End If

    
    'Get total offset vector and use it to translate all ladder positions away from curve.
    dTotalOffset = dWallOffset + dExtraOffset   'Here do not add in SideFrameWidth / 2#
    oVecYRotTopAway.Get dDirX, dDirY, dDirZ
    dTransX = dTotalOffset * dDirX
    dTransY = dTotalOffset * dDirY
    dTransZ = dTotalOffset * dDirZ
    oVecTotalOffset.Set dTransX, dTransY, dTransZ
    'Get translated ladder center origin position.
    dCenterX = dCenterX + dTransX
    dCenterY = dCenterY + dTransY
    dCenterZ = dCenterZ + dTransZ
    oPosCenterOrigin.Set dCenterX, dCenterY, dCenterZ
    'Get translated ladder left support leg position.
    dLeftXroot = dLeftXroot + dTransX
    dLeftYroot = dLeftYroot + dTransY
    dLeftZroot = dLeftZroot + dTransZ
    oPosLeftRoot.Set dLeftXroot, dLeftYroot, dLeftZroot
    'Get translated ladder right support leg position.
    dRightXroot = dRightXroot + dTransX
    dRightYroot = dRightYroot + dTransY
    dRightZroot = dRightZroot + dTransZ
    oPosRightRoot.Set dRightXroot, dRightYroot, dRightZroot


    'Create infinite vertical lines and typed curves through translated ladder left, right support leg positions.
    dDirX = 0#  'Define the components of a vertical up unit vector for the line direction.
    dDirY = 0#
    dDirZ = 1#
    'Create left infinite vertical line through left support leg position.
    Set oLeftLine = oGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, dLeftXroot, dLeftYroot, dLeftZroot, dDirX, dDirY, dDirZ, 1)
    If (oLeftLine Is Nothing) Then
        GoTo ErrorHandler
    End If
    oLeftLine.Infinite = True
    'Put left infinite vertical line into left typed curve.
    Set oLeftCurve = oLeftLine
    If (oLeftCurve Is Nothing) Then
        GoTo ErrorHandler
    End If
    'Create right infinite vertical line through right support leg position.
    Set oRightLine = oGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, dRightXroot, dRightYroot, dRightZroot, dDirX, dDirY, dDirZ, 1)
    If (oRightLine Is Nothing) Then
        GoTo ErrorHandler
    End If
    oRightLine.Infinite = True
    'Put right infinite vertical line into right typed curve.
    Set oRightCurve = oRightLine
    If (oRightCurve Is Nothing) Then
        GoTo ErrorHandler
    End If
    
    
    'Try to get minimum distance between translated ladder left side and left end of sub curve.
    'Try to get MinDistPtTypedCurve between infinite left vertical line/curve and left intersect point.
    'This routine finds the minimum distance between a given GTypedCurve and a given point.
    'Both the minimum distance and the position of closest point on the curve are returned.
    'Necessary because of different Zvalues for left support leg position and left intersect point.
    On Error Resume Next
    oISpsLadderUtil.MinDistPtTypedCurve oLeftCurve, dLeftIntersectX, dLeftIntersectY, dLeftIntersectZ, dLeftMinPosX, dLeftMinPosY, dLeftMinPosZ, dLeftMinDist
    If ((Err.Number <> 0)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    
    
    'Try to get minimum distance between translated ladder right side and right end of sub curve.
    'Try to get MinDistPtTypedCurve between infinite right vertical line/curve and right intersect point.
    'This routine finds the minimum distance between a given GTypedCurve and a given point.
    'Both the minimum distance and the position of closest point on the curve are returned.
    'Necessary because of different Zvalues for right support leg position and right intersect point.
    On Error Resume Next
    oISpsLadderUtil.MinDistPtTypedCurve oRightCurve, dRightIntersectX, dRightIntersectY, dRightIntersectZ, dRightMinPosX, dRightMinPosY, dRightMinPosZ, dRightMinDist
    If ((Err.Number <> 0)) Then
        bRotationOffsets = False 'We cannot do RotationOffsets.
        GoTo RevertToOldLogicIfCannotDoRotationOffsets
    End If
    On Error GoTo ErrorHandler
    
    'Set the calculated left and right support leg offsets based on the left and right minimum distances.
    dLeftSupportLegOffset = dLeftMinDist
    dRightSupportLegOffset = dRightMinDist
    
    'Set the calculated left and right support leg lengths based on the corresponding offsets.
    dLeftSupportLegLength = dLeftSupportLegOffset + dSideFrameWidth / 2#
    dRightSupportLegLength = dRightSupportLegOffset + dSideFrameWidth / 2#
    
    
RevertToOldLogicIfCannotDoRotationOffsets:
    On Error GoTo ErrorHandler  'Make sure any previous error and On Error Resume Next is reset/cleared.
    
    'If we cannot do RotationOffsets, then reset calculated offsets so that we revert to old logic.
    If (False = bRotationOffsets) Then 'If we cannot do RotationOffsets.
        'Retrieve the wall offset again here just to be sure it is correct.
        'Get the WallOffset attribute.
        dWallOffset = GetAttribute(OccAttrs, "WallOffset", PartOccInfoCol)
        'Revert to old logic which based both leg offsets on wall offset with no extra offset.
        dLeftSupportLegOffset = dWallOffset
        dRightSupportLegOffset = dWallOffset
        dExtraOffset = 0#
        dLeftSupportLegLength = dLeftSupportLegOffset + dSideFrameWidth / 2#
        dRightSupportLegLength = dRightSupportLegOffset + dSideFrameWidth / 2#
    End If '(False = bRotationOffsets)


    'Set the attributes for LeftSupportLegLength, RightSupportLegLength, ExtraOffset.
    'Set the LeftSupportLegLength attribute.
    Call SetAttribute(OccAttrs, dLeftSupportLegLength, "LeftSupportLegLength", PartOccInfoCol)
    'Set the RightSupportLegLength attribute.
    Call SetAttribute(OccAttrs, dRightSupportLegLength, "RightSupportLegLength", PartOccInfoCol)
    'Set the ExtraOffset attribute.
    Call SetAttribute(OccAttrs, dExtraOffset, "ExtraOffset", PartOccInfoCol)


    'Reset to nothing these newed and allocated objects and interfaces I added for CR113595.
    Set oVecTmp2 = Nothing
    Set oSPSLadderFactory = Nothing
    Set oISpsLadderUtil = Nothing
    Set oGeometryFactory = Nothing
    Set oGeometryServices = Nothing
    Set oVecXRotTopDir = Nothing
    Set oVecYRotTopAway = Nothing
    Set oVecZRotVertUp = Nothing
    Set oPosOriginPivot = Nothing
    Set oVecLeftNormal = Nothing
    Set oVecRightNormal = Nothing
    Set oVecTemp = Nothing
    Set oPosLeftRoot = Nothing
    Set oPosRightRoot = Nothing
    Set oPosCenterOrigin = Nothing
    Set oPosTemp = Nothing
    Set oLeftPlane = Nothing
    Set oRightPlane = Nothing
    Set oLeftSurface = Nothing
    Set oRightSurface = Nothing
    Set oLeftIntersectElements = Nothing
    Set oLeftIntersectPoint = Nothing
    Set oPosLeftIntersect = Nothing
    Set oRightIntersectElements = Nothing
    Set oRightIntersectPoint = Nothing
    Set oPosRightIntersect = Nothing
    Set oVecFrCtrToMin = Nothing
    Set oVecFrCtrToMax = Nothing
    Set oVecTotalOffset = Nothing
    Set oLeftLine = Nothing
    Set oRightLine = Nothing
    Set oLeftCurve = Nothing
    Set oRightCurve = Nothing
    'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.




    'These newed and allocated objects were already being reset to nothing.
    Set oBottomPlane = Nothing
    Set BottomNormal = Nothing
    Set gPosOrig = Nothing
    Set oMountingWall = Nothing
    Set oTopEdge = Nothing
    Set oRefEdge = Nothing
    Set oRefPlane = Nothing
    Set Matrix = Nothing
    Set startpt = Nothing
    Set endpt = Nothing
    Set oLocalizer = Nothing
    Exit Function
    
RefEdgeMissingHandler:
    Set oLocalizer = Nothing
    SetOriginAndSpan = LADDER_E_SIDE_REF_MISSING
    Err.Raise LADDER_E_SIDE_REF_MISSING
    Exit Function
    
ErrorHandler:
'    Set m_oErrors = New IMSErrorLog.JServerErrors
'    m_oErrors.Add Err.Number, METHOD, Err.Description
    SetOriginAndSpan = E_FAIL
    Err.Description = oLocalizer.GetString(IDS_LADDER_MISSING_INPUTS, "Ladder has missing inputs")
    Err.Raise E_FAIL
    Set oLocalizer = Nothing
End Function 'End SetOriginAndSpan


Public Sub SetAttribute(OccAttrs As IJDAttributes, Value As Variant, strAttribute As String, InfosColl As IJDInfosCol)
On Error GoTo ErrHandler
    
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    
    
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
    If Not InfosColl Is Nothing Then
        iInterfaceCount = InfosColl.Count
        If iInterfaceCount > 0 Then
        For icounter = 1 To iInterfaceCount
            On Error Resume Next
            Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
            If Not AttrInfo Is Nothing And Err.Number = 0 Then
                Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                Set oAttr = Attrcol.Item(strAttribute)
                oAttr.Value = Value
                Exit Sub
            End If
            Set AttrInfo = Nothing
        Next icounter
        End If
    End If
     
    Exit Sub

ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "SetAttribute", Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetAttribute(OccAttrs As IJDAttributes, strAttribute As String, InfosColl As IJDInfosCol)    'this is collection of all interfaces
On Error GoTo ErrHandler

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
     
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
        If Not InfosColl Is Nothing Then
            iInterfaceCount = InfosColl.Count
            If iInterfaceCount > 0 Then
            For icounter = 1 To iInterfaceCount
                On Error Resume Next
                Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
                If Not AttrInfo Is Nothing And Err.Number = 0 Then
                    Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                    Set oAttr = Attrcol.Item(strAttribute)
                    GetAttribute = oAttr.Value
                    Exit Function
                End If
                Set AttrInfo = Nothing
            Next icounter
            End If
        End If
        
    Exit Function
    
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "GetAttribute", Err.Description
    Err.Raise E_FAIL
End Function

Public Function GetAttribute1(OccAttrs As IJDAttributes, strAttribute As String, InfosColl As IJDInfosCol)        'this is collection of all interfaces
On Error GoTo ErrHandler
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
      
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
     
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
        If Not InfosColl Is Nothing Then
            iInterfaceCount = InfosColl.Count
            If iInterfaceCount > 0 Then
            For icounter = 1 To iInterfaceCount
                On Error Resume Next
                Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
                If Not AttrInfo Is Nothing And Err.Number = 0 Then
                    Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                    Set oAttr = Attrcol.Item(strAttribute)
                    GetAttribute1 = oAttr.Value
                    Exit Function
                End If
                Set AttrInfo = Nothing
            Next icounter
            End If
        End If
        
    Exit Function
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "GetAttribute1", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetIntersectPoint(m_pTopEdge As Object, m_pRefEdge As Object) As DPosition
Const METHOD = "GetIntersectPoint"
On Error GoTo ErrorHandler

    Set GetIntersectPoint = New DPosition
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim ClosestParam As Double
    Dim ParmU As Double, ParmV As Double
    Dim oTopface As IJPlane
    Dim oTopSurf As IJSurface
    Dim oTopCurve As IJCurve
    Dim oTopEdge As IJLine
    Dim oRefCurve As IJCurve
    Dim oRefEdge As IJLine
    Dim oRefFace As IJPlane
    Dim oRefSurf As IJSurface
    
    '09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
    'Removed a series of redundantly repeated On Error Resume Next statements from legacy code.
    'Left the first On Error Resume Next in the series in effect to avoid breaking legacy code.
    'Added IsTypeOf checks to try to avoid triggering errors by Set statements expected to fail.
    On Error Resume Next

    If (TypeOf m_pTopEdge Is IJLine) Then
        Set oTopEdge = m_pTopEdge
    End If
    
    If (TypeOf m_pTopEdge Is IJCurve) Then
        Set oTopCurve = m_pTopEdge
    End If
    If oTopCurve Is Nothing Then
        If (TypeOf oTopEdge Is IJCurve) Then
            Set oTopCurve = oTopEdge
        End If
    End If
    
    If (TypeOf m_pTopEdge Is IJPlane) Then
        Set oTopface = m_pTopEdge
    End If

    If (TypeOf m_pTopEdge Is IJSurface) Then
        Set oTopSurf = m_pTopEdge
    End If
    If oTopSurf Is Nothing Then
        If (TypeOf oTopface Is IJSurface) Then
            Set oTopSurf = oTopface
        End If
    End If

    If (TypeOf m_pRefEdge Is IJLine) Then
        Set oRefEdge = m_pRefEdge
    End If

    If (TypeOf m_pRefEdge Is IJCurve) Then
        Set oRefCurve = m_pRefEdge
    End If
    If oRefCurve Is Nothing Then
        If (TypeOf oRefEdge Is IJCurve) Then
            Set oRefCurve = oRefEdge
        End If
    End If
    
    If (TypeOf m_pRefEdge Is IJPlane) Then
        Set oRefFace = m_pRefEdge
    End If

    If (TypeOf m_pRefEdge Is IJSurface) Then
        Set oRefSurf = m_pRefEdge
    End If
    If oRefSurf Is Nothing Then
        If (TypeOf oRefFace Is IJSurface) Then
            Set oRefSurf = oRefFace
        End If
    End If
        
    If Not oTopSurf Is Nothing And Not oRefSurf Is Nothing Then
        GoTo ErrorHandler
    End If
    
    If m_pRefEdge Is Nothing Then
        If Not oTopCurve Is Nothing Then
            oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
            GetIntersectPoint.Set xSt, ySt, zSt
            Exit Function
        End If
    Else
        If oRefSurf Is Nothing And oRefCurve Is Nothing Then   ' i know Top will be curve ' kludge as gridplanes dont support IJsurface
            Set GetIntersectPoint = FindIntersection(oRefFace, oTopCurve)
        Exit Function
        End If
    End If
    
    Dim x1 As Double, y1 As Double, z1 As Double, Mindist As Double
    If Not oRefCurve Is Nothing Or oTopSurf Is Nothing And oRefSurf Is Nothing Then
        oRefCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        If Not oTopCurve Is Nothing Then
            oTopCurve.EndPoints x1, y1, z1, x2, y2, z2  'TR#110689  Null pointer used in GenericSpanAndOrigin.bas, variable oTopCurve
            If Abs(z2 - z1) <= TOL Then
                oTopCurve.DistanceBetween oRefCurve, Mindist, x1, y1, z1, x2, y2, z2
                oTopCurve.Parameter x2, y2, z2, ClosestParam
                oTopCurve.position ClosestParam, x2, y2, z2
            Else
                Dim GeomFactory As New IngrGeom3D.GeometryFactory
                Dim DummyFace As New Plane3d
                Dim oNewBottomSurf As IJSurface
                Dim code As Geom3dIntersectConstants
                Dim temp As IJElements
                Dim dotprod As Double
                Dim NormVec As DVector
                Dim RefVec As DVector
                Dim zVec As DVector
                Set RefVec = New DVector
                Set zVec = New DVector
                Set NormVec = New DVector
                RefVec.Set xEn - xSt, yEn - ySt, zEn - zSt
                RefVec.Length = 1
                zVec.Set 0, 0, 1
                dotprod = RefVec.dot(zVec)
                If (dotprod < 0.000001) And (dotprod > -0.0000001) Then
                    Set NormVec = RefVec.Cross(zVec)
                Else
                    NormVec.Set x2 - x1, y2 - y1, 0
                End If
                Set temp = New JObjectCollection ' elements
                Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, xSt, ySt, zSt, NormVec.x, NormVec.y, NormVec.z)
                Set oNewBottomSurf = DummyFace
                oNewBottomSurf.Intersect oTopCurve, temp, code
                If temp.Count <> 0 Then
                    Dim point As IJPoint
                    Set point = New Point3d
                    Set point = temp.Item(1)
                    point.GetPoint x2, y2, z2
                End If
            End If
        ElseIf Not oTopSurf Is Nothing Then
            oTopSurf.Parameter xEn, yEn, zEn, ParmU, ParmV
            oTopSurf.position ParmU, ParmV, x2, y2, z2
        End If
    ElseIf Not oRefSurf Is Nothing And Not oTopCurve Is Nothing Then 'TR#110689  Null pointer used in GenericSpanAndOrigin.bas, variable oTopCurve
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oRefSurf.Parameter xSt, ySt, zSt, ParmU, ParmV
        oRefSurf.position ParmU, ParmV, x2, y2, z2
    End If
    GetIntersectPoint.Set x2, y2, z2
    
    Exit Function
    
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Function FindIntersection(oPlane As IJPlane, oLine As IJCurve) As DPosition
    On Error GoTo ErrHandler
    Dim v1 As DVector
    Dim v2 As DVector
    Dim StartPoint As DPosition
    Dim EndPoint As DPosition
    Dim PlanePoint As DPosition
    Dim Dot_prod As Double


    Set v1 = New DVector
    Set v2 = New DVector
    Set StartPoint = New DPosition
    Set EndPoint = New DPosition
    Set PlanePoint = New DPosition
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    oLine.EndPoints x, y, z, x1, y1, z1
    StartPoint.Set x, y, z
    EndPoint.Set x1, y1, z1
    oPlane.GetRootPoint x, y, z
    PlanePoint.Set x, y, z
    
    v1.Set EndPoint.x - StartPoint.x, EndPoint.y - StartPoint.y, EndPoint.z - StartPoint.z
    v2.Set PlanePoint.x - StartPoint.x, PlanePoint.y - StartPoint.y, PlanePoint.z - StartPoint.z

    v1.Length = 1
    Dot_prod = v1.dot(v2)

    Dim Intersection As DPosition
    Set Intersection = New DPosition

    Intersection.x = StartPoint.x + (Dot_prod * v1.x)
    Intersection.y = StartPoint.y + (Dot_prod * v1.y)
    Intersection.z = StartPoint.z + (Dot_prod * v1.z)
    Set FindIntersection = Intersection


    Set v1 = Nothing
    Set v2 = Nothing
    Set StartPoint = Nothing
    Set EndPoint = Nothing
    Set PlanePoint = Nothing
    Set Intersection = Nothing

    Exit Function
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "FindIntersection", Err.Description
    Err.Raise E_FAIL
End Function

        
'01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
'Private Function MaxOffset(dist As Double, ByRef Maxdist As Double, ByRef Mindist As Double, _
'                            ByRef position As DPosition, ByVal m_ptop As Object, m_pref As Object)
Private Function MaxOffset(BoolArc As Boolean, dist As Double, ByRef Maxdist As Double, ByRef Mindist As Double, _
                            ByRef position As DPosition, ByVal m_ptop As Object, m_pref As Object)

Const METHOD = "MaxOffset"
On Error GoTo ErrorHandler
    Dim SParamU As Double, SParamV As Double, EParamU As Double, EParamV As Double
    Dim SParam As Double, EParam As Double
    Dim ClosestParam As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim startpt As DPosition
    Set startpt = New DPosition
    Dim endpt As DPosition
    Set endpt = New DPosition
    Dim IntersectPt As DPosition
    Set IntersectPt = New DPosition
    Set IntersectPt = GetIntersectPoint(m_ptop, m_pref)

    Dim oTopEdge As IJLine
    Dim oTopCurve As IJCurve
    Dim oTopface As IJPlane
    Dim oTopSurf As IJSurface
    Dim oProj As IJProjection
    Dim TempArc As IJArc
    Dim Length As Double
    On Error Resume Next
    Set TempArc = m_ptop
    If Not TempArc Is Nothing Then
        Set oTopCurve = TempArc
        Length = oTopCurve.Length
        Mindist = 0#
        Maxdist = Length
        TempArc.GetStartPoint xSt, ySt, zSt
        If dist <> 0# Then
            Set oTopCurve = TempArc
            oTopCurve.ParamRange SParam, EParam
            ClosestParam = ((dist * (EParam - SParam)) / Length) + SParam
            oTopCurve.position ClosestParam, xSt, ySt, zSt
        End If
        position.Set xSt, ySt, zSt
        Exit Function
    End If
    On Error Resume Next
    Set oTopEdge = m_ptop
    On Error Resume Next
    Set oTopCurve = m_ptop
    If oTopCurve Is Nothing Then
        Set oTopCurve = oTopEdge
    End If
    
    On Error Resume Next
    Set oTopface = m_ptop
    On Error Resume Next
    Set oTopSurf = m_ptop
    If oTopSurf Is Nothing Then
        Set oTopSurf = oTopface
    End If
    
    On Error Resume Next
    Set oProj = m_ptop
    If Not oProj Is Nothing Then
        Dim oProjCurve As IJCurve
        Dim IntersectParam As Double
        Dim PtonTopParam  As Double
        Set oProjCurve = oProj.Curve
        oProjCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oProjCurve.ParamRange SParam, EParam
        oProjCurve.Parameter IntersectPt.x, IntersectPt.y, IntersectPt.z, IntersectParam
        Length = oProjCurve.Length
        Mindist = (Abs(IntersectParam - SParam) / (EParam - SParam)) * Length
        Maxdist = (Abs(IntersectParam - EParam) / (EParam - SParam)) * Length
        Dim Tmpdist As Double
        If dist = 0# Then
            oProjCurve.position IntersectParam, xSt, ySt, zSt
        ElseIf dist < 0 Then
            Tmpdist = Mindist - Abs(dist)
            ClosestParam = ((Tmpdist * (EParam - SParam)) / Length) + SParam
            oProjCurve.position ClosestParam, xSt, ySt, zSt
        ElseIf dist > 0 Then
            Tmpdist = Mindist + dist
            ClosestParam = ((Tmpdist * (EParam - SParam)) / Length) + SParam
            oProjCurve.position ClosestParam, xSt, ySt, zSt
        End If
        Mindist = Mindist * -1
        position.Set xSt, ySt, zSt
        Exit Function
    End If
    
    If Not oTopSurf Is Nothing Then
        position.Set 0, 0, 0
        Exit Function
    End If
    
    Dim Direc1 As DVector
    Set Direc1 = New DVector
    Dim Direc2 As DVector
    Set Direc2 = New DVector
    Dim OffsetDir As DVector
    Set OffsetDir = New DVector
    If Not oTopCurve Is Nothing Then
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        position.Set xSt, ySt, zSt
        startpt.Set xSt, ySt, zSt
        endpt.Set xEn, yEn, zEn
        If m_pref Is Nothing Then
            Mindist = 0#
            Maxdist = oTopCurve.Length
        End If

        '01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
        'If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Or oTopCurve.Form = CURVE_FORM_CLOSED_WITH_TANGENT Then
        If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Or oTopCurve.Form = CURVE_FORM_CLOSED_WITH_TANGENT Or BoolArc Then
            oTopCurve.ParamRange SParam, EParam
            Length = oTopCurve.Length
            ClosestParam = ((dist * (EParam - SParam)) / Length) + SParam
            oTopCurve.position ClosestParam, xSt, ySt, zSt
            position.Set xSt, ySt, zSt
            Exit Function
        End If
        Dim dot As Double
        Direc1.Set endpt.x - IntersectPt.x, endpt.y - IntersectPt.y, endpt.z - IntersectPt.z
        Direc2.Set startpt.x - IntersectPt.x, startpt.y - IntersectPt.y, startpt.z - IntersectPt.z
        dot = Direc1.dot(Direc2)
        If dot >= 0# Then
            If IntersectPt.DistPt(startpt) > IntersectPt.DistPt(endpt) Then
                Mindist = IntersectPt.DistPt(endpt)
                Maxdist = IntersectPt.DistPt(startpt)
            Else
                Mindist = IntersectPt.DistPt(startpt)
                Maxdist = IntersectPt.DistPt(endpt)
            End If
        Else
            OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
            OffsetDir.Length = 1 'Unit Directional vector
            Mindist = -1 * (IntersectPt.DistPt(startpt))
            Maxdist = IntersectPt.DistPt(endpt)
        End If
    End If
        
    Exit Function

ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
 
Public Function GetIntersectPointOnSideRef(oLadder As ISPSLadder) As IJDPosition
Const METHOD = "GetIntersectPointOnSideRef"
On Error GoTo ErrorHandler
    Dim oPointOnRefPlane As IJDPosition
    Dim oTopPlane As IJPlane
    Dim oTopSurface As IJSurface
    Dim oMatrix As IJDT4x4
    Dim posX As Double
    Dim posY As Double
    Dim posZ As Double
    Dim tempX As Double
    Dim tempY As Double
    Dim tempZ As Double
    Dim normX As Double
    Dim normY As Double
    Dim normZ As Double
    Dim oGeomFactory As New GeometryFactory
    Dim oLines3D As ILines3d
    Dim oPlanes3D As IPlanes3d
    
    'Initialize vars
    Set oLines3D = oGeomFactory.Lines3d
    Set oPlanes3D = oGeomFactory.Planes3d
    Set oPointOnRefPlane = New DPosition
    
    'Now get the position of the ladder and the matrix
    oLadder.GetPosition posX, posY, posZ
    oLadder.GetMatrix oMatrix
    
    'First lets create a plane from the top ref
    'create a plane with out vector and the ladder position
    Set oTopPlane = oPlanes3D.CreateByPointNormal(Nothing, posX, posY, posZ, oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6))
    Set oTopSurface = oTopPlane
    
    Dim oRef As Object
    Dim oRefPlane As IJPlane
    Dim oRefSurf As IJSurface
    Dim oTopCurve As IJCurve
    
    Set oRef = oLadder.ReferenceEdge
    
    'Now get the intersection point on the sideref plane for different cases
    If oRef Is Nothing Then
        'remember side ref is optional, in this case just take the starting point of the top curve
        Set oTopCurve = oLadder.TopEdge
        
        If Not oTopCurve Is Nothing Then
            Dim xSt As Double
            Dim ySt As Double
            Dim zSt As Double
            Dim xEn As Double
            Dim yEn As Double
            Dim zEn As Double
            oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
            oPointOnRefPlane.Set xSt, ySt, zSt
        End If
    Else
        If TypeOf oRef Is IJLine Then
            Dim oRefLine As IJLine
            Set oRefLine = oRef
            Set oPointOnRefPlane = FindIntersection(oTopPlane, oRefLine)
        Else
            'For this case, we get a curve from the top support and get the min dist/pt from this curve
            If TypeOf oRef Is IJCurve Then
                Dim oRefCurve As IJCurve
                Set oRefCurve = oRef
                
                Dim oTopLine As IJLine
                Dim Mindist As Double
        
                Set oTopLine = oLines3D.CreateByPtVectLength(Nothing, posX, posY, posZ, oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2), 1000#)
                Set oTopCurve = oTopLine
                oTopCurve.DistanceBetween oRefCurve, Mindist, tempX, tempY, tempZ, oPointOnRefPlane.x, oPointOnRefPlane.y, oPointOnRefPlane.z
                
            Else
            'if oRef is any other type, create a surface first which will be sued for intersection
                If TypeOf oRef Is IJPlane Then
                    Dim oTempPlane As IJPlane
                    Set oTempPlane = oRef
                    oTempPlane.GetNormal normX, normY, normZ
                    oTempPlane.GetRootPoint tempX, tempY, tempZ
                    Set oRefPlane = oPlanes3D.CreateByPointNormal(Nothing, tempX, tempY, tempZ, normX, normY, normZ)
                    Set oRefSurf = oRefPlane
                
                ElseIf TypeOf oRef Is IJSurface Then
                    Set oRefSurf = oRef
                End If
                
                If Not oRefSurf Is Nothing Then
                    Dim numPts As Long
                    Dim pnts1() As Double
                    Dim pnts2() As Double
                    Dim pars1() As Double
                    Dim pars2() As Double
                    oRefSurf.DistanceBetween oTopSurface, Mindist, oPointOnRefPlane.x, oPointOnRefPlane.y, oPointOnRefPlane.z, tempX, tempY, tempZ, numPts, pnts1, pnts2, pars1, pars2
                End If
            End If
        End If
    End If

    Set GetIntersectPointOnSideRef = oPointOnRefPlane
    Set oGeomFactory = Nothing
    Exit Function
    
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Function GetDirectionOfTopSupport(oLadder As ISPSLadder) As IJDVector
Const METHOD = "GetDirectionOfTopSupport"
On Error GoTo ErrorHandler
    'Returns direction of the top support.
    ' for line the vector from start to end
    ' for surface, the surface normal at the position of ladder
    ' for curve, the normal at the position of the ladder
    Dim oTop As Object
    Dim oTopCurve As IJCurve
    Dim oTopLine As IJLine
    Dim oTopSurface As IJSurface
    Dim oVecTop As IJDVector
    Dim tempX As Double
    Dim tempY As Double
    Dim tempZ As Double
    
    Set oVecTop = New DVector
    
    Set oTop = oLadder.TopEdge
    
    If TypeOf oTop Is IJLine Then
        Set oTopLine = oTop
        oTopLine.GetDirection tempX, tempY, tempZ
        oVecTop.Set tempX, tempY, tempZ
        
        Dim tempX1 As Double
        Dim tempY1 As Double
        Dim tempZ1 As Double
       oTopLine.GetStartPoint tempX, tempY, tempZ
       oTopLine.GetEndPoint tempX1, tempY1, tempZ1
        oVecTop.Set tempX1 - tempX, tempY1 - tempY, tempZ1 - tempZ
    ElseIf TypeOf oTop Is IJSurface Then
        Dim paramU As Double
        Dim paramV As Double
    
        Set oTopSurface = oTop
        oLadder.GetPosition tempX, tempY, tempZ
        oTopSurface.Parameter tempX, tempY, tempZ, paramU, paramV
        oTopSurface.Normal paramU, paramV, tempX, tempY, tempZ
        oVecTop.Set tempX, tempY, tempZ
    
    ElseIf TypeOf oTop Is IJCurve Then
        Dim param As Double
        Dim xTan1 As Double, yTan1 As Double, zTan1 As Double, xTan2 As Double, yTan2 As Double, zTan2 As Double
                
        Set oTopCurve = oTop
        oLadder.GetPosition tempX, tempY, tempZ
        
        oTopCurve.Parameter tempX, tempY, tempZ, param
        oTopCurve.Evaluate param, tempX, tempY, tempZ, xTan1, yTan1, zTan1, xTan2, yTan2, zTan2
        oVecTop.Set xTan1, yTan1, zTan1
    End If
    
    Set GetDirectionOfTopSupport = oVecTop
    
    Exit Function
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


'Begin 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
Public Function ApplyLadderRotationAboutVertical(dRotation As Double, oVector As DVector) As DVector
Const METHOD = "ApplyLadderRotationAboutVertical"  '09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.
On Error GoTo ErrorHandler

    Dim oRotationMatrix As DT4x4
    Set oRotationMatrix = New DT4x4
    oRotationMatrix.LoadIdentity

    Dim oVecVerticalUp As DVector
    Set oVecVerticalUp = New DVector
    oVecVerticalUp.Set 0, 0, 1

    oRotationMatrix.Rotate dRotation, oVecVerticalUp

    Dim oVecTmp1 As DVector
    Set oVecTmp1 = New DVector

    oVecTmp1.Set oVector.x, oVector.y, oVector.z

    Dim dXtmp1 As Double
    Dim dYtmp1 As Double
    Dim dZtmp1 As Double

    oVecTmp1.Get dXtmp1, dYtmp1, dZtmp1

    Dim oVecTmp2 As DVector
    Set oVecTmp2 = New DVector

    Set oVecTmp2 = oRotationMatrix.TransformVector(oVecTmp1)

    Dim dXtmp2 As Double
    Dim dYtmp2 As Double
    Dim dZtmp2 As Double

    oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2

    oVecTmp2.Length = 1
    oVecTmp2.Get dXtmp2, dYtmp2, dZtmp2

    Set ApplyLadderRotationAboutVertical = oVecTmp2

    Set oVecTmp2 = Nothing
    Set oVecTmp1 = Nothing
    Set oVecVerticalUp = Nothing
    Set oRotationMatrix = Nothing

    Exit Function
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
'End 09/12/2007 C.C.P. CR113595 Ladder on Corrugated Bulkhead.

