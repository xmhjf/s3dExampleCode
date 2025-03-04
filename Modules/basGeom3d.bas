Attribute VB_Name = "basGeom3d"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2007, Intergraph Corporation. All rights reserved.
'
'   basGeom3d.bas
'   Author:
'   Creation Date:
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'6 June 2003 Symbol Team - Added functions RetrieveNozzleParameters, SmallerDim and ReportUnanticipatedError2
'If Symbol uses only basGeom3d.bas file, for error report call function  "ReportUnanticipatedError2"
'If Symbol uses only FullyParametricFun.bas file, for error report call function  "ReportUnanticipatedError3"
'If Symbol uses both basGeom3d.bas and FullyParametricFun.bas files, for error report call function  "ReportUnanticipatedError2"
'12th June 2003 MS  TR 42250: Added the statements "Exit Function" for both RetrieveNozzleParameters and SmallerDim.
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is Added.
'   23.Jul.2003     SymbolTeam(India)       Added PlaceTrapezoidWithPlanes function to create a trapezoid with 6 planes.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   26.Jul.2004     ACM                   DI-61828 changed declaration of pipeport from IJDPipePort to IJCatalogPipePort
'   10.Dec.2004     SymbolTeam(India)      TR 60730: Modified getOrthonormal function to return correct ortho normal, helped solving cone problem.
'   SS Jan/26/2004
'       CR46127: added new methods - PlaceEccentricCone, PlaceSemiEllipsoid, PlaceNnagon, PlaceRectangularTorus
'   10.Oct.2005     KKK,SymbolTeam(India) TR 84724  Need to use a tolerance while comparing double values for equality
'                               added new methods -added new methods - CmpDblEqual,CmpDblGreaterthan,CmpDblGreaterthanOrEqualTo,CmpDblLessThan,CmpDblLessThanOrEqualTo
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007     RRK  CR-123952 Modfied RetrieveParameters, RetrieveParametersForThruBoltedEnds, RetrieveParameterswithPressureRating
'                                 functions to give cptOffset value zero when port termination class is bolted using a new optional
'                                 boolean parameter as argument. Corrected compare double functions return type as boolean.
'   18.Feb.2010     RUK     CR-CP-192433  Update all piping and instrument symbols to support non-circular flange faces
'                           CR-CP-178036  Enhance the piping port graphics symbol to provide more realistic rgraphics
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "Geom3d/Nozzle"
Public Const LINEAR_TOLERANCE = 0.0000001
Public Const ANGULAR_TOLERANCE = 0.00001

Public Function getOrthonormal(xx As Double, yy As Double, zz As Double) As AutoMath.DVector

    Dim vec     As New AutoMath.DVector
    Dim normV   As New AutoMath.DVector
    
    vec.Set xx, yy, zz
    If CmpDblLessThanOrEqualTo(Abs(zz), Abs(xx)) And CmpDblLessThanOrEqualTo(Abs(zz), Abs(yy)) Then
        normV.Set 0, 0, 1
    Else
        If CmpDblLessThanOrEqualTo(Abs(xx), Abs(yy)) Then
            normV.Set 1, 0, 0
        Else
            normV.Set 0, 1, 0
        End If
    End If
    Set getOrthonormal = vec.Cross(normV)
    Set vec = Nothing
    Set normV = Nothing

End Function

'''<{(Complex string curve begin)}>
Public Function PlaceTrCString(ByVal startPosition As AutoMath.DPosition, _
                                        ByVal curves As Collection) As IngrGeom3D.ComplexString3d

''' This function creates transient (non-persistent) complex string curve
''' based on collection of curves (IJElements with StartPoint/EndPoints methods)
''' exapmle of use
'''    Dim lines           As Collection
'''    Dim oLine           As IngrGeom3D.Line3d
'''    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
'''    Dim objCStr         As IngrGeom3D.ComplexString3d
'''
'''    Set lines = New Collection
'''    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, 0.5, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, -0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, -0.5, 0)
'''    lines.Add oLine
'''    stPoint.Set 0, 0, -1
'''    Set objCStr = PlaceTrCString(stPoint, lines)
'''    Set oLine = Nothing
'''    Dim iCount As Integer
'''    For iCount = 1 To lines.Count
'''        lines.Remove 1
'''    Next iCount
'''    Set lines = Nothing
'''    ......     use our complex string (e.g. for projection or revolution)
'''    Set objCStr = Nothing
    
    Const METHOD = "PlaceTrCString:"
    On Error GoTo ErrorHandler
        
    Dim objCString  As IngrGeom3D.ComplexString3d
    Dim curPoint    As New AutoMath.DPosition
    Dim curve       As IngrGeom3D.IJCurve
    Dim trCurve     As IngrGeom3D.IJTransform
    Dim objCurve    As Object
    'Dim Elems       As New DynElements
    Dim vMove       As New AutoMath.DVector
    Dim vCompare    As New AutoMath.DVector
    Dim Elems As IJElements
    Set Elems = New JObjectCollection

    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double
    
    Set curPoint = startPosition
    For Each objCurve In curves
        Set curve = objCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        vCompare.Set x2 - x1, y2 - y1, z2 - z1
        If vCompare.Length < LINEAR_TOLERANCE Then
'            If vCompare.Length = 0 Then
'                MsgBox "Start and end points of a curve are the same"
'            Else
'                MsgBox "Start and end points of a curve are too close"
'            End If
            Exit For
        End If
        vMove.Set curPoint.x - x1, curPoint.y - y1, curPoint.z - z1
        Dim tForm   As New AutoMath.DT4x4
        tForm.Translate vMove
        Set trCurve = objCurve
        trCurve.Transform tForm
        Set tForm = Nothing
        Elems.Add trCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        curPoint.Set x2, y2, z2
    Next objCurve
    Set objCString = geomFactory.ComplexStrings3d.CreateByCurves( _
                                                    Nothing, Elems)
                                                    
    Set PlaceTrCString = objCString
    Set objCString = Nothing
    Set geomFactory = Nothing
    Elems.Clear
    Set Elems = Nothing

    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Complex string curve end)}>
    
'''<{(B-spline curve begin)}>
Public Function PlaceTrBspline(nOrder As Long, ByRef basePoints() As AutoMath.DPosition) As IngrGeom3D.BSplineCurve3d
    
''' This function creates transient (non-persistent) b-spline curve based on array of points
''' Array should consist of at least 4 points
''' Example of call:
''' Dim points(1 to 4)   As new AutoMath.DPosition
''' Dim objBspline  As IngrGeom3D.BSplineCurve3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' points(2).set 1, 1, 0
''' points(3).set 0, 1, 0
''' be carefull as nOrder is not related to array; it's order of polinomonal b-spline curve
''' set objBspline = PlaceTrBspline(4, points)
''' ......... use this object (e.g. to create projection)
''' set objBspline = Nothing


    Const METHOD = "PlaceTrBspline:"
    On Error GoTo ErrorHandler
        
    Dim dblPnts()       As Double
    Dim fKnots()        As Double
    Dim fWeights()      As Double
    Dim iCount          As Integer
    Dim curIndex        As Integer
    Dim nPoints         As Long
    Dim nKnots          As Double
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    
    nPoints = UBound(basePoints) - LBound(basePoints) + 1
    If nPoints < 4 Then
'        MsgBox "Place transient b-spline: at least 4 points needed, cannot create b-spline"
        Exit Function
    End If
    If nOrder < 2 Or nOrder > nPoints Or nOrder > 32 Then
'        MsgBox "Place transient b-spline: Order should be in 2...MAX(nPoints,32) range; set to 4"
        nOrder = 4
    End If
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    ReDim dblPnts(3 * nPoints - 1) As Double
    nKnots = nPoints + nOrder
    ReDim fKnots(nKnots - 1) As Double
    
    For iCount = LBound(basePoints) To UBound(basePoints)
        curIndex = (iCount - LBound(basePoints)) * 3
        dblPnts(curIndex) = basePoints(iCount).x
        dblPnts(curIndex + 1) = basePoints(iCount).y
        dblPnts(curIndex + 2) = basePoints(iCount).z
    Next iCount
    
    ReDim fWeights(nPoints - 1) As Double
    
    For iCount = 0 To nPoints - 1
        fWeights(iCount) = 1
    Next
        
    For iCount = 0 To nOrder - 1
        fKnots(iCount) = 0#
        fKnots(nKnots - iCount - 1) = 1#
    Next
    For iCount = nOrder To nPoints - 1
        fKnots(iCount) = (iCount - nOrder + 1) / (nPoints + 1 - nOrder)
    Next
    
' Create bspline object
    Set oBspline = geomFactory.BSplineCurves3d.CreateByParameters( _
                                                    Nothing, _
                                                    nOrder, nPoints, _
                                                    dblPnts, fWeights, fKnots, False)
    Set PlaceTrBspline = oBspline
    Set oBspline = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(B-spline curve end)}>

'''<{(Cylinder begin)}>

Public Function PlaceCylinder(ByVal objOutputColl As Object, _
                                lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double, _
                                isCapped As Boolean) As Object

''' This function creates persistent projetion of circle
''' based on two points (axis of cylinder) and diameter
''' Example of call:
''' Dim stPoint   As new AutoMath.DPosition
''' Dim enPoint   As new AutoMath.DPosition
''' Dim ldiam     as long
''' Dim objCylinder  As object
''' stPoint.set 0, 0, 0
''' enPoint.set 0, 0, 1
''' lDiam = 1.5
''' set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, lDiam, True)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
''' Set objCylinder = Nothing

    Const METHOD = "PlaceCylinder:"
    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.x - lStartPoint.x, _
                     lEndPoint.y - lStartPoint.y, _
                     lEndPoint.z - lStartPoint.z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
    
' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, isCapped)
    
    Set objCircle = Nothing
    
    Set PlaceCylinder = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Cylinder end)}>

'''<{(Cone begin)}>
Public Function PlaceCone(ByVal objOutputColl As Object, _
                        ByVal centerBase As AutoMath.DPosition, _
                        ByVal centerTop As AutoMath.DPosition, _
                        radiusBase As Double, _
                        radiusTop As Double, Optional isCapped As Boolean = True)
                        
''' This function creates persistent cone
''' based on 2 points (axis of cone) and 2 diameters
''' Example of call:
''' Dim stPoint   As new AutoMath.DPosition
''' Dim enPoint   As new AutoMath.DPosition
''' Dim objCone  As object
''' stPoint.set 0, 0, 0
''' enPoint.set 0, 0, 1
''' set objCone = PlaceCone(m_OutputColl, stPoint, enPoint, 2, 1)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
''' Set objCone = Nothing

' Construction of cone
    Const METHOD = "PlaceCone:"
    On Error GoTo ErrorHandler
        
    Dim startBase   As New AutoMath.DPosition
    Dim vecNorm     As New AutoMath.DVector
    Dim startTop    As New AutoMath.DPosition
    Dim objCone     As IngrGeom3D.Cone3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set vecNorm = getOrthonormal(centerTop.x - centerBase.x, _
                   centerTop.y - centerBase.y, _
                   centerTop.z - centerBase.z)
    vecNorm.Length = radiusBase
    Set startBase = centerBase.Offset(vecNorm)
    vecNorm.Length = radiusTop
    Set startTop = centerTop.Offset(vecNorm)
    Set objCone = geomFactory.Cones3d.CreateBy4Pts(objOutputColl.ResourceManager, _
                                                    centerBase.x, centerBase.y, centerBase.z, _
                                                    centerTop.x, centerTop.y, centerTop.z, _
                                                    startBase.x, startBase.y, startBase.z, _
                                                    startTop.x, startTop.y, startTop.z, _
                                                    isCapped)
    Set PlaceCone = objCone
    Set objCone = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Cone end)}>

'''<{(Revolution begin)}>
Public Function PlaceRevolution(ByVal objOutputColl As Object, _
                        ByVal objCurve As Object, _
                        ByVal axisVector As AutoMath.DVector, _
                        ByVal centerPoint As AutoMath.DPosition, _
                        revAngle As Double, _
                        isCapped As Boolean)

                       
''' This function creates persistent revolution based on curve
''' axis of revolution and angle
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim axis        As new AutoMath.DVector
''' Dim objRevolution  As object
''' centPoint.set 0, 0, 0
''' axis.set 0, 1, 0
''' Dim points(1 to 4)   As new AutoMath.DPosition
''' Dim objBspline  As IngrGeom3D.BSplineCurve3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' points(2).set 1, 1, 0
''' points(3).set 0, 1, 0
''' set objBspline = PlaceTrBspline(points)
''' set objRevolution = PlaceRevolution(m_OutputColl, objBSpline, axis, centPoint, 3.141586/2, True)
''' set objBspline = Nothing
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
''' Set objRevolution = Nothing

' Construction of revolution
    Const METHOD = "PlaceRevolution:"
    On Error GoTo ErrorHandler
        
    Dim objRevolution   As IngrGeom3D.Revolution3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    objCurve, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    revAngle, isCapped)

    
    Set PlaceRevolution = objRevolution
    Set objRevolution = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Revolution end)}>

'''<{(Projection begin)}>
Public Function PlaceProjection(ByVal objOutputColl As Object, _
                        ByVal objCurve As Object, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean)

                       
''' This function creates persistent projection based on curve
''' axis of projection and height
''' Example of call:
''' Dim axis        As new AutoMath.DVector
''' Dim objProjection  As object
''' centPoint.set 0, 0, 0
''' axis.set 0, 0, 1
''' Dim points(1 to 4)   As new AutoMath.DPosition
''' Dim objBspline  As IngrGeom3D.BSplineCurve3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' points(2).set 1, 0, 1
''' points(3).set 0, 0, 1
''' set objBspline = PlaceTrBspline(points)
''' set objRevolution = PlaceProjection(m_OutputColl, objBSpline, axis, 1, True)
''' set objBspline = Nothing
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput),
''' Set objRevolution = Nothing

' Construction of projection
    Const METHOD = "PlaceProjection:"
    On Error GoTo ErrorHandler
        
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objProjection = geomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    objCurve, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)

    
    Set PlaceProjection = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Projection end)}>

'''<{(Torus begin)}>
Public Function PlaceTorus(ByVal objOutputColl As Object, _
                        ByVal centerPoint As AutoMath.DPosition, _
                        ByVal axisVec As AutoMath.DVector, _
                        majorRadius As Double, _
                        minorRadius As Double)
                        
''' This function creates persistent Torus
''' based on center point and 2 diameters
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim axis        As new AutoMath.DVector
''' Dim objTorus    As object
''' centPoint.set 0, 0, 0
''' axis.set 0, 0, 1
''' set objTorus = PlaceTorus(m_OutputColl, centPoint, axis, 1.5, .5)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTorus
''' Set objTorus = Nothing

' Construction of Torus
    Const METHOD = "PlaceTorus:"
    On Error GoTo ErrorHandler
        
    If axisVec.Length < LINEAR_TOLERANCE Then
'        MsgBox "Axis is too short"
        Exit Function
    End If
    
    Dim startBase   As New AutoMath.DPosition
    Dim vecDir      As New AutoMath.DVector
    Dim normV       As New AutoMath.DVector
    Dim startTop    As New AutoMath.DPosition
    Dim objTorus    As IngrGeom3D.Torus3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    
' find orthonormal vector to axisVec
    Set vecDir = getOrthonormal(axisVec.x, axisVec.y, axisVec.z)
    vecDir.Length = 1
    Set objTorus = geomFactory.Tori3d.CreateByAxisMajorMinorRadius( _
                                                    objOutputColl.ResourceManager, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    axisVec.x, axisVec.y, axisVec.z, _
                                                    vecDir.x, vecDir.y, vecDir.z, _
                                                    majorRadius, minorRadius, _
                                                    True)
    Set PlaceTorus = objTorus
    Set objTorus = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Torus end)}>


'''<{(Torus sweep begin)}>
Public Function PlaceTorusSweep(ByVal objOutputColl As Object, _
                        ByVal centerPoint As AutoMath.DPosition, _
                        ByVal axisVec As AutoMath.DVector, _
                        majorRadius As Double, _
                        minorRadius As Double, SweepAngle As Double)
                        
''' This function creates persistent Torus
''' based on center point and 2 diameters
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim axis        As new AutoMath.DVector
''' Dim objTorus    As object
''' centPoint.set 0, 0, 0
''' axis.set 0, 0, 1
''' set objTorus = PlaceTorus(m_OutputColl, centPoint, axis, 1.5, .5)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTorus
''' Set objTorus = Nothing

' Construction of Torus
    Const METHOD = "PlaceTorusSweep:"
    On Error GoTo ErrorHandler
        
    If axisVec.Length < LINEAR_TOLERANCE Then
'        MsgBox "Axis is too short"
        Exit Function
    End If
    
    Dim startBase   As New AutoMath.DPosition
    Dim vecDir      As New AutoMath.DVector
    Dim normV       As New AutoMath.DVector
    Dim startTop    As New AutoMath.DPosition
    Dim objTorus    As IngrGeom3D.Torus3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    
' find orthonormal vector to axisVec
    Set vecDir = getOrthonormal(axisVec.x, axisVec.y, axisVec.z)
    vecDir.Length = 1
    Set objTorus = geomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep( _
                                                    objOutputColl.ResourceManager, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    axisVec.x, axisVec.y, axisVec.z, _
                                                    vecDir.x, vecDir.y, vecDir.z, _
                                                    majorRadius, minorRadius, _
                                                    SweepAngle, True)
    Set PlaceTorusSweep = objTorus
    Set objTorus = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Torus Sweep end)}>
Public Function RetrieveParametersForThruBoltedEnds(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double, ByRef lRaisedFaceOrSocketDiam As Double, _
                                Optional lBoltedPartDimIncludesFlgFaceProj = True)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
                lFlangeThick = oPipePort.FlangeOrHubThickness
                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
                ' Case when part dimension includes face projection and port termination class is 'Bolted'
                If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.TerminationClass = 5 Then
                    lcptoffset = 0
                End If
                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                lRaisedFaceOrSocketDiam = oPipePort.RaisedFaceOrSocketDiameter
                Exit For
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function
'''<{(Sphere begin)}>
Public Function PlaceSphere(ByVal objOutputColl As Object, _
                        ByVal centerPoint As AutoMath.DPosition, _
                        Radius As Double)
                        
''' This function creates persistent Sphere
''' based on center point and radius
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim objSphere   As object
''' CenterPos.set 0, 0, 0
''' set objSphere = PlaceSphere(m_OutputColl, CenterPos, 0.5)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSphere
''' Set objSphere = Nothing

' Construction of Torus
    Const METHOD = "PlaceSphere:"
    On Error GoTo ErrorHandler
        
    Dim objSphere   As IngrGeom3D.Sphere3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set objSphere = geomFactory.Spheres3d.CreateByCenterRadius( _
                                                    objOutputColl.ResourceManager, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    Radius, _
                                                    True)
    Set PlaceSphere = objSphere
    Set objSphere = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Sphere end)}>

'''<{(Box begin)}>
Public Function PlaceBox(ByVal objOutputColl As Object, _
                        point1 As AutoMath.DPosition, _
                        point2 As AutoMath.DPosition) _
                        As IMSSymbolEntities.DSymbol
    
''' this function takes two opposite box corners as input parameters
''' Example of call:
''' Dim pPos1   As new AutoMath.DPosition
''' Dim pPos2   As new AutoMath.DPosition
''' Dim oBox    As object
''' pPos1.set -1, -0.5, -0.1
''' pPos2.set 1, 0.5, 0.1
''' set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
''' Set oBox = Nothing

    Const METHOD = "PlaceBox:"
    On Error GoTo ErrorHandler
    
    ' Get or create the definition
    Dim defColl             As IJDDefinitionCollection
    Dim boxDef              As IJDSymbolDefinition
    Dim definitionParams    As Variant
    Dim oEnv                As IMSSymbolEntities.DSymbol
    Dim newEnumArg          As IJDEnumArgument
    Dim IJEditJDArg         As IJDEditJDArgument
    Dim PC                  As IJDParameterContent
    Dim argument            As IJDArgument
    Dim iCount              As Integer
    Dim oSymbolFactory      As IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set defColl = oSymbolFactory.DefinitionCollection(objOutputColl.ResourceManager)
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)

    Set oEnv = oSymbolFactory.PlaceSymbol(boxDef, objOutputColl.ResourceManager)
    Set oSymbolFactory = Nothing
    Set boxDef = Nothing
    Set defColl = Nothing

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    Dim Coords(1 To 6)  As Double
    
    Coords(1) = point1.x
    Coords(2) = point1.y
    Coords(3) = point1.z
    Coords(4) = point2.x
    Coords(5) = point2.y
    Coords(6) = point2.z

    For iCount = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = Coords(iCount)
        PC.Type = igValue
        PC.UomType = 1
    ' Feed the Argument
        argument.index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next

    oEnv.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oEnv
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    Set PlaceBox = oEnv
    Set oEnv = Nothing
    Set oSymbolFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
    Debug.Assert False

End Function
'''<{(Box end)}>

Public Function CreateNozzle(nozzleIndex As Integer, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                lDir As AutoMath.DVector, _
                                lPos As AutoMath.DPosition) As GSCADNozzleEntities.IJDNozzle

''' This function places Nozzle based on 2 parameters:
''' direction and placePoint
''' 2 first parameters (partInput and objOutPutColl) are from symbol machinery
'''' This is example of nozzles output
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -lFaceToFace / 2 - hcgs, 0, 0
'''    oDir.Set -1, 0, 0
'''
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim NozzleFactory       As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory
    
    Dim objPort As Object
    On Error Resume Next
    Set objPort = CreateObject("SymbolUtilities.SyUtilities")
    Err.Clear
    On Error GoTo ErrHandler
    If Not objPort Is Nothing Then
        Set oNozzle = objPort.CreatePortGraphics(nozzleIndex, partInput, objOutputColl, lDir, lPos)
        Set objPort = Nothing
    Else
        Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, nozzleIndex, _
                                        False, objOutputColl.ResourceManager)
    End If
    Set NozzleFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = oPipePort.FlangeOrHubThickness
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzle = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    Exit Function
    
ErrHandler:
    
 End Function

Public Function RetrieveParameters(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double, Optional lBoltedPartDimIncludesFlgFaceProj = True)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
                lFlangeThick = oPipePort.FlangeOrHubThickness
                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                ' Case when part dimension includes face projection and port termination class is 'Bolted'
                If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.TerminationClass = 5 Then
                                                            
                    lcptoffset = 0
                End If
                Exit For
            End If
    Next pPortIndex

    Set oPipePort = Nothing
    
End Function

Public Function RetrieveParameterswithPressureRating(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double, ByRef lPressureRating As Long, _
                                Optional lBoltedPartDimIncludesFlgFaceProj = True)
    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
                lFlangeThick = oPipePort.FlangeOrHubThickness
                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
                ' Case when part dimension includes face projection and port termination class is 'Bolted'
                If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.TerminationClass = 5 Then
                    lcptoffset = 0
                End If
                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                lPressureRating = oPipePort.PressureRating
                Exit For
            End If
    Next pPortIndex
    Set oPipePort = Nothing
        
End Function

Public Function RetrievePipeOD(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function

'Create Nozzle by defining the length when Length is different than Flange Thickness
Public Function CreateNozzleWithLength _
( _
        nozzleIndex As Integer, _
        ByRef partInput As PartFacelets.IJDPart, _
        ByVal objOutputColl As Object, _
        lDir As AutoMath.DVector, _
        lPos As AutoMath.DPosition, _
        NozzleLength As Double _
) As GSCADNozzleEntities.IJDNozzle

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim NozzleFactory       As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory
    
    Dim objPort As Object
    On Error Resume Next
    Set objPort = CreateObject("SymbolUtilities.SyUtilities")
    Err.Clear
    On Error GoTo ErrHandler
    If Not objPort Is Nothing Then
        Set oNozzle = objPort.CreatePortGraphics(nozzleIndex, partInput, objOutputColl, lDir, lPos, NozzleLength)
        Set objPort = Nothing
    Else
        Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, nozzleIndex, _
                                        False, objOutputColl.ResourceManager)
    End If
    
    Set NozzleFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = NozzleLength
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzleWithLength = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    Exit Function
ErrHandler:
    
 End Function

'''''Function that returns the Outside Body Diameter for Female Reducing Two ports Socket Welded or Threaded component
'''''based on greater Socket Flange Diameter
Public Function BodyOD(ByVal fd As Double, ByVal fd2 As Double) As Double
        If CmpDblGreaterthanOrEqualTo(fd, fd2) Then
            BodyOD = fd
        Else
            BodyOD = fd2
        End If
    
End Function
'''''Function that returns the Insulation Diameter for reducing Socket component
'''''based on greater Socket Flange Diameter
Public Function InsulationDiameter(ByVal fd As Double, ByVal fd2 As Double, ByVal it As Double) As Double
        If CmpDblGreaterthanOrEqualTo(fd, fd2) Then
            InsulationDiameter = fd + it * 2
        Else
            InsulationDiameter = fd2 + it * 2
        End If
    
End Function

'''''Function that returns the Insulation Length for Flanges
'''''based on the greater of the FacetoFace or The Flange Thickness + Insulation Thickness
Public Function InsulationLength(ByVal faceToFace As Double, ByVal FlangeThickness As Double, ByVal InsulationThickness As Double) As Double
        If CmpDblGreaterthanOrEqualTo(faceToFace, FlangeThickness + InsulationThickness) Then
            InsulationLength = faceToFace
        Else
            InsulationLength = FlangeThickness + InsulationThickness
        End If
    
End Function
Public Function CreateNozzleJustaCircle(nozzleIndex As Integer, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                lDir As AutoMath.DVector, _
                                lPos As AutoMath.DPosition) As GSCADNozzleEntities.IJDNozzle

''' This function places Nozzle based on 2 parameters:
''' direction and placePoint
''' 2 first parameters (partInput and objOutPutColl) are from symbol machinery
'''' This is example of nozzles output
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -lFaceToFace / 2 - hcgs, 0, 0
'''    oDir.Set -1, 0, 0
'''
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim NozzleFactory       As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory
    Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, nozzleIndex, _
                                        True, objOutputColl.ResourceManager)
    Set NozzleFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = oPipePort.FlangeOrHubThickness
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzleJustaCircle = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    
 End Function

'''<{(Line begin)}>
Public Function PlaceTrLine(ByRef startPoint As AutoMath.DPosition, _
                            ByRef endPoint As AutoMath.DPosition) _
                            As IngrGeom3D.Line3d
    
''' This function creates transient (non-persistent) Line from 2 points
''' Example of call:
''' Dim points(2)   As new AutoMath.DPosition
''' Dim objLine  As IngrGeom3D.Line3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' set objLine = PlaceTrLine(points(0), points(1))
''' ......... use this object (e.g. to create projection)
''' set objLine = Nothing


    Const METHOD = "PlaceTrLine:"
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the line"
    ' Create Line object
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            startPoint.x, startPoint.y, startPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrLine = oLine
    Set oLine = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Line end)}>


'''<{(Arcbycenternorm begin)}>
Public Function PlaceTrArcByCenterNorm(ByRef startPoint As AutoMath.DPosition, _
                            ByRef endPoint As AutoMath.DPosition, _
                            ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector) _
                            As IngrGeom3D.Arc3d
    
''' This function creates transient (non-persistent) arc by start, end, and center
''' Example of call:
''' Dim points(3)   As new AutoMath.DPosition
''' Dim nvec As New AutoMath.DVector
''' Dim objLine  As IngrGeom3D.Arc3d
''' points(0).set 1, 0, 0
''' points(1).set 0, 1, 0
''' points(2).set 0, 0, 0
''' nvec.set 0, 0, 1
''' set objArc = PlaceTrArcByCenterNorm(points(0), points(1), points(2), nvec)
''' ......... use this object (e.g. to create projection)
''' set objArc = Nothing


    Const METHOD = "PlaceTrArcByCenterNorm:"
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the arc"
    ' Create arc object
    Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                             centerPoint.x, centerPoint.y, centerPoint.z, _
                             normalVector.x, normalVector.y, normalVector.z, _
                             startPoint.x, startPoint.y, startPoint.z, _
                             endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrArcByCenterNorm = oArc
    Set oArc = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Arcbycenternorm end)}>

'''<{(Arcbycenter begin)}>
Public Function PlaceTrArcByCenter(ByRef startPoint As AutoMath.DPosition, _
                            ByRef endPoint As AutoMath.DPosition, _
                            ByRef centerPoint As AutoMath.DPosition) _
                            As IngrGeom3D.Arc3d
    
''' This function creates transient (non-persistent) arc by start, end, and center
''' Example of call:
''' Dim points(3)   As new AutoMath.DPosition
''' Dim objLine  As IngrGeom3D.Arc3d
''' points(0).set 1, 0, 0
''' points(1).set 0, 1, 0
''' points(2).set 0, 0, 0
''' set objArc = PlaceTrArcByCenter(points(0), points(1), points(2))
''' ......... use this object (e.g. to create projection)
''' set objArc = Nothing


    Const METHOD = "PlaceTrArcByCenter:"
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the arc"
    ' Create arc object
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            startPoint.x, startPoint.y, startPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrArcByCenter = oArc
    Set oArc = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Arcbycenter end)}>

'''<{(Arcby3pts begin)}>
Public Function PlaceTrArcBy3Pts(ByRef startPoint As AutoMath.DPosition, _
                            ByRef endPoint As AutoMath.DPosition, _
                            ByRef onarcPoint As AutoMath.DPosition) _
                            As IngrGeom3D.Arc3d
    
''' This function creates transient (non-persistent) arc by start, end, and point on
''' Example of call:
''' Dim points(3)   As new AutoMath.DPosition
''' Dim objArc  As IngrGeom3D.Arc3d
''' points(0).set 1, 0, 0
''' points(1).set 0, 1, 0
''' points(2).set 0, 0, 0
''' set objArc = PlaceTrArcBy3Pts(points(0), points(1), points(2))
''' ......... use this object (e.g. to create projection)
''' set objArc = Nothing


    Const METHOD = "PlaceTrArcBy3Pts:"
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the arc"
    ' Create arc object
    Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                            startPoint.x, startPoint.y, startPoint.z, _
                            onarcPoint.x, onarcPoint.y, onarcPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrArcBy3Pts = oArc
    Set oArc = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Arcby3pts end)}>

'''<{(Simple Complex string curve begin)}>
Public Function PlaceTrCStringNoCheck(ByVal curves As Collection) _
                    As IngrGeom3D.ComplexString3d

''' This function creates transient (non-persistent) complex string curve

    
    Const METHOD = "PlaceTrCStringNoCheck:"
    On Error GoTo ErrorHandler
        
    Dim objCString  As IngrGeom3D.ComplexString3d
    Dim objCurve    As Object
'    Dim Elems As New DynElements
    Dim Elems As IJElements
    Set Elems = New JObjectCollection

    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    For Each objCurve In curves
           Elems.Add objCurve
    Next objCurve
    
    Set objCString = geomFactory.ComplexStrings3d.CreateByCurves( _
                                                    Nothing, Elems)
                                                    
    Set PlaceTrCStringNoCheck = objCString
    
    Set objCString = Nothing
    Set geomFactory = Nothing
    Elems.Clear
    Set Elems = Nothing

    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(simple Complex string curve end)}>

''''Cable Tray Public Functions
''''Retrieve the Cable Tray Port Properties

Public Sub RetrieveCableTrayPortProperties(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByRef dActualWidth As Double, ByRef dActualDepth As Double, _
                                Optional dNominalWidth As Double = 0, Optional dNominalDepth As Double = 0)
    
    Const METHOD = "RetrieveCableTrayPortProperties"
    Dim oTrayPort       As IJCableTrayPort 'catalog port
    Dim oCollection As IJDCollection
    On Error GoTo ErrorLabel
    Set oCollection = partInput.GetNozzles()
    Set oTrayPort = oCollection.Item(index)
    If Not oTrayPort Is Nothing Then
        dActualWidth = oTrayPort.ActualWidth
        dActualDepth = oTrayPort.ActualDepth
        dNominalWidth = oTrayPort.NominalWidth
        dNominalDepth = oTrayPort.NominalDepth
    End If
    Set oTrayPort = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub

''''Create Cable Tray Port
Public Function CreateCableTrayPort(ByRef oPart As PartFacelets.IJDPart, _
    dNozzleIndex As Long, _
    oBasePt As AutoMath.DPosition, _
    oAxis As AutoMath.DVector, oRadial As AutoMath.DVector, _
    ByVal objOutputColl As Object) As GSCADNozzleEntities.IJCableTrayPortOcc
    ' This subroutine creates a Cable Tray Port  and sets it's position and direction
    Dim oNozzleFactory As GSCADNozzleEntities.NozzleFactory
    Set oNozzleFactory = New NozzleFactory
    Dim NullObj As Object
    Dim oDistribPort As IJDistribPort
    Dim oLogicalDistPort As IJLogicalDistPort
    Dim oCableTrayPort As GSCADNozzleEntities.IJCableTrayPortOcc
    
    Const METHOD = "CreateCableTrayPort:"

    On Error GoTo ErrHandler
    Set oCableTrayPort = oNozzleFactory.CreateCableTrayNozzleFromPart(oPart, dNozzleIndex, _
                                                                objOutputColl.ResourceManager)
    Set oLogicalDistPort = oCableTrayPort
    Set oDistribPort = oCableTrayPort
    

    oLogicalDistPort.SetCenterLocation oBasePt
    
    oDistribPort.SetDirectionVector oAxis
    oDistribPort.SetRadialOrient oRadial
    
    Set CreateCableTrayPort = oCableTrayPort
        
    Set oNozzleFactory = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oCableTrayPort = Nothing
    Set oAxis = Nothing
    Set oRadial = Nothing
    Set oBasePt = Nothing
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function

Public Function CreateConduitNozzle(oBasePt As AutoMath.DPosition, oAxis As AutoMath.DVector, ByVal objOutputColl As Object, ByRef oPart As PartFacelets.IJDPart, dNozzleIndex As Long) As GSCADNozzleEntities.IJConduitPortOcc
    ' This subroutine creates a ConduitNozzle  and sets it's position and direction
    Dim oNozzleFactory As GSCADNozzleEntities.NozzleFactory
    Set oNozzleFactory = New NozzleFactory
    Dim NullObj As Object
    Dim oDistribPort As IJDistribPort
    Dim oLogicalDistPort As IJLogicalDistPort
    Dim oConduitNozzle As GSCADNozzleEntities.IJConduitPortOcc
    Const METHOD = "CreateConduitNozzle:"

    On Error GoTo ErrHandler
    
    Set oConduitNozzle = oNozzleFactory.CreateConduitNozzleFromPart(oPart, dNozzleIndex, _
                                                                objOutputColl.ResourceManager)
    Set oLogicalDistPort = oConduitNozzle
    Set oDistribPort = oConduitNozzle
    

    oLogicalDistPort.SetCenterLocation oBasePt
    
    oDistribPort.SetDirectionVector oAxis
    
    Set CreateConduitNozzle = oConduitNozzle
        
    Set oNozzleFactory = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oConduitNozzle = Nothing
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function

    Public Function PlaceFAN(m_oOutputCollection As Object, ByVal oBaseOrigin As AutoMath.DPosition, _
                        ByVal dBladeLength As Double, _
                        Optional dBladeWidth As Double, Optional dBladeThickness As Double, _
                        Optional dNoOfBlades As Double, Optional dRotAboutXaxis As Double, _
                        Optional dRotAboutYaxis As Double, Optional dRotAboutZaxis As Double) As Object
    
    ''EXAMPLE
'    Dim dFanBaseOrigin As New AutoMath.DPosition
'    Dim dBladeLength As Double
'    dBladeLength = 0.6
'    dFanBaseOrigin.Set 0, 0.25, 0.36
'    Set ObjFan = PlaceFANassembly(m_OutputColl, dFanBaseOrigin, dBladeLength, 0.15, 0.05, 4, _
'                                                -3.142 / 2, 0, 0)
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFan
'    Set ObjFan = Nothing
    
    Const METHOD = "PlaceFAN:"
    
    On Error GoTo ErrHandler
    
    If CmpDblLessThanOrEqualTo(dBladeWidth, 0) Then dBladeWidth = 0.02    ''Assumed bladewidth as 20mm
    If CmpDblLessThanOrEqualTo(dBladeThickness, 0) Then dBladeThickness = 0.002 ''Assumed bladethickness as 2mm
    If CmpDblLessThanOrEqualTo(dNoOfBlades, 0) Then dNoOfBlades = 4 '' Assumed Number blades as 4 by default.
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
    Dim oProjection As IngrGeom3D.Projection3d
                                
    ''Create a complex string in the form of fan blades.
    Dim oBladePts() As Double
    ReDim oBladePts(0 To dNoOfBlades * 3 - 1) As Double
    Dim dMinorToMajorRatio As Double
    Dim dEllipseMajorX As Double
    Dim oBladeAng As Double
    Dim iCount As Integer
    Dim PI As Double
    
    PI = 4 * Atn(1)
    oBladeAng = 0
    
    dEllipseMajorX = dBladeLength / 2
    dMinorToMajorRatio = (dBladeWidth / 2) / dEllipseMajorX
    
    ''Calculate Points of blades.
    For iCount = 0 To dNoOfBlades - 1
        ''Mid Point
        oBladePts(3 * iCount) = oBaseOrigin.x - (dBladeLength / 2) * Cos(oBladeAng)
        oBladePts(3 * iCount + 1) = oBaseOrigin.y + (dBladeLength / 2) * Sin(oBladeAng)
        oBladePts(3 * iCount + 2) = oBaseOrigin.z
        oBladeAng = oBladeAng + (2 * PI) / dNoOfBlades
    Next iCount
    
    Dim oCurveCollection As Collection
    Set oCurveCollection = New Collection
    
    oBladeAng = (2 * PI) / dNoOfBlades
    
   ''First Elliptical Arc at 0 deg angle.
   Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    oBladePts(0), oBladePts(1), oBladePts(2), 0, 0, -1, dEllipseMajorX, _
                                    0, 0, dMinorToMajorRatio, PI, PI)
     
   oCurveCollection.Add oEllipticalArc
        
    For iCount = 1 To dNoOfBlades - 1
        Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    oBladePts(3 * iCount), oBladePts(3 * iCount + 1), oBladePts(3 * iCount + 2), _
                    0, 0, -1, -dEllipseMajorX * Cos(oBladeAng), dEllipseMajorX * Sin(oBladeAng), 0, dMinorToMajorRatio, PI, PI)

        oCurveCollection.Add oEllipticalArc

        Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    oBladePts(3 * iCount), oBladePts(3 * iCount + 1), oBladePts(3 * iCount + 2), _
                    0, 0, -1, -dEllipseMajorX * Cos(oBladeAng), dEllipseMajorX * Sin(oBladeAng), 0, dMinorToMajorRatio, 0, PI)

        oCurveCollection.Add oEllipticalArc
        
        oBladeAng = oBladeAng + (2 * PI) / dNoOfBlades
    Next iCount

    ''Last elliptical arc
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                            oBladePts(0), oBladePts(1), oBladePts(2), 0, 0, -1, dEllipseMajorX, _
                            0, 0, dMinorToMajorRatio, 0, PI)
    
    oCurveCollection.Add oEllipticalArc
    
    Dim stPosition As New AutoMath.DPosition

    stPosition.Set oBaseOrigin.x - dBladeLength, oBaseOrigin.y, oBaseOrigin.z
    Set oComplexStr = PlaceTrCString(stPosition, oCurveCollection)
 
    
    For iCount = 1 To oCurveCollection.Count
        oCurveCollection.Remove 1
    Next iCount
    
    Dim oAxisVector As New AutoMath.DVector
    oAxisVector.Set 0, 0, -1
    
    Set oProjection = PlaceProjection(m_oOutputCollection, oComplexStr, oAxisVector, _
                                                                    dBladeThickness, True)
    
 ''This "if" is to check whether transformation is needed or not.
  If Not CmpDblEqual(dRotAboutXaxis, 0) Or Not CmpDblEqual(dRotAboutYaxis, 0) Or _
                                            Not CmpDblEqual(dRotAboutZaxis, 0) Then
    
    ''Transformation starts
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim oMoveMatrix As New AutoMath.DT4x4
    Dim oMoveVector As New AutoMath.DVector
    Dim oDirVector As New AutoMath.DVector
    
''Preparing the Translation Vector to Move the element to origin for performing the transformation.
    oMoveVector.Set -oBaseOrigin.x, -oBaseOrigin.y, -oBaseOrigin.z
    oMoveMatrix.LoadIdentity

''Applying the required translation.
    
    oMoveMatrix.Translate oMoveVector
''Move the element to Origin
    oProjection.Transform oMoveMatrix
 
    ''Loading identity matrix
    oTransformationMat.LoadIdentity
''Apply Rotation
    If Not CmpDblEqual(dRotAboutXaxis, 0) Then
      oDirVector.Set 1, 0, 0
      oTransformationMat.Rotate dRotAboutXaxis, oDirVector
    End If
    If Not CmpDblEqual(dRotAboutYaxis, 0) Then
      oDirVector.Set 0, 1, 0
      oTransformationMat.Rotate dRotAboutYaxis, oDirVector
    End If
    If Not CmpDblEqual(dRotAboutZaxis, 0) Then
      oDirVector.Set 0, 0, 1
      oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    End If
    
''Transform the element at Origin
    oProjection.Transform oTransformationMat

''Move the element back to original position
    oMoveVector.Set oBaseOrigin.x, oBaseOrigin.y, oBaseOrigin.z
    oMoveMatrix.LoadIdentity
    oMoveMatrix.Translate oMoveVector
    oProjection.Transform oMoveMatrix
    
    Set oTransformationMat = Nothing
    Set oMoveMatrix = Nothing
    Set oMoveVector = Nothing
    Set oDirVector = Nothing
  
  End If
      
    Set PlaceFAN = oProjection
    
    ''Removing the Line strings
    Dim objComplexString As IJDObject
    Set objComplexString = oComplexStr
    objComplexString.Remove
    Set objComplexString = Nothing
    
    Set oGeomFactory = Nothing
    Set oComplexStr = Nothing
    Set oEllipticalArc = Nothing
    Set oProjection = Nothing
    Set oAxisVector = Nothing
    Set oCurveCollection = Nothing
    Set stPosition = Nothing
    
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function
Public Function PlaceTrapezoid(objOutputcollection As Object, _
                    oBaseCenterPt As AutoMath.DPosition, dBaseSideLength As Double, _
                    dBaseSideWidth As Double, dTopSideLength As Double, dTopSideWidth As Double, _
                    dHeight As Double, isCapped As Boolean, Optional dRotAboutXaxis As Double, Optional dRotAboutYaxis As Double, _
                    Optional dRotAboutZaxis As Double) As IngrGeom3D.RuledSurface3d
            
''This function Places a Trapezoid with top and bottom surfaces as rectangles.
''1.First we place the Trapezoid assuming it is on the X-Y Plane.
''2.Then we move the Trapezoid to the origin by translation.
''3.Then we apply the required transformation on the element at origin.
''4.Then we translate back the element to the original location.
''Note: Here those optional angles are the angles of the axis of the base of bottom rectangle
''Makes with X,Y,Z angles respectively.Anticlockwise is positive and clockwise is negative.

''EXAMPLE
'    Dim dBaseCenter As New AutoMath.DPosition
'    dim ObjTrapeBody as Object
'    Dim dBLength As Double
'    Dim dBWidth As Double
'    Dim dTLength As Double
'    Dim dTWidth As Double
'    Dim dHeight As Double
'    dBaseCenter.Set 0, 0.25, 0.35
'    dBLength = 0.6
'    dBWidth = 0.4
'    dTLength = 0.4
'    dTWidth = 0.2
'    dHeight = 0.3
'
'    Set ObjTrapeBody = PlaceTrapezoid(m_OutputColl, dBaseCenter, dBLength, dBWidth, _
'                                        dTLength, dTWidth, dHeight, True, 0, 0, 0)
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrapeBody
'    Set ObjTrapeBody = Nothing

Const METHOD = "PlaceTrapezoid:"

Dim oGeomFactory As New IngrGeom3D.GeometryFactory
Dim oRuledSurface As IngrGeom3D.RuledSurface3d
Dim oBaseLinestr As IngrGeom3D.LineString3d
Dim oTopLinestr As IngrGeom3D.LineString3d
Dim dCornerpts(0 To 14) As Double
Dim dNoCorpts As Double

On Error GoTo ErrHandler

dCornerpts(0) = oBaseCenterPt.x - (dBaseSideLength / 2)
dCornerpts(1) = oBaseCenterPt.y - (dBaseSideWidth / 2)
dCornerpts(2) = oBaseCenterPt.z

dCornerpts(3) = oBaseCenterPt.x + (dBaseSideLength / 2)
dCornerpts(4) = oBaseCenterPt.y - (dBaseSideWidth / 2)
dCornerpts(5) = oBaseCenterPt.z

dCornerpts(6) = oBaseCenterPt.x + (dBaseSideLength / 2)
dCornerpts(7) = oBaseCenterPt.y + (dBaseSideWidth / 2)
dCornerpts(8) = oBaseCenterPt.z

dCornerpts(9) = oBaseCenterPt.x - (dBaseSideLength / 2)
dCornerpts(10) = oBaseCenterPt.y + (dBaseSideWidth / 2)
dCornerpts(11) = oBaseCenterPt.z

dCornerpts(12) = dCornerpts(0)
dCornerpts(13) = dCornerpts(1)
dCornerpts(14) = dCornerpts(2)


Set oBaseLinestr = oGeomFactory.LineStrings3d.CreateByPoints(objOutputcollection.ResourceManager, 5, dCornerpts)

''Center of the Top Rectangle differs only in Z-coordinate by just adding height to Z coordinate of base center.
''Remaining coordinates are same.
dCornerpts(0) = oBaseCenterPt.x - (dTopSideLength / 2)
dCornerpts(1) = oBaseCenterPt.y - (dTopSideWidth / 2)
dCornerpts(2) = oBaseCenterPt.z + dHeight

dCornerpts(3) = oBaseCenterPt.x + (dTopSideLength / 2)
dCornerpts(4) = oBaseCenterPt.y - (dTopSideWidth / 2)
dCornerpts(5) = oBaseCenterPt.z + dHeight

dCornerpts(6) = oBaseCenterPt.x + (dTopSideLength / 2)
dCornerpts(7) = oBaseCenterPt.y + (dTopSideWidth / 2)
dCornerpts(8) = oBaseCenterPt.z + dHeight

dCornerpts(9) = oBaseCenterPt.x - (dTopSideLength / 2)
dCornerpts(10) = oBaseCenterPt.y + (dTopSideWidth / 2)
dCornerpts(11) = oBaseCenterPt.z + dHeight

dCornerpts(12) = dCornerpts(0)
dCornerpts(13) = dCornerpts(1)
dCornerpts(14) = dCornerpts(2)

Set oTopLinestr = oGeomFactory.LineStrings3d.CreateByPoints(objOutputcollection.ResourceManager, 5, dCornerpts)

Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(objOutputcollection.ResourceManager, _
                                                oTopLinestr, oBaseLinestr, isCapped)
    
If Not CmpDblEqual(dRotAboutXaxis, 0) Or Not CmpDblEqual(dRotAboutYaxis, 0) Or _
                                         Not CmpDblEqual(dRotAboutZaxis, 0) Then
    
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim oMoveMatrix As New AutoMath.DT4x4
    Dim oMoveVector As New AutoMath.DVector
    Dim oDirVector As New AutoMath.DVector
    
''Preparing the Translation Vector to Move the element to origin for performing the transformation.
    oMoveVector.Set -oBaseCenterPt.x, -oBaseCenterPt.y, -oBaseCenterPt.z
    
''Applying the required translation.

    oMoveMatrix.LoadIdentity
    oMoveMatrix.Translate oMoveVector
    'Move the element to Origin
    oRuledSurface.Transform oMoveMatrix

 ''Loading identity matrix
    oTransformationMat.LoadIdentity

''Applying Rotation
    If Not CmpDblEqual(dRotAboutXaxis, 0) Then
      oDirVector.Set 1, 0, 0
      oTransformationMat.Rotate dRotAboutXaxis, oDirVector
    End If
    If Not CmpDblEqual(dRotAboutYaxis, 0) Then
      oDirVector.Set 0, 1, 0
      oTransformationMat.Rotate dRotAboutYaxis, oDirVector
    End If
    If Not CmpDblEqual(dRotAboutZaxis, 0) Then
      oDirVector.Set 0, 0, 1
      oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    End If
    
''Transform the element at Origin
    oRuledSurface.Transform oTransformationMat

''Move the element back to original position
    oMoveVector.Set oBaseCenterPt.x, oBaseCenterPt.y, oBaseCenterPt.z
    oMoveMatrix.LoadIdentity
    oMoveMatrix.Translate oMoveVector
    oRuledSurface.Transform oMoveMatrix
        
    Set oTransformationMat = Nothing
    Set oDirVector = Nothing
    Set oMoveMatrix = Nothing
    Set oMoveVector = Nothing
    
 End If

Set PlaceTrapezoid = oRuledSurface

''Removing the Line strings
Dim objLineString As IJDObject
Set objLineString = oBaseLinestr
objLineString.Remove
Set objLineString = oTopLinestr
Set oTopLinestr = Nothing
objLineString.Remove

Set oGeomFactory = Nothing
Set oRuledSurface = Nothing
Set oBaseLinestr = Nothing

Exit Function

ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD

End Function

Public Function PlaceTruncatedCylinderByHeights(ByVal objOutputColl As Object, _
                                ByVal lStartPoint As AutoMath.DPosition, _
                                lShortHeight As Double, _
                                lLongHeight As Double, _
                                ByVal lDir As AutoMath.DVector, _
                                lRotation As Double, _
                                lDiameter As Double, _
                                isCapped As Boolean) As Object
                                                                
''' This function creates persistent ruled surface of circle
''' and ellipse to form a truncated cylinder
''' Example of call:
'''    Dim objTruncCylinder  As Object
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim Height1     As Double
'''    Dim Height2     As Double
'''    Dim Dir As New AutoMath.DVector
'''    Dim Rotation As Double
'''    Dim Diam As Double
'''    stPoint.Set 0, 0, 0
'''    Height1 = 0.015
'''    Height2 = 0.025
'''    Dir.Set 1, 1, 0
'''    Rotation = 0
'''    Diam = 0.05
'''    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_OutputColl, stPoint, Height1, Height2, Dir, Rotation, Diam, True)
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTruncCylinder
'''    Set objTruncCylinder = Nothing

    Const METHOD = "PlaceTruncatedCylinderByHeights:"
    On Error GoTo ErrorHandler
    
    Dim objRuledSurface   As IngrGeom3D.RuledSurface3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim slopeAngle As Double
    slopeAngle = Atn((lLongHeight - lShortHeight) / lDiameter)
 
'  Steps used:
'   I. Create Truncated cylinder at Origin(0,0,0) and its length along positive x direction,
'   looking towards negative Z-axis its view is as follows:
'      Y-axis
'      ^
'      |
'      |<-lShortHeight->|
'      |-----------------------|
'      |---> X-axis             |
'      |---------------------------|
'      |<-  lLongHeight   -> |
'
'    a) Circle on the left hand side is at the Origin(0,0,0), and negative X axis is its normal
'    b) Ellipse is on the right hand side, [Cos(slopeAngle), sin(slopeAngle),0] is its normal
'    The Truncated cylinder is created by Ruled Surface using circle and ellipse.
'
'   II. Apply rotation transformations to the object created to account for:
'    a) Given 'lRotation'
'    b) Angle of the cylinder direction vector 'lDir' with respect to positive X-axis
'
'   III. Apply translation to the object from Origin(0,0,0) to its given position (lStartPoint.x,
'        lStartPoint.y,lStartPoint.z) to arrive at final Truncated cylinder object.
'
' Step-I:
'   Construct a circle
    Dim circleCenter    As New AutoMath.DPosition
    Dim circleNormal    As New AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    circleCenter.Set 0, 0, 0
    circleNormal.x = -1
    circleNormal.y = 0
    circleNormal.z = 0
    
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
'   Construct a Ellipse
    Dim MidPtDistance As Double
    Dim ellipseCenterX As Double
    Dim ellipseCenterY As Double
    Dim ellipseCenterZ As Double
    Dim ellipseNormalX As Double
    Dim ellipseNormalY As Double
    Dim ellipseNormalZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim ObjEllipse As IngrGeom3D.Ellipse3d
    
    MidPtDistance = (lShortHeight + lLongHeight) / 2
    ellipseCenterX = circleCenter.x + MidPtDistance
    ellipseCenterY = circleCenter.y
    ellipseCenterZ = circleCenter.z
    
    ellipseNormalX = Cos(slopeAngle)
    ellipseNormalY = Sin(slopeAngle)
    ellipseNormalZ = 0
    
    MajorX = -Tan(slopeAngle) * lDiameter / 2
    MajorY = lDiameter / 2
    MajorZ = 0
    mMRatio = Cos(slopeAngle)
    
    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ellipseCenterX, ellipseCenterY, ellipseCenterZ, _
                                        ellipseNormalX, ellipseNormalY, ellipseNormalZ, _
                                        MajorX, MajorY, MajorZ, mMRatio)

'   Created Ruled surface
    Set objRuledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(objOutputColl.ResourceManager, _
                                                            objCircle, ObjEllipse, isCapped)

' Step-II:
'   Apply rotation transformations to the object created to account for 'lRotation'
    Dim transMatObj     As New AutoMath.DT4x4
    Dim xVec As New AutoMath.DVector
    transMatObj.LoadIdentity
    xVec.Set 1, 0, 0
' Check is made to make sure the angle is more than 0.005 deg.
    If Abs(lRotation) > 0.0001 Then
        transMatObj.Rotate lRotation, xVec
        objRuledSurface.Transform transMatObj
    End If
    
'   Apply rotation transformations to the object created to account for Angle between lDir and positive X axis
    Dim PI As Double
    Dim zVec As New AutoMath.DVector
    Dim xAngle As Double
    Dim normVx As AutoMath.DVector
    PI = 4 * Atn(1)
'   Check if the given vector 'lDir' makes angle approximately Zero or PI angle with positive X-axis
    zVec.Set 0, 0, 1
    xAngle = xVec.Angle(lDir, zVec)
    
'   If xAngle is very small (<= 0.0001), then there is no need for transformation,
'   So consider other cases
    transMatObj.LoadIdentity

    If Abs(xAngle) > 0.0001 Then
        If Abs(xAngle - PI) < 0.0001 Then  'This is PI angle with positive X axis case
            Set normVx = zVec
        Else ' This case would handle all other cases except 0 and 180 deg
            Set normVx = xVec.Cross(lDir)
            xAngle = xVec.Angle(lDir, normVx)
        End If
        transMatObj.Rotate xAngle, normVx
        objRuledSurface.Transform transMatObj
    End If

' Step-III: Apply translation to the object from Origin(0,0,0) to its given position
    Dim transVec As New AutoMath.DVector
    transMatObj.LoadIdentity
    If Abs(lStartPoint.x) > 0 Or Abs(lStartPoint.y) > 0 Or Abs(lStartPoint.z) > 0 Then
        transVec.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
        transMatObj.Translate transVec
        objRuledSurface.Transform transMatObj
    End If
        
    Set objCircle = Nothing
    Set ObjEllipse = Nothing
    
    Set PlaceTruncatedCylinderByHeights = objRuledSurface
    Set objRuledSurface = Nothing
    Set geomFactory = Nothing
    
    Set circleCenter = Nothing
    Set circleNormal = Nothing
    Set transMatObj = Nothing
    Set xVec = Nothing
    Set zVec = Nothing
    Set transVec = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Truncated Cylinder end)}>

Public Function RetrieveNozzleParameters(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef NpdUnit As String)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Const METHOD = "RetrieveNozzleParameters:"
    On Error GoTo ErrHandler
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
'''                lPipeDiam = oPipePort.PipingOutsideDiameter
'''                lFlangeThick = oPipePort.FlangeOrHubThickness
'''                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
'''                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
'''                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                NpdUnit = oPipePort.NPDUnitType
                Exit For
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
    Exit Function

ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function

'''''Function that returns the Outside Body Diameter for Female Reducing Two ports Socket Welded or Threaded component
'''''based on greater Socket Flange Diameter
Public Function SmallerDim(ByVal Dim1 As Double, ByVal Dim2 As Double) As Double
    Const METHOD = "SmallerDim:"
    On Error GoTo ErrHandler
    If CmpDblLessThanOrEqualTo(Dim1, Dim2) Then
        SmallerDim = Dim1
    Else
        SmallerDim = Dim2
    End If
        
   Exit Function
        
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function
'''<{(PlaceTrapezoidWithPlanes begin)}>
Public Function PlaceTrapezoidWithPlanes(ByVal objOutputColl As Object, _
                        ByRef topSurfacePoints() As IJDPosition, _
                        ByRef bottomSurfacePoints() As IJDPosition _
                        ) As Collection
''This function returns six Planes of Trapezoid as "Collection". The order of planes in the
''collection is Top,Bottom and four sides. The top and bottom surface points are taken in
''anti-clockwise direction when viewing from top.
    Const METHOD = "PlaceTrapezoidWithPlanes:"
    On Error GoTo ErrorHandler
    Dim iCount As Integer
    Dim arrayTopPoints(0 To 11) As Double
    Dim arrayBottomPoints(0 To 11) As Double
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim oTmpCollection As New Collection
    
'   These are to obtain the coordinates of various vertices.
    For iCount = 0 To 3
        arrayTopPoints(3 * iCount) = topSurfacePoints(iCount).x
        arrayTopPoints(3 * iCount + 1) = topSurfacePoints(iCount).y
        arrayTopPoints(3 * iCount + 2) = topSurfacePoints(iCount).z
        
        arrayBottomPoints(3 * iCount) = bottomSurfacePoints(iCount).x
        arrayBottomPoints(3 * iCount + 1) = bottomSurfacePoints(iCount).y
        arrayBottomPoints(3 * iCount + 2) = bottomSurfacePoints(iCount).z
    Next iCount
    Set objGeomFactory = New IngrGeom3D.GeometryFactory

    Dim arrayPlanePoints(0 To 11) As Double
    Dim objPlane As IngrGeom3D.Plane3d
    
'   Top plane
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                        4, arrayTopPoints)
    oTmpCollection.Add objPlane
    
'   Bottom plane
'To orient normal of the bottom plane outside the body, the plane points are taken in reverse order.
    Dim botPtCount As Integer
    For iCount = 1 To 4
        botPtCount = 5 - iCount
        arrayPlanePoints(3 * iCount - 3) = arrayBottomPoints(3 * botPtCount - 3)
        arrayPlanePoints(3 * iCount - 2) = arrayBottomPoints(3 * botPtCount - 2)
        arrayPlanePoints(3 * iCount - 1) = arrayBottomPoints(3 * botPtCount - 1)
    Next iCount
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
'   Front plane
    arrayPlanePoints(0) = arrayBottomPoints(0)
    arrayPlanePoints(1) = arrayBottomPoints(1)
    arrayPlanePoints(2) = arrayBottomPoints(2)
    
    arrayPlanePoints(3) = arrayBottomPoints(3)
    arrayPlanePoints(4) = arrayBottomPoints(4)
    arrayPlanePoints(5) = arrayBottomPoints(5)
    
    arrayPlanePoints(6) = arrayTopPoints(3)
    arrayPlanePoints(7) = arrayTopPoints(4)
    arrayPlanePoints(8) = arrayTopPoints(5)
    
    arrayPlanePoints(9) = arrayTopPoints(0)
    arrayPlanePoints(10) = arrayTopPoints(1)
    arrayPlanePoints(11) = arrayTopPoints(2)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Right hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(3)
    arrayPlanePoints(1) = arrayBottomPoints(4)
    arrayPlanePoints(2) = arrayBottomPoints(5)
    
    arrayPlanePoints(3) = arrayBottomPoints(6)
    arrayPlanePoints(4) = arrayBottomPoints(7)
    arrayPlanePoints(5) = arrayBottomPoints(8)
    
    arrayPlanePoints(6) = arrayTopPoints(6)
    arrayPlanePoints(7) = arrayTopPoints(7)
    arrayPlanePoints(8) = arrayTopPoints(8)
    
    arrayPlanePoints(9) = arrayTopPoints(3)
    arrayPlanePoints(10) = arrayTopPoints(4)
    arrayPlanePoints(11) = arrayTopPoints(5)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Rear plane
    arrayPlanePoints(0) = arrayBottomPoints(6)
    arrayPlanePoints(1) = arrayBottomPoints(7)
    arrayPlanePoints(2) = arrayBottomPoints(8)
    
    arrayPlanePoints(3) = arrayBottomPoints(9)
    arrayPlanePoints(4) = arrayBottomPoints(10)
    arrayPlanePoints(5) = arrayBottomPoints(11)
    
    arrayPlanePoints(6) = arrayTopPoints(9)
    arrayPlanePoints(7) = arrayTopPoints(10)
    arrayPlanePoints(8) = arrayTopPoints(11)
    
    arrayPlanePoints(9) = arrayTopPoints(6)
    arrayPlanePoints(10) = arrayTopPoints(7)
    arrayPlanePoints(11) = arrayTopPoints(8)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Left hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(9)
    arrayPlanePoints(1) = arrayBottomPoints(10)
    arrayPlanePoints(2) = arrayBottomPoints(11)
    
    arrayPlanePoints(3) = arrayBottomPoints(0)
    arrayPlanePoints(4) = arrayBottomPoints(1)
    arrayPlanePoints(5) = arrayBottomPoints(2)
    
    arrayPlanePoints(6) = arrayTopPoints(0)
    arrayPlanePoints(7) = arrayTopPoints(1)
    arrayPlanePoints(8) = arrayTopPoints(2)
    
    arrayPlanePoints(9) = arrayTopPoints(9)
    arrayPlanePoints(10) = arrayTopPoints(10)
    arrayPlanePoints(11) = arrayTopPoints(11)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
    Set PlaceTrapezoidWithPlanes = oTmpCollection

    Set oTmpCollection = Nothing
    Set objGeomFactory = Nothing
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
    
End Function
'''<{(PlaceTrapezoidWithPlanes end)}>

'this function places a eccentric cone
Public Function PlaceEccentricCone(oOutputCol As Object, ByVal dHeight As Double, ByVal dBaseRadius As Double, ByVal dTopRadius As Double, ByVal dOffset As Double, ByVal bIsCapped As Boolean) As Object
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    
    Dim oBaseCircle As IJCircle
    Dim oTopCircle As IJCircle
    
    Set oBaseCircle = oGeometryFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0#, 0#, 0#, 1#, 0#, 0#, dBaseRadius)
    Set oTopCircle = oGeometryFactory.Circles3d.CreateByCenterNormalRadius(Nothing, dHeight, -dOffset, 0#, 1#, 0#, 0#, dTopRadius)
    
    Dim oElements As IJElements
    Set oElements = New JObjectCollection
    
    Dim oBase As IJComplexString
    Dim oTop As IJComplexString
    
    oElements.Add oBaseCircle
    Set oBase = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    
    oElements.Clear
    
    oElements.Add oTopCircle
    Set oTop = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    
    Dim oEccentricCone As Object
    Set oEccentricCone = oGeometryFactory.RuledSurfaces3d.CreateByCurves(oOutputCol.ResourceManager, oBase, oTop, bIsCapped)
    
    Set PlaceEccentricCone = oEccentricCone
End Function

'this function places a semi-ellipsoid
Public Function PlaceSemiEllipsoid(oOutputCol As Object, dMajorRadius As Double, dMinorRadius As Double, ByVal bIsCapped As Boolean) As Object
    Dim PI As Double
    PI = 4# * Atn(1)
    
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Dim oQuarterEllipse As Object
    Dim oSemiEllipsoid As IngrGeom3D.Revolution3d
    
    'place a quarter ellipse
    Set oQuarterEllipse = oGeometryFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
        Nothing, _
        0#, 0#, 0#, _
        0#, 0#, -1#, _
        0#, dMajorRadius, 0#, _
        dMinorRadius / dMajorRadius, _
        0#, PI / 2#)
        
    'revolve the quarter ellipse to generate the semi-ellipsoid
    Set oSemiEllipsoid = oGeometryFactory.Revolutions3d.CreateByCurve( _
        oOutputCol.ResourceManager, _
        oQuarterEllipse, _
        1#, 0#, 0#, _
        0#, 0#, 0#, _
        PI * 2#, bIsCapped)

    Set PlaceSemiEllipsoid = oSemiEllipsoid
End Function

'this function places a N-nagon
Public Function PlaceNnagon(oOutputCol As Object, ByVal lNumberOfSides As Long, ByVal dSide As Double, ByVal dLength As Double, ByVal bIsCapped As Boolean) As Object
    Dim PI As Double
    PI = 4# * Atn(1)
    
    Dim dAngle As Double
    dAngle = 2# * PI / CDbl(lNumberOfSides)
        
    Dim dHeight As Double
    dHeight = dSide / (2# * Tan(dAngle / 2#))
        
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    
    Dim oT4x4 As IJDT4x4
    Set oT4x4 = New DT4x4
    
    oT4x4.LoadIdentity
        
    Dim oVec As IJDVector
    Set oVec = New DVector
    
    oVec.Set 1#, 0#, 0#
        
    oT4x4.Rotate dAngle, oVec
    
    Dim oPoint As IJPoint
    Set oPoint = oGeometryFactory.Points3d.CreateByPoint(Nothing, 0#, dHeight, dSide / 2#)
        
    Dim dPoints() As Double
    
    ReDim dPoints(1 To (lNumberOfSides + 1) * 3) As Double
    
    Dim ii As Long
    
    For ii = 0 To lNumberOfSides
        oPoint.GetPoint dPoints(ii * 3 + 1), dPoints(ii * 3 + 2), dPoints(ii * 3 + 3)
        oPoint.Transform oT4x4
    Next ii
    
    Dim oLS3d As IJLineString
    Set oLS3d = oGeometryFactory.LineStrings3d.CreateByPoints(Nothing, lNumberOfSides + 1, dPoints)
        
    Dim oNnagon As IJProjection
    Set oNnagon = oGeometryFactory.Projections3d.CreateByCurve(oOutputCol.ResourceManager, oLS3d, 1#, 0#, 0#, dLength, bIsCapped)
            
    Set PlaceNnagon = oNnagon
End Function

'this function places a rectangular torus
Public Function PlaceRectangularTorus(oOutputCol As Object, ByVal dRadius As Double, ByVal dSweepAngle As Double, dRectangleWidth As Double, ByVal dRectangleHeight As Double, ByVal bIsCapped As Boolean) As Object
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Const lNumberOfSides As Long = 4 'four sides for rectangle
    
    Dim dPoints(1 To (lNumberOfSides + 1) * 3) As Double
    
    dPoints(1) = 0
    dPoints(2) = dRectangleWidth
    dPoints(3) = dRectangleHeight
    
    dPoints(4) = 0
    dPoints(5) = -dRectangleWidth
    dPoints(6) = dRectangleHeight
    
    dPoints(7) = 0
    dPoints(8) = -dRectangleWidth
    dPoints(9) = -dRectangleHeight
    
    dPoints(10) = 0
    dPoints(11) = dRectangleWidth
    dPoints(12) = -dRectangleHeight
    
    dPoints(13) = 0
    dPoints(14) = dRectangleWidth
    dPoints(15) = dRectangleHeight
    
    Dim oLS3d As IJLineString
    Set oLS3d = oGeometryFactory.LineStrings3d.CreateByPoints(Nothing, lNumberOfSides + 1, dPoints)
    
    Dim oRectangularTorus As IngrGeom3D.Revolution3d
    Set oRectangularTorus = oGeometryFactory.Revolutions3d.CreateByCurve(oOutputCol.ResourceManager, oLS3d, 0#, 0#, 1#, 0#, dRadius, 0#, dSweepAngle, bIsCapped)

    Set PlaceRectangularTorus = oRectangularTorus
End Function
Public Function CmpDblEqual(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

    If (LeftVariable >= (RightVariable - Tolerance)) And _
        (LeftVariable <= (RightVariable + Tolerance)) Then
        CmpDblEqual = True
    
    Else
        CmpDblEqual = False
    End If

End Function

Public Function CmpDblGreaterthan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

    If (LeftVariable > (RightVariable - Tolerance)) Then
        CmpDblGreaterthan = True
    
    Else
        CmpDblGreaterthan = False
    End If

End Function
Public Function CmpDblGreaterthanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

    If (LeftVariable >= (RightVariable - Tolerance)) Then
        CmpDblGreaterthanOrEqualTo = True
    
    Else
        CmpDblGreaterthanOrEqualTo = False
    End If

End Function
Public Function CmpDblLessThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

    If (LeftVariable < (RightVariable + Tolerance)) Then
        CmpDblLessThan = True
    
    Else
        CmpDblLessThan = False
    End If

End Function
Public Function CmpDblLessThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

'DblLessThanOrEqualTo = IIf(LeftVariable <= (RightVariable + Tolerance), true, False)
    If (LeftVariable <= (RightVariable + Tolerance)) Then
        CmpDblLessThanOrEqualTo = True
    
    Else
        CmpDblLessThanOrEqualTo = False
    End If

End Function
'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError2(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub


