Attribute VB_Name = "basGeom3d"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
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
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Const MODULE = "Geom3d/Nozzle"
Const TOLERANCE = 0.0000001

Private Function getOrthonormal(xx As Double, yy As Double, zz As Double) As AutoMath.DVector

    Dim vec     As New AutoMath.DVector
    Dim normV   As New AutoMath.DVector
    
    vec.Set xx, yy, zz
' find orthonormal vector to vec
    If (zz <= xx) And (zz <= xx) Then
        normV.Set 0, 0, 1
    Else
        If (xx <= yy) Then
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
    Dim Elems       As New DynElements
    Dim vMove       As New AutoMath.DVector
    Dim vCompare    As New AutoMath.DVector
    
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
        If vCompare.Length < TOLERANCE Then
            If vCompare.Length = 0 Then
                MsgBox "Start and end points of a curve are the same"
            Else
                MsgBox "Start and end points of a curve are too close"
            End If
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
    ReportUnanticipatedError MODULE, METHOD

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
        MsgBox "Place transient b-spline: at least 4 points needed, cannot create b-spline"
        Exit Function
    End If
    If nOrder < 2 Or nOrder > nPoints Or nOrder > 32 Then
        MsgBox "Place transient b-spline: Order should be in 2...MAX(nPoints,32) range; set to 4"
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
  ReportUnanticipatedError MODULE, METHOD

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
    ReportUnanticipatedError MODULE, METHOD
        
End Function
'''<{(Cylinder end)}>

'''<{(Cone begin)}>
Public Function PlaceCone(ByVal objOutputColl As Object, _
                        ByVal centerBase As AutoMath.DPosition, _
                        ByVal centerTop As AutoMath.DPosition, _
                        radiusBase As Double, _
                        radiusTop As Double)
                        
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
                                                    True)
    Set PlaceCone = objCone
    Set objCone = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
        
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
    ReportUnanticipatedError MODULE, METHOD
        
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
    ReportUnanticipatedError MODULE, METHOD
        
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
        
    If axisVec.Length < TOLERANCE Then
        MsgBox "Axis is too short"
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
    ReportUnanticipatedError MODULE, METHOD
        
End Function
'''<{(Torus end)}>

'''<{(Sphere begin)}>
Public Function PlaceSphere(ByVal objOutputColl As Object, _
                        ByVal centerPoint As AutoMath.DPosition, _
                        radius As Double)
                        
''' This function creates persistent Sphere
''' based on center point and radius
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim objSphere   As object
''' centPoint.set 0, 0, 0
''' set objSphere = PlaceSphere(m_OutputColl, centPoint, 0.5)
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
                                                    radius, _
                                                    True)
    Set PlaceSphere = objSphere
    Set objSphere = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
        
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
    ReportUnanticipatedError MODULE, METHOD
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
    Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, nozzleIndex, _
                                            False, objOutputColl.ResourceManager)
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
    
 End Function

Public Function RetrieveParameters(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lSptOffset As Double, _
                                ByRef lDepth As Double)

    Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim NozzleFactory   As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle         As GSCADNozzleEntities.IJDNozzle
    Dim gscadElem       As IJDObject
    
    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory
    Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, index, _
                                            False, objOutputColl.ResourceManager)
    Set NozzleFactory = Nothing
    Set oPipePort = oNozzle
    
    lPipeDiam = oPipePort.PipingOutsideDiameter
    lFlangeThick = oPipePort.FlangeOrHubThickness
    lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
    lSptOffset = oPipePort.FlangeProjectionOrSocketOffset
    lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
    
    Set oPipePort = Nothing
' Release objects
    Set gscadElem = oNozzle
    Set oNozzle = Nothing
    gscadElem.Remove
    
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
  ReportUnanticipatedError MODULE, METHOD

End Function
'''<{(Line end)}>

