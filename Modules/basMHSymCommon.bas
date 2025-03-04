Attribute VB_Name = "basMHSymCommon"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   basMHSymCommon.bas
'   Author:         Neelima BhupatiRaju
'   Creation Date:  09 -Nov-2010
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   09.Nov.2010    Neelima B    Initial creation
'   24-Sep-2011    Neelima B    Added a function to create cylinder
'   15-Oct-2012    Shireesha M  Added two new functions and some golbal declarations.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit


Private Const MODULE = "basMHSymCommon:" 'Used for error messages

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type


Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
Const E_FAIL = -2147467259
Err.Raise E_FAIL
End Sub

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
                                                        dblCylWidth, True)
    
    Set objCircle = Nothing
    
    Set PlaceCylinder = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD
        
End Function
'''<{(Cylinder end)}>

'''<{(PlaceTrapezoidWithTriangles begin)}>
Public Function PlaceTrapezoidWithTriangles(ByVal objOutputColl As Object, _
                        ByRef topSurfacePoints() As IJDPosition, _
                        ByRef bottomSurfacePoints() As IJDPosition _
                        ) As Collection
''This function returns six Planes of Trapezoid as "Collection". The order of planes in the
''collection is Top,Bottom and four sides. The top and bottom surface points are taken in
''anti-clockwise direction when viewing from top.
    Const METHOD = "PlaceTrapezoidWithTriangles:"
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
    
    CreatePlanesCollection arrayBottomPoints(0), arrayBottomPoints(1), arrayBottomPoints(2), arrayBottomPoints(3), arrayBottomPoints(4), arrayBottomPoints(5), arrayTopPoints(3), arrayTopPoints(4), arrayTopPoints(5), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(0), arrayBottomPoints(1), arrayBottomPoints(2), arrayTopPoints(0), arrayTopPoints(1), arrayTopPoints(2), arrayTopPoints(3), arrayTopPoints(4), arrayTopPoints(5), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(3), arrayBottomPoints(4), arrayBottomPoints(5), arrayBottomPoints(6), arrayBottomPoints(7), arrayBottomPoints(8), arrayTopPoints(6), arrayTopPoints(7), arrayTopPoints(8), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(3), arrayBottomPoints(4), arrayBottomPoints(5), arrayTopPoints(3), arrayTopPoints(4), arrayTopPoints(5), arrayTopPoints(6), arrayTopPoints(7), arrayTopPoints(8), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(6), arrayBottomPoints(7), arrayBottomPoints(8), arrayBottomPoints(9), arrayBottomPoints(10), arrayBottomPoints(11), arrayTopPoints(9), arrayTopPoints(10), arrayTopPoints(11), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(6), arrayBottomPoints(7), arrayBottomPoints(8), arrayTopPoints(6), arrayTopPoints(7), arrayTopPoints(8), arrayTopPoints(9), arrayTopPoints(10), arrayTopPoints(11), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(9), arrayBottomPoints(10), arrayBottomPoints(11), arrayBottomPoints(0), arrayBottomPoints(1), arrayBottomPoints(2), arrayTopPoints(0), arrayTopPoints(1), arrayTopPoints(2), oTmpCollection, objOutputColl
    CreatePlanesCollection arrayBottomPoints(9), arrayBottomPoints(10), arrayBottomPoints(11), arrayTopPoints(9), arrayTopPoints(10), arrayTopPoints(11), arrayTopPoints(0), arrayTopPoints(1), arrayTopPoints(2), oTmpCollection, objOutputColl

    Set PlaceTrapezoidWithTriangles = oTmpCollection

    Set oTmpCollection = Nothing
    Set objGeomFactory = Nothing
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    
End Function
'''<{(PlaceTrapezoidWithTriangles end)}>

''''''''''''''''''''''''''''''''''

Private Function CreatePlanesCollection(ByVal firstpos, ByVal secpos, ByVal thirdpos, ByVal forpos, ByVal fifpos, ByVal sixpos, ByVal sevpos, ByVal eigpos, ByVal ninepos, ByRef TempCollection, ByVal objOutputColl)
    Dim arrayPlanePoints(0 To 11) As Double
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    
    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    
    arrayPlanePoints(0) = firstpos
    arrayPlanePoints(1) = secpos
    arrayPlanePoints(2) = thirdpos
    
    arrayPlanePoints(3) = forpos
    arrayPlanePoints(4) = fifpos
    arrayPlanePoints(5) = sixpos
    
    arrayPlanePoints(6) = sevpos
    arrayPlanePoints(7) = eigpos
    arrayPlanePoints(8) = ninepos
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    3, arrayPlanePoints)
    TempCollection.Add objPlane
    Set objPlane = Nothing
    Set objGeomFactory = Nothing
End Function

