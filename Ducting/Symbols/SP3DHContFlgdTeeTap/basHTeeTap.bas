Attribute VB_Name = "basHTeeTap"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basHTeeTap.bas
'   Author:          RUK
'   Creation Date:  Thursday, June 21 2007
'   Description:
'   This symbol is prepared for Cotoured Flanged straight and conical tee taps of McGill Air flow corporation as per CR-120452

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Public Const MODULE = "basHTeeTap"
Private PI       As Double

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

'''<{(Cone begin)}>
Public Function PlaceConeTrans(ByVal centerBase As AutoMath.DPosition, _
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
    Const METHOD = "PlaceConeTrans:"
    On Error GoTo ErrorHandler
        
    Dim startBase   As New AutoMath.DPosition
    Dim vecNorm     As New AutoMath.DVector
    Dim startTop    As New AutoMath.DPosition
    Dim objCone     As IngrGeom3D.Cone3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set vecNorm = getOrthonormal(centerTop.X - centerBase.X, _
                   centerTop.Y - centerBase.Y, _
                   centerTop.Z - centerBase.Z)
'    vecNorm.Set 0, 0, 1
    vecNorm.Length = radiusBase
    Set startBase = centerBase.Offset(vecNorm)
    vecNorm.Length = radiusTop
    Set startTop = centerTop.Offset(vecNorm)
    Set objCone = geomFactory.Cones3d.CreateBy4Pts(Nothing, _
                                                    centerBase.X, centerBase.Y, centerBase.Z, _
                                                    centerTop.X, centerTop.Y, centerTop.Z, _
                                                    startBase.X, startBase.Y, startBase.Z, _
                                                    startTop.X, startTop.Y, startTop.Z, _
                                                    isCapped)
    Set PlaceConeTrans = objCone
    Set objCone = Nothing
    Set geomFactory = Nothing
    
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Cone end)}>

Public Function getOrthonormal(xx As Double, yy As Double, zz As Double) As AutoMath.DVector

    Dim vec     As New AutoMath.DVector
    Dim normV   As New AutoMath.DVector
    
    vec.Set xx, yy, zz
    If (Abs(zz) <= Abs(xx)) And (Abs(zz) <= Abs(yy)) Then
        normV.Set 0, 0, 1
    Else
        If (Abs(xx) <= Abs(yy)) Then
            normV.Set 1, 0, 0
        Else
            normV.Set 0, 1, 0
        End If
    End If
    Set getOrthonormal = vec.Cross(normV)
    Set vec = Nothing
    Set normV = Nothing

End Function

Public Function PlaceCylinderTrans(lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double) As Object

''' This function creates persistent projetion of circle
''' based on two points (axis of cylinder) and diameter
''' Example of call:
''' Dim oStPoint   As new AutoMath.DPosition
''' Dim oEnPoint   As new AutoMath.DPosition
''' Dim ldiam     as long
''' Dim objCylinder  As object
''' oStPoint.set 0, 0, 0
''' oEnPoint.set 0, 0, 1
''' lDiam = 1.5
''' set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, lDiam, True)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
''' Set objCylinder = Nothing

    Const METHOD = "PlaceCylinderTrans:"
    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.X, lStartPoint.Y, lStartPoint.Z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.X - lStartPoint.X, _
                     lEndPoint.Y - lStartPoint.Y, _
                     lEndPoint.Z - lStartPoint.Z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        lDiameter / 2)
    
' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        objCircle, _
                                                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                                                        dblCylWidth, False)
    
    Set objCircle = Nothing
    
    Set PlaceCylinderTrans = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function

'''<{(Circle begin)}>
Public Function PlaceTrCircleByCenter(ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector, _
                            ByRef Radius As Double) _
                            As IngrGeom3D.Circle3d

''' This function creates transient (non-persistent) circle
''' Example of call:
''' Dim point   As new AutoMath.DPosition
''' Dim normal  As new AutoMath.DVector
''' Dim objCircle  As IngrGeom3D.circle3d
''' point.set 0, 0, 0
''' normal.set 0, 0, 1
''' set objCircle = PlaceTrCircleByCenter(point, normal, 2 )
''' ......... use this object (e.g. to create projection)
''' set objCircle = Nothing


    Const METHOD = "PlaceTrCircleByCenter:"
'    On Error GoTo ErrorHandler
        
    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the Circle"
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.X, centerPoint.Y, centerPoint.Z, _
                            normalVector.X, normalVector.Y, normalVector.Z, _
                            Radius)
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function
    
'ErrorHandler:
'  ReportUnanticipatedError METHOD

End Function

Public Function CreFlatOvalBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.X - (Width - Depth) / 2, CP.Y, CP.Z + Depth / 2
    Pt(2).Set CP.X + (Width - Depth) / 2, CP.Y, CP.Z + Depth / 2
    Pt(3).Set CP.X + Width / 2, CP.Y, CP.Z
    Pt(4).Set CP.X + (Width - Depth) / 2, CP.Y, CP.Z - Depth / 2
    Pt(5).Set CP.X - (Width - Depth) / 2, CP.Y, CP.Z - Depth / 2
    Pt(6).Set CP.X - Width / 2, CP.Y, CP.Z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOvalBranch = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    
End Function

Public Function CreFlatOval(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.X, CP.Y - (Width - Depth) / 2, CP.Z + Depth / 2
    Pt(2).Set CP.X, CP.Y + (Width - Depth) / 2, Pt(1).Z
    Pt(3).Set CP.X, CP.Y + Width / 2, CP.Z
    Pt(4).Set CP.X, Pt(2).Y, CP.Z - Depth / 2
    Pt(5).Set CP.X, Pt(1).Y, Pt(4).Z
    Pt(6).Set CP.X, CP.Y - Width / 2, CP.Z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 1, 0, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOval = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    
End Function

Public Function CreSMRectBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreSMRectBranch:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.X - HW, CP.Y + HD, CP.Z
    Pt(2).Set CP.X + HW, CP.Y + HD, CP.Z
    Pt(3).Set CP.X + HW, CP.Y - HD, CP.Z
    Pt(4).Set CP.X - HW, CP.Y - HD, CP.Z

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreSMRectBranch = objCStr
    Set oLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    
    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD
   
End Function



