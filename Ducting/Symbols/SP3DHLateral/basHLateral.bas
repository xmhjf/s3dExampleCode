Attribute VB_Name = "basHLateral"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basHLateral.bas
'   Author:          RRK
'   Creation Date:  Tuesday, June 13 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basHLateral:"
Dim PI As Double
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

Public Function PlaceTrCircleByCenter(ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector, _
                            ByRef Radius As Double) _
                            As IngrGeom3D.Circle3d

    Const METHOD = "PlaceTrCircleByCenter:"
    On Error GoTo ErrorHandler
        
    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            normalVector.x, normalVector.y, normalVector.z, _
                            Radius)
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

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
    PI = 4 * Atn(1)
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x, CP.y - (Width - Depth) / 2, CP.z + Depth / 2
    Pt(2).Set CP.x, CP.y + (Width - Depth) / 2, Pt(1).z
    Pt(3).Set CP.x, CP.y + Width / 2, CP.z
    Pt(4).Set CP.x, Pt(2).y, CP.z - Depth / 2
    Pt(5).Set CP.x, Pt(1).y, Pt(4).z
    Pt(6).Set CP.x, CP.y - Width / 2, CP.z
        
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
    ReportUnanticipatedError2 MODULE, METHOD
    
End Function
Public Function CreFltOvlBranchNormaltoZ(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal Orient As Double) _
                            As IngrGeom3D.ComplexString3d
'This function creates flat oval branch that is normal to Z axis and its depth direction will be along Y-axis
'The center point should be specified by the length of the branch (as Z co-ordinate) e.g:(0,0,Branch length)
    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFltOvlBranchNormaltoZ:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x - (Width - Depth) / 2, CP.y + Depth / 2, CP.z
    Pt(2).Set CP.x + (Width - Depth) / 2, CP.y + Depth / 2, CP.z
    Pt(3).Set CP.x + Width / 2, CP.y, CP.z
    Pt(4).Set CP.x + (Width - Depth) / 2, CP.y - Depth / 2, CP.z
    Pt(5).Set CP.x - (Width - Depth) / 2, CP.y - Depth / 2, CP.z
    Pt(6).Set CP.x - Width / 2, CP.y, CP.z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVectorPlaneofBranch As AutoMath.DVector
    Dim oDirVectorOrient As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVectorPlaneofBranch = New AutoMath.DVector
    Set oDirVectorOrient = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    oDirVectorOrient.Set 0, 1, 0
    
    oTransformationMat.Rotate Orient, oDirVectorOrient
    objCStr.Transform oTransformationMat
    
    Set CreFltOvlBranchNormaltoZ = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
    
End Function

Public Function CreRectBranchNormaltoZ(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                                ByVal Orient As Double) _
                            As IngrGeom3D.ComplexString3d
'This function creates rectangular curve that is normal to Z axis and its depth direction will be along Y-axis
'The center point should be specified by the length of the branch (as Z co-ordinate) e.g:(0,0,Branch length)
    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreRectBranchNormaltoZ:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x - HW, CP.y + HD, CP.z
    Pt(2).Set CP.x + HW, CP.y + HD, CP.z
    Pt(3).Set CP.x + HW, CP.y - HD, CP.z
    Pt(4).Set CP.x - HW, CP.y - HD, CP.z

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)

    Dim oDirVectorPlaneofBranch As AutoMath.DVector
    Dim oDirVectorOrient As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVectorPlaneofBranch = New AutoMath.DVector
    Set oDirVectorOrient = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    
    oDirVectorOrient.Set 0, 1, 0
    
    oTransformationMat.Rotate Orient, oDirVectorOrient
    objCStr.Transform oTransformationMat
    
    Set CreRectBranchNormaltoZ = objCStr
    Set oLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
   
End Function
