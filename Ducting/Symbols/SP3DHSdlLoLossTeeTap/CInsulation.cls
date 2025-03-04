VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   ProgID:         SP3DHSdlLoLossTeeTap.HSdlLLTeeTap
'   Author:         RRK
'   Creation Date:  Tuesday, June 13 2007
'   Description:
'   This symbol is prepared for saddle Lo Loss tap of McGill Air flow corporation as per CR-120452
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
      PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim PortDirection As New AutoMath.DVector
    Dim RadialDirection As New AutoMath.DVector
    Dim iOutput     As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHVACShape As Long
    Dim parInsulationThickness As Double
    
    Dim Inch As Double
    Inch = 0.0254

' Inputs

    Set oPartFclt = arrayOfInputs(1)
    
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHVACShape = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)

 'In case of round shape (where Depth is optional) making Depth equal to Width
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
'In case of round shape (where BDepth is optional) making BDepth equal to BWidth
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
    

    Dim dBrBottomWidth As Double
    Dim dBrBottomDepth As Double
    Dim dTapHeight As Double
    

'Assigning the values of Tap Height based on Branch Width(as below) as given in McGill catalog.

'    parBWidth             dTapHeight
'    (Inches)                (Inches)
'    3 - 8                       4
'    8.5 - 14                    7
'    14.5 - 26                  10
'    27 or Larger               13

    If CmpDblGreaterthan(parBWidth, 3 * Inch) And CmpDblLessThanOrEqualTo(parBWidth, 8 * Inch) Then
        dTapHeight = 4 * Inch
    ElseIf CmpDblGreaterthan(parBWidth, 8 * Inch) And CmpDblLessThanOrEqualTo(parBWidth, 14 * Inch) Then
        dTapHeight = 7 * Inch
    ElseIf CmpDblGreaterthan(parBWidth, 14 * Inch) And CmpDblLessThanOrEqualTo(parBWidth, 26 * Inch) Then
        dTapHeight = 10 * Inch
    ElseIf CmpDblGreaterthan(parBWidth, 26 * Inch) Then
        dTapHeight = 13 * Inch
    End If

'Check to see that Branch Width doesn't exceed Width of the duct
    If CmpDblGreaterthan(parBWidth, parWidth) Then
        parBWidth = parWidth
    End If

'Check to see that Branch Depth doesn't exceed Depth of the duct
    If CmpDblGreaterthan(parBDepth, parDepth) Then
        parBDepth = parDepth
    End If

'Assigning the values of dBrBottomDepth(depth of bottom of the lo loss shape) based on conditions
'given in Mc Gill catalog
    If CmpDblLessThanOrEqualTo(parBDepth, parDepth - (2 * Inch)) Then
        dBrBottomDepth = parBDepth + (2 * Inch)
    Else
        dBrBottomDepth = parBDepth
    End If
    
    If CmpDblEqual(dBrBottomDepth, parDepth) Then
        dBrBottomDepth = dBrBottomDepth - 0.001
    End If
    
'Assigning the values of dBrBottomWidth(Width of bottom of the lo loss shape) based on conditions
'given in Mc Gill catalog
    dBrBottomWidth = parBWidth + dTapHeight
    
    iOutput = 0
    
    Dim dBottomPos As Double
    If parHVACShape = 4 Then 'Round
        dBottomPos = (parWidth / 2) - Sqr((parWidth / 2) ^ 2 - (dBrBottomDepth / 2) ^ 2)
    ElseIf parHVACShape = 1 Then 'Rectangular
        dBottomPos = 0
    ElseIf parHVACShape = 3 Then  'FlatOval
        dBottomPos = (parDepth / 2) - Sqr((parDepth / 2) ^ 2 - (dBrBottomDepth / 2) ^ 2)
    End If
    
' Insert your code for output 6(Insulation for Tap Body)
    Dim objInsTapBody As Object
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim objHeader  As Object
    Dim objBranch  As Object
    Dim objBranch2  As Object
    
    Dim CP As New AutoMath.DPosition
    Dim Dir As New AutoMath.DVector
    Dim oTopCurve As Object
    
    Dim oIJSurfaceHeader As IJSurface
    Dim oIntersectionCurve As IJElements
    Dim IntersectCode As Geom3dIntersectConstants
        
    Dim dCPXPos As Double
    dCPXPos = (dBrBottomWidth - parBWidth) / 2
            
    Dim oIJcurve As IJCurve
    
'Defining curves based on shape
    If parHVACShape = FlatOval Then
        Dim oHeaderCurve As Object
        Dim oBranchCurve As Object
        Dim oBranchCurve2 As Object
        
         If CmpDblGreaterthanOrEqualTo(parInsulationThickness, dTapHeight) Then
            CP.Set 0, -parWidth / 2 - (dTapHeight) - (2 * Inch), -parDepth / 2
            Set oTopCurve = CreFlatOvalBranch(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), 0)
            CP.Set -dCPXPos, -parWidth / 2 + dBottomPos, -parDepth / 2
            Set oIJcurve = CreFlatOvalBranch(CP, dBrBottomWidth + (2 * parInsulationThickness), dBrBottomDepth + (2 * parInsulationThickness), 0)
        Else
         'Creating transient header(shape of the duct)
         CP.Set (parBWidth / 2) + (Inch) + parInsulationThickness, 0, -parDepth / 2
         Set oHeaderCurve = CreFlatOval(CP, parWidth + (2 * parInsulationThickness), parDepth + (2 * parInsulationThickness), 0)
         Set objHeader = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oHeaderCurve, _
                                                     -1, 0, 0, _
                                                     parBWidth + dTapHeight + (2 * Inch) + (2 * parInsulationThickness), False)
         
         'Creating flatoval curve with dimensions of branch width and branch depth
         CP.Set 0, 0, -parDepth / 2
         Set oBranchCurve = CreFlatOvalBranch(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), 0)
         
         'Creating flatoval curve with width and depth dimensions of bottom of insulation of the Lo-Loss shape
         CP.Set -dCPXPos, 0, -parDepth / 2
         Set oBranchCurve2 = CreFlatOvalBranch(CP, dBrBottomWidth + (2 * parInsulationThickness), dBrBottomDepth + (2 * parInsulationThickness), 0)
         
         'Creating transient Branch
         Set objBranch = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oBranchCurve, _
                                                     0, -1, 0, _
                                                     parWidth / 2 + (dTapHeight) + (2 * Inch), False)
         'Creating transient Branch2 with width and depth dimensions(including insulation) of bottom of the lo loss shape
         Set objBranch2 = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oBranchCurve2, _
                                                     0, -1, 0, _
                                                     parWidth / 2 + (dTapHeight), False)
        
         'Creating plane normal to branch
         Dim objPlane As IngrGeom3D.Plane3d
         Set objPlane = m_GeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, -parWidth / 2 - (dTapHeight) - (2 * Inch), -parDepth / 2, 0, -1, 0)
         
         'Getting top curve from the intersection of transient Branch and plane
         Set oIJSurfaceHeader = objBranch
         oIJSurfaceHeader.Intersect objPlane, oIntersectionCurve, IntersectCode
         Set oTopCurve = oIntersectionCurve.Item(1)
        
         'Getting bottom curve from the intersection of transient Branch2 and transient header
         Set oIJSurfaceHeader = objBranch2
         oIJSurfaceHeader.Intersect objHeader, oIntersectionCurve, IntersectCode
         Set oIJcurve = oIntersectionCurve.Item(1)
        End If
    
    ElseIf parHVACShape = Rectangular Then
        
        'Creating Top Curve
        CP.Set 0, 0, dTapHeight + (2 * Inch)
        Set oTopCurve = CreSMRectBranch(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), 0)
        
        'Creating Bottom Curve
        CP.Set -dCPXPos, 0, 0
        Set oIJcurve = CreSMRectBranch(CP, dBrBottomWidth + (2 * parInsulationThickness), dBrBottomDepth + (2 * parInsulationThickness), 0)
                                                    
    ElseIf parHVACShape = 4 Then 'Round
        
        
        'Creating transient Branch2 as ellipse dimensions(including insulation thickness) of bottom of the lo loss shape
        Dim oMajor As New AutoMath.DPosition
        Dim dMMRatio As Double
        Dim oEllipse As IngrGeom3D.Ellipse3d
        
        CP.Set -dCPXPos, 0, -parDepth / 2
        oMajor.Set dBrBottomWidth / 2 + (parInsulationThickness), 0, 0
        dMMRatio = (dBrBottomDepth + (2 * parInsulationThickness)) / (dBrBottomWidth + (2 * parInsulationThickness))
        
        
        If CmpDblGreaterthanOrEqualTo(parInsulationThickness, dTapHeight) Then
            CP.Set 0, 0, dTapHeight + (2 * Inch)
            Dir.Set 0, 0, 1
            Set oTopCurve = PlaceTrCircleByCenter(CP, Dir, (parBWidth + (2 * parInsulationThickness)) / 2)
            CP.Set -dCPXPos, 0, -dBottomPos
            Set oIJcurve = m_GeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                        CP.x, CP.y, CP.z, 0, 0, 1, oMajor.x, oMajor.y, oMajor.z, dMMRatio)
        Else
            'Getting bottom curve from the intersection of transient Branch2 and transient header
            'Creating transient header(round duct shape)
            stPoint.Set (-parBWidth / 2) - dTapHeight - parInsulationThickness - (Inch), 0, -parWidth / 2
            enPoint.Set (parBWidth / 2) + (Inch) + parInsulationThickness, 0, -parWidth / 2
            Set objHeader = PlaceCylinderTrans(stPoint, enPoint, parWidth + (2 * parInsulationThickness))
            
            Set oEllipse = m_GeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                            CP.x, CP.y, CP.z, 0, 0, 1, oMajor.x, oMajor.y, oMajor.z, dMMRatio)
                                          
            Set objBranch2 = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oEllipse, _
                                                        0, 0, 1, _
                                                        parWidth / 2 + (dTapHeight) + (2 * Inch), False)
                                                        
            
            'plane normal to branch
            Set objPlane = m_GeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, 0, dTapHeight + (2 * Inch), 0, 0, 1)
        
        'Creating transient branch
            stPoint.Set 0, 0, -parWidth / 2
            enPoint.Set 0, 0, dTapHeight + (2 * Inch)
            Set objBranch = PlaceCylinderTrans(stPoint, enPoint, parBWidth + (2 * parInsulationThickness))
            
        'Getting top curve from the intersection of transient Branch and plane
            Set oIJSurfaceHeader = objBranch
            oIJSurfaceHeader.Intersect objPlane, oIntersectionCurve, IntersectCode
            Set oTopCurve = oIntersectionCurve.Item(1)
       
            Set oIJSurfaceHeader = objBranch2
            oIJSurfaceHeader.Intersect objHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve = oIntersectionCurve.Item(1)
        End If
    End If
    
    'Creating insulation for tap body by joining top curve and bottom curve
    Set objInsTapBody = m_GeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oTopCurve, oIJcurve, True)
            
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTapBody
    

    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
    Exit Sub

End Sub

Private Sub Class_Terminate()
    Set m_GeomFactory = Nothing
End Sub


Private Function PlaceCylinderTrans(lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double) As Object

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
    Set objProjection = geomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, False)
    
    Set objCircle = Nothing
    
    Set PlaceCylinderTrans = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function

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
    
    Const METHOD = "CreFlatOvalBranch:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x - (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(2).Set CP.x + (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(3).Set CP.x + Width / 2, CP.y, CP.z
    Pt(4).Set CP.x + (Width - Depth) / 2, CP.y, CP.z - Depth / 2
    Pt(5).Set CP.x - (Width - Depth) / 2, CP.y, CP.z - Depth / 2
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
    ReportUnanticipatedError2 MODULE, METHOD
    
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
    ReportUnanticipatedError2 MODULE, METHOD
End Function








