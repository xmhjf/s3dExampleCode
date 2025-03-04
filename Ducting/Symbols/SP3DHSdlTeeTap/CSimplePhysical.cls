VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   ProgID:         SP3DHSdlTeeTap.HSdlTeeTap
'   Author:         RRK
'   Creation Date:  Tuesday, June 13 2007
'   Description:
'   This symbol is prepared for saddle straight and conical tee taps of McGill Air flow corporation as per CR-120452
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
        
'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    Dim dBrBottomWidth As Double
    Dim dBrBottomDepth As Double
    Dim dTapHeight As Double
    
    'Condition for part data basis values of Default(codelist value:1)
    'or Saddle Tee Tap, Straight(codelist value: 35)
    If lPartdatabasis <= 1 Or lPartdatabasis = 35 Then
        'Check to see that Branch Width doesn't exceed Width of the duct
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        'Check to see that Branch Depth doesn't exceed Depth of the duct
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
        dBrBottomWidth = parBWidth
        dBrBottomDepth = parBDepth
        dTapHeight = Inch
    
    'Condition for part data basis value of Saddle Tee Tap, Conical(codelist value:40)
    ElseIf lPartdatabasis = 40 Then
        
        'Check to see that Branch Width doesn't exceed (Width-2inch) as per Mc Gill Catalog
        If CmpDblGreaterthan(parBWidth, parWidth - (2 * Inch)) Then
            parBWidth = parWidth - (2 * Inch)
        End If
        
        'Check to see that Branch Depth doesn't exceed (Depth-2inch) as per Mc Gill Catalog
        If CmpDblGreaterthan(parBDepth, parDepth - (2 * Inch)) Then
            parBDepth = parDepth - (2 * Inch)
        End If
        
        dBrBottomWidth = parBWidth + (1 * Inch)
        dBrBottomDepth = parBDepth + (1 * Inch)
        dTapHeight = 4 * Inch
    End If
    
    If CmpDblEqual(dBrBottomDepth, parDepth) Then
        dBrBottomDepth = dBrBottomDepth - 0.001
    End If
    
    iOutput = 0
    
' Insert your code for output 1(TapBody)
    Dim objTapBody As Object

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
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    Dim oTakeoffCurve As Object
    Dim oStripCurve As Object

    Dim oIJcurve As IJCurve
    
'Defining curves based on shape
    If parHVACShape = FlatOval Then
        Dim oHeaderCurve As Object
        Dim oBranchCurve As Object
        Dim oBranchCurve2 As Object
        
        'Creating transient header(shape of the duct)
        CP.Set -(parBWidth / 2) - Inch, 0, -parDepth / 2
        Set oHeaderCurve = CreFlatOval(CP, parWidth, parDepth, 0)
        Set objHeader = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oHeaderCurve, _
                                                    1, 0, 0, _
                                                    parBWidth + (2 * Inch), False)
        
        'Creating flatoval curve with dimensions of branch width and branch depth
        CP.Set 0, 0, -parDepth / 2
        Set oBranchCurve = CreFlatOvalBranch(CP, parBWidth, parBDepth, 0)
        
        'Creating flatoval curve with dBrBottomWidth and dBrBottomWidth dimensions
        Set oBranchCurve2 = CreFlatOvalBranch(CP, dBrBottomWidth, dBrBottomDepth, 0)
        
        
        'Creating transient Branch
        Set objBranch = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oBranchCurve, _
                                                    0, -1, 0, _
                                                    parWidth / 2 + (dTapHeight), False)
        
        'Creating transient Branch2 with dBrBottomWidth and dBrBottomWidth dimensions
        Set objBranch2 = m_GeomFactory.Projections3d.CreateByCurve(Nothing, oBranchCurve2, _
                                                    0, -1, 0, _
                                                    parWidth / 2 + (dTapHeight), False)
                                                    
        'plane normal to branch
        Dim objPlane As IngrGeom3D.Plane3d
        Set objPlane = m_GeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, -parWidth / 2 - (dTapHeight), -parDepth / 2, 0, -1, 0)
        
        'Getting top curve from the intersection of transient Branch and plane
        Set oIJSurfaceHeader = objBranch
        oIJSurfaceHeader.Intersect objPlane, oIntersectionCurve, IntersectCode
        Set oTopCurve = oIntersectionCurve.Item(1)
        
        'Getting bottom curve from the intersection of transient Branch2 and transient header
        Set oIJSurfaceHeader = objBranch2
        oIJSurfaceHeader.Intersect objHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve = oIntersectionCurve.Item(1)
        
        'Creating Take off curve
        CP.Set 0, -(parWidth / 2) - (dTapHeight), -parDepth / 2
        Set oTakeoffCurve = CreFlatOvalBranch(CP, 1.01 * parBWidth, 1.01 * parBDepth, 0)
        
        'Creating strip curve
        CP.Set 0, -(parWidth / 2) - (dTapHeight) + 0.004, -parDepth / 2
        Set oStripCurve = CreFlatOvalBranch(CP, 1.06 * parBWidth, 1.06 * parBDepth, 0)
        
    ElseIf parHVACShape = Rectangular Then
        
        'Creating Top Curve
        CP.Set 0, 0, dTapHeight
        Set oTopCurve = CreSMRectBranch(CP, parBWidth, parBDepth, 0)
        
        'Creating Bottom Curve
        CP.Set 0, 0, 0
        Set oIJcurve = CreSMRectBranch(CP, dBrBottomWidth, dBrBottomDepth, 0)
        
        'Creating Take off
        CP.Set 0, 0, dTapHeight
        Set oTakeoffCurve = CreSMRectBranch(CP, 1.01 * parBWidth, 1.01 * parBDepth, 0)
        
        'Creating Strip Curve
        CP.Set 0, 0, (dTapHeight) - 0.004
        Set oStripCurve = CreSMRectBranch(CP, 1.06 * parBWidth, 1.06 * parBDepth, 0)
                                                    
    ElseIf parHVACShape = 4 Then 'Round
           
        'Creating transient header(round duct shape)
        stPoint.Set (-parBWidth / 2) - Inch, 0, -parWidth / 2
        enPoint.Set (parBWidth / 2) + Inch, 0, -parWidth / 2
        Set objHeader = PlaceCylinderTrans(stPoint, enPoint, parWidth)
        
        'Creating transient branch(round duct shape)
        stPoint.Set 0, 0, -parWidth / 2
        enPoint.Set 0, 0, dTapHeight
        Set objBranch = PlaceCylinderTrans(stPoint, enPoint, dBrBottomDepth)
        
        'Creating top curve
        CP.Set 0, 0, dTapHeight
        Dir.Set 0, 0, -1
        Set oTopCurve = PlaceTrCircleByCenter(CP, Dir, parBWidth / 2)
        
        'Transformation of top curve by 180 degrees about Z axis is done to get appropriate graphics when ruled surface is used to create tap body
        oTransformationMat.LoadIdentity
        dRotAboutZaxis = PI
        oDirVector.Set 0, 0, 1
        oTransformationMat.Rotate dRotAboutZaxis, oDirVector
        oTopCurve.Transform oTransformationMat
         
        'Creating Take off curve
        Set oTakeoffCurve = PlaceTrCircleByCenter(CP, Dir, 1.01 * parBWidth / 2)
        
        'Creating Strip curve
        CP.Set 0, 0, (dTapHeight) - 0.004
        Set oStripCurve = PlaceTrCircleByCenter(CP, Dir, 1.06 * parBWidth / 2)
        
        'Getting bottom curve from the intersection of transient Branch and transient header
        Set oIJSurfaceHeader = objBranch
        oIJSurfaceHeader.Intersect objHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve = oIntersectionCurve.Item(1)
         
    End If
    
    'Creating tap body by joining top curve and bottom curve
    Set objTapBody = m_GeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oTopCurve, oIJcurve, True)
            
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTapBody
    
    'Creating Output 2(Take-off)
    Dim objTakeoff As Object
    Dim objStrip As Object
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
    oAxisVec.Set 0, 0, 1
    If parHVACShape = FlatOval Then
          oAxisVec.Set 0, -1, 0
    End If
    
    Set objTakeoff = PlaceProjection(m_OutputColl, oTakeoffCurve, oAxisVec, 2 * Inch, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeoff
    Set objTakeoff = Nothing
    
    'Creating output 3 (Header Strip 1)
    Set objStrip = PlaceProjection(m_OutputColl, oStripCurve, oAxisVec, 0.008, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    

'Creating output 4(Flange Body)
    Dim FlangePoints(1) As New AutoMath.DPosition
    Dim objFlangeBody As Object
    Dim objFlangeoutline As Object

    If (parHVACShape = 1) Then
        Set objFlangeoutline = New IngrGeom3D.Line3d

        FlangePoints(0).Set -parBWidth / 2 - (2 * Inch), -parBDepth / 2 - (2 * Inch), 0
        FlangePoints(1).Set FlangePoints(0).x, -FlangePoints(0).y, FlangePoints(0).z

        Set objFlangeoutline = PlaceTrLine(FlangePoints(0), FlangePoints(1))

    Else
        Set objFlangeoutline = New IngrGeom3D.Arc3d

        Dim dAngle1 As Double
        Dim dAngle2 As Double
        Dim dTotAngle As Double

        If CmpDblEqual(parDepth, dBrBottomDepth) Then
            dAngle1 = PI / 2
            dAngle2 = 0
        Else
            dAngle1 = Atn(parBDepth / Sqr(parDepth ^ 2 - dBrBottomDepth ^ 2))
            dAngle2 = (2 * Inch) / (parDepth / 2)
        End If

        dTotAngle = dAngle1 + dAngle2

        FlangePoints(0).Set -parBWidth / 2 - (2 * Inch), -(parDepth / 2) * Sin(dTotAngle), (parDepth / 2) * Cos(dTotAngle) - parDepth / 2
        FlangePoints(1).Set FlangePoints(0).x, -FlangePoints(0).y, FlangePoints(0).z

        CP.Set -parBWidth / 2 - (2 * Inch), 0, -parDepth / 2
        oDirVector.Set -1, 0, 0

        Set objFlangeoutline = PlaceTrArcByCenterNorm(FlangePoints(0), FlangePoints(1), CP, oDirVector)
        
        'In case of flat oval transforming the position flange on to the round portion of falt oval
        If parHVACShape = FlatOval Then
            
            oTransformationMat.LoadIdentity
            dRotAboutZaxis = PI / 2
            oDirVector.Set 1, 0, 0
            oTransPos.Set 0, -parWidth / 2, -parDepth / 2
            
            oTransformationMat.Translate oTransPos
            oTransformationMat.Rotate dRotAboutZaxis, oDirVector
            objFlangeoutline.Transform oTransformationMat
            
        End If
    End If

    oAxisVec.Set 1, 0, 0
    Set objFlangeBody = PlaceProjection(m_OutputColl, objFlangeoutline, oAxisVec, parBWidth + (4 * Inch), True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody

    Set objTapBody = Nothing
    Set m_GeomFactory = Nothing
    Set objFlangeBody = Nothing
    Set objFlangeoutline = Nothing
'
' Place Nozzle 1
'=====================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'=====================================
    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    
    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector

    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1) As Long
    Dim dThickness(1) As Double
    Dim dFlangeWidth(1) As Double
    Dim lFlowDir(1) As Long
    Dim dPortDepth(1) As Double
    Dim dCptOffSet(1) As Double
    Dim dNozzLength(1) As Double
    
    dCornerRadius = 0#
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing

    iPortIndex = 1
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

' Insert your code for output 5(HvacPort)
    iPortIndex = 1
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parBWidth, parBDepth, _
                                             dCornerRadius, bDimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
    
'Nozzle in round and rectangular cases is along Z direction and in  flat oval case it is along Y direction
    If parHVACShape = 4 Or parHVACShape = Rectangular Then
        oPos.Set 0, 0, dTapHeight
        oDir.Set 0, 0, 1
    ElseIf parHVACShape = FlatOval Then
        oPos.Set 0, (-parWidth / 2) - (dTapHeight), -parDepth / 2
        oDir.Set 0, -1, 0
    End If
    Set oDistribPort = oHvacNozzle

    oDistribPort.SetPortLocation oPos
    oDistribPort.SetDirectionVector oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzLength(1)

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing

    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

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





