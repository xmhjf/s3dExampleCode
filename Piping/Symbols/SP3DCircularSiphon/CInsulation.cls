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
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Feb 13 2007
'   Description:
'       Circular Siphon (Source: Form D, www.sika.net/pdf/englisch/PRESS3_39_40.PDF)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   22.May.2007     svsmylav     CR-113569: Created the symbol.
'   7.Nov.2007      RRK          TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const NEGLIGIBLE_VALUE = 0.0001

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam         As Double
    Dim flangeThick      As Double
    Dim flangeDiam1       As Double
    Dim flangeDiam2       As Double
    Dim depth            As Double
    Dim flangeThick2     As Double
    Dim sptOffset       As Double
    
    Dim iOutput     As Double
    
    Dim parInletFacetoOutletSeat As Double
    Dim parFace1toCenter As Double
    Dim parRadius As Double
    Dim parDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFace2toCenter As Double
    Dim parFacetoFace As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInletFacetoOutletSeat = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parRadius = arrayOfInputs(4)
    parDiameter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parFace2toCenter = arrayOfInputs(7)
    parFacetoFace = arrayOfInputs(8)
    
    iOutput = 0

    Dim dInsulationDiameter As Double
    
'   Origin is taken at the center of the circular portion at half of the offset
'   between port 1 and port 2

'   Part data basis value 201:
        'Circular siphon, specified by inlet face to outlet seat and inlet face to center'
'   Part data basis value 202:
        'Circular siphon, specified by inlet face to center and outlet face to center'
'   Part data basis value 204:
        'Circular siphon, specified by face to face and outlet face to center'
        
 'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

'   Do initial calculation for preparing the symbol geometry
    Dim dNozzle1Length As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset, depth

    'Assumption: Offset between inlet and outlet ports are taken to be 1.5 times the pipe diameter
    Dim dOffset_Ports As Double
    dOffset_Ports = 1.5 * pipeDiam
    
'   Define face to center variables and use them for rest of the work
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 201) Then
        dFace1toCenter = parFace1toCenter
        'Compute distance upto the top portion of Clamp Muff Connection
        dFace2toCenter = parInletFacetoOutletSeat - dFace1toCenter + flangeThick2
    ElseIf (lPartDataBasis = 202) Then
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
    ElseIf (lPartDataBasis = 204) Then
        dFace2toCenter = parFace2toCenter
        dFace1toCenter = parFacetoFace - dFace2toCenter
    End If
    
    'Beginning of the bend (located nearer to port 1) is located at a distance D = R0 + R1 where
    'R0 = (parDiameter + pipeDiam) /2  and R1 = parRadius + pipeDiam /2. Distance along X axis will be
    'square rooot of (D^2 - R0^2) = Sqr(R0^2 + R1^2 + 2R0R1 - R1^2) = Sqr(R0^2 + 2R0R1)
'   Assumption: If value of radius is not provided in part data, it will be taken as
'   half of the diameter
    If CmpDblLessThanOrEqualTo(parRadius, 0) Then parRadius = parDiameter / 2

    Dim dR0 As Double
    Dim dR1 As Double
    Dim dArcAngle
    Dim dCentertoCenterDistX As Double 'Center to Center Distance measured along X-axis
    dR0 = (parDiameter + pipeDiam) / 2
    dR1 = parRadius + pipeDiam / 2
    dCentertoCenterDistX = Sqr(dR0 ^ 2 + 2 * dR0 * dR1)
    dNozzle1Length = dFace1toCenter - dCentertoCenterDistX
    If CmpDblLessThan(dNozzle1Length, flangeThick) Then dNozzle1Length = flangeThick

    'Angle of outer arc measured from center of siphon is arc tan of (dR1/dCentertoCenterDistX)
    dArcAngle = Atn(dR1 / dCentertoCenterDistX)

' Insert your code for output 1(Insulation for Siphon Body)
    'A complex string will be created as a sweep curve upto the Outlet Seat
    Dim oBotTangentLine   As IngrGeom3D.Line3d
    Dim oBotArcPort1Side  As IngrGeom3D.Arc3d
    Dim oTopArcPort1Side  As IngrGeom3D.Arc3d
    Dim oBotArcPort2Side  As IngrGeom3D.Arc3d
    Dim oTopArcPort2Side  As IngrGeom3D.Arc3d
    Dim oTopTangentLine   As IngrGeom3D.Line3d
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition

    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Bottom Tangent Line on Port 1 Side
    oStPoint.Set -dFace1toCenter, 0, dOffset_Ports / 2
    oEnPoint.Set -dCentertoCenterDistX, 0, dOffset_Ports / 2
    Set oBotTangentLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)

    'Bottom Arc on Port 1 Side
    oStPoint.Set -dCentertoCenterDistX, 0, dOffset_Ports / 2
    oEnPoint.Set -dR0 * Cos(dArcAngle), -dR0 * Sin(dArcAngle), dOffset_Ports / 2
    oCenPoint.Set -dCentertoCenterDistX, -dR1, dOffset_Ports / 2

    Set oBotArcPort1Side = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)

    'Top Arc on Port 1 Side
    oStPoint.Set -dR0 * Cos(dArcAngle), -dR0 * Sin(dArcAngle), dOffset_Ports / 2
    oEnPoint.Set dR0, 0, dOffset_Ports / 2
    oCenPoint.Set 0, 0, dOffset_Ports / 2
    
    Set oTopArcPort1Side = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    
    'Helical path at center, using ellipse
    Dim dCenX As Double
    Dim dCenY As Double
    Dim dCenZ As Double
    Dim dMajorX As Double
    Dim dMajorY As Double
    Dim dMajorZ As Double
    Dim mMRatio As Double
    Dim dStartAngle As Double
    Dim dSweepAngle As Double
    Dim dNorX As Double
    Dim dNorY As Double
    Dim dNorZ As Double
    Dim oHelixArc As IngrGeom3D.EllipticalArc3d

    'Compute normal vector
    Dim oVecYAxis As AutoMath.DVector
    Dim oVecHelixCenPoints As AutoMath.DVector
    Dim oVecHelixAxis As AutoMath.DVector
    Set oVecYAxis = New AutoMath.DVector
    Set oVecHelixCenPoints = New AutoMath.DVector

    oStPoint.Set dR0, 0, dOffset_Ports / 2 'Start point of helix
    oEnPoint.Set -dR0, 0, -dOffset_Ports / 2 'End point of helix
    oVecHelixCenPoints.Set oEnPoint.x - oStPoint.x, oEnPoint.y - oStPoint.y, oEnPoint.z - oStPoint.z

    oVecYAxis.Set 0, -1, 0
    
    Set oVecHelixAxis = oVecHelixCenPoints.Cross(oVecYAxis)
    oVecHelixAxis.Length = 1 'Normalize
    dNorX = oVecHelixAxis.x
    dNorY = oVecHelixAxis.y
    dNorZ = oVecHelixAxis.z
    
    'Set center
    dCenX = 0
    dCenY = 0
    dCenZ = 0

    'Compute major axis
    dMajorX = oVecHelixCenPoints.x / 2
    dMajorY = oVecHelixCenPoints.y / 2
    dMajorZ = oVecHelixCenPoints.z / 2

    'Compute mMRatio
    Dim dHelixAngle As Double
    dHelixAngle = Atn(dOffset_Ports / (2 * dR0))
    
    mMRatio = Cos(dHelixAngle)
    dStartAngle = PI
    dSweepAngle = PI

    Set oHelixArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            dCenX, dCenY, dCenZ, dNorX, dNorY, dNorZ, dMajorX, dMajorY, dMajorZ, mMRatio, _
                                            dStartAngle, dSweepAngle)

    'Bottom Arc on Port 2 Side
    oStPoint.Set -dR0, 0, -dOffset_Ports / 2
    oEnPoint.Set dR0 * Cos(dArcAngle), -dR0 * Sin(dArcAngle), -dOffset_Ports / 2
    oCenPoint.Set 0, 0, -dOffset_Ports / 2

    Set oBotArcPort2Side = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    
    'Top Arc on Port 2 Side
    oStPoint.Set dR0 * Cos(dArcAngle), -dR0 * Sin(dArcAngle), -dOffset_Ports / 2
    oEnPoint.Set dCentertoCenterDistX, 0, -dOffset_Ports / 2
    oCenPoint.Set dCentertoCenterDistX, -dR1, -dOffset_Ports / 2
    
    Set oTopArcPort2Side = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    
    'Top Tangent Line on Port 2 Side
    oStPoint.Set dCentertoCenterDistX, 0, -dOffset_Ports / 2
    If lPartDataBasis <= 1 Or lPartDataBasis = 201 Then
        oEnPoint.Set parInletFacetoOutletSeat - dFace1toCenter, _
                0, -dOffset_Ports / 2
    ElseIf lPartDataBasis = 202 Or lPartDataBasis = 204 Then
        oEnPoint.Set dFace2toCenter, 0, -dOffset_Ports / 2
    End If
    Set oTopTangentLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)

' I BranchPort2Side
    Dim oSweepCurve         As IngrGeom3D.ComplexString3d
    Dim oSweepCurveCol      As Collection
    Set oSweepCurveCol = New Collection
    oSweepCurveCol.Add oBotTangentLine
    oSweepCurveCol.Add oBotArcPort1Side
    oSweepCurveCol.Add oTopArcPort1Side
    oSweepCurveCol.Add oHelixArc
    oSweepCurveCol.Add oBotArcPort2Side
    oSweepCurveCol.Add oTopArcPort2Side
    oSweepCurveCol.Add oTopTangentLine

    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set -dFace1toCenter, 0, dOffset_Ports / 2
    Set oSweepCurve = PlaceTrCString(StartBC, oSweepCurveCol)
    
    Dim oCircle As IngrGeom3D.Circle3d
    
'   Compute insulation diameter and check to see it does not exceed the circular bend diameter
    dInsulationDiameter = pipeDiam / 2 + parInsulationThickness
    If CmpDblGreaterthan(dInsulationDiameter, parDiameter) Then dInsulationDiameter = parDiameter
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                               -dFace1toCenter, 0, dOffset_Ports / 2, _
                                                1, 0, 0, _
                                                dInsulationDiameter)

    Dim Surfset   As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                     oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)
                       
    Dim ObjSurface As Object
    For Each ObjSurface In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
    Next ObjSurface

    If lPartDataBasis <= 1 Or lPartDataBasis = 201 Then
'  Insert your code for output 10(Clamp muff connection at Port2)
        Dim ObjInsClampMuffConnection As Object
        'Assumptions:
        ' 1. The insulation shape for the Clamp muff connection is considered to be cylindrical
        ' 2. Projection Length is taken as 2 times the flange thickness
        Dim dProjLen As Double
        dProjLen = 2 * flangeThick2
        If CmpDblLessThan(dProjLen, NEGLIGIBLE_VALUE) Then dProjLen = NEGLIGIBLE_VALUE
        
        oStPoint.Set dFace2toCenter, 0, -dOffset_Ports / 2
        oEnPoint.Set oStPoint.x - dProjLen, 0, -dOffset_Ports / 2
        If CmpDblGreaterthanOrEqualTo(flangeDiam2, pipeDiam) Then
            dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
        Else
             dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsClampMuffConnection = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                dInsulationDiameter, True)
    
    '   Set the output
        m_OutputColl.AddOutput "InsClampMuffConnection", ObjInsClampMuffConnection
        Set ObjInsClampMuffConnection = Nothing
    End If

' Insert your code for output 2(Insulation for Port 1)
    Dim objInsPort1 As Object
    oStPoint.Set -dFace1toCenter, 0, dOffset_Ports / 2
    oEnPoint.Set oStPoint.x + flangeThick + parInsulationThickness, 0, dOffset_Ports / 2

    If CmpDblGreaterthan(flangeDiam1, pipeDiam) Then
        dInsulationDiameter = flangeDiam1 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set oStPoint.x + NEGLIGIBLE_VALUE, 0, dOffset_Ports / 2
        
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(dInsulationDiameter, parDiameter) Then dInsulationDiameter = parDiameter
    End If
    
    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for Port 2)
    Dim objInsPort2 As Object
    oStPoint.Set dFace2toCenter, 0, -dOffset_Ports / 2
    oEnPoint.Set oStPoint.x - flangeThick2 - parInsulationThickness, 0, -dOffset_Ports / 2

    If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then
        dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick2, 0) Then
        oEnPoint.Set oStPoint.x - NEGLIGIBLE_VALUE, 0, -dOffset_Ports / 2
        
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(dInsulationDiameter, parDiameter) Then dInsulationDiameter = parDiameter
    End If
    
    Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort2
    Set objInsPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oGeomFactory = Nothing
    Set oVecYAxis = Nothing
    Set oVecHelixCenPoints = Nothing
    Set oSweepCurveCol = Nothing
    Set StartBC = Nothing
    Set oPipeComponent = Nothing
    Set oBotTangentLine = Nothing
    Set oBotArcPort1Side = Nothing
    Set oTopArcPort1Side = Nothing
    Set oVecHelixAxis = Nothing
    Set oHelixArc = Nothing
    Set oBotArcPort2Side = Nothing
    Set oTopArcPort2Side = Nothing
    Set oTopTangentLine = Nothing
    Set oSweepCurve = Nothing
    Set oCircle = Nothing
    Set Surfset = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

