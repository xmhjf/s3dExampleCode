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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   CSimplePhysical.cls
'   Author:         RRK
'   Creation Date:  Tuesday, Aug 7 2007
'   Description:
'   This symbol is prepared for Shut off damper that caters to Round and Rectangular shapes as per CR-123916
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oEditErrors As IJEditErrors
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    Set m_oEditErrors = New JServerErrors
End Sub
Private Sub Class_Terminate()
    Set m_oEditErrors = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parDamperLength As Double
    Dim parActuatorPosition As Double
    Dim parCentertoSlot As Double
    Dim parSlotToLeverEnd As Double
    Dim parActuatorWidth As Double
    Dim parRodExtension As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim CornerRadius As Double
    CornerRadius = 0

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parDamperLength = arrayOfInputs(4)
    parActuatorPosition = arrayOfInputs(5)
    parCentertoSlot = arrayOfInputs(6)
    parSlotToLeverEnd = arrayOfInputs(7)
    parActuatorWidth = arrayOfInputs(8)
    parRodExtension = arrayOfInputs(9)
    parAngle = arrayOfInputs(10)
    parHVACShape = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    'Restricting the angle values to -5 degrees to 90 degress in case of Round and width greater than 600mm and
    ' to 0 to 90 degrees in all other cases.
    If parHVACShape = 1 Or CmpDblLessThan(parWidth, 0.6) Then
            If CmpDblLessThan(parAngle, -LINEAR_TOLERANCE) Or _
                CmpDblGreaterthan(parAngle, PI / 2 + LINEAR_TOLERANCE) Then
                m_oEditErrors.AddFromErr Err, "Invalid Angle Value: " & parAngle * 180 / PI & "; Acceptable Angle values are zero to 90 degrees", _
                METHOD, MODULE
                ReportUnanticipatedError2 MODULE, METHOD
            End If
    Else
        If CmpDblLessThan(parAngle, -5 * PI / 180 - LINEAR_TOLERANCE) Or _
                CmpDblGreaterthan(parAngle, PI / 2 + LINEAR_TOLERANCE) Then
                m_oEditErrors.AddFromErr Err, "Invalid Angle Value: " & parAngle * 180 / PI & "; Acceptable Angle values are -5 to 90 degrees", _
                METHOD, MODULE
                ReportUnanticipatedError2 MODULE, METHOD
        End If
    End If
    
    Dim dActPosFromCenter As Double
    dActPosFromCenter = (parDamperLength / 2) - parActuatorPosition
    
'Creating Damper Body (Output 1)
    Dim objDamperBody As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    If (parHVACShape = 4) Then 'Round
    
        oStPoint.Set -parDamperLength / 2, 0, 0
        oEnPoint.Set parDamperLength / 2, 0, 0
        Set objDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
        
    ElseIf (parHVACShape = Rectangular) Then
        oStPoint.Set -parDamperLength / 2, parDepth / 2, parWidth / 2
        oEnPoint.Set parDamperLength / 2, -parDepth / 2, -parWidth / 2
        Set objDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    
'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
    Set objDamperBody = Nothing

'Creating Cylindrical Rod (Output 2)
    Dim objCylRod As Object
    Dim dCylRodDia As Double
    
    'Assuming The cylindrical rod dia to be 20 mm
    dCylRodDia = 0.02
    
    oStPoint.Set dActPosFromCenter, 0, -(parWidth / 2) - parRodExtension
    oEnPoint.Set dActPosFromCenter, 0, (parWidth / 2) + parActuatorWidth
    
    Set objCylRod = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylRodDia, True)
    
'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylRod
    Set objCylRod = Nothing
    
'Insert your code for output 3 (Frame)
    
    Dim LineStrPoints(0 To 14) As Double
    Dim dFrameCenttoBottom As Double
    dFrameCenttoBottom = 0.038

    'Defining points for outline of the frame
    LineStrPoints(0) = dActPosFromCenter
    LineStrPoints(1) = parCentertoSlot + 0.012
    LineStrPoints(2) = (parWidth / 2) + (parActuatorWidth / 2)
    
    LineStrPoints(3) = dActPosFromCenter - dFrameCenttoBottom
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = LineStrPoints(2)
    
    LineStrPoints(6) = dActPosFromCenter - dFrameCenttoBottom
    LineStrPoints(7) = -dFrameCenttoBottom
    LineStrPoints(8) = LineStrPoints(2)
    
    LineStrPoints(9) = parCentertoSlot + 0.012 + dActPosFromCenter
    LineStrPoints(10) = -dFrameCenttoBottom
    LineStrPoints(11) = LineStrPoints(2)
    
    LineStrPoints(12) = parCentertoSlot + 0.012 + dActPosFromCenter
    LineStrPoints(13) = 0
    LineStrPoints(14) = LineStrPoints(2)
    
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
        
    Dim objCurvesCollection  As Collection
    Set objCurvesCollection = New Collection
    
    'Defining lines (or curves) for the outline of the frame
    Dim oLine1 As IngrGeom3D.Line3d
    Set oLine1 = New IngrGeom3D.Line3d
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(0), LineStrPoints(1), LineStrPoints(2), _
                            LineStrPoints(3), LineStrPoints(4), LineStrPoints(5))
    objCurvesCollection.Add oLine1
    Set oLine1 = Nothing
                            
    Dim oLine2 As IngrGeom3D.Line3d
    Set oLine2 = New IngrGeom3D.Line3d
    Set oLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(3), LineStrPoints(4), LineStrPoints(5), _
                            LineStrPoints(6), LineStrPoints(7), LineStrPoints(8))
    objCurvesCollection.Add oLine2
    Set oLine2 = Nothing
    
    Dim oLine3 As IngrGeom3D.Line3d
    Set oLine3 = New IngrGeom3D.Line3d
    Set oLine3 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(6), LineStrPoints(7), LineStrPoints(8), _
                            LineStrPoints(9), LineStrPoints(10), LineStrPoints(11))
    objCurvesCollection.Add oLine3
    Set oLine3 = Nothing
    
    Dim oLine4 As IngrGeom3D.Line3d
    Set oLine4 = New IngrGeom3D.Line3d
    Set oLine4 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(9), LineStrPoints(10), LineStrPoints(11), _
                            LineStrPoints(12), LineStrPoints(13), LineStrPoints(14))
    objCurvesCollection.Add oLine4
    Set oLine4 = Nothing
    
    Dim oArc As IngrGeom3D.Arc3d
    Set oArc = New IngrGeom3D.Arc3d
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, dActPosFromCenter, 0, LineStrPoints(2), _
                                LineStrPoints(12), LineStrPoints(13), LineStrPoints(14), _
                                LineStrPoints(0), LineStrPoints(1), LineStrPoints(2))
    objCurvesCollection.Add oArc
    Set oArc = Nothing
    
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set LineStrPoints(0), LineStrPoints(1), LineStrPoints(2)
    
    Dim objFrameoutline  As IngrGeom3D.ComplexString3d
    Set objFrameoutline = PlaceTrCString(oStPoint, objCurvesCollection)
    
    Dim oPlane As IngrGeom3D.Plane3d
    Set oPlane = New IngrGeom3D.Plane3d
    Set oPlane = geomFactory.Planes3d.CreateByOuterBdry(Nothing, objFrameoutline)
    
    'Defining points for the outline of the hole on the frame
    Dim LineStrPoints2(0 To 11) As Double
    Dim dSlotGap As Double
    dSlotGap = 0.006
    
    LineStrPoints2(0) = dActPosFromCenter
    LineStrPoints2(1) = parCentertoSlot + dSlotGap
    LineStrPoints2(2) = (parWidth / 2) + (parActuatorWidth / 2)
    
    LineStrPoints2(3) = parCentertoSlot + dSlotGap + dActPosFromCenter
    LineStrPoints2(4) = 0
    LineStrPoints2(5) = LineStrPoints2(2)
    
    LineStrPoints2(6) = parCentertoSlot - dSlotGap + dActPosFromCenter
    LineStrPoints2(7) = 0
    LineStrPoints2(8) = LineStrPoints2(2)
    
    LineStrPoints2(9) = dActPosFromCenter
    LineStrPoints2(10) = parCentertoSlot - dSlotGap
    LineStrPoints2(11) = LineStrPoints2(2)
    
    Dim objCurvesCollection2  As Collection
    Set objCurvesCollection2 = New Collection
    
    'Defining curves for outline of the hole of the frame
    Dim oArc1 As IngrGeom3D.Arc3d
    Set oArc1 = New IngrGeom3D.Arc3d
    Set oArc1 = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, dActPosFromCenter, 0, LineStrPoints2(2), _
                                LineStrPoints2(0), LineStrPoints2(1), LineStrPoints2(2), _
                                LineStrPoints2(3), LineStrPoints2(4), LineStrPoints2(5))
    objCurvesCollection2.Add oArc1
    Set oArc1 = Nothing
    
    Dim oArc2 As IngrGeom3D.Arc3d
    Set oArc2 = New IngrGeom3D.Arc3d
    Set oArc2 = geomFactory.Arcs3d.CreateBy3Points(Nothing, LineStrPoints2(3), LineStrPoints2(4), LineStrPoints2(5), _
                                parCentertoSlot + dActPosFromCenter, -dSlotGap, LineStrPoints2(2), _
                                LineStrPoints2(6), LineStrPoints2(7), LineStrPoints2(8))
    objCurvesCollection2.Add oArc2
    Set oArc2 = Nothing
    
    Dim oArc3 As IngrGeom3D.Arc3d
    Set oArc3 = New IngrGeom3D.Arc3d
    Set oArc3 = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, dActPosFromCenter, 0, LineStrPoints2(2), _
                                LineStrPoints2(6), LineStrPoints2(7), LineStrPoints2(8), _
                                LineStrPoints2(9), LineStrPoints2(10), LineStrPoints2(11))
    objCurvesCollection2.Add oArc3
    Set oArc3 = Nothing
    
    Dim oArc4 As IngrGeom3D.Arc3d
    Set oArc4 = New IngrGeom3D.Arc3d
    Set oArc4 = geomFactory.Arcs3d.CreateBy3Points(Nothing, LineStrPoints2(9), LineStrPoints2(10), LineStrPoints2(11), _
                                                    -dSlotGap + dActPosFromCenter, parCentertoSlot, LineStrPoints2(2), _
                                LineStrPoints2(0), LineStrPoints2(1), LineStrPoints2(2))
    objCurvesCollection2.Add oArc4
    Set oArc4 = Nothing
    
    oStPoint.Set LineStrPoints2(0), LineStrPoints2(1), LineStrPoints2(2)
    Dim objHoleOutline  As IngrGeom3D.ComplexString3d
    Set objHoleOutline = PlaceTrCString(oStPoint, objCurvesCollection2)
    
    oPlane.AddHole objHoleOutline
    Set objCurvesCollection = Nothing
    Set objCurvesCollection2 = Nothing
    Set objHoleOutline = Nothing
    
    Dim oDirVector As AutoMath.DVector
    Set oDirVector = New AutoMath.DVector
    
    Dim oTransformationMat  As AutoMath.DT4x4
    Set oTransformationMat = New AutoMath.DT4x4

'Transforming plane (representing the frame) in case of Rectangular shape
    If (parHVACShape = Rectangular) Then
        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(0) = -1
        oPlane.Transform oTransformationMat
        
        oTransformationMat.LoadIdentity
        oDirVector.Set 2 * dActPosFromCenter, 0, 0
        oTransformationMat.Translate oDirVector
        oPlane.Transform oTransformationMat
    End If
    
    Dim objFrameBody As Object
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
'Creating sweep of plane through the straight line
    Dim oSweepLine As IngrGeom3D.Line3d
    Set oSweepLine = New IngrGeom3D.Line3d
    Set oSweepLine = geomFactory.Lines3d.CreateBy2Points(Nothing, dActPosFromCenter, 0, (parWidth / 2) + (parActuatorWidth / 2), _
                                                               dActPosFromCenter, 0, (parWidth / 2) + (parActuatorWidth / 2) + 0.4 * parActuatorWidth)
    Dim Surfset   As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Set Surfset = geomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                  oSweepLine, oPlane, CircularCorner, 3, stnorm, ednorm, True)
                  
    For Each objFrameBody In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrameBody
    Next objFrameBody
    
'Creating Lever (Output 4)
    
    'Assuming lever wdith to be 25mm
    Dim dLeverWidth As Double
    dLeverWidth = 0.025
    Dim LineStrPoints3(0 To 15) As Double
    
    'Creating points for the cross section of the lever
    LineStrPoints3(0) = dActPosFromCenter - (dLeverWidth / 2) * Sin(parAngle)
    LineStrPoints3(1) = (dLeverWidth / 2) * Cos(parAngle)
    LineStrPoints3(2) = (parWidth / 2) + parActuatorWidth
    
    LineStrPoints3(3) = dActPosFromCenter + (dLeverWidth / 2) * Sin(parAngle)
    LineStrPoints3(4) = -(dLeverWidth / 2) * Cos(parAngle)
    LineStrPoints3(5) = LineStrPoints3(2)
    
    LineStrPoints3(6) = LineStrPoints3(3)
    LineStrPoints3(7) = LineStrPoints3(4)
    LineStrPoints3(8) = LineStrPoints3(2) - 0.1 * parActuatorWidth
    
    LineStrPoints3(9) = -(dLeverWidth / 2) * Sin(parAngle) + dActPosFromCenter
    LineStrPoints3(10) = (dLeverWidth / 2) * Cos(parAngle)
    LineStrPoints3(11) = LineStrPoints3(2) - 0.1 * parActuatorWidth
    
    LineStrPoints3(12) = LineStrPoints3(0)
    LineStrPoints3(13) = LineStrPoints3(1)
    LineStrPoints3(14) = LineStrPoints3(2)
    
    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints3)
    
    Dim objLever As Object
    oAxisVec.Set Cos(parAngle), Sin(parAngle), 0
    Set objLever = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parCentertoSlot + parSlotToLeverEnd, True)
    
    'Transforming lever in case of Rectangular shape
    If (parHVACShape = Rectangular) Then
        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(0) = -1
        objLever.Transform oTransformationMat
        
        oTransformationMat.LoadIdentity
        oDirVector.Set 2 * dActPosFromCenter, 0, 0
        oTransformationMat.Translate oDirVector
        objLever.Transform oTransformationMat
    End If
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLever
    Set objLever = Nothing
    
    '-------------------------------------------------------------------------------------------------
    '==================================================
    'Create HVAC NOZZLE 1
    '==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    
    Dim oPos As AutoMath.DPosition
    Set oPos = New AutoMath.DPosition
    Dim oPortDir As AutoMath.DVector
    Set oPortDir = New AutoMath.DVector
    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector
    

    Dim iPortIndex As Integer

    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    Dim dNozzLength(1 To 2) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    
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
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                parHVACShape, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    Dim x As Double, y As Double, z As Double
    x = -parDamperLength / 2
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oPortDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oPortDir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

    '===================
    'BUILD HVACNOZZLE2
    '===================
    iPortIndex = 2

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                    parHVACShape, lEndPrep(2), dThickness(2), dFlangeWidth(2), lFlowDir(2), _
                    parWidth, parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                    m_OutputColl.ResourceManager)

    
    'Position of the nozzle should be the conenct point of the nozzle
    x = parDamperLength / 2
    y = 0
    z = 0
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oPortDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oPortDir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    
    Set NozzleFactory = Nothing
    Set oPos = Nothing
    Set oPortDir = Nothing
    Set oRadialDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub
