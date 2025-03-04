VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Wednesday, Jan 09 2008
'   Description:
'   This symbol is created for CR-123918 (Water Tight Valve )
'   Source: Water Tight Valve.pdf.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parWidth As Double
    Dim parValveLength As Double
    Dim parHeight As Double
    Dim parHandwheelDia As Double
    Dim parHandletoCenter As Double
    Dim parHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorHeight As Double
    Dim parHeight2 As Double
    Dim parValveCenterToWheelEnd As Double
    Dim parOperatorAngle As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parValveLength = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parHandwheelDia = arrayOfInputs(5)
    parHandletoCenter = arrayOfInputs(6)
    parHeight1 = arrayOfInputs(7)
    parActuatorWidth = arrayOfInputs(8)
    parActuatorHeight = arrayOfInputs(9)
    parHeight2 = arrayOfInputs(10)
    parValveCenterToWheelEnd = arrayOfInputs(11)
    parOperatorAngle = arrayOfInputs(12)
   
    'Origin is taken at the Center of the Valve
    'For Horizontally routed Duct X - Along the flow and Y - towards UP
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for Output 1 (ValveBody1)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim objValveBody As Object
    oStPoint.Set -parValveLength / 2, 0, 0
    oEnPoint.Set parValveLength / 2, 0, 0
    
    Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
    Set objValveBody = Nothing
    
    'Insert your code for Output 2 (ValveBody2)
    oStPoint.Set 0, parWidth / 2, 0
    oEnPoint.Set 0, parHeight1, 0
    
    Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parValveLength, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
    Set objValveBody = Nothing

    'Insert your code for Output 3 (OperatorBody1)
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
        
    dLineStrPoints(0) = -parHandletoCenter / 2
    dLineStrPoints(1) = parHeight1
    dLineStrPoints(2) = parValveCenterToWheelEnd / 2

    dLineStrPoints(3) = -parHandletoCenter / 2
    dLineStrPoints(4) = parHeight1 + (0.8 * parHeight2)
    dLineStrPoints(5) = parValveCenterToWheelEnd / 2

    dLineStrPoints(6) = -parHandletoCenter / 2
    dLineStrPoints(7) = parHeight1 + (0.8 * parHeight2)
    dLineStrPoints(8) = -parValveCenterToWheelEnd / 2

    dLineStrPoints(9) = -parHandletoCenter / 2
    dLineStrPoints(10) = parHeight1
    dLineStrPoints(11) = -parValveCenterToWheelEnd / 2

    dLineStrPoints(12) = -parHandletoCenter / 2
    dLineStrPoints(13) = parHeight1
    dLineStrPoints(14) = parValveCenterToWheelEnd / 2

    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    
    Dim objOpBody As Object
    Set objOpBody = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                            parHandletoCenter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOpBody
    
    'Rotate OperatorBody1
    Dim oRotVector As AutoMath.DVector
    Set oRotVector = New AutoMath.DVector
    Dim oTransMatrix As DT4x4
    Set oTransMatrix = New DT4x4
  
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objOpBody.Transform oTransMatrix

    'Insert your code for Output 4 (OperatorBody2)
    oStPoint.Set 0, parHeight1 + (0.8 * parHeight2), 0
    oEnPoint.Set 0, parHeight1 + parHeight2, 0

    Set objOpBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parValveLength, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOpBody

    'Rotate OperatorBody2
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objOpBody.Transform oTransMatrix
    
    'Insert your code for Output 5 (ActuatorBody1)
    Dim objActuatorBody As Object
    oStPoint.Set 0, parHeight1 + parHeight2 + (0.5 * parActuatorHeight), -parActuatorWidth / 2
    oEnPoint.Set 0, parHeight1 + parHeight2 + (0.5 * parActuatorHeight), parActuatorWidth / 2

    Set objActuatorBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            parActuatorHeight, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objActuatorBody

    'Rotate ActuatorBody1
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objActuatorBody.Transform oTransMatrix
    
    'Insert your code for Output 6 (ActuatorBody2)
    oStPoint.Set 0, parHeight1 + parHeight2 + parActuatorHeight, 0
    oEnPoint.Set 0, (parHeight - 0.5 * parWidth), 0

    Set objActuatorBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            0.8 * parValveLength, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objActuatorBody
    
    'Rotate ActuatorBody2
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objActuatorBody.Transform oTransMatrix
    
    'Insert your code for Output 7 (HandwheelSpoke1)
    Dim objHandwheelSpoke As Object
    oStPoint.Set -parHandletoCenter / 2, parHeight1 + (0.5 * parHeight2), _
                                                    -0.4 * parValveCenterToWheelEnd
    oEnPoint.Set -0.7 * parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                    -0.4 * parValveCenterToWheelEnd

    Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            0.1 * parHeight2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheelSpoke

    'Rotate HandwheelSpoke1
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheelSpoke.Transform oTransMatrix
    
    'Insert your code for Output 8 (HandwheelSpoke2)
    oStPoint.Set -0.7 * parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                        -0.4 * parValveCenterToWheelEnd
    oEnPoint.Set -parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                (0.5 * parHandwheelDia - 0.4 * parValveCenterToWheelEnd)

    Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                0.1 * parHeight2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheelSpoke

    'Rotate HandwheelSpoke2
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheelSpoke.Transform oTransMatrix
    
    'Insert your code for Output 9 (HandwheelSpoke3)
    oStPoint.Set -0.7 * parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                        -0.4 * parValveCenterToWheelEnd
    oEnPoint.Set -parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                            -parValveCenterToWheelEnd

    Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                0.1 * parHeight2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheelSpoke

    'Rotate HandwheelSpoke3
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheelSpoke.Transform oTransMatrix
    
    'Insert your code for Output 10 (HandwheelSpoke4)
    oStPoint.Set -0.7 * parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                        -0.4 * parValveCenterToWheelEnd
    oEnPoint.Set -parHandletoCenter, parHeight1 + (0.5 * parHeight2) + _
                                (0.5 * parHandwheelDia), -0.4 * parValveCenterToWheelEnd

    Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            0.1 * parHeight2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheelSpoke

    'Rotate HandwheelSpoke4
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheelSpoke.Transform oTransMatrix
    
    'Insert your code for Output 11 (HandwheelSpoke5)
    oStPoint.Set -0.7 * parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                    -0.4 * parValveCenterToWheelEnd
    oEnPoint.Set -parHandletoCenter, parHeight1 + (0.5 * parHeight2) - _
                            (0.5 * parHandwheelDia), -0.4 * parValveCenterToWheelEnd

    Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                0.1 * parHeight2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheelSpoke

    'Rotate HandwheelSpoke5
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheelSpoke.Transform oTransMatrix

    'Insert your code for Output 12 (Handwheel)
    Dim objHandwheel As Object
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    Dim dMajorRadius As Double
    Dim dMinorRadius As Double
    oCenPoint.Set -parHandletoCenter, parHeight1 + (0.5 * parHeight2), _
                                                    -0.4 * parValveCenterToWheelEnd
    oAxisVec.Set -1, 0, 0
    dMajorRadius = 0.5 * parHandwheelDia
    dMinorRadius = 0.05 * parHandwheelDia

    Set objHandwheel = PlaceTorus(m_OutputColl, oCenPoint, oAxisVec, dMajorRadius, dMinorRadius)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheel
    
    'Rotate Handwheel
    oRotVector.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parOperatorAngle, oRotVector
    objHandwheel.Transform oTransMatrix

    Set oGeomFactory = Nothing
    Set oLineString = Nothing
    Set objOpBody = Nothing
    Set objActuatorBody = Nothing
    Set objHandwheelSpoke = Nothing
    Set objHandwheel = Nothing
    Set oAxisVec = Nothing
    Set oCenPoint = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    '-------------------------------------------------------------------------------------------------
    '==================================================
    'BUILD HVACNOZZLE 1
    '==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double

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
            dThickness(iPortIndex) = NEGLIGIBLE_THICKNESS
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + NEGLIGIBLE_THICKNESS
        End If
    Next iPortIndex

    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing

    iPortIndex = 1
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Dim dDepth As Double
    If CmpDblEqual(GSCADNozzleEntities.Round, 4) Then dDepth = parWidth

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the Nozzle
    Dim x As Double, y As Double, z As Double
    x = -parValveLength / 2
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

    '==================================================
    'BUILD HVACNOZZLE 2
    '==================================================
    iPortIndex = 2
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(2), _
                                dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                                dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                                m_OutputColl.ResourceManager)

    'Position of the Nozzle
    x = parValveLength / 2
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir

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
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
