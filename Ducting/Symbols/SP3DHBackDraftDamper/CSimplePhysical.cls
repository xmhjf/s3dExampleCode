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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.

'   CSimplePhysical.cls
'   Author:         RRK
'   Creation Date:  Friday, Aug 17 2007
'   Description:
'   This symbol is prepared for Back draft damper that caters to Round and Rectangular shapes as per CR-124909

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   02.SEP.2008     PK      CR-149330 Updated the symbol for the PDB cases 145,146 and 147
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oEditErrors As IJEditErrors
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput                 As Double
    
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parDamperLength         As Double
    Dim parLeverLength          As Double
    Dim parLeverPositionX       As Double
    Dim parLeverPositionY       As Double
    Dim parLeverPositionZ       As Double
    Dim parCounterWeightAngle   As Double
    Dim parHVACShape            As Double
    Dim parInsulationThickness  As Double
    Dim parControlLength        As Double
    Dim parControlWidth         As Double
    Dim parControlHeight        As Double
    
    Dim CornerRadius            As Double
    Dim lHVACShape              As Long
    Dim dControlLength          As Double
    Dim dControlWidth           As Double
    Dim dControlHeight          As Double
    
    CornerRadius = 0

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parWidth = arrayOfInputs(2)
    'parDepth = arrayOfInputs(3)
    'parDamperLength = arrayOfInputs(4)
    'parLeverLength = arrayOfInputs(5)
    'parLeverPositionX = arrayOfInputs(6)
    'parLeverPositionY = arrayOfInputs(7)
    'parLeverPositionZ = arrayOfInputs(8)
    'parCounterWeightAngle = arrayOfInputs(9)
    'parHVACShape = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    'parControlLength = arrayOfInputs(12)
    'parControlWidth = arrayOfInputs(13)
    'parControlHeight = arrayOfInputs(14)
     
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 145               'Default or Back draft damper with counter weight
            parWidth = arrayOfInputs(2)
            parDepth = arrayOfInputs(3)
            parDamperLength = arrayOfInputs(4)
            parLeverLength = arrayOfInputs(5)
            parLeverPositionX = arrayOfInputs(6)
            parLeverPositionY = arrayOfInputs(7)
            parLeverPositionZ = arrayOfInputs(8)
            parCounterWeightAngle = arrayOfInputs(9)
            parHVACShape = arrayOfInputs(10)
            lHVACShape = parHVACShape
        
        Case 146                        'Back draft damper, type A
            parWidth = arrayOfInputs(2)
            parDepth = arrayOfInputs(3)
            parDamperLength = arrayOfInputs(4)
            lHVACShape = 1              'Rectangular Shape
        
        Case 147                        'Back draft damper, type B
            parWidth = arrayOfInputs(2)
            parDepth = arrayOfInputs(3)
            parDamperLength = arrayOfInputs(4)
            parControlLength = arrayOfInputs(12)
            parControlWidth = arrayOfInputs(13)
            parControlHeight = arrayOfInputs(14)
            lHVACShape = 1              'Rectangular Shape
            'Assuming that the Static Pressure Control length will be equal to 0.75 times the damper length
            If CmpDblEqual(parControlLength, 0) Then
                dControlLength = 0.75 * parDamperLength
            Else
                dControlLength = parControlLength
            End If
            'Assuming that the Static Pressure Control width will be equal to 0.2 times the duct depth
            If CmpDblEqual(parControlWidth, 0) Then
                dControlWidth = 0.2 * parDepth
            Else
                dControlWidth = parControlWidth
            End If
            'Assuming that the Static Pressure Control height will be equal to 0.4 times the duct width
            If CmpDblEqual(parControlHeight, 0) Then
                dControlHeight = 0.4 * parWidth
            Else
                dControlHeight = parControlHeight
            End If

    End Select
    
    iOutput = 0
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Select Case lPartDataBasis
    
    Case Is <= 1
        Dim dLeverPosfrmCentX As Double
        Dim dLeverPosfrmCentY As Double
        Dim dLeverPosfrmCentZ As Double
        
        'Length and height of the frame are assumed to be 32mm
        Dim dFrameLength As Double
        dFrameLength = 0.064
        
        dLeverPosfrmCentX = (parDamperLength / 2) - parLeverPositionX
        dLeverPosfrmCentY = (parDepth / 2) - parLeverPositionY
        
        If parHVACShape = 4 Then
            dLeverPosfrmCentZ = Sqr((parWidth / 2) ^ 2 - (dLeverPosfrmCentY + (dFrameLength / 2)) ^ 2) + parLeverPositionZ
        Else
            dLeverPosfrmCentZ = (parWidth / 2) + parLeverPositionZ
        End If

        'Creating Output 1 (Damper Body)
        Dim objDamperBody As Object
        If (parHVACShape = 4) Then 'Round
            oStPoint.Set -parDamperLength / 2, 0, 0
            oEnPoint.Set parDamperLength / 2, 0, 0
            Set objDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
        ElseIf (parHVACShape = Rectangular) Then
            oStPoint.Set -parDamperLength / 2, -parDepth / 2, -parWidth / 2
            oEnPoint.Set parDamperLength / 2, parDepth / 2, parWidth / 2
            Set objDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        End If
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
        Set objDamperBody = Nothing
    
        'Insert your code for output 2 (Frame)
        Dim objFrame As Object
        
        oStPoint.Set -dLeverPosfrmCentX - (dFrameLength / 2), _
                dLeverPosfrmCentY - (dFrameLength / 2), dLeverPosfrmCentZ - parLeverPositionZ
        oEnPoint.Set -dLeverPosfrmCentX + (dFrameLength / 2), _
                dLeverPosfrmCentY + (dFrameLength / 2), dLeverPosfrmCentZ
        
        Set objFrame = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
        'Set the Output
        m_OutputColl.AddOutput "Frame", objFrame
        Set objFrame = Nothing
        
        'Creating Output 3(Lever)
        'Assuming lever dia to be 13mm
        Dim objLever As Object
        
        Dim dLeverDia As Double
        dLeverDia = 0.013
        
        oStPoint.Set -dLeverPosfrmCentX, dLeverPosfrmCentY, dLeverPosfrmCentZ
        oEnPoint.Set oStPoint.x + parLeverLength * Cos(parCounterWeightAngle), _
                oStPoint.y + parLeverLength * Sin(parCounterWeightAngle), oStPoint.z
        
        Set objLever = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dLeverDia, True)
        
        'Set the Output
        m_OutputColl.AddOutput "Lever", objLever
        Set objLever = Nothing
        
        'Creating Counter Weight (Output 4)
        'Assuming Counter Weight dia to be 40mm and counter weight length to 25% of Lever Length
        Dim objCounterWeight As Object
        
        Dim dCounterWtDia As Double
        Dim dCounterWtLength As Double
        
        dCounterWtDia = 0.04
        dCounterWtLength = parLeverLength / 4
        
        oStPoint.Set -dLeverPosfrmCentX + parLeverLength * Cos(parCounterWeightAngle), _
                dLeverPosfrmCentY + parLeverLength * Sin(parCounterWeightAngle), dLeverPosfrmCentZ
        oEnPoint.Set oStPoint.x - (dCounterWtLength * Cos(parCounterWeightAngle)), _
                oStPoint.y - (dCounterWtLength * Sin(parCounterWeightAngle)), oStPoint.z
        
        Set objCounterWeight = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCounterWtDia, True)
        
        'Set the Output
        m_OutputColl.AddOutput "CounterWeight", objCounterWeight
        Set objCounterWeight = Nothing
    
    Case 146, 147
        'Creating Damper Body
        Dim oDamper         As Object
        Dim oGeomFactory    As IngrGeom3D.GeometryFactory
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        oStPoint.Set -parDamperLength / 2, -parDepth / 2, -parWidth / 2
        oEnPoint.Set parDamperLength / 2, parDepth / 2, parWidth / 2
        
        Set oDamper = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oDamper
        Set oDamper = Nothing
        
        If lPartDataBasis = 147 Then
            'Creating the Static Pressure Control body
            Dim oControlString              As Object
            Dim dLineStrPoints(0 To 17)     As Double
            Dim oDir                        As AutoMath.DVector
            Dim oControlBody                As Object
            
            Set oDir = New AutoMath.DVector
            oDir.Set 0, 1, 0
            
            dLineStrPoints(0) = -parDamperLength / 2
            dLineStrPoints(1) = parDepth / 2
            dLineStrPoints(2) = parWidth / 2
            
            dLineStrPoints(3) = -parDamperLength / 2 - 0.3 * dControlLength
            dLineStrPoints(4) = parDepth / 2
            dLineStrPoints(5) = parWidth / 2
            
            dLineStrPoints(6) = -parDamperLength / 2 - dControlLength
            dLineStrPoints(7) = parDepth / 2
            dLineStrPoints(8) = parWidth / 2 - dControlHeight / 2
            
            dLineStrPoints(9) = -parDamperLength / 2 - dControlLength
            dLineStrPoints(10) = parDepth / 2
            dLineStrPoints(11) = parWidth / 2 - dControlHeight
            
            dLineStrPoints(12) = -parDamperLength / 2
            dLineStrPoints(13) = parDepth / 2
            dLineStrPoints(14) = parWidth / 2 - dControlHeight
            
            dLineStrPoints(15) = -parDamperLength / 2
            dLineStrPoints(16) = parDepth / 2
            dLineStrPoints(17) = parWidth / 2
            
            Set oControlString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLineStrPoints)
            
            Set oControlBody = PlaceProjection(m_OutputColl, oControlString, oDir, dControlWidth, True)
            
            m_OutputColl.AddOutput "ControlBody", oControlBody
            Set oControlBody = Nothing
            Set oControlString = Nothing
            Set oDir = Nothing
            
            'Create Control Knob
            Dim oKnob   As Object
            oStPoint.Set -parDamperLength / 2 - dControlLength / 2, _
                         parDepth / 2 + dControlWidth, _
                         parWidth / 2 - 0.6 * dControlHeight
            oEnPoint.Set oStPoint.x, oStPoint.y + 0.25 * dControlWidth, oStPoint.z
            
            Set oKnob = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * dControlHeight, True)
            
            m_OutputColl.AddOutput "Knob", oKnob
            Set oKnob = Nothing
            Set oGeomFactory = Nothing
        End If
                    
    End Select
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
                
    '===================
    'BUILD HVACNOZZLE1
    '===================
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
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                lHVACShape, lEndPrep(1), _
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
                    lHVACShape, lEndPrep(2), dThickness(2), dFlangeWidth(2), lFlowDir(2), _
                    parWidth, parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                    m_OutputColl.ResourceManager)

    
    'Position of the nozzle should be the connect point of the nozzle
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
