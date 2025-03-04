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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          VRK
'   Creation Date:  Friday,June 15 2007
'   Description:    Definition of HVAC Straight Cross 90 Degree
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   05.03.2008      VRK     TR-CP-138099:Orientation of port 3 of SP3DHStCross symbol is incorrect
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const INCH = 0.0254
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parHvacShape As Integer
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBWidth1 As Double
    Dim parBDepth As Double
    Dim parBDepth1 As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBWidth1 = arrayOfInputs(5)
    parDepth = arrayOfInputs(6)
    parBDepth = arrayOfInputs(7)
    parBDepth1 = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    
    'Set the Default Shape as Round
    If Not parHvacShape = 4 And Not parHvacShape = Rectangular Then
        parHvacShape = 4
    End If
    
    Dim oStPoint  As New AutoMath.DPosition
    Dim oEnPoint  As New AutoMath.DPosition
    Dim objTakeOff As Object
    Dim objStrip As Object
    Dim objBody As Object
    Dim dTakeOffLen As Double
    Dim dV As Double
    dTakeOffLen = 2 * INCH
    
    'Check to see that Branch Width doesn't exceed Width of the duct
    If CmpDblGreaterthan(parBWidth, parWidth) Then
        parBWidth = parWidth
    End If
    
    'Check to see that Branch Depth doesn't exceed Depth of the duct
    If CmpDblGreaterthan(parBDepth, parDepth) Then
        parBDepth = parDepth
    End If
    
    'Check to see that Branch Width1 doesn't exceed Width of the duct
    If CmpDblGreaterthan(parBWidth1, parWidth) Then
        parBWidth1 = parWidth
    End If
    
    'Check to see that Branch Depth1 doesn't exceed Depth of the duct
    If CmpDblGreaterthan(parBDepth1, parDepth) Then
        parBDepth1 = parDepth
    End If
    
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Or parAngle = 3 * PI / 2 Then
            If CmpDblGreaterthan(parBDepth, parBWidth1) Then
                dV = parBDepth + 2 * INCH
            Else
                dV = parBWidth1 + 2 * INCH
            End If
        Else
            dV = parBDepth + 2 * INCH
        End If
    ElseIf parHvacShape = 4 Then  'Round
        dV = parBWidth + 2 * INCH
    End If
    
    If parHvacShape = Rectangular Then
        'Create the body of Straight Cross
        oStPoint.Set dV / 2, -(parDepth / 2), -parWidth / 2
        oEnPoint.Set -dV / 2, parDepth / 2, parWidth / 2
        Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2, 0, 0
        oEnPoint.Set -dV / 2, 0, 0
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If

    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set parBWidth1 / 2, parBDepth1 / 2, 0
            oEnPoint.Set -(parBWidth1 / 2), -parBDepth1 / 2, (parWidth / 2 + INCH)
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set parBDepth1 / 2, 0, parBWidth1 / 2
            oEnPoint.Set -(parBDepth1 / 2), (parDepth / 2 + INCH), -parBWidth1 / 2
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
           oStPoint.Set parBWidth1 / 2, parBDepth1 / 2, 0
            oEnPoint.Set -(parBWidth1 / 2), -parBDepth1 / 2, -(parWidth / 2 + INCH)
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        End If

    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, (parWidth / 2 + INCH) * Cos(parAngle - PI), (parWidth / 2 + INCH) * Sin(parAngle - PI)
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth1, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If
    
     
    If parHvacShape = Rectangular Then
        oStPoint.Set parBDepth / 2, 0, parWidth / 2
        oEnPoint.Set -(parBDepth / 2), -(parDepth / 2 + INCH), -parWidth / 2
        Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, -(parWidth / 2 + INCH), 0
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If
    
    'Create the TakeOff 1 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set (dV / 2 + 0.004), 1.01 * parDepth / 2, (parWidth / 2 + 0.01 * parDepth)
        oEnPoint.Set (dV / 2 + 0.004 + dTakeOffLen), -1.01 * parDepth / 2, -(parWidth / 2 + 0.01 * parDepth)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2 + 0.004, 0, 0
        oEnPoint.Set oStPoint.X + dTakeOffLen, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    
    'Create the Strip 1 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set (dV / 2 - 0.004), 1.06 * parDepth / 2, (parWidth / 2 + 0.06 * parDepth / 2)
        oEnPoint.Set (dV / 2 + 0.004), -1.06 * parDepth / 2, -(parWidth / 2 + 0.06 * parDepth / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2 - 0.004, 0, 0
        oEnPoint.Set oStPoint.X + 0.008, 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If
    
            
    'Create the TakeOff 2 Of Straight Cross
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set 1.01 * parBWidth1 / 2, -(parBDepth1 / 2 + 0.01 * parBWidth1), (parWidth / 2 + INCH + 0.004)
            oEnPoint.Set -1.01 * parBWidth1 / 2, (parBDepth1 / 2 + 0.01 * parBWidth1), (parWidth / 2 + INCH + 0.004 + dTakeOffLen)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set 1.01 * parBDepth1 / 2, (parDepth / 2 + INCH + 0.004), -(parBWidth1 / 2 + 0.01 * parBDepth1)
            oEnPoint.Set -1.01 * parBDepth1 / 2, (parDepth / 2 + INCH + 0.004 + dTakeOffLen), (parBWidth1 / 2 + 0.01 * parBDepth1)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
            oStPoint.Set 1.01 * parBWidth1 / 2, -(parBDepth1 / 2 + 0.01 * parBWidth1), -(parWidth / 2 + INCH + 0.004)
            oEnPoint.Set -1.01 * parBWidth1 / 2, (parBDepth1 / 2 + 0.01 * parBWidth1), -(parWidth / 2 + INCH + 0.004 + dTakeOffLen)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        End If
        
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, (parWidth / 2 + INCH + 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004) * Sin(parAngle - PI)
        oEnPoint.Set 0, (parWidth / 2 + INCH + 0.004 + dTakeOffLen) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004 + dTakeOffLen) * Sin(parAngle - PI)
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parBWidth1, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If

    'Create the Strip 2 Of Straight Cross
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set -1.06 * parBWidth1 / 2, -(parBDepth1 / 2 + 0.06 * parBWidth1), (parWidth / 2 + INCH - 0.004)
            oEnPoint.Set 1.06 * parBWidth1 / 2, (parBDepth1 / 2 + 0.06 * parBWidth1), (parWidth / 2 + INCH + 0.004)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set -1.06 * parBDepth1 / 2, (parDepth / 2 + INCH - 0.004), -(parBWidth1 / 2 + 0.06 * parBDepth1)
            oEnPoint.Set 1.06 * parBDepth1 / 2, (parDepth / 2 + INCH + 0.004), (parBWidth1 / 2 + 0.06 * parBDepth1)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
            oStPoint.Set -1.06 * parBWidth1 / 2, -(parBDepth1 / 2 + 0.06 * parBWidth1), -(parWidth / 2 + INCH - 0.004)
            oEnPoint.Set 1.06 * parBWidth1 / 2, (parBDepth1 / 2 + 0.06 * parBWidth1), -(parWidth / 2 + INCH + 0.004)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        End If
            
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, (parWidth / 2 + INCH - 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH - 0.004) * Sin(parAngle - PI)
        oEnPoint.Set 0, (parWidth / 2 + INCH + 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004) * Sin(parAngle - PI)
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parBWidth1, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If

    'Create the TakeOff 3 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set 1.01 * parBDepth / 2, -(parDepth / 2 + INCH + 0.004), -(parBWidth / 2 + 0.01 * parBDepth)
        oEnPoint.Set -1.01 * parBDepth / 2, -(parDepth / 2 + INCH + 0.004 + dTakeOffLen), parBWidth / 2 + 0.01 * parBDepth
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, -(parWidth / 2 + INCH + 0.004), 0
        oEnPoint.Set 0, -(parWidth / 2 + INCH + 0.004 + dTakeOffLen), 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parBWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    

    'Create the Strip 3 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -1.06 * parBDepth / 2, -(parDepth / 2 + INCH - 0.004), -(parBWidth / 2 + 0.06 * parBDepth)
        oEnPoint.Set 1.06 * parBDepth / 2, -(parDepth / 2 + INCH + 0.004), (parBWidth / 2 + 0.06 * parBDepth)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, -(parWidth / 2 + INCH - 0.004), 0
        oEnPoint.Set 0, (oStPoint.Y - 0.008), 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parBWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
   End If
        
        
    'Create the TakeOff 4 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -(dV / 2 + 0.004), 1.01 * parDepth / 2, (parWidth / 2 + 0.01 * parDepth)
        oEnPoint.Set -(dV / 2 + 0.004 + dTakeOffLen), -1.01 * parDepth / 2, -(parWidth / 2 + 0.01 * parDepth)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set -(dV / 2 + 0.004), 0, 0
        oEnPoint.Set -(dV / 2 + 0.004 + dTakeOffLen), 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    
    'Create the Strip 4 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -(dV / 2 - 0.004), 1.06 * parDepth / 2, (parWidth / 2 + 0.06 * parDepth / 2)
        oEnPoint.Set -(dV / 2 + 0.004), -1.06 * parDepth / 2, -(parWidth / 2 + 0.06 * parDepth / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set -(dV / 2 - 0.004), 0, 0
        oEnPoint.Set -(dV / 2 + 0.008), 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
'-------------------------------------------------------------------------------------------------
'==================
'BUILD HVACNOZZLE1
'==================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim oDir As New AutoMath.DVector
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    
    Dim dWidth As Double
    Dim dDepth As Double
    Dim CornerRadius As Double
    Dim lCSType As Long
    
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 4) As Long
    Dim dThickness(1 To 4) As Double
    Dim dFlangeWidth(1 To 4) As Double
    Dim lFlowDir(1 To 4) As Long
    Dim dPortDepth(1 To 4) As Double
    Dim dCptOffSet(1 To 4) As Double
    Dim dNozzLength(1 To 4) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    dNozzLength(3) = 0.001
    dNozzLength(4) = 0.001

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
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    If parHvacShape = Rectangular Then
         dWidth = parWidth
         dDepth = parDepth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHvacShape = 4 Then 'Round
        dDepth = parWidth
        dWidth = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    If parHvacShape = Rectangular Then
        X = dV / 2
        Y = 0#
        Z = 0#
    ElseIf parHvacShape = 4 Then 'Round
        X = dV / 2
        Y = 0#
        Z = 0#
    End If
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    oDistribPort.SetDirectionVector dir
    If parHvacShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

'==================
'BUILD HVACNOZZLE2
'==================
    iPortIndex = 2
    
    If parHvacShape = Rectangular Then
        dWidth = parBWidth1
        dDepth = parBDepth1
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width2
    If parHvacShape = 4 Then 'Round
        dWidth = parBWidth1
        dDepth = parBWidth1
        lCSType = GSCADNozzleEntities.Round
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the conenct point of the nozzle
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            X = 0
            Y = 0
            Z = parWidth / 2 + INCH
        ElseIf parAngle = PI Then
            X = 0#
            Y = parDepth / 2 + INCH
            Z = 0#
        ElseIf parAngle = 3 * PI / 2 Then
            X = 0#
            Y = 0
            Z = -(parWidth / 2 + INCH)
        End If
            
    ElseIf parHvacShape = 4 Then 'Round
        X = 0#
        Y = (parWidth / 2 + INCH) * Cos(parAngle - PI)
        Z = 0#
    End If
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set 0, 1, 0
    oDistribPort.SetDirectionVector dir
    
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            dir.Set 0, 0, 1
            oDir.Set 0, 1, 0
        ElseIf parAngle = PI Then
            dir.Set 0, 1, 0
            oDir.Set 1, 0, 0
        ElseIf parAngle = 3 * PI / 2 Then
            dir.Set 0, 0, -1
            oDir.Set 0, 1, 0
        End If
        oDistribPort.SetDirectionVector dir
        oDistribPort.SetRadialOrient oDir
    End If

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
'==================
'BUILD HVACNOZZLE3
'==================
    iPortIndex = 3
    If parHvacShape = 4 Then 'Round
        dWidth = parBWidth
        dDepth = parBWidth
        lCSType = GSCADNozzleEntities.Round
    End If
    If parHvacShape = Rectangular Then
        dWidth = parBWidth  'dHalfDepth
        dDepth = parBDepth  'parWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
     
    If parHvacShape = Rectangular Then
        X = 0#
        Y = -(parDepth / 2 + INCH)
        Z = 0#
    ElseIf parHvacShape = 4 Then 'Round
        X = 0#
        Y = -(parWidth / 2 + INCH)
        Z = 0#
    End If
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, -1, 0
    oDistribPort.SetDirectionVector dir
    
    If parHvacShape = Rectangular Then
        Dim oRadialDir As AutoMath.DVector
        Set oRadialDir = New AutoMath.DVector
        oDir.Set 0, -1, 0
        oRadialDir.Set 1, 0, 0
        oDistribPort.SetDirectionVector oDir
        oDistribPort.SetRadialOrient oRadialDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    
'==================
'BUILD HVACNOZZLE4
'==================
    iPortIndex = 4
    If parHvacShape = 4 Then 'Round
        dWidth = parWidth
        dDepth = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If
    If parHvacShape = Rectangular Then
        dWidth = parWidth 'dHalfDepth
        dDepth = parDepth   'parWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(4), _
                                            dThickness(4), dFlangeWidth(4), lFlowDir(4), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz4", dPortDepth(4), dCptOffSet(4), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
     
    If parHvacShape = Rectangular Then
        X = -dV / 2
        Y = 0#
        Z = 0#
    ElseIf parHvacShape = 4 Then 'Round
        X = -dV / 2
        Y = 0
        Z = 0#
    End If
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    oDistribPort.SetDirectionVector dir
    If parHvacShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(4)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oDir = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
