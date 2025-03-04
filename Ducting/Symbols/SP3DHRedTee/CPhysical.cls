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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          RUK
'   Creation Date:  Thursday, May 24 2007
'   Description:    HVAC Reducing Tee.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----        ------------------
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
    Dim parWidth2 As Double
    Dim parDepth2 As Double
    Dim parBDepth As Double
    Dim parBWidth As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parWidth2 = arrayOfInputs(5)
    parDepth2 = arrayOfInputs(6)
    parBWidth = arrayOfInputs(7)
    parBDepth = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0
    If CmpDblEqual(parDepth, LINEAR_TOLERANCE) Then
        parDepth = parWidth
    End If
    If CmpDblEqual(parDepth2, LINEAR_TOLERANCE) Then
        parDepth2 = parWidth2
    End If
    
    'Set the Default Shape as Round
    If Not parHvacShape = 4 And Not parHvacShape = Rectangular Then
        parHvacShape = 4
    End If
    'Origin is taken at the intersection of all port axes. i.e mid point of the Header
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
'Create the Header Part 1
    Dim objHeader As Object
    Dim dHeaderLength1 As Double
    
    If parHvacShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            'Assumed that Header Part 1 length is BWidth + 2"
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dHeaderLength1 = parWidth + 2 * INCH
            Else
            dHeaderLength1 = parBWidth + 2 * INCH
            End If
        ElseIf lPartdatabasis = 20 Then
            'Assumed that Header Part 1 length is BWidth + 4"
            'But BWidth <= Width - 2"
            If CmpDblGreaterthan(parBWidth, parWidth - 2 * INCH) Then
                dHeaderLength1 = parWidth - 2 * INCH + 4 * INCH
            Else
                dHeaderLength1 = parBWidth + 4 * INCH
            End If
        End If
        oStPoint.Set -dHeaderLength1 / 2, 0, 0
        oEnPoint.Set dHeaderLength1 / 2, 0, 0
        Set objHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
    ElseIf parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            If CmpDblGreaterthan(parBDepth, parDepth) Then
                dHeaderLength1 = parDepth + 2 * INCH
            Else
                dHeaderLength1 = parBDepth + 2 * INCH
            End If
        ElseIf lPartdatabasis = 20 Then
            If CmpDblGreaterthan(parBDepth, parDepth - 2 * INCH) Then
                dHeaderLength1 = parDepth - 2 * INCH + 4 * INCH
            Else
                dHeaderLength1 = parBDepth + 4 * INCH
            End If
        End If
        oStPoint.Set -dHeaderLength1 / 2, -parDepth / 2, -parWidth / 2
        oEnPoint.Set dHeaderLength1 / 2, parDepth / 2, parWidth / 2
        Set objHeader = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeader
    Set objHeader = Nothing

'Create the TakeOff 1 Of Header
    Dim objTakeOff As Object
    Dim dTakeOffLen As Double
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
        oStPoint.Set -dHeaderLength1 / 2, 0, 0
        oEnPoint.Set -(dHeaderLength1 / 2 + dTakeOffLen), 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -dHeaderLength1 / 2, -1.01 * parDepth / 2, -(parWidth / 2 + 0.01 * parDepth)
        oEnPoint.Set -(dHeaderLength1 / 2 + dTakeOffLen), 1.01 * parDepth / 2, (parWidth / 2 + 0.01 * parDepth / 2)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip 1 Of Header
    Dim objStrip As Object
    If parHvacShape = 4 Then 'Round
        oStPoint.Set -(dHeaderLength1 / 2 - 0.004), 0, 0
        oEnPoint.Set -(dHeaderLength1 / 2 + 0.004), 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -(dHeaderLength1 / 2 - 0.004), -1.06 * parDepth / 2, -(parWidth / 2 + 0.06 * parDepth / 2)
        oEnPoint.Set -(dHeaderLength1 / 2 + 0.004), 1.06 * parDepth / 2, (parWidth / 2 + 0.06 * parDepth / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
'Create the Header Part 2
    Dim dHeaderLength2 As Double
    If parHvacShape = 4 Then 'Round
        'Assumed that Header Part 2 length is Width - Width2.
        'Header Part 2 Length must be in between 1" and 12"
        dHeaderLength2 = parWidth - parWidth2
        If CmpDblLessThan(dHeaderLength2, 1 * INCH) Then
            dHeaderLength2 = 1 * INCH
        End If
        If CmpDblGreaterthan(dHeaderLength2, 12 * INCH) Then
            dHeaderLength2 = 12 * INCH
        End If
        oStPoint.Set dHeaderLength1 / 2, 0, 0
        oEnPoint.Set oStPoint.X + dHeaderLength2, oStPoint.Y, oStPoint.Z
        Set objHeader = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parWidth / 2, parWidth2 / 2, True)
    End If
    If parHvacShape = Rectangular Then
        'Assumed that Header Part 2 length is Depth - Depth2.
        'Header Part 2 Length must be in between 1" and 12"
        dHeaderLength2 = parDepth - parDepth2
        If CmpDblLessThan(dHeaderLength2, 1 * INCH) Then
            dHeaderLength2 = 1 * INCH
        End If
        If CmpDblGreaterthan(dHeaderLength2, 12 * INCH) Then
            dHeaderLength2 = 12 * INCH
        End If
        oStPoint.Set dHeaderLength1 / 2, 0, 0
        Set objHeader = PlaceTrapezoid(m_OutputColl, oStPoint, parDepth, parWidth, parDepth2, parWidth2, _
                                    dHeaderLength2, True, PI / 2, PI / 2, 0)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeader
    Set objHeader = Nothing
    
'Create the TakeOff 2 Of Header
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2, 0, 0
        oEnPoint.Set oStPoint.X + dTakeOffLen, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2, -1.01 * parDepth2 / 2, -(parWidth2 / 2 + 0.01 * parDepth2)
        oEnPoint.Set oStPoint.X + dTakeOffLen, 1.01 * parDepth2 / 2, (parWidth2 / 2 + 0.01 * parDepth2 / 2)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip 2 Of Header
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2 - 0.004, 0, 0
        oEnPoint.Set oStPoint.X + 0.008, 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2 - 0.004, -1.06 * parDepth2 / 2, -(parWidth2 / 2 + 0.06 * parDepth2 / 2)
        oEnPoint.Set oStPoint.X + 0.008, 1.06 * parDepth2 / 2, (parWidth2 / 2 + 0.06 * parDepth2 / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
'Create the Branch
    'For Part Databasis = 1 create Cylider
    'For PartDataBasis = 5 create cone
    'Assumed that Branch is at the middle of the Header Part1.
    Dim objBranch As Object
    Dim dBranchDia As Double            'Used for Round Shape and PartdataBasis = 1
    Dim dBranchDepth As Double          'Used for Rectangular Shape and PartDataBasis = 1
    Dim dBranchWidth As Double          'Used for Rectangular Shape and PartDataBasis = 1
    Dim dBranchBottomDia As Double      'Used for Round Shape and PartDataBasis = 5
    Dim dBranchTopDia As Double         'Used for Round Shape and PartDataBasis = 5
    Dim dBranchBottomDepth As Double    'Used for Rectangular Shape and PartDataBasis = 5
    Dim dBranchTopDepth As Double       'Used for Rectangular Shape and PartDataBasis = 5
    
    If parHvacShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            'Assumed that Branch Diameter(BWidth) does not exceeds Width"
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dBranchDia = parWidth
            Else
                dBranchDia = parBWidth
            End If
            Dim dBranchStartY As Double
            dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchDia / 2) ^ 2)
            oStPoint.Set 0, -dBranchStartY, 0
            oEnPoint.Set oStPoint.X, -(parWidth / 2 + 1 * INCH), oStPoint.Z
            Set objBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBranchDia, True)
        End If
        If lPartdatabasis = 20 Then
            'Assumed that Branch Width(BWidth) does not exceeds Width - 2"
            If CmpDblGreaterthan(parBWidth, parWidth - 2 * INCH) Then
                dBranchTopDia = parWidth - 2 * INCH
            Else
                dBranchTopDia = parBWidth
            End If
            dBranchBottomDia = dBranchTopDia + 1 * INCH
            dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchBottomDia / 2) ^ 2)
            oStPoint.Set 0, -dBranchStartY, 0
            oEnPoint.Set oStPoint.X, -(parWidth / 2 + 4 * INCH), oStPoint.Z
            Set objBranch = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dBranchBottomDia / 2, dBranchTopDia / 2, True)
        End If
    End If
    
    If parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            'Assumed that Branch Depth(BDepth) does not exceeds Width"
            If CmpDblGreaterthan(parBDepth, parDepth) Then
                dBranchDepth = parDepth
            Else
                dBranchDepth = parBDepth
            End If
            dBranchWidth = parBWidth
            oStPoint.Set -dBranchDepth / 2, parDepth / 2, -dBranchWidth / 2
            oEnPoint.Set dBranchDepth / 2, (parDepth / 2 + 1 * INCH), dBranchWidth / 2
            Set objBranch = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        End If
        If lPartdatabasis = 20 Then
            'Assumed that Branch Depth does not exceeds Width - 2"
            If CmpDblGreaterthan(parBDepth, parDepth - 2 * INCH) Then
                dBranchTopDepth = parDepth - 2 * INCH
            Else
                dBranchTopDepth = parBDepth
            End If
            dBranchBottomDepth = dBranchTopDepth + 1 * INCH
            dBranchWidth = parBWidth
            oStPoint.Set 0, parDepth / 2, 0
            Set objBranch = PlaceTrapezoid(m_OutputColl, oStPoint, dBranchBottomDepth, dBranchWidth + 1 * INCH, _
                                        dBranchTopDepth, dBranchWidth, 4 * INCH, True, -PI / 2, 0, 0)
        End If
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranch
    Set objBranch = Nothing
    
'Create the TakeOff Of Branch
    If parHvacShape = 4 Then 'Round
        Dim dBTakeOffDia As Double
        dBTakeOffDia = dBranchDia + dBranchTopDia
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set 0, -(parWidth / 2 + 1 * INCH), 0
            oEnPoint.Set oStPoint.X, oStPoint.Y - dTakeOffLen, 0
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set 0, -(parWidth / 2 + 4 * INCH), 0
            oEnPoint.Set oStPoint.X, oStPoint.Y - dTakeOffLen, 0
        End If
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * dBTakeOffDia, True)
    ElseIf parHvacShape = Rectangular Then
        Dim dBTDepth As Double
        dBTDepth = dBranchDepth + dBranchTopDepth
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set -1.01 * dBTDepth / 2, parDepth / 2 + 1 * INCH, -(dBranchWidth / 2 + 0.01 * dBTDepth)
            oEnPoint.Set 1.01 * dBTDepth / 2, (parDepth / 2 + 1 * INCH + dTakeOffLen), dBranchWidth / 2 + 0.01 * dBTDepth
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set -1.01 * dBTDepth / 2, parDepth / 2 + 4 * INCH, -(dBranchWidth / 2 + 0.01 * dBTDepth)
            oEnPoint.Set 1.01 * dBTDepth / 2, (parDepth / 2 + 4 * INCH + dTakeOffLen), dBranchWidth / 2 + 0.01 * dBTDepth
        End If
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip Of Branch
    If parHvacShape = 4 Then 'Round
        Dim dBStripDia As Double
        dBStripDia = dBranchDia + dBranchTopDia
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set 0, -(parWidth / 2 + 1 * INCH - 0.004), 0
            oEnPoint.Set 0, -(parWidth / 2 + 1 * INCH + 0.004), 0
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set 0, -(parWidth / 2 + 4 * INCH - 0.004), 0
            oEnPoint.Set 0, -(parWidth / 2 + 4 * INCH + 0.004), 0
        End If
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * dBStripDia, True)
    ElseIf parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set -1.06 * dBTDepth / 2, parDepth / 2 + 1 * INCH - 0.004, -(dBranchWidth / 2 + 0.06 * dBTDepth)
            oEnPoint.Set 1.06 * dBTDepth / 2, parDepth / 2 + 1 * INCH + 0.004, (dBranchWidth / 2 + 0.06 * dBTDepth)
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set -1.06 * dBTDepth / 2, parDepth / 2 + 4 * INCH - 0.004, -(dBranchWidth / 2 + 0.06 * dBTDepth)
            oEnPoint.Set 1.06 * dBTDepth / 2, parDepth / 2 + 4 * INCH + 0.004, (dBranchWidth / 2 + 0.06 * dBTDepth)
        End If
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
'-------------------------------------------------------------------------------------------------
'==================================================
'BUILD HVACNOZZLE1 at the begining of the Header
'==================================================
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
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    Dim dNozzLength(1 To 3) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    dNozzLength(3) = 0.001
    
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
    If parHvacShape = Rectangular Then
        dDepth = parDepth
        dWidth = parWidth
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
    X = -dHeaderLength1 / 2
    Y = 0#
    Z = 0#
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
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

'==================================================
'BUILD HVACNOZZLE2 at the end of the Reducing Tee
'==================================================
    iPortIndex = 2
    If parHvacShape = Rectangular Then
        dDepth = parDepth2
        dWidth = parWidth2
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width2
    If parHvacShape = 4 Then 'Round
        dDepth = parWidth2
        dWidth = parWidth2
        lCSType = GSCADNozzleEntities.Round
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    X = dHeaderLength1 / 2 + dHeaderLength2
    Y = 0#
    Z = 0#
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
    iNozzle.Length = dNozzLength(2)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
'============================================
'BUILD HVACNOZZLE3 at the end of the Branch
'============================================
    iPortIndex = 3
    If parHvacShape = Rectangular Then
        'if PartDataBasis = 5 then dBranchTopDepth becomes zero.
        'if PartDataBasis = 10 then dBranchDepth becomes zero.
        dDepth = dBranchDepth + dBranchTopDepth
        dWidth = dBranchWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    If parHvacShape = 4 Then 'Round
        'if PartDataBasis = 5 then dBranchTopDia becomes zero.
        'if PartDataBasis = 10 then dBranchDia becomes zero.
        dDepth = dBranchDia + dBranchTopDia
        dWidth = dBranchDia + dBranchTopDia
        lCSType = GSCADNozzleEntities.Round
    End If
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    ' Y Should be positive for Rectangular and negitive for Round
    X = 0#
    If parHvacShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            Y = -((parDepth / 2) + 1 * INCH)
        ElseIf lPartdatabasis = 20 Then
            Y = -((parDepth / 2) + 4 * INCH)
        End If
    ElseIf parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            Y = ((parDepth / 2) + 1 * INCH)
        ElseIf lPartdatabasis = 20 Then
            Y = ((parDepth / 2) + 4 * INCH)
        End If
    End If
    Z = 0#
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
        oDir.Set 0, 1, 0
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
    Set NozzleFactory = Nothing
    Set oDir = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

