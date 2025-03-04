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
'   Author:          RUK
'   Creation Date:  Thursday, May 24 2007
'   Description:    HVAC Reducing Tee.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages
Private PI       As Double
Private Const INCH = 0.0254

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
    'Origin is taken at the intersection of all port axes. i.e mid point of the Header
    
    'Set the Default Shape as Round
    If Not parHvacShape = 4 And Not parHvacShape = Rectangular Then
        parHvacShape = 4
    End If
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
    
    'Assumed that Header Part 1 length is BWidth + 4"
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
        Set objHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        parWidth + 2 * parInsulationThickness, True)
    End If
    If parHvacShape = Rectangular Then
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
        oStPoint.Set -dHeaderLength1 / 2, -parDepth / 2 - parInsulationThickness, _
                                        -parWidth / 2 - parInsulationThickness
        oEnPoint.Set dHeaderLength1 / 2, parDepth / 2 + parInsulationThickness, _
                                            parWidth / 2 + parInsulationThickness
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
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        1.01 * parWidth + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -dHeaderLength1 / 2, -1.01 * parDepth / 2 - parInsulationThickness, _
                        -(parWidth / 2 + 0.01 * parDepth + parInsulationThickness)
        oEnPoint.Set -(dHeaderLength1 / 2 + dTakeOffLen), _
                        1.01 * parDepth / 2 + parInsulationThickness, _
                        (parWidth / 2 + 0.01 * parDepth / 2 + parInsulationThickness)
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
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        1.06 * parWidth + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -(dHeaderLength1 / 2 - 0.004), _
                        -(1.06 * parDepth / 2 + parInsulationThickness), _
                        -(parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
        oEnPoint.Set -(dHeaderLength1 / 2 + 0.004), _
                        1.06 * parDepth / 2 + parInsulationThickness, _
                        (parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
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
        Set objHeader = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidth / 2 + parInsulationThickness, _
                                    parWidth2 / 2 + parInsulationThickness, True)
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
        Set objHeader = PlaceTrapezoid(m_OutputColl, oStPoint, parDepth + 2 * parInsulationThickness, _
                            parWidth + 2 * parInsulationThickness, parDepth2 + 2 * parInsulationThickness, _
                            parWidth2 + 2 * parInsulationThickness, dHeaderLength2, True, PI / 2, PI / 2, 0)
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
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            1.01 * parWidth2 + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2, _
                        -(1.01 * parDepth2 / 2 + parInsulationThickness), _
                        -(parWidth2 / 2 + 0.01 * parDepth2 + parInsulationThickness)
        oEnPoint.Set oStPoint.X + dTakeOffLen, _
                        1.01 * parDepth2 / 2 + parInsulationThickness, _
                        (parWidth2 / 2 + 0.01 * parDepth2 / 2 + parInsulationThickness)
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
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        1.06 * parWidth2 + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2 - 0.004, _
                    -(1.06 * parDepth2 / 2 + parInsulationThickness), _
                    -(parWidth2 / 2 + 0.06 * parDepth2 / 2 + parInsulationThickness)
        oEnPoint.Set oStPoint.X + 0.008, _
                    1.06 * parDepth2 / 2 + parInsulationThickness, _
                        (parWidth2 / 2 + 0.06 * parDepth2 / 2 + parInsulationThickness)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
'Create the Branch
    'For Part Databasis = 1 create Cylider
    'For Part DataBasis = 5 create cone
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
            oEnPoint.Set oStPoint.X, -(parDepth / 2 + 1 * INCH), oStPoint.Z
            Set objBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBranchDia + 2 * parInsulationThickness, True)
        End If
        If lPartdatabasis = 20 Then
            'Assumed that Branch Depth does not exceeds Width - 2"
            If CmpDblGreaterthan(parBWidth, parWidth - 2 * INCH) Then
                dBranchTopDia = parWidth - 2 * INCH
            Else
                dBranchTopDia = parBWidth
            End If
            dBranchBottomDia = dBranchTopDia + 1 * INCH
            dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchBottomDia / 2) ^ 2)
            oStPoint.Set 0, -dBranchStartY, 0
            oEnPoint.Set oStPoint.X, -(parWidth / 2 + 4 * INCH), oStPoint.Z
            Set objBranch = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dBranchBottomDia / 2 + parInsulationThickness, _
                                        dBranchTopDia / 2 + parInsulationThickness, True)
        End If
    End If
    
    If parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            'Assumed that Branch Depth does not exceeds Depth"
            If CmpDblGreaterthan(parBDepth, parDepth) Then
                dBranchDepth = parDepth
            Else
                dBranchDepth = parBDepth
            End If
            dBranchWidth = parBWidth
            oStPoint.Set -(dBranchDepth / 2 + parInsulationThickness), parDepth / 2, _
                            -(dBranchWidth / 2 + parInsulationThickness)
            oEnPoint.Set dBranchDepth / 2 + parInsulationThickness, (parDepth / 2 + 1 * INCH), _
                            dBranchWidth / 2 + parInsulationThickness
            Set objBranch = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        End If
        If lPartdatabasis = 20 Then
            'Assumed that Bottom Width does not exceeds Depth - 2"
            If CmpDblGreaterthan(parBDepth, parDepth - 2 * INCH) Then
                dBranchTopDepth = parDepth - 2 * INCH
            Else
                dBranchTopDepth = parBDepth
            End If
            dBranchBottomDepth = dBranchTopDepth + 1 * INCH
            dBranchWidth = parBWidth
            oStPoint.Set 0, parDepth / 2, 0
            Set objBranch = PlaceTrapezoid(m_OutputColl, oStPoint, dBranchBottomDepth + 2 * parInsulationThickness, _
                                    dBranchWidth + 1 * INCH + 2 * parInsulationThickness, _
                                    dBranchTopDepth + 2 * parInsulationThickness, _
                                    dBranchWidth + 2 * parInsulationThickness, _
                                    4 * INCH, True, -PI / 2, 0, 0)
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
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * dBTakeOffDia + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        Dim dBTDepth As Double
        dBTDepth = dBranchDepth + dBranchTopDepth
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set -(1.01 * dBTDepth / 2 + parInsulationThickness), _
                        parDepth / 2 + 1 * INCH, _
                        -(dBranchWidth / 2 + 0.01 * dBTDepth + parInsulationThickness)
            oEnPoint.Set 1.01 * dBTDepth / 2 + parInsulationThickness, _
                    (parDepth / 2 + 1 * INCH + dTakeOffLen), _
                    dBranchWidth / 2 + 0.01 * dBTDepth + parInsulationThickness
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set -(1.01 * dBTDepth / 2 + parInsulationThickness), _
                        parDepth / 2 + 4 * INCH, _
                        -(dBranchWidth / 2 + 0.01 * dBTDepth + parInsulationThickness)
            oEnPoint.Set 1.01 * dBTDepth / 2 + parInsulationThickness, _
                    (parDepth / 2 + 4 * INCH + dTakeOffLen), _
                    dBranchWidth / 2 + 0.01 * dBTDepth + parInsulationThickness
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
            oEnPoint.Set 0, -(parWidth / 2 + 1 * INCH + 0.008), 0
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set 0, -(parWidth / 2 + 4 * INCH - 0.004), 0
            oEnPoint.Set 0, -(parWidth / 2 + 4 * INCH + 0.008), 0
        End If
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                         1.06 * dBStripDia + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 15 Then
            oStPoint.Set -(1.06 * dBTDepth / 2 + parInsulationThickness), _
                    parDepth / 2 + 1 * INCH - 0.004, _
                    -(dBranchWidth / 2 + 0.06 * dBTDepth + parInsulationThickness)
            oEnPoint.Set 1.06 * dBTDepth / 2 + parInsulationThickness, _
                    parDepth / 2 + 1 * INCH + 0.004, _
                    (dBranchWidth / 2 + 0.06 * dBTDepth + parInsulationThickness)
        ElseIf lPartdatabasis = 20 Then
            oStPoint.Set -(1.06 * dBTDepth / 2 + parInsulationThickness), _
                    parDepth / 2 + 4 * INCH - 0.004, _
                    -(dBranchWidth / 2 + 0.06 * dBTDepth + parInsulationThickness)
            oEnPoint.Set 1.06 * dBTDepth / 2 + parInsulationThickness, _
                    parDepth / 2 + 4 * INCH + 0.004, _
                    (dBranchWidth / 2 + 0.06 * dBTDepth + parInsulationThickness)
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
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


