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
'   Author:          VRK
'   Creation Date:  Wednesday,June 13 2007
'   Description:    Definition of HVAC Bull Head Tee
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
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
    Dim parSA As Double
    Dim parSC As Double
    Dim parSD As Double
    Dim parInsulationThickness As Double
    Dim parNoOfVanes As Integer

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parNoOfVanes = arrayOfInputs(4)
    parDepth = arrayOfInputs(5)
    parSA = arrayOfInputs(6)
    parSC = arrayOfInputs(7)
    parSD = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0
    
     'Set the Default Shape as Round
    If Not parHvacShape = 4 And Not parHvacShape = Rectangular Then
        parHvacShape = 4
    End If
    iOutput = 0
  
    Dim oStPoint  As New AutoMath.DPosition
    Dim oEnPoint  As New AutoMath.DPosition
    Dim objInsTakeOff As Object
    Dim objInsStrip As Object
    Dim objInsBody As Object
    Dim dTakeOffLen As Double
    Dim dV As Double
    Dim dHalfWidth As Double
    dTakeOffLen = 2 * INCH
    dV = parWidth + 2 * INCH
    dHalfWidth = parWidth / 2
    
   If parHvacShape = Rectangular Then
        'Create the body of Bull Head Tee
        oStPoint.Set dHalfWidth / 2 + parInsulationThickness, -(parDepth / 2 + parInsulationThickness), -(dHalfWidth + parSD)
        oEnPoint.Set -(dHalfWidth / 2 + parInsulationThickness), parDepth / 2 + parInsulationThickness, (dHalfWidth + parSC)
        Set objInsBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
        Set objInsBody = Nothing
        
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, dV / 2
        oEnPoint.Set 0, 0, -dV / 2
        Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
        Set objInsBody = Nothing
    End If
    
    If parHvacShape = Rectangular Then
        oStPoint.Set dHalfWidth / 2, parDepth / 2 + parInsulationThickness, _
                        (dHalfWidth + parInsulationThickness)
        oEnPoint.Set (dHalfWidth / 2 + parSA), -(parDepth / 2 + parInsulationThickness), _
                        -(dHalfWidth + parInsulationThickness)
        Set objInsBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
        Set objInsBody = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, 0
        oEnPoint.Set parWidth / 2 + INCH, 0, 0
        Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
        Set objInsBody = Nothing
    End If
        
    'Create the TakeOff 1 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set (dHalfWidth / 2 + parSA), _
                        (parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness), _
                            1.01 * dHalfWidth + parInsulationThickness
        oEnPoint.Set (dHalfWidth / 2 + parSA + dTakeOffLen), _
                        -(parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness), _
                            -(1.01 * dHalfWidth + parInsulationThickness)
        Set objInsTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set parWidth / 2 + INCH + 0.004, 0, 0
        oEnPoint.Set oStPoint.X + dTakeOffLen, 0, 0
        Set objInsTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    End If
    
    'Create the Strip 1 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set (dHalfWidth / 2 + parSA - 0.004), _
                    (parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), _
                        1.06 * dHalfWidth + parInsulationThickness
        oEnPoint.Set (dHalfWidth / 2 + parSA + 0.004), _
                    -(parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), _
                    -(1.06 * dHalfWidth + parInsulationThickness)
        Set objInsStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set parWidth / 2 + INCH - 0.004, 0, 0
        oEnPoint.Set oStPoint.X + 0.008, 0, 0
        Set objInsStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
    End If

    'Create the TakeOff 2 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set 1.01 * dHalfWidth / 2 + parInsulationThickness, -(parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness), dHalfWidth + parSC + 0.004
        oEnPoint.Set -(1.01 * dHalfWidth / 2 + parInsulationThickness), parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness, (dHalfWidth + parSC + 0.004 + dTakeOffLen)
        Set objInsTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, parWidth / 2 + INCH + 0.004
        oEnPoint.Set 0, 0, parWidth / 2 + INCH + 0.004 + dTakeOffLen
        Set objInsTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    End If

    'Create the Strip 2 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set -(1.06 * dHalfWidth / 2 + parInsulationThickness), -(parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), dHalfWidth + parSC - 0.004
        oEnPoint.Set 1.06 * dHalfWidth / 2 + parInsulationThickness, (parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), dHalfWidth + parSC + 0.004
        Set objInsStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, parWidth / 2 + INCH - 0.004
        oEnPoint.Set 0, 0, parWidth / 2 + INCH + 0.004
        Set objInsStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
    End If


    'Create the TakeOff 3 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set 1.01 * dHalfWidth / 2 + parInsulationThickness, -(parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness), -(dHalfWidth + parSD + 0.004)
        oEnPoint.Set -(1.01 * dHalfWidth / 2 + parInsulationThickness), parDepth / 2 + 0.01 * dHalfWidth + parInsulationThickness, -(dHalfWidth + parSD + 0.004 + dTakeOffLen)
        Set objInsTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, -(parWidth / 2 + INCH + 0.004)
        oEnPoint.Set 0, 0, -(parWidth / 2 + INCH + 0.004 + dTakeOffLen)
        Set objInsTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
        Set objInsTakeOff = Nothing
    End If
    

    'Create the Strip 3 Of Bull Head Tee
    If parHvacShape = Rectangular Then
        oStPoint.Set -(1.06 * dHalfWidth / 2 + parInsulationThickness), -(parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), -(dHalfWidth + parSD - 0.004)
        oEnPoint.Set 1.06 * dHalfWidth / 2 + parInsulationThickness, (parDepth / 2 + 0.06 * dHalfWidth + parInsulationThickness), -(dHalfWidth + parSD + 0.004)
        Set objInsStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, -(parWidth / 2 + INCH - 0.004)
        oEnPoint.Set 0, 0, -(parWidth / 2 + INCH + 0.004)
        Set objInsStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
        Set objInsStrip = Nothing
   End If
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub



