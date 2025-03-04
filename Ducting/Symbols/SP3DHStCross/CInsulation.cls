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
'   Creation Date:  Friday,June 15 2007
'   Description:    Definition of HVAC Straight Cross 90 Degree
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
        oStPoint.Set dV / 2, -(parDepth / 2 + parInsulationThickness), -(parWidth / 2 + parInsulationThickness)
        oEnPoint.Set -dV / 2, parDepth / 2 + parInsulationThickness, parWidth / 2 + parInsulationThickness
        Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2, 0, 0
        oEnPoint.Set -dV / 2, 0, 0
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If

    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set parBWidth1 / 2 + parInsulationThickness, _
                                    parBDepth1 / 2 + parInsulationThickness, 0
            oEnPoint.Set -(parBWidth1 / 2 + parInsulationThickness), -(parBDepth1 / 2 + parInsulationThickness), _
                                        (parWidth / 2 + INCH)
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set parBDepth1 / 2 + parInsulationThickness, 0, parBWidth1 / 2 + parInsulationThickness
            oEnPoint.Set -(parBDepth1 / 2 + parInsulationThickness), (parDepth / 2 + INCH), _
                                    -(parBWidth1 / 2 + parInsulationThickness)
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
           oStPoint.Set parBWidth1 / 2 + parInsulationThickness, _
                                        parBDepth1 / 2 + parInsulationThickness, 0
            oEnPoint.Set -(parBWidth1 / 2 + parInsulationThickness), _
                                    -(parBDepth1 / 2 + parInsulationThickness), _
                                    -(parWidth / 2 + INCH)
            Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set objBody = Nothing
        End If

    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, (parWidth / 2 + INCH) * Cos(parAngle - PI), (parWidth / 2 + INCH) * Sin(parAngle - PI)
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth1 + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If
    
     
    If parHvacShape = Rectangular Then
        oStPoint.Set parBDepth / 2 + parInsulationThickness, 0, _
                                    parWidth / 2 + parInsulationThickness
        oEnPoint.Set -(parBDepth / 2 + parInsulationThickness), -(parDepth / 2 + INCH), _
                                    -(parWidth / 2 + parInsulationThickness)
        Set objBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, -(parWidth / 2 + INCH), 0
        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
    End If
    
    'Create the TakeOff 1 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set (dV / 2 + 0.004), (1.01 * parDepth / 2 + parInsulationThickness), _
                            (parWidth / 2 + 0.01 * parDepth + parInsulationThickness)
        oEnPoint.Set (dV / 2 + 0.004 + dTakeOffLen), _
                        -(1.01 * parDepth / 2 + parInsulationThickness), _
                        -(parWidth / 2 + 0.01 * parDepth + parInsulationThickness)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2 + 0.004, 0, 0
        oEnPoint.Set oStPoint.X + dTakeOffLen, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            1.01 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    
    'Create the Strip 1 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set (dV / 2 - 0.004), (1.06 * parDepth / 2 + parInsulationThickness), _
                            (parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
        oEnPoint.Set (dV / 2 + 0.004), _
                            -(1.06 * parDepth / 2 + parInsulationThickness), _
                            -(parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set dV / 2 - 0.004, 0, 0
        oEnPoint.Set oStPoint.X + 0.008, 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                1.06 * parWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If
    
            
    'Create the TakeOff 2 Of Straight Cross
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set (1.01 * parBWidth1 / 2 + parInsulationThickness), _
                        -(parBDepth1 / 2 + 0.01 * parBWidth1 + parInsulationThickness), _
                        (parWidth / 2 + INCH + 0.004)
            oEnPoint.Set -(1.01 * parBWidth1 / 2 + parInsulationThickness), _
                            (parBDepth1 / 2 + 0.01 * parBWidth1 + parInsulationThickness), _
                            (parWidth / 2 + INCH + 0.004 + dTakeOffLen)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set (1.01 * parBDepth1 / 2 + parInsulationThickness), _
                            (parDepth / 2 + INCH + 0.004), _
                            -(parBWidth1 / 2 + 0.01 * parBDepth1 + parInsulationThickness)
            oEnPoint.Set -(1.01 * parBDepth1 / 2 + parInsulationThickness), _
                            (parDepth / 2 + INCH + 0.004 + dTakeOffLen), _
                            (parBWidth1 / 2 + 0.01 * parBDepth1 + parInsulationThickness)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
            oStPoint.Set 1.01 * parBWidth1 / 2 + parInsulationThickness, _
                        -(parBDepth1 / 2 + 0.01 * parBWidth1 + parInsulationThickness), _
                        -(parWidth / 2 + INCH + 0.004)
            oEnPoint.Set -(1.01 * parBWidth1 / 2 + parInsulationThickness), _
                            (parBDepth1 / 2 + 0.01 * parBWidth1 + parInsulationThickness), _
                            -(parWidth / 2 + INCH + 0.004 + dTakeOffLen)
            Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
            Set objTakeOff = Nothing
        End If
        
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, (parWidth / 2 + INCH + 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004) * Sin(parAngle - PI)
        oEnPoint.Set 0, (parWidth / 2 + INCH + 0.004 + dTakeOffLen) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004 + dTakeOffLen) * Sin(parAngle - PI)
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        1.01 * parBWidth1 + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If

    'Create the Strip 2 Of Straight Cross
    If parHvacShape = Rectangular Then
        If parAngle = PI / 2 Then
            oStPoint.Set -(1.06 * parBWidth1 / 2 + parInsulationThickness), _
                            -(parBDepth1 / 2 + 0.06 * parBWidth1 + parInsulationThickness), _
                            (parWidth / 2 + INCH - 0.004)
            oEnPoint.Set (1.06 * parBWidth1 / 2 + parInsulationThickness), _
                            (parBDepth1 / 2 + 0.06 * parBWidth1 + parInsulationThickness), _
                            (parWidth / 2 + INCH + 0.004)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        ElseIf parAngle = PI Then
            oStPoint.Set -(1.06 * parBDepth1 / 2 + parInsulationThickness), _
                            (parDepth / 2 + INCH - 0.004), _
                            -(parBWidth1 / 2 + 0.06 * parBDepth1 + parInsulationThickness)
            oEnPoint.Set (1.06 * parBDepth1 / 2 + parInsulationThickness), _
                            (parDepth / 2 + INCH + 0.004), _
                            (parBWidth1 / 2 + 0.06 * parBDepth1 + parInsulationThickness)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        ElseIf parAngle = 3 * PI / 2 Then
            oStPoint.Set -(1.06 * parBWidth1 / 2 + parInsulationThickness), _
                            -(parBDepth1 / 2 + 0.06 * parBWidth1 + parInsulationThickness), _
                                -(parWidth / 2 + INCH - 0.004)
            oEnPoint.Set (1.06 * parBWidth1 / 2 + parInsulationThickness), _
                            (parBDepth1 / 2 + 0.06 * parBWidth1 + parInsulationThickness), _
                            -(parWidth / 2 + INCH + 0.004)
            Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
            Set objStrip = Nothing
        End If
            
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, (parWidth / 2 + INCH - 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH - 0.004) * Sin(parAngle - PI)
        oEnPoint.Set 0, (parWidth / 2 + INCH + 0.004) * Cos(parAngle - PI), (parWidth / 2 + INCH + 0.004) * Sin(parAngle - PI)
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    (1.06 * parBWidth1 + 2 * parInsulationThickness), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If

    'Create the TakeOff 3 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set (1.01 * parBDepth / 2 + parInsulationThickness), _
                        -(parDepth / 2 + INCH + 0.004), _
                            -(parBWidth / 2 + 0.01 * parBDepth + parInsulationThickness)
        oEnPoint.Set -(1.01 * parBDepth / 2 + parInsulationThickness), _
                        -(parDepth / 2 + INCH + 0.004 + dTakeOffLen), _
                            (parBWidth / 2 + 0.01 * parBDepth + parInsulationThickness)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, -(parWidth / 2 + INCH + 0.004), 0
        oEnPoint.Set 0, -(parWidth / 2 + INCH + 0.004 + dTakeOffLen), 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            (1.01 * parBWidth + 2 * parInsulationThickness), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    

    'Create the Strip 3 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -(1.06 * parBDepth / 2 + parInsulationThickness), _
                    -(parDepth / 2 + INCH - 0.004), _
                    -(parBWidth / 2 + 0.06 * parBDepth + parInsulationThickness)
        oEnPoint.Set (1.06 * parBDepth / 2 + parInsulationThickness), _
                        -(parDepth / 2 + INCH + 0.004), _
                        (parBWidth / 2 + 0.06 * parBDepth + parInsulationThickness)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set 0, -(parWidth / 2 + INCH - 0.004), 0
        oEnPoint.Set 0, (oStPoint.Y - 0.008), 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                1.06 * parBWidth + 2 * parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
   End If
        
        
    'Create the TakeOff 4 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -(dV / 2 + 0.004), (1.01 * parDepth / 2 + parInsulationThickness), _
                        (parWidth / 2 + 0.01 * parDepth + parInsulationThickness)
        oEnPoint.Set -(dV / 2 + 0.004 + dTakeOffLen), _
                        -(1.01 * parDepth / 2 + parInsulationThickness), _
                        -(parWidth / 2 + 0.01 * parDepth + parInsulationThickness)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set -(dV / 2 + 0.004), 0, 0
        oEnPoint.Set -(dV / 2 + 0.004 + dTakeOffLen), 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            (1.01 * parWidth + 2 * parInsulationThickness), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
        Set objTakeOff = Nothing
    End If
    
    'Create the Strip 4 Of Straight Cross
    If parHvacShape = Rectangular Then
        oStPoint.Set -(dV / 2 - 0.004), (1.06 * parDepth / 2 + parInsulationThickness), _
                        (parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
        oEnPoint.Set -(dV / 2 + 0.004), -(1.06 * parDepth / 2 + parInsulationThickness), _
                        -(parWidth / 2 + 0.06 * parDepth / 2 + parInsulationThickness)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    ElseIf parHvacShape = 4 Then  'Round
        oStPoint.Set -(dV / 2 - 0.004), 0, 0
        oEnPoint.Set -(dV / 2 + 0.008), 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        (1.06 * parWidth + 2 * parInsulationThickness), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
        Set objStrip = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub



