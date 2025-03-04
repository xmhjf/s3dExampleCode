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
'   Author:         VRK
'   Creation Date:  Wednesday, June 06 2007
'   Description:
'       Reducing Tangential Tee 90 Degree
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation" 'Used for error messages
Private Const INCH = 0.0254
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
   
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parHvacShape As Integer
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parWidth2 As Double
    Dim parDepth As Double
    Dim parBDepth As Double
    Dim parDepth2 As Double
    Dim parAngle As Double
    Dim dHeaderLength As Double
    Dim dLength As Double
    Dim dTakeOff As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parDepth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    parWidth2 = arrayOfInputs(7)
    parDepth2 = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
        
    iOutput = 0
    If parHvacShape = Rectangular Then
        dHeaderLength = parBWidth + 2 * INCH
    ElseIf parHvacShape = 4 Then
        dHeaderLength = parBWidth + 2 * INCH
    End If
    
    dTakeOff = 2 * INCH ' 2 Inches
    dLength = 0.5 * dHeaderLength + 0.1 * INCH
    
    If parHvacShape = Rectangular Then
        Dim ObjInsBody1 As Object
        Dim oStPoint   As New AutoMath.DPosition
        Dim oEnPoint   As New AutoMath.DPosition
        oStPoint.Set -(dHeaderLength / 2 + 2 * INCH), -(parDepth / 2 + parInsulationThickness), -(parWidth / 2 + parInsulationThickness)
        oEnPoint.Set dHeaderLength / 2, (parDepth / 2 + parInsulationThickness), (parWidth / 2 + parInsulationThickness)
        Set ObjInsBody1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody1", ObjInsBody1
        Set ObjInsBody1 = Nothing
        
    ElseIf parHvacShape = 4 Then
        oStPoint.Set dHeaderLength / 2, 0, 0
        oEnPoint.Set -(dHeaderLength / 2 + 2 * INCH), 0, 0
        Set ObjInsBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness, True)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody1", ObjInsBody1
        Set ObjInsBody1 = Nothing
    End If
    
    '========================
    'BUILD Body Of Conical
    '========================
    Dim objInsCone As Object
    Dim dHeader2Length As Double
    If parHvacShape = Rectangular Then
        dHeader2Length = parDepth - parDepth2
        If CmpDblGreaterthan(dHeader2Length, 12 * INCH) Then
            dHeader2Length = 12 * INCH
        ElseIf CmpDblLessThan(dHeader2Length, 1 * INCH) Then
            dHeader2Length = 1 * INCH
        End If
        oStPoint.Set dHeaderLength / 2, 0, 0
        Set objInsCone = PlaceTrapezoid(m_OutputColl, oStPoint, _
                                        parDepth + 2 * parInsulationThickness, _
                                        parWidth + 2 * parInsulationThickness, _
                                        parDepth2 + 2 * parInsulationThickness, _
                                        parWidth2 + 2 * parInsulationThickness, _
                                    dHeader2Length, True, PI / 2, PI / 2, 0)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjInsCone", objInsCone
        Set objInsCone = Nothing
    
    ElseIf parHvacShape = 4 Then 'Round
        dHeader2Length = parWidth - parWidth2
        If CmpDblGreaterthan(dHeader2Length, 12 * INCH) Then
            dHeader2Length = 12 * INCH
        ElseIf CmpDblLessThan(dHeader2Length, 1 * INCH) Then
            dHeader2Length = 1 * INCH
        End If
        oStPoint.Set dHeaderLength / 2, 0, 0
        oEnPoint.Set (dHeaderLength / 2 + dHeader2Length), 0, 0
        Set objInsCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                parWidth / 2 + parInsulationThickness, _
                                parWidth2 / 2 + parInsulationThickness, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjInsCone", objInsCone
        Set objInsCone = Nothing
    
    End If
    '========================
    'BUILD Strip Of Header 2
    '========================
    Dim objStrip2 As Object
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength / 2 + dHeader2Length - 0.004, 0, 0
        oEnPoint.Set oStPoint.x + 0.008, 0, 0
        Set objStrip2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    1.06 * parWidth2 + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength / 2 + dHeader2Length - 0.004, -(1.06 * parDepth2 / 2 + parInsulationThickness), _
                                        -(parWidth2 / 2 + 0.06 * parDepth2 / 2 + parInsulationThickness)
        oEnPoint.Set oStPoint.x + 0.008, 1.06 * parDepth2 / 2 + parInsulationThickness, _
                                    (parWidth2 / 2 + 0.06 * parDepth2 / 2 + parInsulationThickness)
        Set objStrip2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "InsStrip2", objStrip2
    Set objStrip2 = Nothing
     
    '===========================
    'BUILD TakeOff 2 Of Lateral
    '===========================
    Dim dTakeOffLen As Double
    Dim objTakeOff As Object
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength / 2 + dHeader2Length, 0, 0
        oEnPoint.Set oStPoint.x + dTakeOffLen, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    1.01 * parWidth2 + 2 * parInsulationThickness, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength / 2 + dHeader2Length, _
                            -(1.01 * parDepth2 / 2 + parInsulationThickness), _
                                -(parWidth2 / 2 + 0.01 * parDepth2 + parInsulationThickness)
        oEnPoint.Set oStPoint.x + dTakeOffLen, _
                            1.01 * parDepth2 / 2 + parInsulationThickness, _
                            (parWidth2 / 2 + 0.01 * parDepth2 / 2 + parInsulationThickness)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsTakeOff2", objTakeOff
    Set objTakeOff = Nothing
    
    
    If parAngle = PI / 2 Then
        If parHvacShape = Rectangular Then
            Dim ObjInsBody2 As Object
            oStPoint.Set 0.5 * parBWidth + parInsulationThickness, parDepth / 2 + parInsulationThickness, 3 * INCH + 0.5 * parWidth
            oEnPoint.Set -(0.5 * parBDepth + parInsulationThickness), parDepth / 2 - parBDepth - parInsulationThickness, 0.5 * parWidth
            Set ObjInsBody2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            iOutput = iOutput + 1
            'Set the output
            m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
            Set ObjInsBody2 = Nothing
    
        ElseIf parHvacShape = 4 Then
            Dim YLength As Double, x As Double, y As Double
            y = parWidth / 2 - parBWidth
            x = Sqr((0.5 * parWidth) ^ 2 - (y * y))
            YLength = x + INCH
            oStPoint.Set 0, parWidth / 2 - parBWidth / 2, 0
            oEnPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 2 * INCH
            Set ObjInsBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth + 2 * parInsulationThickness, True)
            iOutput = iOutput + 1
            'Set the output
            m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
            Set ObjInsBody2 = Nothing
        End If
    
    ElseIf parAngle = 3 * PI / 2 Then
        If parHvacShape = Rectangular Then
                oStPoint.Set 0.5 * parBWidth + parInsulationThickness, -(parDepth / 2 + parInsulationThickness), 3 * INCH + 0.5 * parWidth
                oEnPoint.Set -(0.5 * parBDepth + parInsulationThickness), -(parDepth / 2 - parBDepth - parInsulationThickness), 0.5 * parWidth
                Set ObjInsBody2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
                iOutput = iOutput + 1
                'Set the output
                m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
                Set ObjInsBody2 = Nothing
        
        ElseIf parHvacShape = 4 Then
                y = parWidth / 2 - parBWidth
                x = Sqr((0.5 * parWidth) ^ 2 - (y * y))
                YLength = x + INCH
                oStPoint.Set 0, -(parWidth / 2 - parBWidth / 2), 0
                oEnPoint.Set 0, -(parWidth / 2 - parBWidth / 2), YLength + 2 * INCH
                Set ObjInsBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBWidth + 2 * parInsulationThickness, True)
                iOutput = iOutput + 1
                'Set the output
                m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
                Set ObjInsBody2 = Nothing
        End If
    
   End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub



