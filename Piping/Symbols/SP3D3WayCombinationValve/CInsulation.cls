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
'   Author:          RRK
'   Creation Date:  Friday, Aug 31 2007
'   Description:    This symbol is prepared for 3 way diverter combination valve
'                   of Tyco Flow Control as per CR-126718
'   Source:         http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages


Private Sub Class_Initialize()

'''

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim flangeDiam3      As Double
    Dim depth3           As Double
    
    Dim iOutput     As Double

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parPorttoPortOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double

    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
 
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parPorttoPortOffset = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    iOutput = 0
    
    If CmpDblLessThanOrEqualTo(parFace1toCenter, 0) Then
        parFace1toCenter = parFace3toCenter
    End If
    
' Insert your code for output 11(Insulation for Left Valve Body)
    Dim objInsLeftValveBody As Object
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    
    Dim dInsulationDia As Double
    dInsulationDia = pipeDiam1 + (2 * parInsulationThickness)
    
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set 0, 0, 0
    
    Set objInsLeftValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLeftValveBody
    Set objInsLeftValveBody = Nothing
    
' Insert your code for output 12(Insulation for Vertical Valve Body)
    Dim objInsVertValveBody As Object
    
    'Assumption: Height of the cylinder(valve body 2) is assumed to be twice the Face 2 to center
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
    
    oStPoint.Set 0, -parFace2toCenter, 0
    oEnPoint.Set 0, parPorttoPortOffset + (pipeDiam3 / 2), 0
        
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    dInsulationDia = pipeDiam2 + (2 * parInsulationThickness)

    Set objInsVertValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsVertValveBody
    Set objInsVertValveBody = Nothing
    
' Insert your code for output 13(Insulation Body for two Flanges)
    Dim objInsFlange As Object
    
    'Assumptions:     1.)The diameter of flanges is assumed to be 1.5 times the smaller of
    '                 parFace1toCenter and parFace3toCenter
    '                 2.)Total Flange Thickness of twp flanges is calculated as:
    '                   parPorttoPortOffset - (pipeDiam1 / 2 + pipeDiam3 / 2 + 0.1 * parPorttoPortOffset)
                      
    Dim dTotalFlangeThk As Double
    Dim dFlangeDia As Double
    
    dTotalFlangeThk = parPorttoPortOffset - (pipeDiam1 / 2 + pipeDiam3 / 2 + 0.1 * parPorttoPortOffset)
    'Check in case dTotalFlangeThk becomes zero or negative
    If CmpDblLessThanOrEqualTo(dTotalFlangeThk, 0) Then dTotalFlangeThk = 0.02
    
    dFlangeDia = 1.5 * SmallerDim(parFace3toCenter, parFace1toCenter)
    
    oStPoint.Set 0, (pipeDiam1 / 2) + (0.05 * parPorttoPortOffset) - parInsulationThickness, 0
    oEnPoint.Set 0, oStPoint.y + dTotalFlangeThk + parInsulationThickness, 0
    
    dInsulationDia = dFlangeDia + (2 * parInsulationThickness)
    Set objInsFlange = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlange
    Set objInsFlange = Nothing
        
' Insert your code for output 14(Insulation for Right Valve Body)
    Dim objInsRightValveBody As Object
    
    oStPoint.Set parFace3toCenter, parPorttoPortOffset, 0
    oEnPoint.Set 0, parPorttoPortOffset, 0
    
    dInsulationDia = pipeDiam3 + (2 * parInsulationThickness)
    
    Set objInsRightValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsRightValveBody
    Set objInsRightValveBody = Nothing
        
' Insert your code for output 15(Insulated Port 1)
    
    Dim objInsPort1 As Object
    
    Dim dFlangeThick As Double
    dFlangeThick = flangeThick1
    
    If CmpDblLessThanOrEqualTo(dFlangeThick, 0) Then
        dFlangeThick = 0.001
    End If
    
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set -parFace1toCenter + dFlangeThick + parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
        dInsulationDia = flangeDiam1 + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam1 + (2 * parInsulationThickness)
    End If
    
    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing
    
' Insert your code for output 16(Insulated Port 2)
    
    Dim objInsPort2 As Object
    
    dFlangeThick = flangeThick2
    
    If CmpDblLessThanOrEqualTo(dFlangeThick, 0) Then
        dFlangeThick = 0.001
    End If
    
    oStPoint.Set 0, -parFace2toCenter, 0
    oEnPoint.Set 0, -parFace2toCenter + dFlangeThick + parInsulationThickness, 0
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dInsulationDia = flangeDiam2 + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam2 + (2 * parInsulationThickness)
    End If
    
    Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort2
    Set objInsPort2 = Nothing

' Insert your code for output 17(Insulated Port 3)
    
    Dim objInsPort3 As Object
    
    dFlangeThick = flangeThick3
    
    If CmpDblLessThanOrEqualTo(dFlangeThick, 0) Then
        dFlangeThick = 0.001
    End If
    
    oStPoint.Set parFace3toCenter, parPorttoPortOffset, 0
    oEnPoint.Set parFace3toCenter - dFlangeThick - parInsulationThickness, _
                parPorttoPortOffset, 0
    
    If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then
        dInsulationDia = flangeDiam3 + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam3 + (2 * parInsulationThickness)
    End If
    
    Set objInsPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort3
    Set objInsPort3 = Nothing
    
    
    Exit Sub

ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Sub
