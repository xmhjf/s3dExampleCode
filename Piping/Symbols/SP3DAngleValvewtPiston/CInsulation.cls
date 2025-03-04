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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Jun 10 2008
'   Description:
'    The symbol is prepared based on 120 deg angle valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-144044  Provide a 120 deg angle valve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Integer
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim parFace1toFace2 As Double
    Dim parFace2toCenter As Double
    Dim parValveHeight As Double
    Dim parValveDiamter As Double
    Dim ObjInsBodyCylinder1 As Object
    Dim ObjInsBodyCylinder2 As Object
    Dim objInsFlangePort1 As Object
    Dim objInsFlangePort2 As Object
    
    
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toFace2 = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parValveHeight = arrayOfInputs(4)
    parValveDiamter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
 
    iOutput = 0

    'Insert your code for output 1(Cylinder Port 2 of Body)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'Anti-Clock wise angle measurement is taken to be positive
    Dim dValveAng As Double
    dValveAng = 4 * PI / 3
    Dim dPort1XOffset As Double
    Dim dPort2XOffset As Double
    
    dPort1XOffset = parFace1toFace2 - parFace2toCenter * Tan(PI / 6)
    dPort2XOffset = parFace2toCenter * Tan(PI / 6)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set dPort2XOffset, -parFace2toCenter, 0
    
    Set ObjInsBodyCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodyCylinder1
    Set ObjInsBodyCylinder1 = Nothing
    
    'Insert your code for output 2(Cylinder Port 1 of Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    
    oStPoint.Set -dPort1XOffset, 0, 0
    oEnPoint.Set (parValveHeight - dPort1XOffset), 0, 0
    
    Set ObjInsBodyCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parValveDiamter + 2 * parInsulationThickness, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodyCylinder2
    Set ObjInsBodyCylinder2 = Nothing

    ' Insert your code for output 3 (Insulated Flange at Port 1)
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
    oStPoint.Set -dPort1XOffset, 0, 0
    oEnPoint.Set -dPort1XOffset + flangeThick + parInsulationThickness, 0, 0
    Dim dFlangeInsulationDiameter As Double
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Else
        dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
    End If

    Set objInsFlangePort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        dFlangeInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort1
    Set objInsFlangePort1 = Nothing
    
    ' Insert your code for output 4 (Insulated Flange at Port 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
    oStPoint.Set dPort2XOffset, -parFace2toCenter, 0
    oEnPoint.Set dPort2XOffset - (flangeThick + parInsulationThickness) * Sin(PI / 6), _
                                -parFace2toCenter + (flangeThick + parInsulationThickness) * Cos(PI / 6), 0
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Else
        dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
    End If

    Set objInsFlangePort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        dFlangeInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort2
    Set objInsFlangePort2 = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
