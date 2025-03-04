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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:  VRK
'   Creation Date: Wednesnday, Jun 25 2008
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   25.Jun.2008     VRK     CR-144628:Provide 30 degree lateral and 60 degree lateral symbols
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI  As Double

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
    Dim oCenterPos      As New AutoMath.DPosition
    Dim iOutput         As Integer

    oCenterPos.Set 0, 0, 0
    
    Dim ObjInsulatedBody1 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedBody2 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort3 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter  As Double
    Dim parInsulationDiameter2  As Double
    Dim parAngle As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parAngle = arrayOfInputs(6)
    
    iOutput = 0

    ' Insert your code for output 1(Insulated Body (Port1 side))
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick), oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing
    
    ' Insert your code for output 2(Insulated Port1)
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set oCenterPos.x - parFace1toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick - parInsulationThickness), _
                oCenterPos.y, oCenterPos.z
    
    'Check to see that header insulation doesn't exceed face to face dimension
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, (parFace1toCenter + parFace2toCenter)) Then
        oEnPoint.Set parFace2toCenter, 0, 0
    End If
    
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then _
                    parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
    'Insert your code for output 3(Insulated Body (Port2 side))
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick, oCenterPos.y, oCenterPos.z
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter2 = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter / 2, _
                                      parInsulationDiameter2 / 2, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing
    
    'Insert your code for output 4(Insulated Port2)
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set oCenterPos.x + parFace2toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick - parInsulationThickness, _
                oCenterPos.y, oCenterPos.z
    
    'Check to see that header insulation doesn't exceed face to face dimension
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, (parFace1toCenter + parFace2toCenter)) Then
        oEnPoint.Set -parFace1toCenter, 0, 0
    End If
    
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then _
                    parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
    'Insert your code for output 5(Insulated Branch (Port3 side))
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set oCenterPos.x + (parFace3toCenter - flangeThick) * Cos(parAngle), _
                oCenterPos.y + (parFace3toCenter - flangeThick) * Sin(parAngle), oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
    
    ' Insert your code for output 6(Insulated Port3)
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set oCenterPos.x + parFace3toCenter * Cos(parAngle), _
                oCenterPos.y + parFace3toCenter * Sin(parAngle), oCenterPos.z
    oEnPoint.Set oCenterPos.x + (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                oCenterPos.y + (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), _
                oCenterPos.z
    
    'Check to see that branch insulation doesn't exceed face to center 3 dimension
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, parFace3toCenter) Then
        oEnPoint.Set 0, 0, 0
    End If
    
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then _
                    parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenterPos = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


