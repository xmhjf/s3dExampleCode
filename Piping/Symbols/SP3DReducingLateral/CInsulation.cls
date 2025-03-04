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
'   Author:          KKC
'   Creation Date:  Tuesday, January 28 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for Insulation aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   23.JAN.2007     KKC       Created :CR-112137 Provide additional symbols required for Victaulic General Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
    Dim oCenterPos       As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    
    oCenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
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

    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter  As Double
    Dim parInsulationDiameter2  As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

' Insert your code for output 1(Insulated Body (Port1 side))
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick), oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing
    
' Insert your code for output 2(Insulated Port1)
    oStPoint.Set oCenterPos.x - parFace1toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick - parInsulationThickness), _
                oCenterPos.y, oCenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 3(Insulated Body (Port2 side))
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick, oCenterPos.y, oCenterPos.z
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter2 = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBody2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter / 2, _
                                      parInsulationDiameter2 / 2, True)
                                      
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing
    
' Insert your code for output 4(Insulated Port2)
    oStPoint.Set oCenterPos.x + parFace2toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick - parInsulationThickness, _
                oCenterPos.y, oCenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
' Insert your code for output 5(Insulated Branch (Port3 side))
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    parFace3toCenter = parFace2toCenter
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set oCenterPos.x + (parFace3toCenter - flangeThick) * Cos(PI / 4), _
                oCenterPos.y + (parFace3toCenter - flangeThick) * Sin(PI / 4), oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
    
' Insert your code for output 6(Insulated Port3)
    oStPoint.Set oCenterPos.x + parFace3toCenter * Cos(PI / 4), _
                oCenterPos.y + parFace3toCenter * Sin(PI / 4), oCenterPos.z
    oEnPoint.Set oCenterPos.x + (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 4), _
                oCenterPos.y + (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 4), _
                oCenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsDiameterPort, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set oCenterPos = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
