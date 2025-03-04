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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Dec 28 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who               change description
'   -----------     -----             ------------------
'   09.Jul.2003     SymbolTeam(India) Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav          TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                     cylinder to avoid PlaceCylinder function failure.
'   08.SEP.2006     KKC  DI-95670     Replace names with initials in all revision history sheets and symbols
'   20.DEC.2007     KKC               CR-130970 Enhance Lateral symbols for underground routing to be trimmable for sloped routing
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI                  As Double


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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
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
    
'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
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
    
    If UBound(arrayOfInputs) >= 6 Then
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 4 'To suppport the prior to V7 instances
    End If
    
    iOutput = 0

' Insert your code for output 1(Insulated Body (Port1 side))
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set CenterPos.x - (parFace1toCenter - flangeThick), CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing
    
' Insert your code for output 2(Insulated Port1)
    stPoint.Set CenterPos.x - parFace1toCenter, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x - (parFace1toCenter - flangeThick - parInsulationThickness), _
                CenterPos.y, CenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 3(Insulated Body (Port2 side))
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    enPoint.Set CenterPos.x + parFace2toCenter - flangeThick, CenterPos.y, CenterPos.z
    stPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter2 = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, parInsulationDiameter / 2, _
                                      parInsulationDiameter2 / 2, False)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing
    
' Insert your code for output 4(Insulated Port2)
    stPoint.Set CenterPos.x + parFace2toCenter, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x + parFace2toCenter - flangeThick - parInsulationThickness, _
                CenterPos.y, CenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
' Insert your code for output 5(Insulated Branch (Port3 side))
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set CenterPos.x + (parFace3toCenter - flangeThick) * Cos(parAngle), _
                CenterPos.y + (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
    
' Insert your code for output 6(Insulated Port3)
    stPoint.Set CenterPos.x + parFace3toCenter * Cos(parAngle), _
                CenterPos.y + parFace3toCenter * Sin(parAngle), CenterPos.z
    enPoint.Set CenterPos.x + (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                CenterPos.y + (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), _
                CenterPos.z
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
