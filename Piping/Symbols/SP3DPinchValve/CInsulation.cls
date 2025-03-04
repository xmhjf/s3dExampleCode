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
'   Author:          RRK/KKC
'   Creation Date:  Thursday, Apr 5 2007
'   Description:
'   This is symbol is created for Pinch valves of Linatex
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulationCylinder As Object
    Dim ObjInsulationBox As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsPort2 As Object
    
    Dim parFacetoFace As Double
    Dim parValveWidth As Double
    Dim parValveHeight As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim lPartDataBasis As Long
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dInsulationDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parValveWidth = arrayOfInputs(3)
    parValveHeight = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    
    If (lPartDataBasis <= 1) Or (lPartDataBasis = 280) Then
     ' Insert your code for output 5(Insulation Cylinder)
        oStPoint.Set -parFacetoFace / 2 + parInsulationThickness + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - parInsulationThickness - flangeThick, 0, 0
        
        dInsulationDia = pipeDiam + 2 * parInsulationThickness
        
        Set ObjInsulationCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationCylinder
        Set ObjInsulationCylinder = Nothing
    
    ' Insert your code for output 6(Insulation Box)
        Dim dBoxLength As Double
        dBoxLength = parFacetoFace / 10
        
        oStPoint.Set dBoxLength / 2 + parInsulationThickness, _
                    -0.4 * parValveHeight - parInsulationThickness, _
                    parValveWidth / 2 + parInsulationThickness
                    
        oEnPoint.Set -dBoxLength / 2 - parInsulationThickness, _
                    0.6 * parValveHeight + parInsulationThickness, _
                    -parValveWidth / 2 - parInsulationThickness
                    
        Set ObjInsulationBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBox
        Set ObjInsulationBox = Nothing
    ElseIf lPartDataBasis = 285 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
     ' Insert your code for output 6(Insulation Body)
        oStPoint.Set -parFacetoFace / 2 + parInsulationThickness + flangeThick, _
                      parValveHeight / 2 + parInsulationThickness, _
                     -parValveWidth / 2 - parInsulationThickness
                        
        oEnPoint.Set parFacetoFace / 2 - parInsulationThickness - flangeThick, _
                     -parValveHeight / 2 - parInsulationThickness, _
                      parValveWidth / 2 + parInsulationThickness
        
        Set ObjInsulationBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBox
        Set ObjInsulationBox = Nothing
    End If
 ' Insert your code for output 7(ObjInsPort1)
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set -parFacetoFace / 2 + parInsulationThickness + flangeThick, 0, 0
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing

 ' Insert your code for output 8(ObjInsPort2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0

    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    Exit Sub
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
