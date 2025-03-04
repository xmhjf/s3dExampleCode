VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav 
'   Creation Date:  Tuesday, Oct 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IEAM1,
'   PDMS Category Name=IVCM). The symbol consists of Twelve outputs, Eleven physical
'   outputs and One Insulation Output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjStem1 As Object
    Dim ObjMotor1 As Object
    Dim ObjMotor2 As Object
    Dim ObjMotor3 As Object
    Dim ObjStem2 As Object
    Dim ObjHandwheel As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight  As Double
    Dim parActuatorHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorOffset As Double
    Dim parMotor1Diameter As Double
    Dim parMotor2Diameter As Double
    Dim parActuatorLength1 As Double
    Dim parMotor3Diameter As Double
    Dim parActuatorHeight1 As Double
    Dim parHandWheelDiameter As Double
    Dim parActuatorOffset1 As Double
    Dim parMotor3EndtoCenter As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim TorOrigin As New AutoMath.DPosition
    
    Dim StemDia As Double
    Dim HandWheelCSDia As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'D2
    parFacetoFace = arrayOfInputs(3)     'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parActuatorHeight = arrayOfInputs(5)     'D5
    parActuatorLength = arrayOfInputs(6)     'D6
    parActuatorOffset = arrayOfInputs(7)     'D7
    parMotor1Diameter = arrayOfInputs(8)     'D8
    parMotor2Diameter = arrayOfInputs(9)     'D9
    parActuatorLength1 = arrayOfInputs(10)   'D10
    parMotor3Diameter = arrayOfInputs(11)    'D11
    parActuatorHeight1 = arrayOfInputs(12)    'D12
    parHandWheelDiameter = arrayOfInputs(13)     'D13
    parActuatorOffset1 = arrayOfInputs(14)   'D14
    parMotor3EndtoCenter = arrayOfInputs(15)     'D15
    parInsulationThickness = arrayOfInputs(16)
    
    iOutput = 0
    
'   Stem diameter and Hand wheel minimum Thickness are taken as 0.05
    StemDia = 0.05
    HandWheelCSDia = 0.05

' Insert your code for output 2(Valve Cone1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 3(Body Cone2 of Valve)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 4(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 '  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
' Insert your code for output 5(Stem 1)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight, 0
    Set ObjStem1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem1
    Set ObjStem1 = Nothing
    
' Insert your code for output 6(Motor 1)
    stPoint.Set -parActuatorLength, _
                parActuatorHeight - parActuatorOffset, _
                parActuatorOffset1
    enPoint.Set 0, parActuatorHeight - parActuatorOffset, parActuatorOffset1
    Set ObjMotor1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1
    Set ObjMotor1 = Nothing
    
' Insert your code for output 7(Motor 2)
    stPoint.Set 0, parActuatorHeight, parActuatorOffset1
    enPoint.Set parActuatorLength1, parActuatorHeight, parActuatorOffset1
    Set ObjMotor2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor2Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2
    Set ObjMotor2 = Nothing
    
' Insert your code for output 8(Motor 3)
    stPoint.Set 0, parActuatorHeight, _
                -(parMotor2Diameter / 2 - parActuatorOffset1)
    enPoint.Set 0, parActuatorHeight, _
                parMotor3EndtoCenter
    Set ObjMotor3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor3Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor3
    Set ObjMotor3 = Nothing
    
' Insert your code for output 9(Stem 2)
    stPoint.Set 0, parActuatorHeight, 0
    enPoint.Set 0, parActuatorHeight1, 0
    Set ObjStem2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem2
    Set ObjStem2 = Nothing
    
' Insert your code for output 10(Handwheel)
    TorOrigin.Set 0, parActuatorHeight1 - HandWheelCSDia / 2, 0
    axisVect.Set 0, 1, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, axisVect, _
                        parHandWheelDiameter / 2 - HandWheelCSDia / 2, _
                        HandWheelCSDia / 2)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
