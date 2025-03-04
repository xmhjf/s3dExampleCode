VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          SSP
'   Creation Date:  Friday, 5 Sep 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=ISSPS
'   This symbol is Custom INtools Instrument with Nozzles are fully Parametric.
'   The symbol consists of 2 cones, 1 valve Cylinder, 1 revolution object covering 2 actuator cylinders and
'   2 Nozzles as part of Physical aspect.One Cylinder for valve body as a part Insulation aspect.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   05.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Operation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    
    Dim ObjOpEnvelope As Object

    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInsulationThickness As Double
    
    'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parPortIndex2 As Long
    Dim parNpd2 As Double
    Dim parEndPreparation2 As Long
    Dim parScheduleThickness2 As Long
    Dim parEndStandard2 As Long
    Dim parPressureRating2 As Long
    Dim parFlowDirection2 As DistribFlow
    Dim parId2 As String
    Dim parNPDUnitType2 As String
    
    Dim parOperationHeight As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)         'D2
    parFacetoFace = arrayOfInputs(3)         'D3
    parCylHeight = arrayOfInputs(4)         'D4
    parActuatorHeight = arrayOfInputs(5)           'D5
    parActuatorDiameter = arrayOfInputs(6)           'D6
    parInstrumentHeight = arrayOfInputs(7)           'D19
    parInsulationThickness = arrayOfInputs(8)
    parNPD = arrayOfInputs(9)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(10)
    parScheduleThickness = arrayOfInputs(11)
    parEndStandard = arrayOfInputs(12)
    parPressureRating = arrayOfInputs(13)
    parFlowDirection = arrayOfInputs(14)
    parPortIndex1 = arrayOfInputs(15)
    parNpd1 = arrayOfInputs(16)
    parEndPreparation1 = arrayOfInputs(17)
    parScheduleThickness1 = arrayOfInputs(18)
    parEndStandard1 = arrayOfInputs(19)
    parPressureRating1 = arrayOfInputs(20)
    parFlowDirection1 = arrayOfInputs(21)
    parPortIndex2 = arrayOfInputs(22)
    parNpd2 = arrayOfInputs(23)
    parEndPreparation2 = arrayOfInputs(24)
    parScheduleThickness2 = arrayOfInputs(25)
    parEndStandard2 = arrayOfInputs(26)
    parPressureRating2 = arrayOfInputs(27)
    parFlowDirection2 = arrayOfInputs(28)
    parOperationHeight = arrayOfInputs(29)     'D7
    parOperationDiameter = arrayOfInputs(30) 'D17
    parOperationDiameter1 = arrayOfInputs(31) 'D18
    
    parId1 = arrayOfInputs(32)
    parId2 = arrayOfInputs(33)
    parNPDUnitType = arrayOfInputs(34)
    parNPDUnitType1 = arrayOfInputs(35)
    parNPDUnitType2 = arrayOfInputs(36)
    
    iOutput = 0


 ' Insert your code for output (Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim dOpDiameter As Double
    
    stPoint.Set 0, parInstrumentHeight + parActuatorHeight, 0
    enPoint.Set 0, parInstrumentHeight + parActuatorHeight + parOperationHeight, 0
    
    If parOperationDiameter >= parOperationDiameter1 Then
        dOpDiameter = parOperationDiameter
    Else
        dOpDiameter = parOperationDiameter1
    End If
    
    Set ObjOpEnvelope = PlaceCylinder(m_OutputColl, stPoint, enPoint, dOpDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpEnvelope
    Set ObjOpEnvelope = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
