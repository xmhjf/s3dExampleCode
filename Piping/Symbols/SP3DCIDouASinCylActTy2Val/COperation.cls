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
'   Creation Date:  Moday, Sep 15 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA2,
'   PDMS Category Name=IVCJ). Physical aspect consist of two cones and a cylinder as a part
'   valve body, Stem cylinder, Two boxes , two cylinders as part of operator and Two Nozzles.
'   Insulation aspect consist of single cylinder for valve body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
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
    Dim parActuatorHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parActuatorCyl1Diameter As Double
    Dim parActuatorCyl1Length As Double
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
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parActuatorHeight = arrayOfInputs(5)       'D5
    parActuatorHeight1 = arrayOfInputs(6)      'D6
    parActuatorWidth = arrayOfInputs(7)    'D7
    parActuatorWidth1 = arrayOfInputs(8)       'D8
    parActuatorLength = arrayOfInputs(9)       'D9
    parActuatorLength1 = arrayOfInputs(10)      'D10
    parActuatorHeight2 = arrayOfInputs(11)      'D11
    parActuatorLength2 = arrayOfInputs(12)      'D12
    parActuatorLength3 = arrayOfInputs(13)     'D13
    parActuatorWidth2 = arrayOfInputs(14)      'D14
    parActuatorWidth3 = arrayOfInputs(15)      'D15
    parActuatorOffset = arrayOfInputs(16)      'D16
    parActuatorCylDiameter = arrayOfInputs(17)    'D17
    parActuatorCylLength = arrayOfInputs(18)      'D18
    parActuatorHeight3 = arrayOfInputs(19)     'D19
    parActuatorCyl1Diameter = arrayOfInputs(20)    'D20
    parActuatorCyl1Length = arrayOfInputs(21)      'D21
    parInsulationThickness = arrayOfInputs(22)
    parNPD = arrayOfInputs(23)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(24)
    parScheduleThickness = arrayOfInputs(25)
    parEndStandard = arrayOfInputs(26)
    parPressureRating = arrayOfInputs(27)
    parFlowDirection = arrayOfInputs(28)
    parPortIndex1 = arrayOfInputs(29)
    parNpd1 = arrayOfInputs(30)
    parEndPreparation1 = arrayOfInputs(31)
    parScheduleThickness1 = arrayOfInputs(32)
    parEndStandard1 = arrayOfInputs(33)
    parPressureRating1 = arrayOfInputs(34)
    parFlowDirection1 = arrayOfInputs(35)
    parPortIndex2 = arrayOfInputs(36)
    parNpd2 = arrayOfInputs(37)
    parEndPreparation2 = arrayOfInputs(38)
    parScheduleThickness2 = arrayOfInputs(39)
    parEndStandard2 = arrayOfInputs(40)
    parPressureRating2 = arrayOfInputs(41)
    parFlowDirection2 = arrayOfInputs(42)
    parOperationHeight = arrayOfInputs(43)     'D22
    
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0

 ' Insert your code for output (Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim dOpWidth As Double
    
    dOpWidth = 0.5 * parActuatorCylDiameter
    If dOpWidth < parActuatorWidth Then dOpWidth = parActuatorWidth
    
    stPoint.Set -parActuatorCylLength - parActuatorLength, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2, _
                        -parActuatorWidth1 + parActuatorOffset
    enPoint.Set parActuatorLength1 + parActuatorCyl1Length, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2 + parOperationHeight, _
                        parActuatorOffset + dOpWidth
                        
    Set ObjOpEnvelope = PlaceBox(m_OutputColl, stPoint, enPoint)
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
