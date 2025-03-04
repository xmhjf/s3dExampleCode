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
'   Author:          BG
'   Creation Date:  Friday, September 19 - 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA3,
'   PDMS Category Name=IVCE). This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. The Physical aspect of symbol consists 2cones as valve body,
'   2 cylinders one as Stem and the other at valve bottom, 3 boxes and Two Nozzles. Insulation aspect consist of
'   a single cylinder for valve body portion.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
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
    parCylHeight = arrayOfInputs(4)       'D4
    parActuatorHeight = arrayOfInputs(5)        'D5
    parActuatorHeight1 = arrayOfInputs(6)       'D6
    parActuatorWidth = arrayOfInputs(7)         'D7
    parActuatorWidth1 = arrayOfInputs(8)        'D8
    parActuatorLength = arrayOfInputs(9)        'D9
    parActuatorLength1 = arrayOfInputs(10)       'D10
    parActuatorHeight2 = arrayOfInputs(11)       'D11
    parActuatorLength2 = arrayOfInputs(12)       'D12
    parActuatorLength3 = arrayOfInputs(13)      'D13
    parActuatorWidth2 = arrayOfInputs(14)       'D14
    parActuatorWidth3 = arrayOfInputs(15)       'D15
    parActuatorOffset = arrayOfInputs(16)       'D16
    parActuatorCylDiameter = arrayOfInputs(17)  'D17
    parActuatorCylLength = arrayOfInputs(18)    'D18
    parActuatorHeight3 = arrayOfInputs(19)      'D19
    parInsulationThickness = arrayOfInputs(20)
    parNPD = arrayOfInputs(21)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(22)
    parScheduleThickness = arrayOfInputs(23)
    parEndStandard = arrayOfInputs(24)
    parPressureRating = arrayOfInputs(25)
    parFlowDirection = arrayOfInputs(26)
    parPortIndex1 = arrayOfInputs(27)
    parNpd1 = arrayOfInputs(28)
    parEndPreparation1 = arrayOfInputs(29)
    parScheduleThickness1 = arrayOfInputs(30)
    parEndStandard1 = arrayOfInputs(31)
    parPressureRating1 = arrayOfInputs(32)
    parFlowDirection1 = arrayOfInputs(33)
    parPortIndex2 = arrayOfInputs(34)
    parNpd2 = arrayOfInputs(35)
    parEndPreparation2 = arrayOfInputs(36)
    parScheduleThickness2 = arrayOfInputs(37)
    parEndStandard2 = arrayOfInputs(38)
    parPressureRating2 = arrayOfInputs(39)
    parFlowDirection2 = arrayOfInputs(40)
    parOperationHeight = arrayOfInputs(41)
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parNPDUnitType = arrayOfInputs(44)
    parNPDUnitType1 = arrayOfInputs(45)
    parNPDUnitType2 = arrayOfInputs(46)
    
    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
'Insert your code for output  (Relief Valve Operation aspect)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    Dim ObjOpEnvelope As Object
    Dim opEnvWidth As Double        'greater of D7 and D17/2
    
    If parActuatorWidth > parActuatorCylDiameter / 2 Then
        opEnvWidth = parActuatorWidth
    Else
        opEnvWidth = parActuatorCylDiameter / 2
    End If
    
    stPoint.Set -parActuatorLength - parActuatorCylLength, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2, _
                        -opEnvWidth - parActuatorOffset
    enPoint.Set parActuatorLength1, stPoint.y + parOperationHeight, _
                        parActuatorWidth1 - parActuatorOffset 'D8-D16
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
