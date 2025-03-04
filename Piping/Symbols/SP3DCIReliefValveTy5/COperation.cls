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
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT5, PDMS Category Name=IVPR)
'   The symbol consists of seven outputs, six physical and one insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
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
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentRadius As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentRadius1 As Double
    Dim parInstrumentHeight4 As Double
    Dim parInstrumentHeight5 As Double
    Dim parInstrumentRadius2 As Double
    Dim parInstrumentHeight6 As Double
    Dim parInstrumentHeight7 As Double
    Dim parCylOffset As Double
    Dim parCylOffset1 As Double
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

    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentRadius = arrayOfInputs(10)       'D10
    parInstrumentHeight2 = arrayOfInputs(11)        'D11
    parInstrumentHeight3 = arrayOfInputs(12)        'D12
    parInstrumentRadius1 = arrayOfInputs(13)      'D13
    parInstrumentHeight4 = arrayOfInputs(14)        'D14
    parInstrumentHeight5 = arrayOfInputs(15)        'D15
    parInstrumentRadius2 = arrayOfInputs(16)  'D16
    parInstrumentHeight6 = arrayOfInputs(17)    'D17
    parInstrumentHeight7 = arrayOfInputs(18)    'D18
    parCylOffset = arrayOfInputs(19)        'D22
    parCylOffset1 = arrayOfInputs(20)    'D23
    parInsulationThickness = arrayOfInputs(21)
    parNPD = arrayOfInputs(22)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(23)
    parScheduleThickness = arrayOfInputs(24)
    parEndStandard = arrayOfInputs(25)
    parPressureRating = arrayOfInputs(26)
    parFlowDirection = arrayOfInputs(27)
    parPortIndex1 = arrayOfInputs(28)
    parNpd1 = arrayOfInputs(29)
    parEndPreparation1 = arrayOfInputs(30)
    parScheduleThickness1 = arrayOfInputs(31)
    parEndStandard1 = arrayOfInputs(32)
    parPressureRating1 = arrayOfInputs(33)
    parFlowDirection1 = arrayOfInputs(34)
    parPortIndex2 = arrayOfInputs(35)
    parNpd2 = arrayOfInputs(36)
    parEndPreparation2 = arrayOfInputs(37)
    parScheduleThickness2 = arrayOfInputs(38)
    parEndStandard2 = arrayOfInputs(39)
    parPressureRating2 = arrayOfInputs(40)
    parFlowDirection2 = arrayOfInputs(41)
    
    parOperationLength = arrayOfInputs(42)      'D24
    parOperationHeight = arrayOfInputs(43)      'D25
    parOperationWidth = arrayOfInputs(44)       'D26
    parOperationWidth1 = arrayOfInputs(45)      'D27

'Comment to make test with text inputs
    parId1 = arrayOfInputs(46)
    parId2 = arrayOfInputs(47)
    parNPDUnitType = arrayOfInputs(48)
    parNPDUnitType1 = arrayOfInputs(49)
    parNPDUnitType2 = arrayOfInputs(50)
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output 1(Relief Valve Operation aspect)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    Dim ObjOpEnvelope As Object
    stPoint.Set -parFace1toCenter, -parFace2toCenter, -parOperationWidth
    enPoint.Set parOperationLength, parOperationHeight, parOperationWidth1
    Set ObjOpEnvelope = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpEnvelope
    Set ObjOpEnvelope = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    iOutput = 0

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
    
End Sub



