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
'   Author:          MS
'   Creation Date:  20 June 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRXPA).
'   Physical aspect consist 1 Cylinder, 2 Cones ,1 stem ,3 boxes,2 nozzles and 1 insulation aspect outputs.
'   The nozzles are fully Parametric.
'   Added Operation Aspect to the symbol. Added OperationHeight as input parameters
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
     
    Dim parFace1toCenter As Double              'D2
    Dim parFacetoFace As Double                 'D3
    Dim parCylHeight As Double               'D4
    Dim parActuatorHeight As Double             'D5
    Dim parActuatorHeight1 As Double            'D6
    Dim parActuatorHeight2 As Double            'D7
    Dim parActuatorLength As Double             'D8
    Dim parActuatorLength1 As Double            'D9
    Dim parActuatorWidth As Double              'D10
    Dim parActuatorWidth1 As Double             'D11
    Dim parActuatorLength2 As Double            'D12
    Dim parActuatorLength3 As Double            'D13
    Dim parActuatorWidth2 As Double             'D14
    Dim parActuatorWidth3 As Double             'D15
    Dim parActuatorWidth4 As Double             'D16
    Dim parActuatorLength4 As Double            'D17
    Dim parActuatorLength5 As Double            'D18
    Dim parActuatorHeight3 As Double            'D19
    Dim parActuatorHeight4 As Double            'D20
    Dim parActuatorHeight5 As Double            'D21
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
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorHeight1 = arrayOfInputs(6)
    parActuatorHeight2 = arrayOfInputs(7)
    parActuatorLength = arrayOfInputs(8)
    parActuatorLength1 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorLength2 = arrayOfInputs(12)
    parActuatorLength3 = arrayOfInputs(13)
    parActuatorWidth2 = arrayOfInputs(14)
    parActuatorWidth3 = arrayOfInputs(15)
    parActuatorWidth4 = arrayOfInputs(16)
    parActuatorLength4 = arrayOfInputs(17)
    parActuatorLength5 = arrayOfInputs(18)
    parActuatorHeight3 = arrayOfInputs(19)
    parActuatorHeight4 = arrayOfInputs(20)
    parActuatorHeight5 = arrayOfInputs(21)
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
    parOperationHeight = arrayOfInputs(43)   'D22
    'Text Inputs
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0

 ' Insert your code for output 11(Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    stPoint.Set -parActuatorLength, _
                        parActuatorHeight + parActuatorHeight1 + parActuatorHeight2, _
                        parActuatorWidth + parActuatorWidth4
    enPoint.Set parActuatorLength1, _
                        parActuatorHeight + parActuatorHeight1 + parActuatorHeight2 + parOperationHeight, _
                        -parActuatorWidth1
    
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
    Resume Next
    
End Sub
