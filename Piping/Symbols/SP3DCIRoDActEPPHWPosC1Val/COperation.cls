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
'   COperationcls
'   Author:          MS
'   Creation Date:  Sunday 28 Sept 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVJU).
'    this Symbol is same as (PDMS Category Name=IVJS).'   This symbol is Custom INtools Instrument with
'   Nozzles are fully Parametric.Physical aspect consist 2 cones, 2 Cylinders of stem Dia, 2 Cylinder , 1 box,
'   1 torus and a cone with elliptical dome.  2 Nozzles. Insulation Aspect consist of a simple Cylinder for the valve Body.
'   Operation Aspect consist of a simple Cylinder
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   28.Sep.2003    MS TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' This class module is the place for user to implement graphical part of VBSymbol for this aspect

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
    Dim parFace1toCenter As Double         'D2
    Dim parFacetoFace     As Double     'D3
    Dim parCylHeight      As Double     'D4
    Dim parActuatorWidth     As Double      'D5
    Dim parActuatorDiameter     As Double       'D6
    Dim parActuatorHeight     As Double          'D7
    Dim parHandWheelLength    As Double         'D8
    Dim parHandWheelDiameter      As Double         'D9
    Dim parActuatorHeight1     As Double            'D10
    Dim parActuatorHeight2     As Double            'D11
    Dim parActuatorLength     As Double             'D12
    Dim parActuatorLength1    As Double             'D14
    Dim parActuatorHeight3     As Double            'D16
    Dim parActuatorHeight4     As Double            'D19
    Dim parHandWheelOffset   As Double          'D20
    Dim parActuatorWidth1       As Double           'D25
    Dim parActuatorWidth2     As Double             'D26
    Dim parActuatorHeight5     As Double            'D27
    Dim parActuatorLength2  As Double           'D28
    Dim parInsulationThickness As Double
    Dim parOperationDiameter As Double      'D17
    Dim parOperationDiameter1 As Double      'D18
    
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)         'D2
    parFacetoFace = arrayOfInputs(3)         'D3
    parCylHeight = arrayOfInputs(4)         'D4
    parActuatorWidth = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorHeight = arrayOfInputs(7)    'D7
    parHandWheelLength = arrayOfInputs(8)    'D8
    parHandWheelDiameter = arrayOfInputs(9)     'D9
    parActuatorHeight1 = arrayOfInputs(10)    'D10
    parActuatorHeight2 = arrayOfInputs(11)    'D11
    parActuatorLength = arrayOfInputs(12)    'D12
    parActuatorLength1 = arrayOfInputs(13)    'D14
    parActuatorHeight3 = arrayOfInputs(14)    'D16
    parActuatorHeight4 = arrayOfInputs(15)    'D19
    parHandWheelOffset = arrayOfInputs(16)    'D20
    parActuatorWidth1 = arrayOfInputs(17)    'D25
    parActuatorWidth2 = arrayOfInputs(18)    'D26
    parActuatorHeight5 = arrayOfInputs(19)    'D27
    parActuatorLength2 = arrayOfInputs(20)    'D28
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
    parOperationDiameter = arrayOfInputs(42)
    parOperationDiameter = arrayOfInputs(43)
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0

 ' Insert your code for output (Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim OpDiameter As Double
    If parOperationDiameter1 > parOperationDiameter Then
        OpDiameter = parOperationDiameter1
    Else
        OpDiameter = parOperationDiameter
    End If
    'The Operation cylinder height is assumed to be parActuatorWidth(D5)
    stPoint.Set parActuatorLength2, parActuatorHeight, -parActuatorWidth
    enPoint.Set parActuatorLength2, parActuatorHeight, -parActuatorWidth - parActuatorWidth
    
    Set ObjOpEnvelope = PlaceCylinder(m_OutputColl, stPoint, enPoint, OpDiameter, True)
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
