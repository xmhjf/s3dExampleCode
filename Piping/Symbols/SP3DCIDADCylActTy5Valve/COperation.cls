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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA5,
'   PDMS Category Name=IVCH). The symbol consists of 12 outputs,Ten physical outputs,
'   one Insulation Output and one operation output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
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
    Dim parCylHeight  As Double
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
    Dim parActuatorOffset1 As Double
    Dim parInsulationThickness As Double
    
'   Dynamic Port Input parameters variable declaration
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
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFacetoFace = arrayOfInputs(3) 'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parActuatorHeight = arrayOfInputs(5) 'D5
    parActuatorHeight1 = arrayOfInputs(6) 'D6
    parActuatorWidth = arrayOfInputs(7) 'D7
    parActuatorWidth1 = arrayOfInputs(8) 'D8
    parActuatorLength = arrayOfInputs(9) 'D9
    parActuatorLength1 = arrayOfInputs(10) 'D10
    parActuatorHeight2 = arrayOfInputs(11) 'D11
    parActuatorLength2 = arrayOfInputs(12) 'D12
    parActuatorLength3 = arrayOfInputs(13) 'D13
    parActuatorWidth2 = arrayOfInputs(14) 'D14
    parActuatorWidth3 = arrayOfInputs(15) 'D15
    parActuatorOffset = arrayOfInputs(16) 'D16
    parActuatorCylDiameter = arrayOfInputs(17) 'D17
    parActuatorCylLength = arrayOfInputs(18) 'D18
    parActuatorHeight3 = arrayOfInputs(19) 'D19
    parActuatorCyl1Diameter = arrayOfInputs(20) 'D20
    parActuatorCyl1Length = arrayOfInputs(21) 'D21
    parActuatorOffset1 = arrayOfInputs(22) 'D23
    parInsulationThickness = arrayOfInputs(23)
    parNPD = arrayOfInputs(24)
    parEndPreparation = arrayOfInputs(25)
    parScheduleThickness = arrayOfInputs(26)
    parEndStandard = arrayOfInputs(27)
    parPressureRating = arrayOfInputs(28)
    parFlowDirection = arrayOfInputs(29)
    parPortIndex1 = arrayOfInputs(30)
    parNpd1 = arrayOfInputs(31)
    parEndPreparation1 = arrayOfInputs(32)
    parScheduleThickness1 = arrayOfInputs(33)
    parEndStandard1 = arrayOfInputs(34)
    parPressureRating1 = arrayOfInputs(35)
    parFlowDirection1 = arrayOfInputs(36)
    parPortIndex2 = arrayOfInputs(37)
    parNpd2 = arrayOfInputs(38)
    parEndPreparation2 = arrayOfInputs(39)
    parScheduleThickness2 = arrayOfInputs(40)
    parEndStandard2 = arrayOfInputs(41)
    parPressureRating2 = arrayOfInputs(42)
    parFlowDirection2 = arrayOfInputs(43)
    parOperationHeight = arrayOfInputs(44)     'D22
    
    parId1 = arrayOfInputs(45)
    parId2 = arrayOfInputs(46)
    parNPDUnitType = arrayOfInputs(47)
    parNPDUnitType1 = arrayOfInputs(48)
    parNPDUnitType2 = arrayOfInputs(49)
        
    iOutput = 0


 ' Insert your code for output (Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set -parActuatorLength - parActuatorCylLength, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2, _
                         -parActuatorOffset - 0.5 * parActuatorCyl1Diameter
    enPoint.Set parActuatorLength1 + parActuatorCyl1Length, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2 + parOperationHeight, _
                        parActuatorOffset1 + 0.5 * parActuatorCylDiameter
                        
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
