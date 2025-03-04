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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:           ACM
'   Creation Date:  Tuesday, Dec 31 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ISSD3),
'   The symbol consists of 11 outputs,8 physical outputs,
'    2 Insulation Output and one Operation output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP   TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDia(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    Dim flangeDia(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim depth(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1Ins As Object
    Dim ObjBodyCone2Ins As Object
    
    Dim parFace1toCenter As Double ''D2
    Dim parFacetoFace As Double ''D3
    Dim parFace3toCenter As Double ''D4
    Dim parActuatorHeight As Double ''D5
    Dim parActuatorDiameter As Double ''D6
    Dim parActuatorHeight1 As Double ''D19
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parOperationHeight As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight1 = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    parNPD = arrayOfInputs(9)  'Input parameters for Dynamic Nozzles
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
    
    parPortIndex3 = arrayOfInputs(29)
    parNpd3 = arrayOfInputs(30)
    parEndPreparation3 = arrayOfInputs(31)
    parScheduleThickness3 = arrayOfInputs(32)
    parEndStandard3 = arrayOfInputs(33)
    parPressureRating3 = arrayOfInputs(34)
    parFlowDirection3 = arrayOfInputs(35)
    
    parOperationHeight = arrayOfInputs(36)     'D7
    parOperationDiameter = arrayOfInputs(37) 'D17
    parOperationDiameter1 = arrayOfInputs(38) 'D18
    
'Comment to make test with text inputs
    parId1 = arrayOfInputs(39)
    parId2 = arrayOfInputs(40)
    parId3 = arrayOfInputs(41)
    parNPDUnitType = arrayOfInputs(42)
    parNPDUnitType1 = arrayOfInputs(43)
    parNPDUnitType2 = arrayOfInputs(44)
    parNPDUnitType3 = arrayOfInputs(45)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
 ' Insert your code for output 6(InsulationBody cone 1)
   
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                        parEndPreparation, parScheduleThickness, parEndStandard, _
                                        parPressureRating, parFlowDirection, parNpd1, _
                                        parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                        parEndStandard1, parPressureRating1, parFlowDirection1, _
                                        parId1, m_OutputColl, m_oCodeListMetadata, _
                                        False, pipeDia, flangeThick, _
                                        flangeDia, sptOffset, depth

    
    Dim CylEdgeToCen As Double
    Dim CylRad As Double
    
    CylRad = parFace1toCenter / 2
    CylEdgeToCen = Sqr(parActuatorHeight1 ^ 2 + CylRad ^ 2)
    
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDia(parPortIndex1) / 2, pipeDia(parPortIndex2) / 2, _
                            flangeDia(parPortIndex1) / 2, flangeDia(parPortIndex2) / 2, CylEdgeToCen)
    
    InsulationDia = maxRad * 2 + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 7(Body Cone 2 Insulation)
    
    InsulationDia = pipeDia(parPortIndex3) + parInsulationThickness * 2
    If flangeDia(parPortIndex3) > pipeDia(parPortIndex3) Then InsulationDia = flangeDia(parPortIndex3) + 2 * parInsulationThickness
       
    stPoint.Set 0, -parFace3toCenter, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2Ins
    Set ObjBodyCone2Ins = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double
    
    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue

End Function

