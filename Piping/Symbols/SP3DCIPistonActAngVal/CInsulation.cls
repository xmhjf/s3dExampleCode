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
'   Author:          BG
'   Creation Date:  Friday 10 Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=ISSPA
'  This symbol is Custom INtools Instrument with Nozzles are fully Parametric.
'   The symbol consists of 2 Cones, 1 Revolution object as Actuator cum Valve Cylinder, and 2 Nozzles as
'   part of physical aspect. Insulation aspect consist of 2 Cylinders for valve body
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1Ins  As Object
    Dim ObjBodyCone2Ins  As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
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
    parFace1toCenter = arrayOfInputs(2)          'D2
    parFace2toCenter = arrayOfInputs(3)          'D3
    parActuatorHeight = arrayOfInputs(4)           'D5
    parActuatorDiameter = arrayOfInputs(5)           'D6
    parInstrumentHeight = arrayOfInputs(6)           'D19
    parInsulationThickness = arrayOfInputs(7)
    parNPD = arrayOfInputs(8)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(9)
    parScheduleThickness = arrayOfInputs(10)
    parEndStandard = arrayOfInputs(11)
    parPressureRating = arrayOfInputs(12)
    parFlowDirection = arrayOfInputs(13)
    parPortIndex1 = arrayOfInputs(14)
    parNpd1 = arrayOfInputs(15)
    parEndPreparation1 = arrayOfInputs(16)
    parScheduleThickness1 = arrayOfInputs(17)
    parEndStandard1 = arrayOfInputs(18)
    parPressureRating1 = arrayOfInputs(19)
    parFlowDirection1 = arrayOfInputs(20)
    parPortIndex2 = arrayOfInputs(21)
    parNpd2 = arrayOfInputs(22)
    parEndPreparation2 = arrayOfInputs(23)
    parScheduleThickness2 = arrayOfInputs(24)
    parEndStandard2 = arrayOfInputs(25)
    parPressureRating2 = arrayOfInputs(26)
    parFlowDirection2 = arrayOfInputs(27)
    
    parOperationHeight = arrayOfInputs(28)
    parOperationDiameter = arrayOfInputs(29)
    parOperationDiameter1 = arrayOfInputs(30)
    
    parId1 = arrayOfInputs(31)
    parId2 = arrayOfInputs(32)
    parNPDUnitType = arrayOfInputs(33)
    parNPDUnitType1 = arrayOfInputs(34)
    parNPDUnitType2 = arrayOfInputs(35)
    
    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 9(Body Cone 1 Insulation)
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        InsulationDia = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 10(Body Cone 2 Insulation)
  '  RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        InsulationDia = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, parInstrumentHeight + parInsulationThickness, 0
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

End Sub

