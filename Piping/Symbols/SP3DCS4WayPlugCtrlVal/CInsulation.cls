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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 23 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = PLU4W,
'   SN = V19 at page no 242 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    Dim iOutput     As Double
    Dim ObjInsulCylinder1 As Object
    Dim ObjInsulCylinder2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'A
    parFace2toCenter = arrayOfInputs(3)     'B
    parFace3toCenter = arrayOfInputs(4)     'C
    parFace4toCenter = arrayOfInputs(5)     'C
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    parNpd = arrayOfInputs(8)
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
    parPortIndex3 = arrayOfInputs(28)
    parNpd3 = arrayOfInputs(29)
    parEndPreparation3 = arrayOfInputs(30)
    parScheduleThickness3 = arrayOfInputs(31)
    parEndStandard3 = arrayOfInputs(32)
    parPressureRating3 = arrayOfInputs(33)
    parFlowDirection3 = arrayOfInputs(34)
    parPortIndex4 = arrayOfInputs(35)
    parNpd4 = arrayOfInputs(36)
    parEndPreparation4 = arrayOfInputs(37)
    parScheduleThickness4 = arrayOfInputs(38)
    parEndStandard4 = arrayOfInputs(39)
    parPressureRating4 = arrayOfInputs(40)
    parFlowDirection4 = arrayOfInputs(41)
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parId3 = arrayOfInputs(44)
    parId4 = arrayOfInputs(45)
    parNpdUnitType = arrayOfInputs(46)
    parNpdUnitType1 = arrayOfInputs(47)
    parNpdUnitType2 = arrayOfInputs(48)
    parNpdUnitType3 = arrayOfInputs(49)
    parNpdUnitType4 = arrayOfInputs(50)
    
    iOutput = 0
    
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

 ' Insert your code for output 7(Insul Cylinder1)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double
    
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0

    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        InsulationDiameter1 = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    Else
        InsulationDiameter1 = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
    
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        InsulationDiameter2 = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    Else
        InsulationDiameter2 = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    
    If InsulationDiameter1 > InsulationDiameter2 Then
        InsulationDiameter = InsulationDiameter1
    Else
        InsulationDiameter = InsulationDiameter2
    End If
    
    Set ObjInsulCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder1
    Set ObjInsulCylinder1 = Nothing
    
' Insert your code for output 8(Insul Cylinder2)
    Dim InsulationDiameter3 As Double
    Dim InsulationDiameter4 As Double
    
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, -parFace4toCenter, 0
    
    If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then
        InsulationDiameter3 = flangeDiam(parPortIndex3) + parInsulationThickness * 2
    Else
        InsulationDiameter3 = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    End If
       
    If flangeDiam(parPortIndex4) > pipeDiam(parPortIndex4) Then
        InsulationDiameter4 = flangeDiam(parPortIndex4) + parInsulationThickness * 2
    Else
        InsulationDiameter4 = pipeDiam(parPortIndex4) + parInsulationThickness * 2
    End If
    
    If InsulationDiameter3 > InsulationDiameter4 Then
        InsulationDiameter = InsulationDiameter3
    Else
        InsulationDiameter = InsulationDiameter4
    End If
    
    Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder2
    Set ObjInsulCylinder2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
