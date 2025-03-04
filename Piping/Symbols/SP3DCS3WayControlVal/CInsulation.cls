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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Wednesday, Jan 22 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = I3AT or I3AS,
'   SN = I3AZ at page no 353 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who                   change description
'   -----------     -----                 ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav              Defined new input parameter HandwheelAngle.
'   06.Dec.2004     SymbolTeam(India)     Removed the small value (0.0000001) that was
'                                         added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG                   TR-124959  Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Integer
    Dim ObjInsulatedCone1 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedCone2 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedCone3 As Object
    Dim ObjInsulatedPort3 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parAngle As Double
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
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    parNPD = arrayOfInputs(8)
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
    parId1 = arrayOfInputs(35)
    parId2 = arrayOfInputs(36)
    parId3 = arrayOfInputs(37)
    parNPDUnitType = arrayOfInputs(38)
    parNPDUnitType1 = arrayOfInputs(39)
    parNPDUnitType2 = arrayOfInputs(40)
    parNPDUnitType3 = arrayOfInputs(41)

    
    iOutput = 0

' Insert your code for output 4(Insulation Cone1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia   As Double

    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick(parPortIndex1), 0) Then
        flangeThick(parPortIndex1) = NEGLIGIBLE_THICKNESS
    End If
    If CmpDblEqual(flangeThick(parPortIndex2), 0) Then
        flangeThick(parPortIndex2) = NEGLIGIBLE_THICKNESS
    End If
    If CmpDblEqual(flangeThick(parPortIndex3), 0) Then
         flangeThick(parPortIndex3) = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set CenterPos.x - (parFace1toCenter - flangeThick(parPortIndex1)), CenterPos.y, CenterPos.z
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjInsulatedCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, InsulationDia / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone1
    Set ObjInsulatedCone1 = Nothing
    
' Insert your code for output 5(Insulation Port1)
    stPoint.Set CenterPos.x - parFace1toCenter, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x - (parFace1toCenter - flangeThick(parPortIndex1) - parInsulationThickness), _
                CenterPos.y, _
                CenterPos.z
    InsulationDia = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1)) Then InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 6(Insulation Cone2)

    stPoint.Set CenterPos.x + parFace2toCenter - flangeThick(parPortIndex2), CenterPos.y, CenterPos.z
' PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
' value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjInsulatedCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, InsulationDia / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone2
    Set ObjInsulatedCone2 = Nothing
    
' Insert your code for output 7(Insulation Port2)
    stPoint.Set CenterPos.x + parFace2toCenter, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x + parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness, _
                CenterPos.y, _
                CenterPos.z
    InsulationDia = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam(parPortIndex2), flangeDiam(parPortIndex2)) Then InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
' Insert your code for output 8(Insulation Cone3)
    stPoint.Set CenterPos.x + (parFace3toCenter - flangeThick(parPortIndex3)) * Cos(parAngle), _
                CenterPos.y + (parFace3toCenter - flangeThick(parPortIndex3)) * Sin(parAngle), _
                CenterPos.z
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    InsulationDia = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    Set ObjInsulatedCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, InsulationDia / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone3
    Set ObjInsulatedCone3 = Nothing
    
' Insert your code for output 9(Insulation Port3)
    stPoint.Set CenterPos.x + parFace3toCenter * Cos(parAngle), _
                CenterPos.y + parFace3toCenter * Sin(parAngle), CenterPos.z
    enPoint.Set CenterPos.x + (parFace3toCenter - flangeThick(parPortIndex3) - parInsulationThickness) * Cos(parAngle), _
                CenterPos.y + (parFace3toCenter - flangeThick(parPortIndex3) - parInsulationThickness) * Sin(parAngle), _
                CenterPos.z
    InsulationDia = flangeDiam(parPortIndex3) + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam(parPortIndex3), flangeDiam(parPortIndex3)) Then InsulationDia = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
