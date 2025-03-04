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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          KKC
'   Creation Date:  Monday, Jun 23 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-137825  Piping instrument catalog should be updated to current Intools Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
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

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorDiameter As Double
    Dim parOperationHeight As Double
    Dim parHandWheelLength As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parPositionerHeight As Double
    Dim parPositionerOffset As Double
    Dim parPostionerLength As Double
    Dim parPositionerOffset1 As Double
    Dim parPositionerOffset2 As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parRotation As Double
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

    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFacetoFace = arrayOfInputs(3)    'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parActuatorHeight = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)  'D6
    parOperationHeight = arrayOfInputs(7)   'D7
    parHandWheelLength = arrayOfInputs(8)  'D8
    parHandWheelDiameter = arrayOfInputs(9) 'D9
    parPositionerHeight = arrayOfInputs(10) 'D10
    parPositionerOffset = arrayOfInputs(11) ' D11
    parPositionerOffset1 = arrayOfInputs(12)  'D12
    parActuatorHeight1 = arrayOfInputs(13) 'D13
    parPositionerOffset2 = arrayOfInputs(14)  'D14
    parHandWheelOffset = arrayOfInputs(15) 'D16
    parOperationDiameter = arrayOfInputs(16)  'D17
    parOperationDiameter1 = arrayOfInputs(17)  ' D18
    parPostionerLength = arrayOfInputs(18)  ' D20
    parRotation = arrayOfInputs(19)  ' D25
    parInsulationThickness = arrayOfInputs(20)
    parNPD = arrayOfInputs(21)
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
    parId1 = arrayOfInputs(41)
    parId2 = arrayOfInputs(42)
    parNPDUnitType = arrayOfInputs(43)
    parNPDUnitType1 = arrayOfInputs(44)
    parNPDUnitType2 = arrayOfInputs(45)
    
    iOutput = 0

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    Dim InsulationDia As Double
 ' Insert your code for output 8(Valve Body Insulation)
    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                            
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    
    CylDia = 0.4 * parActuatorDiameter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
   
'maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam(parPortIndex1) / 2, pipeDiam(parPortIndex2) / 2, flangeDiam(parPortIndex1) / 2, flangeDiam(parPortIndex2) / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Dim ObjValveBodyIns As Object
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


