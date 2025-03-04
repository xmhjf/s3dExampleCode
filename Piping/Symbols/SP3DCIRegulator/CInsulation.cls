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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          KKC
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(SREG1)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
    
    PI = 4 * Atn(1)

End Sub
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim pipeDiam(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
        
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjValveBodyIns As Object
    Dim ObjVertCylinsIns As Object
    Dim ObjHoriCylindersIns  As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim iOutput     As Integer
    
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
    parFace1toCenter = arrayOfInputs(2) 'V1
    parFacetoFace = arrayOfInputs(3)        'V2
    parCylHeight = arrayOfInputs(4)          'V3
    parInstrumentLength = arrayOfInputs(5)   'A3
    parInstrumentHeight = arrayOfInputs(6)   'A2
    parInstrumentDiameter = arrayOfInputs(7)     'A4
    parInstrumentLength1 = arrayOfInputs(8)    'A5
    parInstrumentLength2 = arrayOfInputs(9)    'A7
    parInstrumentDiameter1 = arrayOfInputs(10)        'A6
    parInsulationThickness = arrayOfInputs(11)
    parRotation = arrayOfInputs(12)
    parNPD = arrayOfInputs(13)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parPortIndex1 = arrayOfInputs(19)
    parNpd1 = arrayOfInputs(20)
    parEndPreparation1 = arrayOfInputs(21)
    parScheduleThickness1 = arrayOfInputs(22)
    parEndStandard1 = arrayOfInputs(23)
    parPressureRating1 = arrayOfInputs(24)
    parFlowDirection1 = arrayOfInputs(25)
    parPortIndex2 = arrayOfInputs(26)
    parNpd2 = arrayOfInputs(27)
    parEndPreparation2 = arrayOfInputs(28)
    parScheduleThickness2 = arrayOfInputs(29)
    parEndStandard2 = arrayOfInputs(30)
    parPressureRating2 = arrayOfInputs(31)
    parFlowDirection2 = arrayOfInputs(32)
    parOperationLength = arrayOfInputs(33) 'I2
    parOperationHeight = arrayOfInputs(34) 'I1
'Text Inputs
    parId1 = arrayOfInputs(35)
    parId2 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 5 (Valve Body Insulation)
     CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam(parPortIndex1), pipeDiam(parPortIndex2), _
                            flangeDiam(parPortIndex1), flangeDiam(parPortIndex2))
    InsulationDia = maxDia + 2 * parInsulationThickness
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
' Insert your code for output 3(Bottom Valve Cylinder)
    Dim dCylDia As Double
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -parCylHeight - parInsulationThickness, 0
    dCylDia = parFace1toCenter + 2 * parInsulationThickness
    Dim ObjValCylIns
    Set ObjValCylIns = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCylIns
    Set ObjValCylIns = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
