VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          VM
'   Creation Date:  Oct 1, 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on PDS 4BOX4 model code
'   The symbol consists of 8 outputs; six physical, one insulation and one Operation output.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   07.Oct.2005     kkk   CR-60561  Vivek provided the symbol and we did Minor Changes.
'
'   08.SEP.2006     KKC   DI-95670  Replace names with initials in all revision history sheets and symbols.
'   28.Sep.2006     dkl   TR-90449  Changed the code for placing Nozzle 4.
'                   dkl             Edited Port 4 creation to enable rotation values of 0 and 315 degrees.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI           As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    Dim ObjInstrumentBody As Object
        
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parPortRotation4 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
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
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)   'A
    parInstrumentWidth = arrayOfInputs(3)   'B
    parInstrumentHeight = arrayOfInputs(4)   'C
    parInstrumentLength1 = arrayOfInputs(5)
    parInstrumentLength2 = arrayOfInputs(6)
    parPortRotation4 = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parNPD = arrayOfInputs(9)                       'Input parameters for Dynamic Nozzles
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
    parPortIndex4 = arrayOfInputs(36)
    parNpd4 = arrayOfInputs(37)
    parEndPreparation4 = arrayOfInputs(38)
    parScheduleThickness4 = arrayOfInputs(39)
    parEndStandard4 = arrayOfInputs(40)
    parPressureRating4 = arrayOfInputs(41)
    parFlowDirection4 = arrayOfInputs(42)
    
    'Text Inputs
    parId1 = arrayOfInputs(43)
    parId2 = arrayOfInputs(44)
    parId3 = arrayOfInputs(45)
    parId4 = arrayOfInputs(46)
    parNPDUnitType = arrayOfInputs(47)
    parNPDUnitType1 = arrayOfInputs(48)
    parNPDUnitType2 = arrayOfInputs(49)
    parNPDUnitType3 = arrayOfInputs(50)
    parNPDUnitType4 = arrayOfInputs(51)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
     
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    
    oPlacePoint.Set -(parInstrumentLength * 0.5) - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex2, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNPDUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parInstrumentLength * 0.5) + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 3
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNPDUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parInstrumentLength * 0.5 + parInstrumentLength1, -parInstrumentWidth * 0.5 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 4
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNPDUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

' This is based on parPortRotation4 measured from +Y direction
    Dim dLocy As Double
    Dim dLocz As Double
    Dim dLimitingAngle As Double
    dLimitingAngle = Atn(parInstrumentHeight / parInstrumentWidth)
    Dim dAngle As Double
    'NOTE: LINEAR_TOLERANCE is added or subtracted accordingly to enable correct comparison of doubles with integers
    If (CmpDblGreaterthan(parPortRotation4, 2 * PI + LINEAR_TOLERANCE) Or _
    CmpDblLessThan(parPortRotation4, -LINEAR_TOLERANCE)) Then GoTo ErrorLabel:

'Conventions used:
'       +Y
'  2----|----1
'   |   |    |
'+Z-----0-------
'   |   |    |
'  3---------4
'       |
'A) 1-0-2 forms the first quadrant, 2-0-3 forms the second quadrant and so on.
'B) When parPortRotation4=45 deg the port alligns with face 2-3, when it is 135 deg the port alligns with 3-4. This
'   order (counter-clockwise) is followed for 225 deg and 315 deg as well.

'Checking to find out the quadrant in which quadrant parPortRotation4 lies.
   '  Checking First Quadrant
    If (CmpDblGreaterthanOrEqualTo(parPortRotation4, (2 * PI - dLimitingAngle)) Or _
    CmpDblLessThan(parPortRotation4, dLimitingAngle - LINEAR_TOLERANCE)) Then
        dLocy = parInstrumentWidth / 2
        oDir.Set 0, 1, 0
        dLocz = (parInstrumentWidth / 2) * Tan(parPortRotation4)
    ' Checking Second Quadrant
    ElseIf (CmpDblGreaterthanOrEqualTo(parPortRotation4, dLimitingAngle) _
        And CmpDblLessThan(parPortRotation4, (PI - dLimitingAngle) - LINEAR_TOLERANCE)) Then
        dLocz = parInstrumentHeight / 2
        oDir.Set 0, 0, 1
        If CmpDblEqual(parPortRotation4, PI / 2) Then
            dLocy = 0
        Else
            dLocy = (parInstrumentHeight / 2) * Tan(PI / 2 - parPortRotation4)
        End If
    ' Checking Third Quadrant
    ElseIf (CmpDblGreaterthanOrEqualTo(parPortRotation4, (PI - dLimitingAngle)) _
        And CmpDblLessThan(parPortRotation4, PI + dLimitingAngle - LINEAR_TOLERANCE)) Then
        dLocy = -parInstrumentWidth / 2
        oDir.Set 0, -1, 0
        dLocz = (parInstrumentWidth / 2) * Tan(PI - parPortRotation4)
    ' Checking Fourth Quadrant
    ElseIf (CmpDblGreaterthanOrEqualTo(parPortRotation4, (PI + dLimitingAngle)) _
        And CmpDblLessThan(parPortRotation4, 2 * PI - dLimitingAngle - LINEAR_TOLERANCE)) Then
        dLocz = -parInstrumentHeight / 2
        oDir.Set 0, 0, -1
        If CmpDblEqual(parPortRotation4, 3 * PI / 2) Then
            dLocy = 0
        Else
            dLocy = -(parInstrumentHeight / 2) * Tan(3 * PI / 2 - parPortRotation4)
        End If
    End If
    
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle

      
    oPlacePoint.Set -parInstrumentLength * 0.5 + parInstrumentLength2, dLocy, dLocz
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Insert your code for output 5(Instrument Body)
    stPoint.Set -parInstrumentLength * 0.5, -parInstrumentWidth * 0.5, -parInstrumentHeight * 0.5
    enPoint.Set parInstrumentLength * 0.5, parInstrumentWidth * 0.5, parInstrumentHeight * 0.5
    Set ObjInstrumentBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstrumentBody
    Set ObjInstrumentBody = Nothing
    
 
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
 
