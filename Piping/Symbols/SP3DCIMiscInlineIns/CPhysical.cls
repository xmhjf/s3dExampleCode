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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author       :   KKC
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on (SMISC1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()

'''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjFlowHoriCylin As Object
    Dim ObjVerticalCylin As Object
    Dim ObjInstruHorBox As Object
    Dim ObjVerticalCylin1 As Object
    Dim ObjInstruHorBox1 As Object
    
    Dim parFacetoFace  As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
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

    parFacetoFace = arrayOfInputs(2)  'M1
    parFace2toCenter = arrayOfInputs(3)   'A1
    parFlowDiameter = arrayOfInputs(4) 'M2
    parInstrumentHeight = arrayOfInputs(5)       'A11
    parInstrumentHeight1 = arrayOfInputs(6)       'A21
    parInstrumentLength = arrayOfInputs(7)      'A12
    parInstrumentLength1 = arrayOfInputs(8)     'A13
    parInstrumentLength2 = arrayOfInputs(9)   'A22
    parInstrumentLength3 = arrayOfInputs(10)   'A23
    parInstrumentHeight2 = arrayOfInputs(11)       'A14
    parInstrumentHeight3 = arrayOfInputs(12)       'A24
    parInstrumentWidth = arrayOfInputs(13)    'A15
    parInstrumentWidth1 = arrayOfInputs(14)   'A25
    parInsulationThickness = arrayOfInputs(15)
    parNPD = arrayOfInputs(16)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(17)
    parScheduleThickness = arrayOfInputs(18)
    parEndStandard = arrayOfInputs(19)
    parPressureRating = arrayOfInputs(20)
    parFlowDirection = arrayOfInputs(21)
    parPortIndex1 = arrayOfInputs(22)
    parNpd1 = arrayOfInputs(23)
    parEndPreparation1 = arrayOfInputs(24)
    parScheduleThickness1 = arrayOfInputs(25)
    parEndStandard1 = arrayOfInputs(26)
    parPressureRating1 = arrayOfInputs(27)
    parFlowDirection1 = arrayOfInputs(28)
    parPortIndex2 = arrayOfInputs(29)
    parNpd2 = arrayOfInputs(30)
    parEndPreparation2 = arrayOfInputs(31)
    parScheduleThickness2 = arrayOfInputs(32)
    parEndStandard2 = arrayOfInputs(33)
    parPressureRating2 = arrayOfInputs(34)
    parFlowDirection2 = arrayOfInputs(35)
    parId1 = arrayOfInputs(36)
    parId2 = arrayOfInputs(37)
    parNPDUnitType = arrayOfInputs(38)
    parNPDUnitType1 = arrayOfInputs(39)
    parNPDUnitType2 = arrayOfInputs(40)
    
    iOutput = 0
' Place Nozzle 1 for Output 1 (Nozzle1)
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    If CmpDblEqual(parFlowDiameter, 0) Then parFlowDiameter = pipeDiam(parPortIndex1)
    If CmpDblEqual(parFace2toCenter, 0) Then parFace2toCenter = parFacetoFace / 2
    If CmpDblEqual(parInstrumentLength1, 0) Then parInstrumentLength1 = parInstrumentLength
    If CmpDblEqual(parInstrumentLength3, 0) Then parInstrumentLength3 = parInstrumentLength2
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -(parFacetoFace - parFace2toCenter) - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    oNozzle.Length = flangeThick(parPortIndex2)

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace2toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' Insert your code for output 3(Vortex Flowmeter HorizntalCylinder)
    Dim oStpoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStpoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
'Origin of the symbol is assumed at the intersection of vertical cylinder axis and nozzles  axes
    oStpoint.Set -(parFacetoFace - parFace2toCenter), 0, 0
    oEnPoint.Set parFace2toCenter, 0, 0
    Set ObjFlowHoriCylin = PlaceCylinder(m_OutputColl, oStpoint, oEnPoint, parFlowDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowHoriCylin
    Set ObjFlowHoriCylin = Nothing

 ' Insert your code for output 4(Vortex Flowmeter Vertical Cylinder)
 'Vertical Cylinder Diameter is assumed as 0.1 times of Flow Diameter
    Dim dVerCylinDia As Double
    dVerCylinDia = 0.1 * parFlowDiameter
    oStpoint.Set 0, Sqr((parFlowDiameter / 2) * _
                                            (parFlowDiameter / 2) - _
                                            (dVerCylinDia / 2) * _
                                            (dVerCylinDia / 2)), 0
    oEnPoint.Set 0, parInstrumentHeight - parInstrumentHeight2 / 2, 0
    Set ObjVerticalCylin = PlaceCylinder(m_OutputColl, oStpoint, oEnPoint, dVerCylinDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerticalCylin
    Set ObjVerticalCylin = Nothing
    
' Insert your code for output 5(Horizontal Instrument Box)
    oStpoint.Set -parInstrumentLength, _
                       parInstrumentHeight - parInstrumentHeight2 / 2, _
                       parInstrumentWidth / 2
    oEnPoint.Set parInstrumentLength1, _
                   parInstrumentHeight + parInstrumentHeight2 / 2, _
                       -parInstrumentWidth / 2
    Set ObjInstruHorBox = PlaceBox(m_OutputColl, oStpoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstruHorBox
    Set ObjInstruHorBox = Nothing

' Insert your code for output 6(Vortex Flowmeter Vertical Cylinder -ive Z-axis direction)
    oStpoint.Set 0, -Sqr((parFlowDiameter / 2) * _
                                            (parFlowDiameter / 2) - _
                                            (dVerCylinDia / 2) * _
                                            (dVerCylinDia / 2)), 0
    oEnPoint.Set 0, -parInstrumentHeight1 + parInstrumentHeight3 / 2, 0
    Set ObjVerticalCylin1 = PlaceCylinder(m_OutputColl, oStpoint, oEnPoint, dVerCylinDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerticalCylin1
    Set ObjVerticalCylin1 = Nothing
    
' Insert your code for output 7(Instrument Box -ive Z-axis Side)
    oStpoint.Set -parInstrumentLength2, _
                     -parInstrumentHeight1 - parInstrumentHeight3 / 2, _
                       parInstrumentWidth1 / 2
    oEnPoint.Set parInstrumentLength3, _
                     -parInstrumentHeight1 + parInstrumentHeight3 / 2, _
                        -parInstrumentWidth1 / 2
    Set ObjInstruHorBox1 = PlaceBox(m_OutputColl, oStpoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstruHorBox1
    Set ObjInstruHorBox1 = Nothing
    Set oStpoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
