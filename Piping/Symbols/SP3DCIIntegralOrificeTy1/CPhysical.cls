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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIOR),
'   Symbol consists of 12 outputs in which 1 insulation, 11 physical aspect outputs.
'   Insulation aspect consist of a single cylinder.
'   The nozzles are fully Parametric.
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    
    Dim iOutput     As Double
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim ObjBox4 As Object
    Dim ObjBox5 As Object
    Dim ObjBox6 As Object
    Dim ObjBox7 As Object
    Dim ObjCylinder1 As Object
    Dim ObjCylinder2 As Object
    
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentWidth2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentHeight4 As Double
    Dim parInstrumentHeight5 As Double
    Dim parInstrumentWidth3 As Double
    Dim parInstrumentHeight6 As Double
    Dim parInstrumentHeight7 As Double
    Dim parInstrumentHeight8 As Double
    Dim parInstrumentWidth4 As Double
    Dim parInstrumentOffset As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentLength4 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim stemDia As Double
    Dim ProjVector As New AutoMath.DVector
    Dim height As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 5) As Double
'    Dim geomFactory As New IngrGeom3D.GeometryFactory

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
    parInstrumentWidth = arrayOfInputs(2) ''D2
    parInstrumentWidth1 = arrayOfInputs(3) ''D3
    parInstrumentHeight = arrayOfInputs(4) ''D4
    parInstrumentHeight1 = arrayOfInputs(5) ''D5
    parInstrumentHeight2 = arrayOfInputs(6) ''D6
    parInstrumentWidth2 = arrayOfInputs(7) ''D7
    parInstrumentHeight3 = arrayOfInputs(8) ''D8
    parInstrumentHeight4 = arrayOfInputs(9) ''D9
    parInstrumentHeight5 = arrayOfInputs(10) ''D10
    parInstrumentWidth3 = arrayOfInputs(11) ''D11
    parInstrumentHeight6 = arrayOfInputs(12) ''D12
    parInstrumentHeight7 = arrayOfInputs(13) ''D13
    parInstrumentHeight8 = arrayOfInputs(14) ''D14
    parInstrumentWidth4 = arrayOfInputs(15) ''D15
    parInstrumentOffset = arrayOfInputs(16) ''D16
    parInstrumentDiameter = arrayOfInputs(17) ''D17
    parInstrumentLength = arrayOfInputs(18) ''D18
    parInstrumentLength1 = arrayOfInputs(19) ''D19
    parInstrumentLength2 = arrayOfInputs(20) ''D20
    parInstrumentLength3 = arrayOfInputs(21) 'D21
    parInstrumentLength4 = arrayOfInputs(22) ''D22
    parInsulationThickness = arrayOfInputs(23)
    parNPD = arrayOfInputs(24)                       'Input parameters for Dynamic Nozzles
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
    'Text Inputs
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0
''Origin is taken at the intersection of CP1-CP2 axis and vertical axis.
' Insert your code for output 4(Box1)

    stPoint.Set -parInstrumentLength1 / 2, -parInstrumentHeight, parInstrumentWidth1
    enPoint.Set parInstrumentLength1 / 2, parInstrumentHeight1, -parInstrumentWidth
    Set ObjBox1 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 5(Box2)

     
     stPoint.Set -parInstrumentLength3 / 2, parInstrumentHeight1, -parInstrumentWidth3 / 2
     enPoint.Set parInstrumentLength3 / 2, parInstrumentHeight2 - parInstrumentHeight3, _
                                                parInstrumentWidth3 / 2

      
     Set ObjBox2 = PlaceBox(m_OutputColl, stPoint, enPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing
    
'' Insert your code for output 6(Box3)'
    
    stPoint.Set parInstrumentLength3 / 2, parInstrumentHeight2 - parInstrumentHeight3, parInstrumentWidth2 / 2
    enPoint.Set -parInstrumentLength3 / 2, parInstrumentHeight2, -parInstrumentWidth2 / 2
      
     Set ObjBox3 = PlaceBox(m_OutputColl, stPoint, enPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing
    
' Insert your code for output 7(Box4)
        
    stPoint.Set parInstrumentLength3 / 2, parInstrumentHeight2 + parInstrumentHeight5, parInstrumentWidth2 / 2
    enPoint.Set -parInstrumentLength3 / 2, parInstrumentHeight2, -parInstrumentWidth2 / 2
      
     Set ObjBox4 = PlaceBox(m_OutputColl, stPoint, enPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox4
    Set ObjBox4 = Nothing
    
' Insert your code for output 8(Box5)
    
    stPoint.Set parInstrumentLength3 / 2, parInstrumentHeight2 + parInstrumentHeight5, parInstrumentWidth3 / 2
    enPoint.Set -parInstrumentLength3 / 2, parInstrumentHeight2 + parInstrumentHeight4 - parInstrumentHeight6, _
                                                                -parInstrumentWidth3 / 2
      
     Set ObjBox5 = PlaceBox(m_OutputColl, stPoint, enPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox5
    Set ObjBox5 = Nothing

' Insert your code for output 9(Box6)
    
    stPoint.Set parInstrumentLength3 / 2, parInstrumentHeight2 + parInstrumentHeight4 - parInstrumentHeight6, _
                                                                parInstrumentWidth4 / 2
    enPoint.Set -parInstrumentLength3 / 2, parInstrumentHeight2 + parInstrumentHeight4, -parInstrumentWidth4 / 2
      
     Set ObjBox6 = PlaceBox(m_OutputColl, stPoint, enPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox6
    Set ObjBox6 = Nothing
    
' Insert your code for output 10(Box7)

    stPoint.Set parInstrumentLength4 / 2, parInstrumentHeight2 + parInstrumentHeight4, _
                                                                    parInstrumentWidth4 / 2
    
    enPoint.Set -parInstrumentLength4 / 2, _
                        parInstrumentHeight2 + parInstrumentHeight4 + parInstrumentHeight7, _
                                                                    -parInstrumentWidth4 / 2
    
    Set ObjBox7 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox7
    Set ObjBox7 = Nothing
    
' Insert your code for output 11(Cylinder1)
    Dim MeetRad As Double
    
    stemDia = 0.05
    MeetRad = Sqr((parInstrumentDiameter / 2) ^ 2 - (stemDia / 2) ^ 2)
    
    stPoint.Set 0, parInstrumentHeight2 + parInstrumentHeight4 + parInstrumentHeight7 - parInstrumentHeight8, _
                                                                 -parInstrumentWidth4 / 2
    
    enPoint.Set 0, parInstrumentHeight2 + parInstrumentHeight4 + parInstrumentHeight7 - parInstrumentHeight8, _
                      -(parInstrumentOffset - MeetRad)
                                           
    
    Set ObjCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
    
' Insert your code for output 12(Cylinder2)

    stPoint.Set parInstrumentLength4 / 2, parInstrumentHeight2 + parInstrumentHeight4 + parInstrumentHeight7 - parInstrumentHeight8, _
                                                              -parInstrumentOffset
    
    enPoint.Set -parInstrumentLength4 / 2, parInstrumentHeight2 + parInstrumentHeight4 + parInstrumentHeight7 - parInstrumentHeight8, _
                                                              -parInstrumentOffset
    
    Set ObjCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
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
    
    oNozzle.Length = parInstrumentLength
    
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -(parInstrumentLength + parInstrumentLength1 / 2) - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    
    oNozzle.Length = parInstrumentLength2

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parInstrumentLength2 + parInstrumentLength1 / 2) + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ProjVector = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
