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
'   Creation Date:  Tursday, Dec 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT3,
'   PDMS Category Name=IVPP). Physical aspect consist of Relief valve body of Box shape, 1 Cylinder
'   Nozzle-1 facing (-)ive side of X-Axis and Nozzle-2 facing (-)ive side of Y-axis.
'   Insulation aspect consist of a single box place covering the whole valve.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
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
    
    Dim ObjRefValBody As Object
    Dim ObjVerCylin As Object
        
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
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
    
    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)            ''D2
    parFace2toCenter = arrayOfInputs(3)            ''D3
    parInstrumentHeight = arrayOfInputs(4)          'D4
    parInstrumentHeight1 = arrayOfInputs(5)         'D5
    parInstrumentLength = arrayOfInputs(6)          'D6
    parInstrumentLength1 = arrayOfInputs(7)         'D7
    parInstrumentWidth = arrayOfInputs(8)           'D8
    parInstrumentWidth1 = arrayOfInputs(9)          'D9
    parInstrumentHeight2 = arrayOfInputs(10)        'D10
    parInstrumentDiameter = arrayOfInputs(11)       'D11
    parInsulationThickness = arrayOfInputs(12)
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
    'Operation Aspect
    parOperationLength = arrayOfInputs(33)
    parOperationHeight = arrayOfInputs(34)
    parOperationWidth = arrayOfInputs(35)
    parOperationWidth1 = arrayOfInputs(36)
    'Text Inputs
    parId1 = arrayOfInputs(37)
    parId2 = arrayOfInputs(38)
    parNPDUnitType = arrayOfInputs(39)
    parNPDUnitType1 = arrayOfInputs(40)
    parNPDUnitType2 = arrayOfInputs(41)
    

    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output 1(Relief Valve Body)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    stPoint.Set parInstrumentLength, parInstrumentHeight1, -parInstrumentWidth
    enPoint.Set -parInstrumentLength1, -parInstrumentHeight, parInstrumentWidth1
    Set ObjRefValBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefValBody
    Set ObjRefValBody = Nothing
    
' Insert your code for output 2 (Vertical Cylinder)
    
    stPoint.Set 0, parInstrumentHeight1, 0
    enPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight2, 0
    
    Set ObjVerCylin = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin
    Set ObjVerCylin = Nothing

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
 
    oNozzle.Length = parFace1toCenter - parInstrumentLength1       'D2-D7
    
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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

    oNozzle.Length = parFace2toCenter - parInstrumentHeight       'D3-D4
' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset(parPortIndex2) + depth(parPortIndex2), 0                       'D3
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

