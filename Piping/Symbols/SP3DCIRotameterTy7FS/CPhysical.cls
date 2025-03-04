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
'   Author:          MS
'   Creation Date:  Thursday, 2 Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR7, PDMS Category Name=IFNM)
'   The symbol consists of six outputs, four physical and two insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
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
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptoffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
   
    Dim iOutput     As Double
    
    Dim ObjRotameterBody As Object
    Dim ObjFlowSwitchBody As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parFace1toCenter As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentWidth As Double
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
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)            'D2
    parFlowDiameter = arrayOfInputs(3)          'D3
    parFace1toCenter = arrayOfInputs(4)         'D4
    parInstrumentDiameter = arrayOfInputs(5)    'D5
    parInstrumentWidth = arrayOfInputs(6)       'D6
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
'Comment to make test with text inputs
    parId1 = arrayOfInputs(28)
    parId2 = arrayOfInputs(29)
    parNPDUnitType = arrayOfInputs(30)
    parNPDUnitType1 = arrayOfInputs(31)
    parNPDUnitType2 = arrayOfInputs(32)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' The origin is considered to be along the flow axis and at parFace1toCenter from nozzle 1
 ' Insert your code for output 1(Rotameter Body)
    stPoint.Set 0, -parFace1toCenter, 0
    enPoint.Set 0, parFacetoFace - parFace1toCenter, 0
    Set ObjRotameterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotameterBody
    Set ObjRotameterBody = Nothing
    
 ' Insert your code for output 2(Flow Switch Body)
    stPoint.Set 0, 0, parInstrumentWidth
    enPoint.Set 0, 0, parFlowDiameter / 2
    Set ObjFlowSwitchBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowSwitchBody
    Set ObjFlowSwitchBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
'CreateRetrieveDynamicNozzle
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptoffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace1toCenter - sptoffset(parPortIndex1) + depth(parPortIndex1), 0
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
                                            flangeDiam, sptoffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex2)

' Direction of the Nozzle
    oDir.Set 0, 1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, parFacetoFace - parFace1toCenter + sptoffset(parPortIndex2) - depth(parPortIndex2), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
        
    
End Sub

