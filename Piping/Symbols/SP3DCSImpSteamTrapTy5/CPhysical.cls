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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: S64BC, Page No. D-106 PDS Piping Component Data Reference Guide.
' SP3DImpSteamTrapTy5 Symbol is created with Eight Outputs - Five are Insulation aspect outputs.
' Other three are physical aspect outputs created as follows:
'   ObjTrapBody - Using 'PlaceCylinder' function and
'   two ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   14.Sep.2004     MS     TR 63296 Corrected the arrayofInputs index number for PressureRating of Port1
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
    Dim ObjTrapBody  As Object
        
    Dim parNozz1toCentre As Double          'D2
    Dim parNozz2toCentre As Double          'D3
    Dim parOffset1 As Double                    'D4
    Dim parOffset2 As Double                    'D5
    Dim parTrapBodyDiameter As Double      'D6
    Dim parTrapHeightBottoCen As Double   'D7
    Dim parTrapHeightToptoCen As Double   'D8
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

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozz1toCentre = arrayOfInputs(2)  'P2
    parNozz2toCentre = arrayOfInputs(3)  'P3
    parOffset1 = arrayOfInputs(4)       'P4
    parOffset2 = arrayOfInputs(5)       'P5
    parTrapBodyDiameter = arrayOfInputs(6)      'P6
    parTrapHeightBottoCen = arrayOfInputs(7)     'P7
    parTrapHeightToptoCen = arrayOfInputs(8)     'P8
    parInsulationThickness = arrayOfInputs(9)
    parNPD = arrayOfInputs(10)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(11)
    parScheduleThickness = arrayOfInputs(12)
    parEndStandard = arrayOfInputs(13)
    parPressureRating = arrayOfInputs(14)
    parFlowDirection = arrayOfInputs(15)
    parPortIndex1 = arrayOfInputs(16)
    parNpd1 = arrayOfInputs(17)
    parEndPreparation1 = arrayOfInputs(18)
    parScheduleThickness1 = arrayOfInputs(19)
    parEndStandard1 = arrayOfInputs(20)
    parPressureRating1 = arrayOfInputs(21)
    parFlowDirection1 = arrayOfInputs(22)
    parPortIndex2 = arrayOfInputs(23)
    parNpd2 = arrayOfInputs(24)
    parEndPreparation2 = arrayOfInputs(25)
    parScheduleThickness2 = arrayOfInputs(26)
    parEndStandard2 = arrayOfInputs(27)
    parPressureRating2 = arrayOfInputs(28)
    parFlowDirection2 = arrayOfInputs(29)
    'Text Inputs
    parId1 = arrayOfInputs(30)
    parId2 = arrayOfInputs(31)
    parNPDUnitType = arrayOfInputs(32)
    parNPDUnitType1 = arrayOfInputs(33)
    parNPDUnitType2 = arrayOfInputs(34)
    
    iOutput = 0

' Insert your code for output 4(Body of Trap)
    stPoint.Set 0, -parTrapHeightBottoCen, 0
    enPoint.Set 0, parTrapHeightToptoCen, 0
    Set ObjTrapBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTrapBodyDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrapBody
    Set ObjTrapBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
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
    
    oNozzle.Length = parNozz1toCentre
     
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parNozz1toCentre - sptOffset(parPortIndex1) + depth(parPortIndex1), -parOffset1, 0
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
    
    oNozzle.Length = parNozz2toCentre

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parNozz2toCentre - sptOffset(parPortIndex2) + depth(parPortIndex2), parOffset2, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

