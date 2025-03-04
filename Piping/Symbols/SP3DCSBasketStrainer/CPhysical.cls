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
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details are available in PDS Piping Data Reference Guide at page no D-106 and SN=S53AC
'   and S53AZ
'   Consist of Physical and Insulation Aspects. Physical aspect is made up of Two Nozzles, One Collar of Cylinder,
'   one Collar extension of Box and Filter body of truncated cone. Insulation aspect consist of a Cylinder
'   for Collar Insulation.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   26.OCT.2004     MS     TR 65841 Corrected the code to generate Collar Cylinder. Corrected the code to be as per PDS Eden
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()

''''

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
    Dim ObjStrainer As Object
    Dim ObjCollar As Object
    Dim ObjCollarExtn As Object
    Dim ObjCollarIns As Object
    
    Dim parCollarThickness As Double
    Dim parBasketLength As Double
    Dim parBasketEndDiameter As Double
    Dim parCollarEndtoCenter As Double
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
    parCollarThickness = arrayOfInputs(2)   'P2
    parBasketLength = arrayOfInputs(3)          'P3
    parBasketEndDiameter = arrayOfInputs(4)     'P4
    parCollarEndtoCenter = arrayOfInputs(5)         'P5
    parInsulationThickness = arrayOfInputs(6)
    parNPD = arrayOfInputs(7)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(8)
    parScheduleThickness = arrayOfInputs(9)
    parEndStandard = arrayOfInputs(10)
    parPressureRating = arrayOfInputs(11)
    parFlowDirection = arrayOfInputs(12)
    parPortIndex1 = arrayOfInputs(13)
    parNpd1 = arrayOfInputs(14)
    parEndPreparation1 = arrayOfInputs(15)
    parScheduleThickness1 = arrayOfInputs(16)
    parEndStandard1 = arrayOfInputs(17)
    parPressureRating1 = arrayOfInputs(18)
    parFlowDirection1 = arrayOfInputs(19)
    parPortIndex2 = arrayOfInputs(20)
    parNpd2 = arrayOfInputs(21)
    parEndPreparation2 = arrayOfInputs(22)
    parScheduleThickness2 = arrayOfInputs(23)
    parEndStandard2 = arrayOfInputs(24)
    parPressureRating2 = arrayOfInputs(25)
    parFlowDirection2 = arrayOfInputs(26)
    'Text Inputs
    parId1 = arrayOfInputs(27)
    parId2 = arrayOfInputs(28)
    parNPDUnitType = arrayOfInputs(29)
    parNPDUnitType1 = arrayOfInputs(30)
    parNPDUnitType2 = arrayOfInputs(31)
    
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
    oPlacePoint.Set -parCollarThickness / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    oPlacePoint.Set parCollarThickness / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
 ' Insert your code for output 3(Strainer)
    Dim baseCenPt As AutoMath.DPosition
    Dim topCenPt As AutoMath.DPosition
    Set baseCenPt = New AutoMath.DPosition
    Set topCenPt = New AutoMath.DPosition
        

    baseCenPt.Set -parCollarThickness / 2, 0, 0
    topCenPt.Set parCollarThickness / 2 + parBasketLength, 0, 0
    Set ObjStrainer = PlaceCone(m_OutputColl, baseCenPt, topCenPt, pipeDiam(parPortIndex1) / 2, parBasketEndDiameter / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainer
    Set ObjStrainer = Nothing
    
    Set baseCenPt = Nothing
    Set topCenPt = Nothing
 
 ' Insert your code for output 4(Collar)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set -parCollarThickness / 2, 0, 0
    endPt.Set parCollarThickness / 2, 0, 0
    
    'If the flangeDiam(parPortIndex1) is zero, then set the collar diameter as pipediam
    Dim CollarDiameter As Double
    CollarDiameter = flangeDiam(parPortIndex1)
    If flangeDiam(parPortIndex1) = 0 Then CollarDiameter = pipeDiam(parPortIndex1)
    
    Set ObjCollar = PlaceCylinder(m_OutputColl, startPt, endPt, CollarDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollar
    Set ObjCollar = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
 
 ' Insert your code for output 5(CollarExtn)
    Dim lowerPt As AutoMath.DPosition
    Dim upperPt As AutoMath.DPosition
    Set lowerPt = New AutoMath.DPosition
    Set upperPt = New AutoMath.DPosition
    
    Dim factor As Double
    ' The factors have been taken similar to that in the PDS Eden code. S53AC
    
    If (pipeDiam(parPortIndex1) < 0.1) Then  'Assuming 3in pipediam not more than 0.1m and 10in pipediam not more than 0.3m
        factor = 1.25
        If flangeDiam(parPortIndex1) = 0 Then factor = 0.75
    Else
        If (pipeDiam(parPortIndex1) > 0.1 And pipeDiam(parPortIndex1) < 0.3) Then
            factor = 0.55
        Else
            factor = 0.3
        End If
    End If
    
    
    Dim CollarExtnWidth As Double
    Dim CollarMeetHeight As Double
    CollarExtnWidth = pipeDiam(parPortIndex1) * factor
    
    CollarMeetHeight = Sqr(CollarDiameter / 2 * CollarDiameter / 2 - CollarExtnWidth / 2 * CollarExtnWidth / 2)
    
    lowerPt.Set -parCollarThickness / 2, CollarMeetHeight, -CollarExtnWidth / 2
    upperPt.Set parCollarThickness / 2, parCollarEndtoCenter, CollarExtnWidth / 2
    
    Set ObjCollarExtn = PlaceBox(m_OutputColl, lowerPt, upperPt)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarExtn
    Set ObjCollarExtn = Nothing
    Set lowerPt = Nothing
    Set upperPt = Nothing

    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
