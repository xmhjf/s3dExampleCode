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
' Symbol Model No.: S51AC or S51AZ, Page No. D-105 PDS Piping Component Data Reference Guide.
' Symbol is created with 13 Outputs - Eight are Insulation aspect outputs.
' Other Five are physical aspect outputs created as follows:
'   ObjBodyBranch, ObjBodyStrainer - Using 'PlaceCylinder' function and
'   three ObjNozzle objects  by using 'CreateNozzleWithLength' function.
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
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


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
    Dim ObjBodyBranch As Object
    Dim ObjBodyStrainer As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parStrToptoCenter As Double
    Dim parStrDiameter As Double
    Dim parBranchDiameter As Double
    Dim parOffset As Double
    Dim parStrWidth As Double
    Dim parFace3toStrTop As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint                 As New AutoMath.DPosition
    Dim enPoint                 As New AutoMath.DPosition
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


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'P2
    parFace2toCenter = arrayOfInputs(3)  'P3
    parStrToptoCenter = arrayOfInputs(4) 'P4
    parStrDiameter = arrayOfInputs(5)    'P5
    parBranchDiameter = arrayOfInputs(6) 'P6
    parOffset = arrayOfInputs(7)         'P7
    parStrWidth = arrayOfInputs(8)       'P8
    parFace3toStrTop = arrayOfInputs(9)  'P9
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
    parPortIndex3 = arrayOfInputs(30)
    parNpd3 = arrayOfInputs(31)
    parEndPreparation3 = arrayOfInputs(32)
    parScheduleThickness3 = arrayOfInputs(33)
    parEndStandard3 = arrayOfInputs(34)
    parPressureRating3 = arrayOfInputs(35)
    parFlowDirection3 = arrayOfInputs(36)
    parInsulationThickness = arrayOfInputs(37)
    'Text Inputs
    parId1 = arrayOfInputs(38)
    parId2 = arrayOfInputs(39)
    parId3 = arrayOfInputs(40)
    parNPDUnitType = arrayOfInputs(41)
    parNPDUnitType1 = arrayOfInputs(42)
    parNPDUnitType2 = arrayOfInputs(43)
    parNPDUnitType3 = arrayOfInputs(44)
    
    iOutput = 0
    
' Insert your code for output 9(Body of Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set (parStrToptoCenter - parStrWidth) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth) * Sin(PI / 4), 0
    Set ObjBodyBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBranchDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyBranch
    Set ObjBodyBranch = Nothing

' Insert your code for output 10(Body of Striner)
    stPoint.Set (parStrToptoCenter - parStrWidth) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth) * Sin(PI / 4), 0
    enPoint.Set parStrToptoCenter * Cos(PI / 4), parStrToptoCenter * Sin(PI / 4), 0
    Set ObjBodyStrainer = PlaceCylinder(m_OutputColl, stPoint, enPoint, parStrDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyStrainer
    Set ObjBodyStrainer = Nothing
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
    
    oNozzle.Length = parFace1toCenter
     
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFace1toCenter + sptOffset(parPortIndex1) - depth(parPortIndex1), 0, 0
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
    
    oNozzle.Length = parFace2toCenter

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
    
    oNozzle.Length = parFace3toStrTop

' Direction of the Nozzle
    oDir.Set Cos(PI / 4), Sin(PI / 4), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parStrToptoCenter + parFace3toStrTop + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Cos(PI / 4) - parOffset * Sin(PI / 4), _
                    (parStrToptoCenter + parFace3toStrTop + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Sin(PI / 4) + parOffset * Cos(PI / 4), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
