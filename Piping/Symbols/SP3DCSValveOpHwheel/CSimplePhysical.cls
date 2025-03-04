VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         dkl
'   Creation Date:  Tuesday, Jan 23 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is based on a functionality of a SP3DCSValve with an in-built Handewheel Operator.
'   The in-built operator is based on SP3DOP3.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   23.Jan.2007     dkl    CR-113194. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    Set m_oGeomHelper = New SymbolServices
  
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
    Set m_oGeomHelper = Nothing
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
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
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
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOperatorHeight = arrayOfInputs(4)
    parOperatorDiameter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    parNPD = arrayOfInputs(7)                       'Input parameters for Dynamic Nozzles
    parNPDUnitType = arrayOfInputs(8)
    parEndPreparation = arrayOfInputs(9)
    parScheduleThickness = arrayOfInputs(10)
    parEndStandard = arrayOfInputs(11)
    parPressureRating = arrayOfInputs(12)
    parFlowDirection = arrayOfInputs(13)
    parId1 = arrayOfInputs(14)
    parPortIndex1 = arrayOfInputs(15)
    parNpd1 = arrayOfInputs(16)
    parNPDUnitType1 = arrayOfInputs(17)
    parEndPreparation1 = arrayOfInputs(18)
    parScheduleThickness1 = arrayOfInputs(19)
    parEndStandard1 = arrayOfInputs(20)
    parPressureRating1 = arrayOfInputs(21)
    parFlowDirection1 = arrayOfInputs(22)
    parId2 = arrayOfInputs(23)
    parPortIndex2 = arrayOfInputs(24)
    parNpd2 = arrayOfInputs(25)
    parNPDUnitType2 = arrayOfInputs(26)
    parEndPreparation2 = arrayOfInputs(27)
    parScheduleThickness2 = arrayOfInputs(28)
    parEndStandard2 = arrayOfInputs(29)
    parPressureRating2 = arrayOfInputs(30)
    parFlowDirection2 = arrayOfInputs(31)
    
    iOutput = 0
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    
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
    
    oNozzle.Length = flangeThick(parPortIndex1)

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
    
'Insert your code for output 3(Left Body Cone)
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
    oStPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    oEndPoint.Set 0, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCone arrayOfOutputs(iOutput), oStPoint, oEndPoint, pipeDiam(parPortIndex1), 0

'Insert your code for output 4(Right Body Cone)
    oStPoint.Set parFace2toCenter - flangeThick(parPortIndex2), 0, 0
    oEndPoint.Set 0, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCone arrayOfOutputs(iOutput), oStPoint, oEndPoint, pipeDiam(parPortIndex2), 0
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    
' Code for Operator - Handwheel
    Dim dStemMinimumDiameter As Double
    Dim dActuatorWheelMinimumThickness As Double
    Dim dStemExtension As Double
    
    dStemMinimumDiameter = parOperatorDiameter * 0.1
    ' Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(dStemMinimumDiameter, 0.015) Then dStemMinimumDiameter = 0.015
    If CmpDblGreaterthan(dStemMinimumDiameter, 0.05) Then dStemMinimumDiameter = 0.05
   
    dActuatorWheelMinimumThickness = dStemMinimumDiameter
    dStemExtension = parOperatorHeight * 0.1
    
' Insert your code for output 5(Stem)
    Dim oCylBotOrigin    As AutoMath.DPosition
    Set oCylBotOrigin = New AutoMath.DPosition
    Dim oCylTopOrigin    As AutoMath.DPosition
    Set oCylTopOrigin = New AutoMath.DPosition
    
    oCylBotOrigin.Set 0, 0, 0
    oCylTopOrigin.Set oCylBotOrigin.x, oCylBotOrigin.y + parOperatorHeight, oCylBotOrigin.z
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oCylBotOrigin, oCylTopOrigin, dStemMinimumDiameter
    
' Insert your code for output 6(Handwheel)
    Dim dTorPriRad       As Double
    Dim dTorSecRad       As Double
    Dim oTorOrigin       As AutoMath.DPosition
    Set oTorOrigin = New AutoMath.DPosition
    Dim oTorAxis         As AutoMath.DVector
    Set oTorAxis = New AutoMath.DVector
    Dim oMajorAxis As AutoMath.DVector
    Set oMajorAxis = New AutoMath.DVector
    
    dTorSecRad = dActuatorWheelMinimumThickness / 2
    dTorPriRad = parOperatorDiameter / 2 - dTorSecRad
    oTorAxis.Set 0, 1, 0
    oMajorAxis.Set parOperatorDiameter - dTorSecRad, parOperatorHeight - dStemExtension - dActuatorWheelMinimumThickness, 0
    oTorOrigin.Set oCylTopOrigin.x, oCylTopOrigin.y - dStemExtension - dTorSecRad, oCylTopOrigin.z
   
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateTorus arrayOfOutputs(iOutput), oTorOrigin, oTorAxis, oMajorAxis, dTorPriRad, 360, dTorSecRad * 2
    Set oTorOrigin = Nothing
    Set oTorAxis = Nothing
    Set oCylBotOrigin = Nothing
    Set oCylTopOrigin = Nothing
    Set oMajorAxis = Nothing
    
    Exit Sub
        
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
    
End Sub
