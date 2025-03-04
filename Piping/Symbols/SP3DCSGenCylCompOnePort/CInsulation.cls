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
'   Creation Date:  Tuesday, Nov 06 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who          change description
'   -----------    -----        ------------------
'   06.Nov.2007     KKC       CR-126830  Provide one port cylindrical & rectangular on-the-fly piping specialty symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const NEGLIGIBLE_THICKNESS = 0.0001


Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)     As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoEnd As Double
    Dim parBodyDiameter As Double
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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parBodyDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parId1 = arrayOfInputs(18)
    parNPDUnitType = arrayOfInputs(19)
    parNPDUnitType1 = arrayOfInputs(20)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
     CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                            
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -parFacetoEnd / 2, 0, 0
    oEnPoint.Set parFacetoEnd / 2 + parInsulationThickness, 0, 0
    
    Dim InsulationDia As Double
    InsulationDia = parBodyDiameter + 2 * parInsulationThickness
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
' Place Insulation for the Port
    If CmpDblGreaterthan(flangeDiam(1), LINEAR_TOLERANCE) Then
        oStPoint.Set -parFacetoEnd / 2, 0, 0
        If CmpDblLessThanOrEqualTo(flangeThick(1), 0) Then flangeThick(1) = NEGLIGIBLE_THICKNESS
        oEnPoint.Set -parFacetoEnd / 2 + flangeThick(1) + parInsulationThickness, 0, 0
        Dim dInsulationRadius As Double
        dInsulationRadius = flangeDiam(1) / 2 + parInsulationThickness
        Dim objInsulatedPort As Object
        Set objInsulatedPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            2 * dInsulationRadius, True)
        
        'Set the output
        m_OutputColl.AddOutput "PortInsulation", objInsulatedPort
        Set objInsulatedPort = Nothing
    End If
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
