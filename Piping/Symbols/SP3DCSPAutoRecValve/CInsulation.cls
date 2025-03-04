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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly I30A  Auto Recirculation Valve Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
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
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object
    Dim ObjInsulBranch As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
       
' Dynamic Port Input parameters variable declaration
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
    parFacetoCenter = arrayOfInputs(2)  'A
    parFace1toCenter = arrayOfInputs(3) 'B
    parFace2toCenter = arrayOfInputs(4) 'C
    parInsulationThickness = arrayOfInputs(5)
    
    parNPD = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
    
    parPortIndex3 = arrayOfInputs(26)
    parNpd3 = arrayOfInputs(27)
    parEndPreparation3 = arrayOfInputs(28)
    parScheduleThickness3 = arrayOfInputs(29)
    parEndStandard3 = arrayOfInputs(30)
    parPressureRating3 = arrayOfInputs(31)
    parFlowDirection3 = arrayOfInputs(32)
    
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parId3 = arrayOfInputs(35)
    parNPDUnitType = arrayOfInputs(36)
    parNPDUnitType1 = arrayOfInputs(37)
    parNPDUnitType2 = arrayOfInputs(38)
    parNPDUnitType3 = arrayOfInputs(39)
    
    iOutput = 0
       
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

 ' Insert your code for output 7(Insulation Cylinder)
 
   CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                        parEndPreparation, parScheduleThickness, parEndStandard, _
                                        parPressureRating, parFlowDirection, parNpd1, _
                                        parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                        parEndStandard1, parPressureRating1, parFlowDirection1, _
                                        parId1, m_OutputColl, m_oCodeListMetadata, _
                                        False, pipeDiam, flangeThick, _
                                        flangeDiam, sptOffset, depth
                                        
 
    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim dInsulationDia As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    oStPoint.Set -parFacetoCenter, 0, 0
    oEndPoint.Set parFace1toCenter, 0, 0
    
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        dInsulationDia1 = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    Else
        dInsulationDia1 = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        dInsulationDia2 = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    Else
        dInsulationDia2 = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    If dInsulationDia1 > dInsulationDia2 Then
        dInsulationDia = dInsulationDia1
    Else
        dInsulationDia = dInsulationDia2
    End If
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, dInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    
 ' Insert your code for output 5(Insulation Cylinder)
 
    If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then
        dInsulationDia = flangeDiam(parPortIndex3) + parInsulationThickness * 2
    Else
        dInsulationDia = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    End If
    oStPoint.Set 0, 0, 0
    oEndPoint.Set 0, parFace2toCenter, 0
    Set ObjInsulBranch = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, dInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulBranch
    Set ObjInsulBranch = Nothing
    
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
       
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
