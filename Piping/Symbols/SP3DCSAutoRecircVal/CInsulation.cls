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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:           ACM
'   Creation Date:  Friday, Jan 24 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    It's an Auto Recirculation Valve (PDS Symbol) Custon Speciality Symbol.
'    It has Nine Outputs among them Three are nozzles and Two are insulation bodies.
'    The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
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
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
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
        
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    parNpd = arrayOfInputs(7)
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
    parPortIndex3 = arrayOfInputs(27)
    parNpd3 = arrayOfInputs(28)
    parEndPreparation3 = arrayOfInputs(29)
    parScheduleThickness3 = arrayOfInputs(30)
    parEndStandard3 = arrayOfInputs(31)
    parPressureRating3 = arrayOfInputs(32)
    parFlowDirection3 = arrayOfInputs(33)
    parId1 = arrayOfInputs(34)
    parId2 = arrayOfInputs(35)
    parId3 = arrayOfInputs(36)
    parNpdUnitType = arrayOfInputs(37)
    parNpdUnitType1 = arrayOfInputs(38)
    parNpdUnitType2 = arrayOfInputs(39)
    parNpdUnitType3 = arrayOfInputs(40)

    
    iOutput = 0
       
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

 ' Insert your code for output 4(Insulation Cylinder)
 
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
    oStPoint.Set -parFace1toCenter, 0, 0
    oEndPoint.Set parFace2toCenter, 0, 0
    
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
    oEndPoint.Set 0, parFace3toCenter, 0
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
    Resume Next
    
End Sub
