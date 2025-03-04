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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Jan 23 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on PDS Instrument Specialty Symbol SN=I16AZ,
'   MC=MC=I16AT or I16AS.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav               Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC                    DI-95670 Replace names with initials in all revision history sheets and symbols
'   25.Sep.2006     KKC                    TR-102869 Modified symbol code to use both Face to Face
'                                          and Face to Center Dimensions
'  2.Jul.2007       KKC                     TR-123363: Corrected the symbol not to fail the V6.1 Instances
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

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
    Dim ObjInsulCylinder As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    
'   Dynamic Port Input parameters variable declaration
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
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double

    Dim parHandwheelAngle   As Double
    Dim lPartDataBasis As Long
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
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
    parPortIndex2 = arrayOfInputs(18)
    parNpd2 = arrayOfInputs(19)
    parEndPreparation2 = arrayOfInputs(20)
    parScheduleThickness2 = arrayOfInputs(21)
    parEndStandard2 = arrayOfInputs(22)
    parPressureRating2 = arrayOfInputs(23)
    parFlowDirection2 = arrayOfInputs(24)
'    parFacetoCenter = arrayOfInputs(25)
'    parId1 = arrayOfInputs(26)
'    parId2 = arrayOfInputs(27)
'    parNPDUnitType = arrayOfInputs(28)
'    parNPDUnitType1 = arrayOfInputs(29)
'    parNPDUnitType2 = arrayOfInputs(30)
'   Check to ensure both V6.1 and V7 symbol definitions work
    If UBound(arrayOfInputs) >= 30 Then
        'In V7, Index 26 onwards are textual inputs
        parId1 = arrayOfInputs(26)
        parId2 = arrayOfInputs(27)
        parNPDUnitType = arrayOfInputs(28)
        parNPDUnitType1 = arrayOfInputs(29)
        parNPDUnitType2 = arrayOfInputs(30)
    Else
        'In V6.1,Index 25 onwards are textual inputs
        parId1 = arrayOfInputs(25)
        parId2 = arrayOfInputs(26)
        parNPDUnitType = arrayOfInputs(27)
        parNPDUnitType1 = arrayOfInputs(28)
        parNPDUnitType2 = arrayOfInputs(29)
    End If
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then
       parFacetoFace = arrayOfInputs(2)
    ElseIf lPartDataBasis = 10 Then
       parFacetoCenter = arrayOfInputs(25)
       parFacetoFace = 2 * parFacetoCenter
    End If
    
 ' Insert your code for output 9(Insul Cylinder)
     CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0


'   Maximum of pipeDiam, flangeDiam of ports 1 and 2 is  used to compute insulation diameter
    InsulationDia = ReturnMax4(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1), _
                            pipeDiam(parPortIndex2), flangeDiam(parPortIndex2)) + _
                            2 * parInsulationThickness
                            
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub
