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
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S53A  Cone Strainer Symbol.
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
    Dim ObjCollarIns As Object
    
    Dim parCollarThickness As Double
    Dim parBasketLength As Double
    Dim parBasketEndDiameter As Double
    Dim parCollarEndtoCenter As Double
    Dim parInsulationThickness As Double
    Dim CollarExtnWidth As Double
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
    parCollarThickness = arrayOfInputs(2)   'A
    parBasketLength = arrayOfInputs(3)      'B
    parBasketEndDiameter = arrayOfInputs(4) 'C
    parCollarEndtoCenter = arrayOfInputs(5) 'D
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

 ' Insert your code for output 6(CollarIns)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    Dim dInsulationDiameter As Double
    
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                                
'   Compute largest of the Pipe outside diameter, check if flange diameter is zero
    Dim POD As Double
    Dim flangeDia As Double
    Dim fdmin As Double
    Dim fd As Double
    
    POD = BodyOD(pipeDiam(parPortIndex1), pipeDiam(parPortIndex2))
    flangeDia = BodyOD(flangeDiam(parPortIndex1), flangeDiam(parPortIndex2))
    If flangeDia = 0 Then
        fd = POD * 1.7
    Else
        fdmin = POD * 1.7
        fd = BodyOD(flangeDiam(parPortIndex1), fdmin)
    End If

    dInsulationDiameter = fd + parInsulationThickness * 2
    startPt.Set -parCollarThickness / 2 - parInsulationThickness, 0, 0
    endPt.Set parCollarThickness / 2 + parInsulationThickness, 0, 0
    Set ObjCollarIns = PlaceCylinder(m_OutputColl, startPt, endPt, dInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarIns
    Set ObjCollarIns = Nothing
    
    Set m_oCodeListMetadata = Nothing
    
 ' Insert your code for output 7(CollarIns)
'    CollarExtnWidth = 0.5 * pipeDiam(parPortIndex1)
'    As per EDEN (S53A) CollarExtnWidth is as follows:
    CollarExtnWidth = parCollarThickness * 2
    If CollarExtnWidth > fd Then CollarExtnWidth = fd / 2
    
    startPt.Set -parCollarThickness / 2 - parInsulationThickness, Sqr((fd / 2) ^ 2 - (CollarExtnWidth / 2) ^ 2) - parInsulationThickness, _
                    -CollarExtnWidth / 2 - parInsulationThickness
    endPt.Set parCollarThickness / 2 + parInsulationThickness, parCollarEndtoCenter + parInsulationThickness, _
                    CollarExtnWidth / 2 + parInsulationThickness
    Dim ObjCollarExtnIns As Object
    Set ObjCollarExtnIns = PlaceBox(m_OutputColl, startPt, endPt)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarExtnIns
    Set ObjCollarExtnIns = Nothing
    Set startPt = Nothing
    Set endPt = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
