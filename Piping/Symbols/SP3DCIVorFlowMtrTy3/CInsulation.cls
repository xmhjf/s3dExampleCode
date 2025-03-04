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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Tuesday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM3,
'   PDMS Category Name=IFNE).  This symbol is Custom INtools Instrument with Nozzles are fully Parametric.
'   Physical aspect consist of one horizontal cylinder, one top cylinder  along Y axis, a top box and
'   two nozzles one facing -ive direction direction of X-axis and the other in  positive
'   direction of X-axis. Insulation aspect consist of a cylinders along X-axis and a box to cover both
'   top cylinder and top box
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim iOutput     As Double
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim ObjHoriCylinIns As Object
    Dim ObjTopBoxIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
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
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentHeight = arrayOfInputs(5)  'D5
    parInstrumentHeight1 = arrayOfInputs(6) 'D6
    parInstrumentWidth = arrayOfInputs(7)   'D7
    parInstrumentWidth1 = arrayOfInputs(8)  'D8
    parInstrumentLength = arrayOfInputs(9)  'D9
    parInstrumentLength1 = arrayOfInputs(10) 'D10
    parInsulationThickness = arrayOfInputs(11)
    parNPD = arrayOfInputs(12)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(13)
    parScheduleThickness = arrayOfInputs(14)
    parEndStandard = arrayOfInputs(15)
    parPressureRating = arrayOfInputs(16)
    parFlowDirection = arrayOfInputs(17)
    parPortIndex1 = arrayOfInputs(18)
    parNpd1 = arrayOfInputs(19)
    parEndPreparation1 = arrayOfInputs(20)
    parScheduleThickness1 = arrayOfInputs(21)
    parEndStandard1 = arrayOfInputs(22)
    parPressureRating1 = arrayOfInputs(23)
    parFlowDirection1 = arrayOfInputs(24)
    parPortIndex2 = arrayOfInputs(25)
    parNpd2 = arrayOfInputs(26)
    parEndPreparation2 = arrayOfInputs(27)
    parScheduleThickness2 = arrayOfInputs(28)
    parEndStandard2 = arrayOfInputs(29)
    parPressureRating2 = arrayOfInputs(30)
    parFlowDirection2 = arrayOfInputs(31)
    parId1 = arrayOfInputs(32)
    parId2 = arrayOfInputs(33)
    parNPDUnitType = arrayOfInputs(34)
    parNPDUnitType1 = arrayOfInputs(35)
    parNPDUnitType2 = arrayOfInputs(36)
    iOutput = 0

 ' Insert your code for output 6(Horizontal Cylinder Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Dim insulationDia As Double
    Dim MaxDia As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
'   Check for max among Pipes, flanges and parFlowDiameter values
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    MaxDia = pipeDiam(parPortIndex1)
    If MaxDia < pipeDiam(parPortIndex2) Then MaxDia = pipeDiam(parPortIndex2)
    If MaxDia < flangeDiam(parPortIndex1) Then MaxDia = flangeDiam(parPortIndex1)
    If MaxDia < flangeDiam(parPortIndex2) Then MaxDia = flangeDiam(parPortIndex2)
    If MaxDia < parFlowDiameter Then MaxDia = parFlowDiameter
    
    insulationDia = MaxDia + 2 * parInsulationThickness
    startPoint.Set -parFace1toCenter - parInsulationThickness, 0, 0
    endPoint.Set parFace2toCenter + parInsulationThickness, 0, 0
    Set ObjHoriCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinIns
    Set ObjHoriCylinIns = Nothing
    
' Insert your code for output 6(Top Cylinder and Top Box Insulation )
    startPoint.Set -parInstrumentLength - parInsulationThickness, _
                    parFlowDiameter / 2, _
                    -parInstrumentWidth1 - parInsulationThickness
    endPoint.Set parInstrumentLength1 + parInsulationThickness, _
                 parInstrumentHeight + parInsulationThickness, _
                 parInstrumentWidth + parInsulationThickness
    Set ObjTopBoxIns = PlaceBox(m_OutputColl, startPoint, endPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBoxIns
    Set ObjTopBoxIns = Nothing
    
    Set endPoint = Nothing
    Set startPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
