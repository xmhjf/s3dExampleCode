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
'   Author:          Babu Govindarajan
'   Creation Date:  Thusday, Dec 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM1,
'   PDMS Category Name=IFNC). This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. Physical aspect consist of one horizontal cylinder, one vertical cylinder,
'   another horizontal cylinder as Instrument cylinder and two nozzles one facing -ive direction direction
'   of X-axis and the other in +ive direction of X-axis.
'   Insulation aspect consist of a cylinders one along X-axis and a box to cover both  vertical  and horizontal
'   Instrument cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    
    Dim ObjVerCylinIns As Object
    Dim ObjHoriCylinIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
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
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)  'D3
    parFlowDiameter = arrayOfInputs(4) 'D4
    parInstrumentHeight = arrayOfInputs(5)       'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentDiameter = arrayOfInputs(8)    'D8
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
    parId1 = arrayOfInputs(30)
    parId2 = arrayOfInputs(31)
    parNPDUnitType = arrayOfInputs(32)
    parNPDUnitType1 = arrayOfInputs(33)
    parNPDUnitType2 = arrayOfInputs(34)
    iOutput = 0

 ' Insert your code for output 6(Horizontal Cylinder Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Dim insulationDia As Double
    Dim MaxDia  As Double
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    'Check for max amoung Pipes, flanges and parFlowDiameter values
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
    
' Insert your code for output 6(Vertical Cylinder and Hori Cylinder Insulation )
    startPoint.Set -parInstrumentLength - parInsulationThickness, _
                            parFlowDiameter / 2, _
                            -parInstrumentDiameter / 2 - parInsulationThickness

    endPoint.Set parInstrumentLength1 + parInsulationThickness, _
                            parInstrumentHeight + parInstrumentDiameter / 2 + parInsulationThickness, _
                            parInstrumentDiameter / 2 + parInsulationThickness
                            
    Set ObjVerCylinIns = PlaceBox(m_OutputColl, startPoint, endPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylinIns
    Set ObjVerCylinIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
