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
'   Creation Date:  Tuesday, Dec 19 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG1,
'   PDMS Category Name=IVRF). This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. Physical aspect consist of Two cones one converging other diverging
'   and a vertical Cylinder with varying diameter, two nozzles one facing -ive direction direction of
'   X-axis and the other in +ive direction of X-axis.
'   Insulation aspect consist of one Cylinder along X-axis and another cylinder along Y-axis
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
    Dim pipeDiam(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjConeIns As Object
    Dim ObjRegVerCylinIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
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
    parFaceToFace = arrayOfInputs(3)        'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight = arrayOfInputs(5)      'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentDiameter = arrayOfInputs(8)     'D8
    parInstrumentDiameter1 = arrayOfInputs(9)        'D9
    parInsulationThickness = arrayOfInputs(10)          'Insulation
    parNPD = arrayOfInputs(11)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(12)
    parScheduleThickness = arrayOfInputs(13)
    parEndStandard = arrayOfInputs(14)
    parPressureRating = arrayOfInputs(15)
    parFlowDirection = arrayOfInputs(16)
    parPortIndex1 = arrayOfInputs(17)
    parNpd1 = arrayOfInputs(18)
    parEndPreparation1 = arrayOfInputs(19)
    parScheduleThickness1 = arrayOfInputs(20)
    parEndStandard1 = arrayOfInputs(21)
    parPressureRating1 = arrayOfInputs(22)
    parFlowDirection1 = arrayOfInputs(23)
    parPortIndex2 = arrayOfInputs(24)
    parNpd2 = arrayOfInputs(25)
    parEndPreparation2 = arrayOfInputs(26)
    parScheduleThickness2 = arrayOfInputs(27)
    parEndStandard2 = arrayOfInputs(28)
    parPressureRating2 = arrayOfInputs(29)
    parFlowDirection2 = arrayOfInputs(30)
    parId1 = arrayOfInputs(31)
    parId2 = arrayOfInputs(32)
    parNPDUnitType = arrayOfInputs(33)
    parNPDUnitType1 = arrayOfInputs(34)
    parNPDUnitType2 = arrayOfInputs(35)
    
    iOutput = 0

 ' Insert your code for output 6(Horizontal Cylinder Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Dim insulationDia As Double
    Dim MaxDia As Double
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
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
        
    insulationDia = MaxDia + 2 * parInsulationThickness
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set parFaceToFace - parFace1toCenter, 0, 0
    Set ObjConeIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeIns
    Set ObjConeIns = Nothing
    
' Insert your code for output 6(Vertical Cylinder )
   'Bottom Cylinder Diameter is taken as equal to Pipe 1 Diameter as per EDEN
   Dim CylDia As Double
    CylDia = pipeDiam(parPortIndex1)
    MaxDia = parInstrumentDiameter1
    If MaxDia < parInstrumentDiameter Then MaxDia = parInstrumentDiameter
    If MaxDia < CylDia Then MaxDia = CylDia
    startPoint.Set 0, -parCylHeight - parInsulationThickness, 0
    endPoint.Set 0, parInstrumentHeight + parInsulationThickness, 0
    insulationDia = MaxDia + 2 * parInsulationThickness
    Set ObjRegVerCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRegVerCylinIns
    Set ObjRegVerCylinIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
