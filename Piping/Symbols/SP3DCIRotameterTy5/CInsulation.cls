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
'   Author:          SSP
'   Creation Date:  Friday, Dec 27 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR5,
' PDMS Category Name=IFNK). Physical aspect consist of one vertical cylinder, One Horizontal Box
' and two nozzles, one facing -ive and the other facing +ive direction of X-axis.
' Insulation aspect consist of a cylinder along Y-axis and the Box along X-axis
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim ObjVerCylinIns As Object
    Dim ObjHoriBoxIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
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
    parInstrumentHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight1 = arrayOfInputs(5)      'D5
    parInstrumentDiameter = arrayOfInputs(6)      'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentHeight3 = arrayOfInputs(8)      'D8
    parInstrumentWidth = arrayOfInputs(9)        'D9
    parInstrumentWidth1 = arrayOfInputs(10)      'D10
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
    'Text Inputs
    parId1 = arrayOfInputs(32)
    parId2 = arrayOfInputs(33)
    parNPDUnitType = arrayOfInputs(34)
    parNPDUnitType1 = arrayOfInputs(35)
    parNPDUnitType2 = arrayOfInputs(36)
    
    iOutput = 0

 ' Insert your code for output 3(Horizontal Box Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parFace1toCenter - parInsulationThickness, _
                            -parInstrumentHeight2 - parInsulationThickness, _
                            parInstrumentWidth + parInsulationThickness
    endPoint.Set parFace2toCenter + parInsulationThickness, _
                            parInstrumentHeight3 + parInsulationThickness, _
                            -parInstrumentWidth1 - parInsulationThickness
    Set ObjHoriBoxIns = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriBoxIns
    Set ObjHoriBoxIns = Nothing
    
' Insert your code for output 4(Vertical Cylinder Insulation)
    startPoint.Set 0, -parInstrumentHeight - parInsulationThickness, 0
    endPoint.Set 0, parInstrumentHeight1 - parInstrumentHeight + parInsulationThickness, 0
    Set ObjVerCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            parInstrumentDiameter + 2 * parInsulationThickness, True)
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
