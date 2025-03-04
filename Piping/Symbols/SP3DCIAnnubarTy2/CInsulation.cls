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
'   Author:          svsmylav
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IA2,
'   PDMS Category Name=IFAH). The symbol consists of five oupputs: Three physical,
'   and two insulation aspect outputs.
'
'   Note: This symbol expects the user to place an Olet component on the pipe and this (Annubar) should be placed
'   at the connect point2  of the Olet.The parameters values are to be taken from connect point2 of the Olet.'
'   The port for the symbol is fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjInsVerBody As Object
    Dim ObjInsHorBody As Object
    
    Dim parInstrumentLength As Double
    Dim parFacetoEnd        As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1    As Double
    Dim parInsulationThickness As Double
    
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

    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 17) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
    Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)  'D2
    parFacetoEnd = arrayOfInputs(3)         'D3
    parInstrumentHeight = arrayOfInputs(4)  'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInsulationThickness = arrayOfInputs(6)
    parNPD = arrayOfInputs(7)
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
    parOperationHeight = arrayOfInputs(20)
    parId1 = arrayOfInputs(21)
    parNPDUnitType = arrayOfInputs(22)
    parNPDUnitType1 = arrayOfInputs(23)

    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

' Insert your code for output 4(Insulation for Horizontal Body)
'   Reference: PDS EDEN symbol code(IA2.mg)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    LineStringPoints(0) = parFacetoEnd + parInsulationThickness
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = parFacetoEnd + parInsulationThickness
    LineStringPoints(4) = pipeDiam(parPortIndex1) / 2 + parInsulationThickness
    LineStringPoints(5) = 0
        
'   Point 3
'   Flange thickness = D2 * 0.125 (as per EDEN)
'   NOTE: As per EDEN Olet body Length is taken as D2/2, but in the current symbol code D2 dimension
'   represents lower portion of annubar WITHOUT Olet. So all dimensions from EDEN which are derived
'   based on D2 dimension need to be multiplied by 2.
    Dim modFactor As Integer
    modFactor = 2
    LineStringPoints(6) = parInstrumentLength + (parInstrumentLength * 0.125 * modFactor) + _
                                parInsulationThickness
    LineStringPoints(7) = pipeDiam(parPortIndex1) / 2 + parInsulationThickness
    LineStringPoints(8) = 0
    
'   Point 4
'   Assumption: Flange diameter = pipeDiam * 2.52
    LineStringPoints(9) = parInstrumentLength + (parInstrumentLength * 0.125 * modFactor) + _
                                parInsulationThickness
    LineStringPoints(10) = (pipeDiam(parPortIndex1) / 2) * 2.52 + parInsulationThickness
    LineStringPoints(11) = 0
    
'   Point 5
    LineStringPoints(12) = 0
    LineStringPoints(13) = (pipeDiam(parPortIndex1) / 2) * 2.52 + parInsulationThickness
    LineStringPoints(14) = 0
    
'   Point 6
    LineStringPoints(15) = 0
    LineStringPoints(16) = 0
    LineStringPoints(17) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set (parFacetoEnd) / 2, 0, 0
    Set ObjInsHorBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHorBody
    Set ObjInsHorBody = Nothing
    
' Insert your code for output 5(Insulation for Vertical Body)
'   Assumptions:
'       Tee height =  pipeDiam * 1.25
'       Valve height = D5 * 0.2
'       Valve cone diameter =  0.015

    Dim TeeHeight As Double
    Dim ValveHeight As Double
    Dim ValveConeDiameter As Double
    
    TeeHeight = pipeDiam(parPortIndex1) * 1.25
    ValveHeight = parInstrumentHeight1 * 0.2
    ValveConeDiameter = 0.015
    
'   Point 1
    LineStringPoints(0) = parFacetoEnd
    LineStringPoints(1) = parInstrumentHeight + parInsulationThickness
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = parFacetoEnd + pipeDiam(parPortIndex1) / 2 + parInsulationThickness
    LineStringPoints(4) = parInstrumentHeight + parInsulationThickness
    LineStringPoints(5) = 0
    
'   Point 3
    LineStringPoints(6) = parFacetoEnd + pipeDiam(parPortIndex1) / 2 + parInsulationThickness
    LineStringPoints(7) = -parInstrumentHeight1 - parInsulationThickness
    LineStringPoints(8) = 0
    
'   Point 4
    LineStringPoints(9) = parFacetoEnd
    LineStringPoints(10) = -parInstrumentHeight1 - parInsulationThickness
    LineStringPoints(11) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStringPoints)
    axisVect.Set 0, 1, 0
    centerPoint.Set parFacetoEnd, 0, 0
    Set ObjInsVerBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVerBody
    Set ObjInsVerBody = Nothing
    
    Set centerPoint = Nothing
    Set axisVect = Nothing
 
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
 
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
