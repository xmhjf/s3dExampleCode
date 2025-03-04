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
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG5.
'   Physical aspect consist Two cones, one converging and other  diverging, 2-cylinders along X-axis,
'   1-Cylinder along Y-axis and two nozzles. Insulation aspect consist of same number of outputs except for the
'   Nozzles.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    
    PI = 4 * Atn(1)

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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
        
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjValveBodyIns As Object
    Dim ObjVerCylinIns As Object
    Dim ObjHorCylins  As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFaceToFace = arrayOfInputs(3)        'D3
    parInstrumentHeight = arrayOfInputs(4) 'D4
    parInstrumentDiameter = arrayOfInputs(5)      'D5
    parInstrumentDiameter1 = arrayOfInputs(6)  'D6
    parInstrumentLength = arrayOfInputs(7)     'D7
    parInstrumentLength1 = arrayOfInputs(8)    'D8
    parInstrumentLength2 = arrayOfInputs(9)       'D9
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 ' Insert your code for output 5(Valve Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxDia As Double
    Dim InsulationDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    InsulationDia = maxDia + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFaceToFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing

'Insert your code for output 6 (Vertical Cylinder Insulation)
    Dim CylDia As Double
    'Stem Diameter is assumed as 0.05  as given in EDEN of this Symbol
    CylDia = 0.05
    stPoint.Set 0, 0, 0
    enPoint.Set 0, (parInstrumentHeight - Sqr((parInstrumentDiameter / 2) ^ 2 - (CylDia / 2) ^ 2)), 0
    InsulationDia = CylDia + 2 * parInsulationThickness
    Set ObjVerCylinIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylinIns
    Set ObjVerCylinIns = Nothing
   Set stPoint = Nothing
   Set enPoint = Nothing
   
' Insert your code for output 7 (Horizontal Cylinders along X-axis)
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 17) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector

    'Point1
    dblLinePts(0) = -parInstrumentLength1 - parInsulationThickness
    dblLinePts(1) = parInstrumentHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = -parInstrumentLength1 - parInsulationThickness
    dblLinePts(4) = parInstrumentHeight - parInstrumentDiameter / 2 - parInsulationThickness
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = parInstrumentLength2 + parInsulationThickness
    dblLinePts(7) = parInstrumentHeight - parInstrumentDiameter / 2 - parInsulationThickness
    dblLinePts(8) = 0
    'Point4
    dblLinePts(9) = parInstrumentLength2 + parInsulationThickness
    dblLinePts(10) = parInstrumentHeight - parInstrumentDiameter1 / 2 - parInsulationThickness
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = parInstrumentLength + parInsulationThickness
    dblLinePts(13) = parInstrumentHeight - parInstrumentDiameter1 / 2 - parInsulationThickness
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = parInstrumentLength + parInsulationThickness
    dblLinePts(16) = parInstrumentHeight
    dblLinePts(17) = 0
    
    Set objLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dblLinePts)
    revVector.Set 1, 0, 0
    centerPt.Set 0, parInstrumentHeight, 0
    Set ObjHorCylins = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHorCylins
    Set ObjHorCylins = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    Set objLineStr = Nothing
    tmpLineStr.Remove
    Set geomFactory = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
