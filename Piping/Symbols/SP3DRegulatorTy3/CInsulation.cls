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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG3,
'   PDMS Category Name=IVRH). Physical aspect consist Two cones, one converging and other
'   diverging, 2 cylinders along Y-axis,  a line string revolved about X-axis that takes care of Two Cylinders
'   and 2 Nozzles.
'   Insulation for the Valve body as a Cylinder, for Vertical cylinders a single cylinder along Y-axis
'   and a line string revolution about X-axis that takes care of Two Cylinders.
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
'''
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
    Dim ObjVertCylinsIns As Object
    Dim ObjHoriCylinIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parCylHeight As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFaceToFace = arrayOfInputs(3)        'D3
    parCylHeight = arrayOfInputs(4)  'D4
    parInstrumentLength = arrayOfInputs(5)      'D5
    parInstrumentHeight = arrayOfInputs(6)  'D6
    parInstrumentDiameter = arrayOfInputs(7)     'D7
    parInstrumentLength1 = arrayOfInputs(8)    'D8
    parInstrumentDiameter1 = arrayOfInputs(9)        'D9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
    
 ' Insert your code for output 6(Valve Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    InsulationDia = maxDia + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFaceToFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
'Insert your code for the Output 7 (Insulation Valve Cylinders)
'For Insulation Valve bottom Cylinder diameter is adopted for the Valve Top Cylinder Diameter.
    Dim CylDia  As Double
    Dim meetRadius  As Double
    Dim miniCylDia As Double
    
    miniCylDia = 0.05
    meetRadius = Sqr((parInstrumentDiameter / 2) ^ 2 - (miniCylDia / 2) ^ 2)
    CylDia = parFace1toCenter
    If CylDia < miniCylDia Then CylDia = miniCylDia
    InsulationDia = CylDia + parInsulationThickness * 2
    stPoint.Set 0, -parCylHeight, 0
    enPoint.Set 0, parInstrumentHeight - meetRadius, 0
    Set ObjVertCylinsIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylinsIns
    Set ObjVertCylinsIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 8(Vertical Cylinder and Hori Cylinder Insulation)
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 17) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector
    
    'Point1
    dblLinePts(0) = -parInstrumentLength1 / 2 - parInsulationThickness
    dblLinePts(1) = parInstrumentHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = -parInstrumentLength1 / 2 - parInsulationThickness
    dblLinePts(4) = parInstrumentHeight - parInstrumentDiameter / 2 - parInsulationThickness
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = parInstrumentLength1 / 2 + parInsulationThickness
    dblLinePts(7) = parInstrumentHeight - parInstrumentDiameter / 2 - parInsulationThickness
    dblLinePts(8) = 0
    'Point4
    dblLinePts(9) = parInstrumentLength1 / 2 + parInsulationThickness
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
    Set ObjHoriCylinIns = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinIns
    Set ObjHoriCylinIns = Nothing
    Set geomFactory = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    Set objLineStr = Nothing
    tmpLineStr.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
