VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jun 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol is Victaulic pressfit swaged (Style 505) coupling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
      PI = 4 * Atn(1)
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim ObjCoupling As Object
    Dim iOutput     As Double
    Dim parFacetoFace As Double
    Dim dTakeout As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

'   Origin is assumed to be at center of coupling.

' Insert your code for output 1(Coupling)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    dTakeout = parFacetoFace - 2 * depth

    Dim dPoints(1 To 9) As AutoMath.DPosition
    Dim iCount As Integer
    For iCount = 1 To 9
        Set dPoints(iCount) = New AutoMath.DPosition
    Next iCount

'   Gap between pipe surface and Coupling-end over the O-ring is taken as half the differnce between
'   pipe and flage diameters.
    Dim dGap As Double

    dGap = Abs(flangeDiam - pipeDiam) / 2
    
    dPoints(1).x = -parFacetoFace / 2
    dPoints(1).y = pipeDiam / 2 + dGap
    dPoints(1).z = 0
    
'   Assumption:- Section of the Coupling over the O-ring is taken to be a semi-circle with radius  equal to
'   0.05 times End-to-End dimension
    Dim dCouplingSecRadius  As Double
    
    dCouplingSecRadius = 0.05 * parFacetoFace
    
    dPoints(2).x = -parFacetoFace / 2 + dCouplingSecRadius
    dPoints(2).y = pipeDiam / 2 + dGap + dCouplingSecRadius
    dPoints(2).z = 0
    
    dPoints(3).x = -parFacetoFace / 2 + 2 * dCouplingSecRadius
    dPoints(3).y = pipeDiam / 2 + dGap
    dPoints(3).z = 0
    
    dPoints(4).x = -dTakeout / 2
    dPoints(4).y = pipeDiam / 2 + dGap
    dPoints(4).z = 0
    
    dPoints(5).x = 0
    dPoints(5).y = pipeDiam / 2 + dGap - dTakeout * 0.3
    dPoints(5).z = 0

    dPoints(6).x = dTakeout / 2
    dPoints(6).y = pipeDiam / 2 + dGap
    dPoints(6).z = 0

    dPoints(7).x = parFacetoFace / 2 - 2 * dCouplingSecRadius
    dPoints(7).y = pipeDiam / 2 + dGap
    dPoints(7).z = 0
    
    dPoints(8).x = parFacetoFace / 2 - dCouplingSecRadius
    dPoints(8).y = pipeDiam / 2 + dGap + dCouplingSecRadius
    dPoints(8).z = 0

    dPoints(9).x = parFacetoFace / 2
    dPoints(9).y = pipeDiam / 2 + dGap
    dPoints(9).z = 0

    Dim oLeftArcAboveOring  As IngrGeom3D.Arc3d
    Dim oLeftLine As IngrGeom3D.Line3d
    Dim oPipeStopArc  As IngrGeom3D.Arc3d
    Dim oRightLine As IngrGeom3D.Line3d
    Dim oRightArcAboveOring  As IngrGeom3D.Arc3d

    Set oLeftArcAboveOring = PlaceTrArcBy3Pts(dPoints(1), dPoints(3), dPoints(2))

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Set oLeftLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dPoints(3).x, dPoints(3).y, dPoints(3).z, _
                            dPoints(4).x, dPoints(4).y, dPoints(4).z)
    

    Set oPipeStopArc = PlaceTrArcBy3Pts(dPoints(4), dPoints(6), dPoints(5))
    

    Set oRightLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dPoints(6).x, dPoints(6).y, dPoints(6).z, _
                            dPoints(7).x, dPoints(7).y, dPoints(7).z)

    Set oRightArcAboveOring = PlaceTrArcBy3Pts(dPoints(7), dPoints(9), dPoints(8))
    
    Dim oCouplingCurves           As Collection
    Dim objComplexeCurve         As IngrGeom3D.ComplexString3d
    Set oCouplingCurves = New Collection

    oCouplingCurves.Add oLeftArcAboveOring
    oCouplingCurves.Add oLeftLine
    oCouplingCurves.Add oPipeStopArc
    oCouplingCurves.Add oRightLine
    oCouplingCurves.Add oRightArcAboveOring
    Set objComplexeCurve = PlaceTrCString(dPoints(1), oCouplingCurves)
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    oCenterPoint.Set 0, 0, 0

    Set ObjCoupling = PlaceRevolution(m_OutputColl, objComplexeCurve, oAxisVec, oCenterPoint, 2 * PI, True)

    'Remove complex curve arcs/lines
    Dim ObjRemoveCurve As IJDObject
    Set ObjRemoveCurve = objComplexeCurve
    ObjRemoveCurve.Remove
    Set ObjRemoveCurve = Nothing
    
    For iCount = 1 To oCouplingCurves.Count
        oCouplingCurves.Remove 1
    Next iCount
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCoupling
    Set ObjCoupling = Nothing
    Set oLeftArcAboveOring = Nothing
    Set oLeftLine = Nothing
    Set oPipeStopArc = Nothing
    Set oRightLine = Nothing
    Set oRightArcAboveOring = Nothing
    Set oGeomFactory = Nothing
    Set oCouplingCurves = Nothing
    Set oAxisVec = Nothing
    Set oCenterPoint = Nothing
    Set objComplexeCurve = Nothing
    
    For iCount = 1 To 9
        Set dPoints(iCount) = Nothing
    Next iCount

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
