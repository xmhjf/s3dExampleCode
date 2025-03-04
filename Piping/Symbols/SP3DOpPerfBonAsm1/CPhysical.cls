VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DOpPerfBonAsm1.OpPerfBonAsm1
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday,22 Oct 2007
'   Description     :  PES – Performance Bonnet Assembly, Type 1
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   22.OCT.2007      PK        CR-127653:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parOperatorHeight    As Double
    Dim parWidth             As Double
    Dim parOperatorDiameter  As Double
    Dim parPositionerOffset  As Double
    Dim iOutput              As Integer
    Dim iCount               As Integer
    Dim oCentPos             As AutoMath.DPosition
    Dim dActOperatorHeight   As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parWidth = arrayOfInputs(4)
    parPositionerOffset = arrayOfInputs(5)
    
    Set oCentPos = New AutoMath.DPosition
    iOutput = 0
    oCentPos.Set 0, parPositionerOffset, 0
    dActOperatorHeight = parOperatorHeight - parPositionerOffset
    
    'Insert your code for Body(Output 1)
    Dim oRevolution     As Object
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As AutoMath.DVector
    Dim oStPoint        As AutoMath.DPosition
    Dim opoints(10)     As AutoMath.DPosition
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    
    Set oAxis = New AutoMath.DVector
    For iCount = 0 To 10
       Set opoints(iCount) = New AutoMath.DPosition
    Next
    Set oStPoint = New AutoMath.DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    opoints(0).Set oCentPos.x, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z
    opoints(1).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(2).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(3).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(4).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.77 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(5).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.77 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(6).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.87 * dActOperatorHeight, _
                                                                oCentPos.z
    opoints(7).Set oCentPos.x + 0.25 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    opoints(8).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    opoints(9).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + dActOperatorHeight, _
                                                                oCentPos.z
    opoints(10).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
       
    Set oBspline = PlaceTrBspline(2, opoints)
    oAxis.Set 0, 1, 0
    Set oRevolution = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'set output1(Revolution)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRevolution
    Set oAxis = Nothing
    For iCount = 0 To 10
        Set opoints(iCount) = Nothing
    Next iCount
    Set oArc = Nothing
    Set oBspline = Nothing
    Set oRevolution = Nothing

    'Code for placing transition
    Dim oLine   As Object
    Dim lines   As Collection
    Dim oSquare As IngrGeom3D.ComplexString3d
    
    Set oSquare = New IngrGeom3D.ComplexString3d
    Set lines = New Collection
    
    'code for forming square plane
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                            oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * _
                            parWidth, oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * _
                            dActOperatorHeight, oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                            oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * _
                            parWidth, oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * _
                            dActOperatorHeight, oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                            oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * _
                            parWidth, oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * _
                            dActOperatorHeight, oCentPos.z + 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                            oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * _
                            parWidth, oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * _
                            dActOperatorHeight, oCentPos.z + 0.5 * parWidth)
    lines.Add oLine

    oStPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z _
                            + 0.5 * parWidth
    Set oSquare = PlaceTrCString(oStPoint, lines)
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set oLine = Nothing
    Set lines = Nothing
    
    'code for forming circular plane
    Dim arcs    As Collection
    Dim oCircle As IngrGeom3D.ComplexString3d
    Set oCircle = New IngrGeom3D.ComplexString3d
    Set arcs = New Collection
    
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z + 0.15 * parOperatorDiameter, oCentPos.x + _
                          0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x + 0.15 * parOperatorDiameter, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z - 0.15 * parOperatorDiameter)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z - 0.15 * parOperatorDiameter, oCentPos.x _
                          - 0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x - 0.15 * parOperatorDiameter, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z + 0.15 * parOperatorDiameter)
    arcs.Add oArc
    
    oStPoint.Set oCentPos.x, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z + 0.15 * parOperatorDiameter
    Set oCircle = PlaceTrCString(oStPoint, arcs)
    For iCount = 1 To arcs.Count
        arcs.Remove 1
    Next iCount
    Set oArc = Nothing
    Set arcs = Nothing
    
    'code for making transition element between square plane and circular plane
    Dim oTransition As Object
    Set oTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                        oSquare, oCircle, True)
    
    'Set Output2(Transition)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTransition
    Set oTransition = Nothing
    Set oCircle = Nothing
    
    'Insert your code for Base
    Dim oBase  As Object
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, -1, 0
    Set oBase = PlaceProjection(m_OutputColl, oSquare, oAxis, 0.2 * dActOperatorHeight, _
                                                                                        True)
    'Set Output3(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oSquare = Nothing
    Set oBase = Nothing
    Set oAxis = Nothing
    
    'Insert your code for Projection
    Dim oProjection As Object
    Dim oCString    As IngrGeom3D.ComplexString3d
    Set oAxis = New AutoMath.DVector
    Set arcs = New Collection

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oCentPos.x - 0.2 * _
                        dActOperatorHeight, oCentPos.y + 0.2 * dActOperatorHeight, _
                        oCentPos.z + 0.5 * parWidth, oCentPos.x, oCentPos.y + 0.45 * _
                        dActOperatorHeight, oCentPos.z + 0.5 * parWidth, oCentPos.x + _
                        0.2 * dActOperatorHeight, oCentPos.y + 0.2 * dActOperatorHeight, _
                        oCentPos.z + 0.5 * parWidth)
    arcs.Add oArc
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.2 * _
                        dActOperatorHeight, oCentPos.y + 0.2 * dActOperatorHeight, _
                        oCentPos.z + 0.5 * parWidth, oCentPos.x - 0.2 * dActOperatorHeight, _
                        oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth)
    arcs.Add oLine
    oAxis.Set 0, 0, -1
    oStPoint.Set oCentPos.x - 0.2 * dActOperatorHeight, oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth
    Set oCString = PlaceTrCString(oStPoint, arcs)
    Set oProjection = PlaceProjection(m_OutputColl, oCString, oAxis, parWidth, True)
    
    'Set Output4(Projection)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
    Set oAxis = Nothing
    Set oCString = Nothing
    Set oArc = Nothing
    Set arcs = Nothing
    Set oLine = Nothing
    Set oProjection = Nothing
    Set oStPoint = Nothing
    Set oCentPos = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

