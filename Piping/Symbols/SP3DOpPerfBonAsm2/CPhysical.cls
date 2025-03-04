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

'   ProgID          :  SP3DOpPerfBonAsm2.OpPerfBonAsm2
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday,22 Oct 2007
'   Description     :  PES – Performance Bonnet Assembly, Type 2
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
    
    Dim parOperatorHeight   As Double
    Dim parWidth            As Double
    Dim parOperatorDiameter As Double
    Dim parPositionerOffset As Double
    
    Dim oCentPos            As AutoMath.DPosition
    Dim iOutput             As Integer
    Dim iCount              As Integer
    Dim dActOperatorHeight  As Double
   
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
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oAxis           As New AutoMath.DVector
    Dim oBody           As Object
    Dim oPoints(9)      As AutoMath.DPosition
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    
    For iCount = 0 To 9
       Set oPoints(iCount) = New AutoMath.DPosition
    Next
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    oPoints(0).Set oCentPos.x, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(2).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(3).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(4).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.87 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(5).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.87 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(6).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(7).Set oCentPos.x + 0.25 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(8).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(9).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    
       
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    
    For iCount = 0 To 9
        Set oPoints(iCount) = Nothing
    Next iCount
    Set oAxis = Nothing
    Set oBspline = Nothing
    Set oBody = Nothing
    
    'Insert your code for Transition(Output2)
    Dim oSquare     As IngrGeom3D.ComplexString3d
    Dim lines       As Collection
    Dim oLine       As Object
    Dim oStPoint    As AutoMath.DPosition
    
    Set oSquare = New IngrGeom3D.ComplexString3d
    Set oStPoint = New AutoMath.DPosition
    
    'Insert your code for Square Plane
    Set lines = New Collection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                         oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth, _
                         oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                         oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                         oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth, _
                         oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                         oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                          oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth, _
                          oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                          oCentPos.z + 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                          oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth, _
                          oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                          oCentPos.z + 0.5 * parWidth)
    lines.Add oLine
    oStPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                            oCentPos.z + 0.5 * parWidth
    Set oSquare = PlaceTrCString(oStPoint, lines)
    
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set oLine = Nothing
    Set lines = Nothing
    
    ''Insert your code for Circular Plane
    Dim arcs            As Collection
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oCircle         As IngrGeom3D.ComplexString3d
    
    Set oCircle = New IngrGeom3D.ComplexString3d
    Set arcs = New Collection
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z + 0.15 * parOperatorDiameter, oCentPos.x _
                          + 0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x + 0.15 * parOperatorDiameter, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z - 0.15 * parOperatorDiameter)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z - 0.15 * parOperatorDiameter, oCentPos.x - _
                          0.15 * parOperatorDiameter, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
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
    
    'Insert your code for transition element between square plane and circular plane
    Dim oTransition As Object
    Set oTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                            oSquare, oCircle, True)
                            
    'code for output2(Transition)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTransition
    
    Set oTransition = Nothing
    Set oCircle = Nothing

    'Insert your code for Base
    Dim oBase As Object
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, -1, 0
    
    Set oBase = PlaceProjection(m_OutputColl, oSquare, oAxis, 0.2 * dActOperatorHeight, True)
    
    'code for output3(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    
    Set oBase = Nothing
    Set oAxis = Nothing
    Set oSquare = Nothing
    
    'Insert your code for Bolts
    Dim oBolt       As Object
    Dim oEnPoint    As AutoMath.DPosition
    Dim oTransMat   As AutoMath.DT4x4
    
    Set oEnPoint = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, 1, 0
    
    oStPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.35 * dActOperatorHeight, _
                                                        oCentPos.z + 0.34 * parWidth
    oEnPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                                                        oCentPos.z + 0.34 * parWidth
                                
    Set oTransMat = New AutoMath.DT4x4
    
    'Set Output4(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            0.15 * parWidth, True)
        oTransMat.LoadIdentity
        oTransMat.Rotate iCount * PI / 2, oAxis
        oBolt.Transform oTransMat
        m_OutputColl.AddOutput "Bolts_", oBolt
        Set oBolt = Nothing
    Next
    
    Set oTransMat = Nothing
    Set oAxis = Nothing
    Set oEnPoint = Nothing
    
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
    oStPoint.Set oCentPos.x - 0.2 * dActOperatorHeight, oCentPos.y + 0.2 * dActOperatorHeight, _
                                                                oCentPos.z + 0.5 * parWidth
    Set oCString = PlaceTrCString(oStPoint, arcs)
    Set oProjection = PlaceProjection(m_OutputColl, oCString, oAxis, parWidth, True)
    
    'Set Output5(Projection)
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

