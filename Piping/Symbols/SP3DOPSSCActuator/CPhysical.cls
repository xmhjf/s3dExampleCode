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

'   ProgID          :  SP3DOPSSCActuator.OPSSCActuator
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Thursday,27 Sep 2007
'   Description     :  SSC Type Actuator
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   27-SEP-2007      PK        CR-127652:Created the symbol.
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
    Dim parPositionerOffset  As Double
    Dim dActOperatorHeight   As Double
    Dim dOperatorDiameter    As Double
    Dim iOutput              As Integer
    Dim iCount               As Integer
    Dim oCentPos             As AutoMath.DPosition
    Dim oStPoint             As AutoMath.DPosition
    Dim oEnPoint             As AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parPositionerOffset = arrayOfInputs(4)
        
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0
    oCentPos.Set 0, parPositionerOffset, 0
    dOperatorDiameter = 0.7 * parOperatorHeight
    dActOperatorHeight = parOperatorHeight - parPositionerOffset
      
    'Insert your code for Body(Output 1)
    Dim oGeomFactory  As IngrGeom3D.GeometryFactory
    Dim oBody As Object
    Dim oPoints(7) As AutoMath.DPosition
    Dim oBspline As IngrGeom3D.BSplineCurve3d
    Dim oAxis As AutoMath.DVector
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    For iCount = 0 To 7
        Set oPoints(iCount) = New AutoMath.DPosition
    Next
    
    oPoints(0).Set oCentPos.x, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + 0.4 * dOperatorDiameter, oCentPos.y + 0.45 * _
                                                    dActOperatorHeight, oCentPos.z
    oPoints(2).Set oCentPos.x + 0.475 * dOperatorDiameter, oCentPos.y + 0.55 * _
                                                    dActOperatorHeight, oCentPos.z
    oPoints(3).Set oCentPos.x + 0.5 * dOperatorDiameter, oCentPos.y + 0.65 * _
                                                    dActOperatorHeight, oCentPos.z
    oPoints(4).Set oCentPos.x + 0.5 * dOperatorDiameter, oCentPos.y + 0.95 * _
                                                    dActOperatorHeight, oCentPos.z
    oPoints(5).Set oCentPos.x + 0.1 * dOperatorDiameter, oCentPos.y + 0.95 * _
                                                    dActOperatorHeight, oCentPos.z
    oPoints(6).Set oCentPos.x + 0.1 * dOperatorDiameter, oCentPos.y + dActOperatorHeight, _
                                                                        oCentPos.z
    oPoints(7).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set Output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oBspline = Nothing
    For iCount = 0 To 7
        Set oPoints(iCount) = Nothing
    Next
    Set oBody = Nothing
   
    'Insert your code for Projection
    Dim oCString    As IngrGeom3D.ComplexString3d
    Dim oProjection As Object
    Dim lines       As Collection
    Dim oLine       As Object
    Set oCString = New IngrGeom3D.ComplexString3d
    
    'Insert your code for Complex String
    Set lines = New Collection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.375 * dOperatorDiameter, oCentPos.x + 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.375 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.375 * dOperatorDiameter, oCentPos.x + 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.125 * dOperatorDiameter, oCentPos.x + 0.375 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.375 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.125 * dOperatorDiameter, oCentPos.x + 0.375 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.375 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.125 * dOperatorDiameter, oCentPos.x + 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.125 * dOperatorDiameter, oCentPos.x + 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.375 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.375 * dOperatorDiameter, oCentPos.x - 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.375 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.375 * dOperatorDiameter, oCentPos.x - 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.125 * dOperatorDiameter, oCentPos.x - 0.375 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.375 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z - _
                    0.125 * dOperatorDiameter, oCentPos.x - 0.375 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.375 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.125 * dOperatorDiameter, oCentPos.x - 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.125 * _
                    dOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.125 * _
                    dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + _
                    0.125 * dOperatorDiameter, oCentPos.x - 0.125 * dOperatorDiameter, _
                    oCentPos.y + 0.45 * dActOperatorHeight, oCentPos.z + 0.375 * _
                    dOperatorDiameter)
    lines.Add oLine
    
    oStPoint.Set oCentPos.x - 0.125 * dOperatorDiameter, oCentPos.y + 0.45 * dActOperatorHeight, _
                    oCentPos.z + 0.375 * dOperatorDiameter
    Set oCString = PlaceTrCString(oStPoint, lines)
        
    oAxis.Set 0, -1, 0
    Set oProjection = PlaceProjection(m_OutputColl, oCString, oAxis, 0.3 * dActOperatorHeight, True)
    
    'Set Output2(Projection)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set oLine = Nothing
    Set lines = Nothing
    Set oProjection = Nothing
    Set oCString = Nothing
    Set oAxis = Nothing
    
    'Insert your code for Base
    Dim oBase As Object
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.15 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y, oCentPos.z + 0.5 * parWidth
    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set Output3(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    
    'Insert your code for Bolts
    Dim oBolt As Object
    Dim oTransMat As AutoMath.DT4x4
    Set oAxis = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    oAxis.Set 0, 1, 0
    oStPoint.Set oCentPos.x + 0.4 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                                                            oCentPos.z + 0.4 * parWidth
    oEnPoint.Set oCentPos.x + 0.4 * parWidth, oCentPos.y + 0.15 * dActOperatorHeight, _
                                                            oCentPos.z + 0.4 * parWidth
    
    'Set Output4(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            0.05 * parWidth, True)
        oTransMat.LoadIdentity
        oTransMat.Rotate iCount * PI / 2, oAxis
        oBolt.Transform oTransMat
        m_OutputColl.AddOutput "Bolts_", oBolt
        Set oBolt = Nothing
    Next
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oTransMat = Nothing
    Set oAxis = Nothing
    Set oCentPos = Nothing
      
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
