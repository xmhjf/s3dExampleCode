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

'   ProgID          :  SP3DOPECActuator.OPECActuator
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday,28 Sep 2007
'   Description     :  EC Type Actuator
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   28.Sep.2007     PK          CR-127652:Created the symbol.
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
    Dim parOperatorDiameter  As Double
    Dim parWidth             As Double
    Dim parPositionerOffset  As Double
    Dim dActOperatorHeight   As Double
    Dim iOutput              As Integer
    Dim iCount               As Integer
    Dim oCentPos             As AutoMath.DPosition
    Dim oStPoint             As AutoMath.DPosition
    Dim oEnPoint             As AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parWidth = arrayOfInputs(4)
    parPositionerOffset = arrayOfInputs(5)
        
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    iOutput = 0
    oCentPos.Set 0, parPositionerOffset, 0
    dActOperatorHeight = parOperatorHeight - parPositionerOffset
      
    'Insert your code for Body(Output 1)
    Dim oGeomFactory  As IngrGeom3D.GeometryFactory
    Dim oBody         As Object
    
    oStPoint.Set oCentPos.x, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOperatorDiameter, True)
    
    'Set Output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oBody = Nothing
   
    'Insert your code for Projection
    Dim oProjection As Object
    Dim oLine       As Object
    Dim lines       As Collection
    Dim oAxis       As AutoMath.DVector
    Dim oCString    As IngrGeom3D.ComplexString3d
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oCString = New IngrGeom3D.ComplexString3d
    Set oAxis = New AutoMath.DVector
    
    'Insert your code for Complex String
    Set lines = New Collection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z - _
                     0.35 * parOperatorDiameter, oCentPos.x - 0.35 * parOperatorDiameter, _
                     oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z - 0.35 * _
                     parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z - _
                     0.35 * parOperatorDiameter, oCentPos.x - 0.35 * parOperatorDiameter, _
                     oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z, _
                     oCentPos.x - 0.35 * parOperatorDiameter, oCentPos.y + 0.3 * _
                     dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.3 * dActOperatorHeight, oCentPos.z, _
                     oCentPos.x - 0.35 * parOperatorDiameter, oCentPos.y + 0.3 * _
                     dActOperatorHeight, oCentPos.z + 0.35 * parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.3 * dActOperatorHeight, oCentPos.z + _
                     0.35 * parOperatorDiameter, oCentPos.x - 0.35 * parOperatorDiameter, _
                     oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + 0.35 * _
                     parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.35 * _
                     parOperatorDiameter, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + _
                     0.35 * parOperatorDiameter, oCentPos.x - 0.35 * parOperatorDiameter, _
                     oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z - 0.35 * _
                     parOperatorDiameter)
    lines.Add oLine
    
    oStPoint.Set oCentPos.x - 0.35 * parOperatorDiameter, oCentPos.y + 0.1 * dActOperatorHeight, _
                     oCentPos.z - 0.35 * parOperatorDiameter
    Set oCString = PlaceTrCString(oStPoint, lines)
    oAxis.Set 1, 0, 0
    Set oProjection = PlaceProjection(m_OutputColl, oCString, oAxis, 0.7 * parOperatorDiameter, True)
    
    'Set Ouptut2(Projection)
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
    
    'Insert your code for partition box
    Dim oBox As Object
    oStPoint.Set oCentPos.x - 0.1 * parOperatorDiameter, oCentPos.y + 0.3 * dActOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z + 0.35 * parOperatorDiameter
    
    Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set Output3(Partition Box)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
    'Insert your code for Base
    Dim oBase As Object
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y, oCentPos.z + 0.5 * parWidth
    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set Output4(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    
    'Insert your code for Bolts
    Dim oBolt     As Object
    Dim oTransMat As AutoMath.DT4x4
    Set oAxis = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    oAxis.Set 0, 1, 0
    
    oStPoint.Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.35 * _
                                    dActOperatorHeight, oCentPos.z + 0.2 * parOperatorDiameter
    oEnPoint.Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.25 * _
                                    dActOperatorHeight, oCentPos.z + 0.2 * parOperatorDiameter
    
    'Set Output5(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 1
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            0.1 * parOperatorDiameter, True)
        oTransMat.LoadIdentity
        oTransMat.Rotate iCount * 3 * PI / 2, oAxis
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



