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

'   ProgID          :  SP3DOPECXActuator.OPECXActuator
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday, 1 Oct 2007
'   Description     :  ECX Type Operator
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   01-OCT-2007      PK        CR-127652:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
    Dim oBody         As Object
    Dim oAxis         As AutoMath.DVector
    Dim oPoints(9)    As AutoMath.DPosition
    Dim oBspline      As IngrGeom3D.BSplineCurve3d
    Dim oGeomFactory  As IngrGeom3D.GeometryFactory
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    For iCount = 0 To 9
        Set oPoints(iCount) = New AutoMath.DPosition
    Next
    
    oPoints(0).Set oCentPos.x, oCentPos.y + 0.44 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + parOperatorDiameter * (0.5 - 0.03 / Tan(PI / 6) - _
                    0.15 / Tan(5 * PI / 12)), oCentPos.y + 0.44 * dActOperatorHeight, oCentPos.z
    oPoints(2).Set oCentPos.x + parOperatorDiameter * (0.5 - 0.03 / Tan(PI / 6)), _
                    oCentPos.y + 0.59 * dActOperatorHeight, oCentPos.z
    oPoints(3).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.62 * _
                    dActOperatorHeight, oCentPos.z
    oPoints(4).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.82 * _
                    dActOperatorHeight, oCentPos.z
    oPoints(5).Set oCentPos.x + parOperatorDiameter * (0.5 - 0.03 / Tan(PI / 6)), _
                    oCentPos.y + 0.85 * dActOperatorHeight, oCentPos.z
    oPoints(6).Set oCentPos.x + parOperatorDiameter * (0.5 - 0.03 / Tan(PI / 6) - _
                    0.15 / Tan(5 * PI / 12)), oCentPos.y + dActOperatorHeight, oCentPos.z
    oPoints(7).Set oCentPos.x + 0.08 * parOperatorDiameter, oCentPos.y + _
                    dActOperatorHeight, oCentPos.z
    oPoints(8).Set oCentPos.x + 0.08 * parOperatorDiameter, oCentPos.y + _
                    dActOperatorHeight + 0.01, oCentPos.z
    oPoints(9).Set oCentPos.x, oCentPos.y + dActOperatorHeight + 0.01, oCentPos.z
    
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set Output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oBspline = Nothing
    For iCount = 0 To 9
        Set oPoints(iCount) = Nothing
    Next
    Set oBody = Nothing
    Set oAxis = Nothing
    
    'Insert your code for Cylinder
    Dim oCylinder As Object
    oStPoint.Set oCentPos.x, oCentPos.y + 0.44 * dActOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + 0.14 * dActOperatorHeight, oCentPos.z
    
    'Insert your code for Complex String
    Set oCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.6 * parWidth, True)
    
    'Set Output2(Cylinder)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCylinder
    Set oCylinder = Nothing
       
    'Insert your code for Base
    Dim oBase As Object
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.14 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y, oCentPos.z + 0.5 * parWidth
    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set Output3(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    
    'Insert your code for bolts
    Dim oBolt     As Object
    Dim oTransMat As AutoMath.DT4x4
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, 1, 0
    
    oStPoint.Set oCentPos.x + 0.4 * parWidth, oCentPos.y + 0.18 * dActOperatorHeight, _
                                                            oCentPos.z + 0.4 * parWidth
    oEnPoint.Set oCentPos.x + 0.4 * parWidth, oCentPos.y + 0.13 * dActOperatorHeight, _
                                                            oCentPos.z + 0.4 * parWidth
    Set oTransMat = New AutoMath.DT4x4
    
    'Set Output4(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parWidth, True)
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


