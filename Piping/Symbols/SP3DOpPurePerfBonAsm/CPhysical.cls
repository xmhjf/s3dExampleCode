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

'   ProgID          :  SP3DOpPurePerfBonAsm.OpPPBonAsm
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Wednesday,17 Oct 2007
'   Description     :  Pure Performance Bonnet Assembly
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   17.SEP.2007      PK        CR-127653:Created the symbol.
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
    Dim oBody           As Object
    Dim oAxis           As AutoMath.DVector
    Dim oStPoint        As AutoMath.DPosition
    Dim oPoints(9)      As AutoMath.DPosition
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    
    Set oAxis = New AutoMath.DVector
    Set oStPoint = New AutoMath.DPosition
    For iCount = 0 To 9
       Set oPoints(iCount) = New AutoMath.DPosition
    Next

    oPoints(0).Set oCentPos.x, oCentPos.y + 0.08 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + 0.25 * parOperatorDiameter, oCentPos.y + 0.08 * _
                                                dActOperatorHeight, oCentPos.z
    oPoints(2).Set oCentPos.x + 0.25 * parOperatorDiameter, oCentPos.y + 0.14 * _
                                                dActOperatorHeight, oCentPos.z
    oPoints(3).Set oCentPos.x + 0.3 * parOperatorDiameter, oCentPos.y + 0.14 * _
                                                dActOperatorHeight, oCentPos.z
    oPoints(4).Set oCentPos.x + 0.3 * parOperatorDiameter, oCentPos.y + 0.2 * _
                                                parOperatorHeight, oCentPos.z
    oPoints(5).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.2 * _
                                                parOperatorHeight, oCentPos.z
    oPoints(6).Set oCentPos.x + 0.5 * parOperatorDiameter + 0.8 * dActOperatorHeight _
                                                * Tan(PI / 36), oCentPos.y + 0.95 * _
                                                dActOperatorHeight, oCentPos.z
    oPoints(7).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + 0.95 * _
                                                dActOperatorHeight, oCentPos.z
    oPoints(8).Set oCentPos.x + 0.15 * parOperatorDiameter, oCentPos.y + _
                                                dActOperatorHeight, oCentPos.z
    oPoints(9).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
       
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set outout1(Revolution)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    
    Set oAxis = Nothing
    For iCount = 0 To 9
        Set oPoints(iCount) = Nothing
    Next iCount
    Set oBspline = Nothing
    Set oBody = Nothing

    'Insert your code for Base
    Dim oBase As Object
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.08 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y, oCentPos.z + 0.5 * parWidth
    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
                                                         
    'code for output2(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    
    'code for placing bolts
    Dim oBolt       As Object
    Dim oTransMat   As AutoMath.DT4x4
    Dim oPoint      As AutoMath.DPosition
    Set oPoint = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, 1, 0
    
    oPoint.Set oCentPos.x + 0.35 * parWidth, oCentPos.y + 0.08 * dActOperatorHeight, _
                                                       oCentPos.z + 0.35 * parWidth
    Set oTransMat = New AutoMath.DT4x4
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceSphere(m_OutputColl, oPoint, 0.05 * parWidth)
        oTransMat.LoadIdentity
        oTransMat.Rotate iCount * PI / 2, oAxis
        oBolt.Transform oTransMat
        m_OutputColl.AddOutput "Bolts_", oBolt
        Set oBolt = Nothing
    Next
    Set oTransMat = Nothing
    Set oAxis = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPoint = Nothing
    Set oCentPos = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub



