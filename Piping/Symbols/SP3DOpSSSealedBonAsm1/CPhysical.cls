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

'   ProgID          :  SP3DOpSSSealedBonAsm1.OpSSSBonAsm1
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Thursday,18 Oct 2007
'   Description     :  Stainless Steel Sealed Bonnet Assembly,Type 1
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   18.OCT.2007      PK        CR-127653:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart

    Dim parOperatorHeight   As Double
    Dim parOperatorDiameter As Double
    Dim parWidth            As Double
    Dim parPositionerOffset As Double

    Dim iOutput             As Integer
    Dim iCount              As Integer
    Dim dActOperatorHeight  As Double

    Dim oCentPos            As AutoMath.DPosition

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

    'Insert your code for Wheel (Output 1)
    Dim oWheel              As Object
    Dim oAxis               As AutoMath.DVector
    Dim oStPoint            As AutoMath.DPosition
    Dim oGeomFactory        As IngrGeom3D.GeometryFactory
    Dim oOctString          As IngrGeom3D.ComplexString3d
    Dim oLine               As Object
    Dim lines               As Collection
    Dim dValue              As Double

    Set oAxis = New AutoMath.DVector
    Set lines = New Collection
    Set oStPoint = New AutoMath.DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    dValue = parOperatorDiameter / (1 + 2 * Cos(3 * PI / 8))
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * dValue, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z + 0.5 * _
                    parOperatorDiameter, oCentPos.x + 0.5 * dValue, oCentPos.y + _
                    0.83 * dActOperatorHeight, oCentPos.z + 0.5 * parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * dValue, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z + 0.5 * _
                    parOperatorDiameter, oCentPos.x + 0.5 * parOperatorDiameter, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z + 0.5 * dValue)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.83 * dActOperatorHeight, _
                    oCentPos.z + 0.5 * dValue, oCentPos.x + 0.5 * parOperatorDiameter, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z - 0.5 * dValue)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z - _
                    0.5 * dValue, oCentPos.x + 0.5 * dValue, oCentPos.y + 0.83 * _
                    dActOperatorHeight, oCentPos.z - 0.5 * parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * dValue, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z - 0.5 * _
                    parOperatorDiameter, oCentPos.x - 0.5 * dValue, oCentPos.y + 0.83 * _
                    dActOperatorHeight, oCentPos.z - 0.5 * parOperatorDiameter)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * dValue, _
                    oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z - 0.5 * _
                    parOperatorDiameter, oCentPos.x - 0.5 * parOperatorDiameter, oCentPos.y + _
                    0.83 * dActOperatorHeight, oCentPos.z - 0.5 * dValue)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z - _
                    0.5 * dValue, oCentPos.x - 0.5 * parOperatorDiameter, oCentPos.y + 0.83 * _
                    dActOperatorHeight, oCentPos.z + 0.5 * dValue)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z + _
                    0.5 * dValue, oCentPos.x - 0.5 * dValue, oCentPos.y + 0.83 * _
                    dActOperatorHeight, oCentPos.z + 0.5 * parOperatorDiameter)
    lines.Add oLine

    oStPoint.Set oCentPos.x - 0.5 * dValue, oCentPos.y + 0.83 * dActOperatorHeight, _
                    oCentPos.z + parOperatorDiameter / 2
    Set oOctString = PlaceTrCString(oStPoint, lines)
    oAxis.Set 0, 1, 0

    Set oWheel = PlaceProjection(m_OutputColl, oOctString, oAxis, 0.12 * dActOperatorHeight, True)
    
    'Set Ouput1(Wheel)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oWheel

    Set oWheel = Nothing
    Set oOctString = Nothing
    Set oLine = Nothing
    Set lines = Nothing
    
    'Insert your code for Revolution
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    Dim oPoints(3)      As AutoMath.DPosition
    Dim oRevolution     As Object
    
    For iCount = 0 To 3
        Set oPoints(iCount) = New AutoMath.DPosition
    Next

    oPoints(0).Set oCentPos.x, oCentPos.y + 0.95 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.95 * _
                                        dActOperatorHeight, oCentPos.z
    oPoints(2).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + dActOperatorHeight, _
                                        oCentPos.z
    oPoints(3).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oRevolution = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set Output2(Revolution)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRevolution
    Set oRevolution = Nothing
    Set oBspline = Nothing

    'Insert your code for output
    Dim oBody           As Object
    Dim oBsplineCurve   As IngrGeom3D.BSplineCurve3d
    Dim oRpoints(8)     As AutoMath.DPosition
    
    For iCount = 0 To 8
        Set oRpoints(iCount) = New AutoMath.DPosition
    Next
    
    oRpoints(0).Set oCentPos.x, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z
    oRpoints(1).Set oCentPos.x + 0.34 * parOperatorDiameter, oCentPos.y + 0.1 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(2).Set oCentPos.x + 0.34 * parOperatorDiameter, oCentPos.y + 0.32 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(3).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.35 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(4).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.35 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(5).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.55 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(6).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.55 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(7).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.83 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(8).Set oCentPos.x, oCentPos.y + 0.83 * dActOperatorHeight, oCentPos.z

    Set oBsplineCurve = PlaceTrBspline(2, oRpoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBsplineCurve, oAxis, oCentPos, 2 * PI, True)

    'Set Output3(body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oAxis = Nothing
    For iCount = 0 To 8
        Set oRpoints(iCount) = Nothing
    Next iCount
    Set oBsplineCurve = Nothing
    Set oBody = Nothing

    'Insert your code for Base
    Dim oBase       As Object
    Dim oEnPoint    As AutoMath.DPosition
    
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth

    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set Output4(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    Set oAxis = Nothing
    
    'Insert your code for Bolts
    Dim oBolt       As Object
    Dim oTransMat   As AutoMath.DT4x4
    
    Set oAxis = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    oAxis.Set 0, 1, 0

    oStPoint.Set oCentPos.x + 0.36 * parWidth, oCentPos.y + 0.15 * dActOperatorHeight, _
                                                            oCentPos.z + 0.36 * parWidth
    oEnPoint.Set oCentPos.x + 0.36 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, _
                                                            oCentPos.z + 0.36 * parWidth
    
    'Set Output5(Bolts)
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

    Set oBolt = Nothing
    Set oEnPoint = Nothing
    Set oStPoint = Nothing
    Set oCentPos = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

