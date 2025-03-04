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

'   ProgID          :  SP3DOpSSBonAsm1.OpSSBonAsm1
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Tuesday,23 OCT 2007
'   Description     :  Stainless Steel Bonnet Assembly, Type 1
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   23.OCT.2007      PK        CR-127653:Created the symbol.
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

    'Insert your code for Body
    Dim oBody           As Object
    Dim oCString        As IngrGeom3D.ComplexString3d
    Dim oAxis           As AutoMath.DVector
    Dim oArc            As Object
    Dim oLine           As Object
    Dim lines           As Collection
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oStPoint        As AutoMath.DPosition
    
    Set oAxis = New AutoMath.DVector
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set lines = New Collection
    oAxis.Set 0, 1, 0
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x, oCentPos.y + 0.1 _
                    * dActOperatorHeight, oCentPos.z, oCentPos.x + 0.3 * parOperatorDiameter, _
                    oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.3 * _
                    parOperatorDiameter, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.3 * parOperatorDiameter, oCentPos.y + 0.25 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x + 0.21 * _
                    parOperatorDiameter, oCentPos.y + 0.25 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.3 * parOperatorDiameter, oCentPos.y + 0.25 * _
                    dActOperatorHeight, oCentPos.z, oCentPos.x + 0.21 * parOperatorDiameter, _
                    oCentPos.y + 0.34 * dActOperatorHeight, oCentPos.z)
    lines.Add oArc
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.21 * _
                    parOperatorDiameter, oCentPos.y + 0.34 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.34 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.1 * _
                    parOperatorDiameter, oCentPos.y + 0.34 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.4 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.1 * _
                    parOperatorDiameter, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.13 * parOperatorDiameter, oCentPos.y + 0.4 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.13 * _
                    parOperatorDiameter, oCentPos.y + 0.4 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.13 * parOperatorDiameter, oCentPos.y + 0.55 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.13 * _
                    parOperatorDiameter, oCentPos.y + 0.55 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.17 * parOperatorDiameter, oCentPos.y + 0.55 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.17 * _
                    parOperatorDiameter, oCentPos.y + 0.55 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.17 * parOperatorDiameter, oCentPos.y + 0.8 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.17 * _
                    parOperatorDiameter, oCentPos.y + 0.8 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.8 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.8 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.87 * _
                    dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * _
                    parOperatorDiameter, oCentPos.y + 0.87 * dActOperatorHeight, oCentPos.z, _
                    oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + dActOperatorHeight, _
                    oCentPos.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.2 * _
                    parOperatorDiameter, oCentPos.y + dActOperatorHeight, oCentPos.z, _
                    oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z)
    lines.Add oLine
    
    oStPoint.Set oCentPos.x, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z
    Set oCString = PlaceTrCString(oStPoint, lines)
    Set oBody = PlaceRevolution(m_OutputColl, oCString, oAxis, oCentPos, 2 * PI, True)
    
    'Set Output1(body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oAxis = Nothing
    Set oArc = Nothing
    Set oLine = Nothing
    Set lines = Nothing
    Set oCString = Nothing
    Set oBody = Nothing
    Set oGeomFactory = Nothing
    
    'Insert your code for Base
    Dim oBase       As Object
    Dim oEnPoint    As AutoMath.DPosition
   
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth

    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set Output2(Base)
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
    
    'Set Output3(Bolts)
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
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

