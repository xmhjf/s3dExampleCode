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

'   ProgID          :  SP3DOpSSSealedBonAsm2.OpSSSBonAsm2
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday,19 Oct 2007
'   Description     :  Stainless Steel Sealed Bonnet Assembly,Type 2
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   19.OCT.2007      PK        CR-127653:Created the symbol.
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
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    Dim oRpoints(11)    As AutoMath.DPosition
    Dim oAxis           As AutoMath.DVector
    
    Set oAxis = New AutoMath.DVector
    For iCount = 0 To 11
        Set oRpoints(iCount) = New AutoMath.DPosition
    Next
    
    oRpoints(0).Set oCentPos.x, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z
    oRpoints(1).Set oCentPos.x + 0.32 * parOperatorDiameter, oCentPos.y + 0.1 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(2).Set oCentPos.x + 0.32 * parOperatorDiameter, oCentPos.y + 0.32 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(3).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.35 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(4).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.35 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(5).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.55 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(6).Set oCentPos.x + 0.13 * parOperatorDiameter, oCentPos.y + 0.55 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(7).Set oCentPos.x + 0.13 * parOperatorDiameter, oCentPos.y + 0.85 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(8).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.87 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(9).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(10).Set oCentPos.x + 0.13 * parOperatorDiameter, oCentPos.y + _
                                                dActOperatorHeight, oCentPos.z
    oRpoints(11).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    
    Set oBspline = PlaceTrBspline(2, oRpoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)

    'Set Output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oAxis = Nothing
    For iCount = 0 To 11
        Set oRpoints(iCount) = Nothing
    Next iCount
    Set oBspline = Nothing
    Set oBody = Nothing

    'Insert your code for Base
    Dim oBase       As Object
    Dim oStPoint    As AutoMath.DPosition
    Dim oEnPoint    As AutoMath.DPosition
    
    Set oEnPoint = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth
    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set Output2(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing

    'Insert your code for Bolts
    Dim oBolt       As Object
    Dim oTransMat   As AutoMath.DT4x4
    
    Set oAxis = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    oAxis.Set 0, 1, 0

    oStPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.15 * dActOperatorHeight, _
                                                            oCentPos.z + 0.34 * parWidth
    oEnPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, _
                                                            oCentPos.z + 0.34 * parWidth

    'Set Output3(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * parWidth, True)
        oTransMat.LoadIdentity
        oTransMat.Rotate iCount * PI / 2, oAxis
        oBolt.Transform oTransMat
        m_OutputColl.AddOutput "Bolts_", oBolt
        Set oBolt = Nothing
    Next
    
    Set oAxis = Nothing
    Set oTransMat = Nothing
    Set oEnPoint = Nothing
    Set oStPoint = Nothing
    Set oCentPos = Nothing
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

