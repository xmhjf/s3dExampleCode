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

'   ProgID          :  SP3DOpWEpoxyBonAsm1.OpWEPBonAsm1
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Tuesday, 23 Oct 2007
'   Description     :  White Epoxy Bonnet Assembly, Type 1
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

    Dim oCentPos    As AutoMath.DPosition
    Dim oTransMat   As AutoMath.DT4x4
    
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
    Dim oBsplineCurve   As IngrGeom3D.BSplineCurve3d
    Dim oRpoints(10)    As AutoMath.DPosition
    Dim oAxis           As AutoMath.DVector
    
    For iCount = 0 To 10
        Set oRpoints(iCount) = New AutoMath.DPosition
    Next
    Set oAxis = New AutoMath.DVector
    
    oRpoints(0).Set oCentPos.x, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z
    oRpoints(1).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.1 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(2).Set oCentPos.x + 0.18 * parOperatorDiameter, oCentPos.y + 0.4 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(3).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.4 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(4).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + 0.6 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(5).Set oCentPos.x + 0.24 * parOperatorDiameter, oCentPos.y + 0.6 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(6).Set oCentPos.x + 0.24 * parOperatorDiameter, oCentPos.y + 0.83 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(7).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.83 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(8).Set oCentPos.x + 0.5 * parOperatorDiameter, oCentPos.y + 0.9 * _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(9).Set oCentPos.x + 0.2 * parOperatorDiameter, oCentPos.y + _
                                        dActOperatorHeight, oCentPos.z
    oRpoints(10).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
    
    Set oBsplineCurve = PlaceTrBspline(2, oRpoints)
    oAxis.Set 0, 1, 0
    Set oBody = PlaceRevolution(m_OutputColl, oBsplineCurve, oAxis, oCentPos, 2 * PI, True)

    'Set Output1(Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oAxis = Nothing
    For iCount = 0 To 10
        Set oRpoints(iCount) = Nothing
    Next iCount
    Set oBsplineCurve = Nothing
    Set oBody = Nothing
    
    'Create Plate around the Bottom part of the Body
    Dim oGeomFactory        As IngrGeom3D.GeometryFactory
    Dim objPlate            As Object
    Dim dLinePts(0 To 14)   As Double
    Dim oLineStr            As IngrGeom3D.LineString3d
    
    Set oGeomFactory = New GeometryFactory
    Set oLineStr = New LineString3d
    
    dLinePts(0) = oCentPos.x + 0.18 * parOperatorDiameter
    dLinePts(1) = oCentPos.y + 0.1 * dActOperatorHeight
    dLinePts(2) = oCentPos.z - 0.001
    
    dLinePts(3) = oCentPos.x + 0.34 * parOperatorDiameter
    dLinePts(4) = dLinePts(1)
    dLinePts(5) = dLinePts(2)
    
    dLinePts(6) = dLinePts(3)
    dLinePts(7) = oCentPos.y + 0.3 * dActOperatorHeight
    dLinePts(8) = dLinePts(2)
    
    dLinePts(9) = dLinePts(0)
    dLinePts(10) = oCentPos.y + 0.4 * dActOperatorHeight
    dLinePts(11) = dLinePts(2)
    
    dLinePts(12) = dLinePts(0)
    dLinePts(13) = dLinePts(1)
    dLinePts(14) = dLinePts(2)
    
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePts)
    
    Dim oRotAxis As AutoMath.DVector
    Dim dRotAng As Double
        
    Set oRotAxis = New DVector
    Set oTransMat = New DT4x4
    Set oAxis = New DVector
    oRotAxis.Set 0, 1, 0
    
    iOutput = iOutput + 1
    For iCount = 1 To 6
        oTransMat.LoadIdentity
        oTransMat.Rotate dRotAng, oRotAxis
        
        oAxis.Set 0, 0, -1
        Set objPlate = PlaceProjection(m_OutputColl, oLineStr, oAxis, 0.002, True)
        objPlate.Transform oTransMat
        
        'Set Output2(Plates)
        m_OutputColl.AddOutput "Plate_", objPlate
        Set objPlate = Nothing
        
        dRotAng = dRotAng + 2 * PI / 6
    Next iCount
    Set oLineStr = Nothing
    Set oRotAxis = Nothing
    Set oTransMat = Nothing

    'Insert your code for Base
    Dim oBase       As Object
    Dim oStPoint    As AutoMath.DPosition
    Dim oEnPoint    As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set oCentPos.x - 0.5 * parWidth, oCentPos.y, oCentPos.z - 0.5 * parWidth
    oEnPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth

    Set oBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set Output3(Base)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
    Set oBase = Nothing
    Set oAxis = Nothing
    
    'Insert your code for Bolts
    Dim oBolt  As Object
    
    Set oAxis = New AutoMath.DVector
    Set oTransMat = New AutoMath.DT4x4
    oAxis.Set 0, 1, 0

    oStPoint.Set oCentPos.x + 0.36 * parWidth, oCentPos.y + 0.15 * dActOperatorHeight, _
                                                            oCentPos.z + 0.36 * parWidth
    oEnPoint.Set oCentPos.x + 0.36 * parWidth, oCentPos.y + 0.1 * dActOperatorHeight, _
                                                            oCentPos.z + 0.36 * parWidth
    
    'Set Output4(Bolts)
    iOutput = iOutput + 1
    For iCount = 0 To 3
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * parWidth, True)
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

