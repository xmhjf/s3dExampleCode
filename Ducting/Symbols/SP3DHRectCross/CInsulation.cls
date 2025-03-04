VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DHRectCross.HRectCross
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'    Definition of HVAC Rectangular-Reducing Cross
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Const EPSILON = 0.000001

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim ObjHvacHeaderIns As Object
    Dim ObjHvacBranch1Ins As Object
    Dim ObjHvacBranch2Ins As Object
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHLength As Double
    Dim parBLength As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHLength = arrayOfInputs(6)
    parBLength = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parPlaneOfBranch = arrayOfInputs(9)

    iOutput = 0

'   If this symbol is placed on a round duct then Depth parameter is ZERO.
'   To guard the symbol depth is made equal to width.
    If parDepth = 0 Then parDepth = parWidth
    If parBDepth = 0 Then parBDepth = parBWidth

' Define default minimum Tee or Reducing Tee header and branch lengths
    Dim dMinHL As Double
    Dim dMinBL As Double

    dMinHL = parWidth + 0.05
    dMinBL = dMinHL / 2
    If CmpDblLessThan(parHLength, dMinHL) Then
        parHLength = dMinHL
    End If
    If CmpDblLessThan(parBLength, dMinBL) Then
        parBLength = dMinBL
    End If

' Insert your code for output 4(Hvac Header Insulation)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = -parHLength / 2
    dY = parWidth / 2 + parInsulationThickness
    dZ = parDepth / 2 + parInsulationThickness

    Dim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = dX
    dlinepts(7) = -dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = dX
    dlinepts(10) = -dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Dim oAxisVec As New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    Set ObjHvacHeaderIns = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, parHLength, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacHeaderIns
    Set ObjHvacHeaderIns = Nothing

' Insert your code for output 5(Hvac Branch 1 Insulation)
    Dim dProjLength As Double
    If CmpDblLessThan(Abs(parPlaneOfBranch), EPSILON) Or CmpDblLessThan(Abs((Abs(parPlaneOfBranch) - PI)), EPSILON) Then
        dX = (parBWidth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBDepth / 2 + parInsulationThickness)
        dProjLength = parBLength - parWidth / 2 - parInsulationThickness
        
    Else
        dX = (parBDepth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBWidth / 2 + parInsulationThickness)
        dProjLength = parBLength - parDepth / 2 - parInsulationThickness
    End If

    'Ensure minimum projection length for third nozzle insulation
    If CmpDblLessThan(dProjLength, EPSILON) Then dProjLength = EPSILON

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = -dX
    dlinepts(7) = dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = -dX
    dlinepts(10) = dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    oAxisVec.Set 0, -1, 0
    Set ObjHvacBranch1Ins = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dProjLength, True)

'   Rotation
    Dim oTransMat As IJDT4x4
    Dim oAboutVec As IJDVector
    If parPlaneOfBranch <> 0 Then
        Set oTransMat = New AutoMath.DT4x4
        Set oAboutVec = New DVector

'       Axis of Rotation is header Axis
        oAboutVec.Set 1, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Rotate parPlaneOfBranch, oAboutVec
        ObjHvacBranch1Ins.Transform oTransMat
    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacBranch1Ins
    Set ObjHvacBranch1Ins = Nothing
' Insert your code for output 6(Hvac Branch 2 Insulation)
    
    Dim parOppPlaneOfBranch As Double

    parOppPlaneOfBranch = parPlaneOfBranch + PI

    If CmpDblLessThan(Abs(parOppPlaneOfBranch), EPSILON) Or CmpDblLessThan(Abs((Abs(parOppPlaneOfBranch) - PI)), EPSILON) Then
        dX = (parBWidth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBDepth / 2 + parInsulationThickness)
        dProjLength = parBLength - parWidth / 2 - parInsulationThickness
        
    Else
        dX = (parBDepth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBWidth / 2 + parInsulationThickness)
        dProjLength = parBLength - parDepth / 2 - parInsulationThickness
    End If

    'Ensure minimum projection length for third nozzle insulation
    If CmpDblLessThan(dProjLength, EPSILON) Then dProjLength = EPSILON

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = -dX
    dlinepts(7) = dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = -dX
    dlinepts(10) = dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    oAxisVec.Set 0, -1, 0
    Set ObjHvacBranch2Ins = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dProjLength, True)

'   Rotation
        If parOppPlaneOfBranch <> 0 Then
        Set oTransMat = New AutoMath.DT4x4
        Set oAboutVec = New DVector

'       Axis of Rotation is header Axis
        oAboutVec.Set 1, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Rotate parOppPlaneOfBranch, oAboutVec
        ObjHvacBranch2Ins.Transform oTransMat
    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacBranch2Ins
    Set ObjHvacBranch2Ins = Nothing
    Set oAboutVec = Nothing
    Set oAxisVec = Nothing
    Set oTransMat = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler

    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d

    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = geomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)

    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set geomFactory = Nothing
 
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function
