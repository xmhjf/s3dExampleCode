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
'   ProgID:         SP3DHRec2RecAdapter.HRec2RecAdapter
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

    Dim ObjHeader1Ins As Object
    Dim ObjHeader2Ins As Object
    
    Dim Width As Double
    Dim Depth As Double
    Dim Width2 As Double
    Dim Depth2 As Double
    Dim Length As Double
    Dim InsulationThickness As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Width2 = arrayOfInputs(4)
    Depth2 = arrayOfInputs(5)
    Length = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    

    iOutput = 0

'   If this symbol is placed on a round duct then Depth parameter is ZERO.
'   To guard the symbol depth is made equal to width.
    If Depth = 0 Then Depth = Width
    If Depth2 = 0 Then Depth2 = Width2

' Insert your code for output 4(Hvac Header Insulation - 1)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = -Length / 2
    dZ = Width / 2 + InsulationThickness
    dY = Depth / 2 + InsulationThickness

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
    Set ObjHeader1Ins = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, Length / 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHeader1Ins
    Set ObjHeader1Ins = Nothing

' Insert your code for output 5(Hvac Header Insulation - 2)
    
    dX = 0#
    dZ = Width2 / 2 + InsulationThickness
    dY = Depth2 / 2 + InsulationThickness

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

    
    oAxisVec.Set 1, 0, 0
    Set ObjHeader2Ins = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, Length / 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHeader2Ins
    Set ObjHeader2Ins = Nothing
    Set oAxisVec = Nothing
    
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
