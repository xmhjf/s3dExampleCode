VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File            :  COperation.cls
'   Author          :  PK
'   Creation Date   :  Friday,August 10 2007
'   Description     :  HVAC Weather Tight Cover
'   Reference       :  STD-M4.pdf
'   CR number       :  CR-124219
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Operation" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
        
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), _
    arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim parWidth            As Double
    Dim parDepth            As Double
    Dim parCoverLength      As Double
    Dim parCoverWidth       As Double
    Dim parCoverThickness   As Double
    Dim parCoverEdgeHeight  As Double
    Dim iOutput             As Integer
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parCoverLength = arrayOfInputs(4)
    parCoverWidth = arrayOfInputs(5)
    parCoverThickness = arrayOfInputs(6)
    parCoverEdgeHeight = arrayOfInputs(7)
   
    iOutput = 0
    
    Dim oStPoint As New AutoMath.DPosition
   
    'Insert your code for output 8(Operational Cover)
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oCStr           As IngrGeom3D.ComplexString3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As New AutoMath.DVector
    Dim oProjection     As Object
    Dim opoints(3)      As New AutoMath.DPosition
    Dim iCount          As Integer
    
    opoints(0).Set -parCoverThickness, -0.5 * _
        parCoverLength, -0.5 * parCoverWidth
    opoints(1).Set -parCoverThickness, 1.5 * _
        parCoverLength, -0.5 * parCoverWidth
    opoints(2).Set -parCoverThickness - parCoverLength, _
        0.5 * parCoverLength, -0.5 * parCoverWidth
 
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oArc = PlaceTrArcBy3Pts(opoints(0), opoints(1), opoints(2))
    lines.Add oArc
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        -parCoverThickness, 1.5 * parCoverLength, -0.5 * parCoverWidth, _
        -parCoverThickness, -0.5 * parCoverLength, -0.5 * parCoverWidth)
    lines.Add oLine
    oStPoint.Set -parCoverThickness, -0.5 * parCoverLength, -0.5 * parCoverWidth
    Set oCStr = PlaceTrCString(oStPoint, lines)
        
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
     
    oAxis.Set 0, 0, 1
    Set oProjection = PlaceProjection(m_OutputColl, oCStr, oAxis, parCoverWidth, _
        True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
    
    Set oStPoint = Nothing
    Set oAxis = Nothing
    For iCount = 0 To 2
        Set opoints(iCount) = Nothing
    Next iCount
    Set oLine = Nothing
    Set oArc = Nothing
    Set lines = Nothing
    Set oCStr = Nothing
    Set oProjection = Nothing
    Set oGeomFactory = Nothing
      
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
