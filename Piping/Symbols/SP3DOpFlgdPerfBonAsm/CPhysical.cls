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

'   ProgID          :  SP3DOpFlgdPerfBonAsm.OpFPBonAsm
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday,01 Oct 2007
'   Description     :  Flanged Performance Bonnet Assembly
'   Reference       :
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   01-OCT-2007      PK         CR-127653:Created the symbol..
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
    
    Dim parOperatorHeight   As Double
    Dim parWidth            As Double
    Dim parOperatorDiameter As Double
    Dim parPositionerOffset As Double
    
    Dim oCentPos            As AutoMath.DPosition
    Dim iOutput             As Integer
    Dim iCount              As Integer
    Dim dActOperatorHeight  As Double
   
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
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As New AutoMath.DVector
    Dim oRevolution     As Object
    Dim oPoints(10)     As AutoMath.DPosition
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    
    For iCount = 0 To 10
       Set oPoints(iCount) = New AutoMath.DPosition
    Next
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    oPoints(0).Set oCentPos.x, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z
    oPoints(1).Set oCentPos.x + 0.2 * parWidth, oCentPos.y + 0.47 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(2).Set oCentPos.x + 0.2 * parWidth, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(3).Set oCentPos.x + 0.3 * parWidth, oCentPos.y + 0.67 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(4).Set oCentPos.x + 0.3 * parWidth, oCentPos.y + 0.77 * dActOperatorHeight, _
                                                                oCentPos.z
    oPoints(5).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.77 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(6).Set oCentPos.x + parOperatorDiameter / 2, oCentPos.y + 0.87 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(7).Set oCentPos.x + 0.25 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(8).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + 0.97 * _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(9).Set oCentPos.x + 0.1 * parOperatorDiameter, oCentPos.y + _
                                                                dActOperatorHeight, oCentPos.z
    oPoints(10).Set oCentPos.x, oCentPos.y + dActOperatorHeight, oCentPos.z
       
    Set oBspline = PlaceTrBspline(2, oPoints)
    oAxis.Set 0, 1, 0
    Set oRevolution = PlaceRevolution(m_OutputColl, oBspline, oAxis, oCentPos, 2 * PI, True)
    
    'Set output1(Revolution)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRevolution
    
    For iCount = 0 To 10
        Set oPoints(iCount) = Nothing
    Next iCount
    Set oAxis = Nothing
    Set oBspline = Nothing
    Set oRevolution = Nothing
    
    'Code for placing transition
    Dim oSquare     As IngrGeom3D.ComplexString3d
    Dim lines       As Collection
    Dim oLine       As Object
    Dim oStPoint    As AutoMath.DPosition
    
    Set oSquare = New IngrGeom3D.ComplexString3d
    Set oStPoint = New AutoMath.DPosition
    
    'code for forming square plane
    Set lines = New Collection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                         oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth, _
                         oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                         oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + 0.5 * parWidth, _
                         oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth, _
                         oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                         oCentPos.z - 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                          oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z - 0.5 * parWidth, _
                          oCentPos.x - 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                          oCentPos.z + 0.5 * parWidth)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * parWidth, _
                          oCentPos.y + 0.2 * dActOperatorHeight, oCentPos.z + 0.5 * parWidth, _
                          oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                          oCentPos.z + 0.5 * parWidth)
    lines.Add oLine
    oStPoint.Set oCentPos.x + 0.5 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                            oCentPos.z + 0.5 * parWidth
    Set oSquare = PlaceTrCString(oStPoint, lines)
    
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set oLine = Nothing
    Set lines = Nothing
    
    'code for forming circular plane
    Dim arcs    As Collection
    Dim oCircle As IngrGeom3D.ComplexString3d
    Set oCircle = New IngrGeom3D.ComplexString3d
    Set arcs = New Collection
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z + 0.2 * parWidth, oCentPos.x _
                          + 0.2 * parWidth, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x + 0.2 * parWidth, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z - 0.2 * parWidth)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z - 0.2 * parWidth, oCentPos.x - _
                          0.2 * parWidth, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z)
    arcs.Add oArc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oCentPos.x, oCentPos.y + _
                          0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x - 0.2 * parWidth, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z, oCentPos.x, _
                          oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z + 0.2 * parWidth)
    arcs.Add oArc
    oStPoint.Set oCentPos.x, oCentPos.y + 0.47 * dActOperatorHeight, oCentPos.z + 0.2 * parWidth
    Set oCircle = PlaceTrCString(oStPoint, arcs)
    
    For iCount = 1 To arcs.Count
        arcs.Remove 1
    Next iCount
    Set oArc = Nothing
    Set arcs = Nothing
    
    'code for making transition element between square plane and circular plane
    Dim oTransition As Object
    Set oTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                            oSquare, oCircle, True)
                            
    'code for output2(Transition)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTransition
    
    Set oTransition = Nothing
    Set oCircle = Nothing

    'code for placing a box type projection
    Dim oProjection As Object
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, -1, 0
    
    Set oProjection = PlaceProjection(m_OutputColl, oSquare, oAxis, 0.2 * dActOperatorHeight, _
                            True)
    
    'code for output3(Projection)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
    
    Set oProjection = Nothing
    Set oAxis = Nothing
    Set oGeomFactory = Nothing
    Set oSquare = Nothing
    
    'Code for providing bolts
    Dim oBolt       As Object
    Dim oEnPoint    As AutoMath.DPosition
    Dim oTransMat   As AutoMath.DT4x4
    
    Set oEnPoint = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, 1, 0
    
    oStPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.35 * dActOperatorHeight, _
                                                        oCentPos.z + 0.34 * parWidth
    oEnPoint.Set oCentPos.x + 0.34 * parWidth, oCentPos.y + 0.2 * dActOperatorHeight, _
                                                        oCentPos.z + 0.34 * parWidth
                                
    Set oTransMat = New AutoMath.DT4x4
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
    Set oTransMat = Nothing
    Set oAxis = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCentPos = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub



