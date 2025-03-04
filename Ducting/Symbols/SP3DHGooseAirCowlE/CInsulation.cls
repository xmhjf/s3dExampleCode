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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          RUK
'   Creation Date:  Friday, Aug 17 2007
'   Description:    Goose Type Air Cowl Type E.
'   Source:         STD-M12C.pdf
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parBWidth As Double             'B
    Dim parBDepth As Double             'a
    Dim parTransitionLength As Double   'D
    Dim parHeight As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parBDepth = arrayOfInputs(4)
    parTransitionLength = arrayOfInputs(5)
    parHeight = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)

    'Origin is taken at the Port 1
    'For vertically routed Duct X-Axis towards UP (Elevation), Y is towards East
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oVector = New AutoMath.DVector
    
    Dim oCollection As Collection
    Dim iCount As Integer
    'Create the Transition
                    '---Ruled surface between circle and rectangle
    Dim objInsTransition As Object
    'Circle
        'To get ruled surface between circle and rectangle circle should be created as a
        'complex string having 4 arcs

    Dim oCirCompStr As IngrGeom3D.ComplexString3d
    Dim oArc As Object

    Set oCirCompStr = New IngrGeom3D.ComplexString3d
    Set oCollection = New Collection

    'Arc 1
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, parWidth / 2 + parInsulationThickness, 0
    oEnPoint.Set 0, 0, -(parWidth / 2 + parInsulationThickness)
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing


    'Arc 2
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, 0, -(parWidth / 2 + parInsulationThickness)
    oEnPoint.Set 0, -(parWidth / 2 + parInsulationThickness), 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    'Arc 3
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, -(parWidth / 2 + parInsulationThickness), 0
    oEnPoint.Set 0, 0, (parWidth / 2 + parInsulationThickness)
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    'Arc 4
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, 0, (parWidth / 2 + parInsulationThickness)
    oEnPoint.Set 0, (parWidth / 2 + parInsulationThickness), 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set 0, (parWidth / 2 + parInsulationThickness), 0
    Set oCirCompStr = PlaceTrCString(oStPoint, oCollection)

    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount

    'Rectangle
    Dim oRectCompStr As IngrGeom3D.ComplexString3d
    Dim oLine As Object
    
    Set oRectCompStr = New IngrGeom3D.ComplexString3d
    Set oCollection = New Collection
    
    'Line 1
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                        (parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parBWidth / 2 + parInsulationThickness), _
                        (parBDepth / 2 + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, (parBWidth / 2 + parInsulationThickness), _
                            (parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2 + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength, (parBWidth / 2 + parInsulationThickness), _
                        -(parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set oStPoint.x, -(parBWidth / 2 + parInsulationThickness), _
                        -(parBDepth / 2 + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                    -(parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set oStPoint.x, -(parBWidth / 2 + parInsulationThickness), _
                        (parBDepth / 2 + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                            (parBDepth / 2 + parInsulationThickness)
    Set oRectCompStr = PlaceTrCString(oStPoint, oCollection)
   
    Set objInsTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                  oCirCompStr, oRectCompStr, True)
   
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTransition
    Set objInsTransition = Nothing
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    Set oCirCompStr = Nothing
    Set oRectCompStr = Nothing
    
    'Create the Gooseneck 1
    Dim objInsGooseNeck As Object
    Dim oCompStr As IngrGeom3D.ComplexString3d
    
    Set oCompStr = New ComplexString3d
    Set oCollection = New Collection
    
    'Line 1
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                            0
    oEnPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2 + parInsulationThickness)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set parTransitionLength + parHeight, -(parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2 + parInsulationThickness)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength + parHeight, -(parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2 + parInsulationThickness)
    oEnPoint.Set parTransitionLength + parHeight, -(parBWidth / 2 + parInsulationThickness), _
                            -(parBDepth / 2)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength + parHeight, _
                -(parBWidth / 2 + parInsulationThickness), _
                -(parBDepth / 2)
    oEnPoint.Set parTransitionLength + parHeight + (parBDepth + parInsulationThickness) * Cos(PI / 6), _
                -(parBWidth / 2 + parInsulationThickness), _
                -(parBDepth / 2 + (parBDepth + parInsulationThickness) * Sin(PI / 6))
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 5
    oStPoint.Set parTransitionLength + parHeight + (parBDepth + parInsulationThickness) * Cos(PI / 6), _
                -(parBWidth / 2 + parInsulationThickness), _
                -(parBDepth / 2 + (parBDepth + parInsulationThickness) * Sin(PI / 6))
    oEnPoint.Set oStPoint.x + (parBDepth / 2) * Tan(PI / 6) * Sin(PI / 6), _
                    -(parBWidth / 2 + parInsulationThickness), _
                    -(-oStPoint.z - (parBDepth / 2) * Tan(PI / 6) * Cos(PI / 6))

    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Arc
    oCenter.Set parTransitionLength + parHeight + (parBDepth / 2) / Cos(PI / 6), _
                 -(parBWidth / 2 + parInsulationThickness), _
                 -(parBDepth / 2)
    oStPoint.Set oStPoint.x + (parBDepth / 2) * Tan(PI / 6) * Sin(PI / 6), _
                    -(parBWidth / 2 + parInsulationThickness), _
                    -(-oStPoint.z - (parBDepth / 2) * Tan(PI / 6) * Cos(PI / 6))
    Dim dTemp As Double
    dTemp = Sqr((parBDepth / 2 + parInsulationThickness) ^ 2 - (parBDepth / 2) ^ 2)
    oEnPoint.Set oCenter.x + dTemp, -(parBWidth / 2 + parInsulationThickness), _
                    (0)
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line 6
    oStPoint.Set parTransitionLength + parHeight + (parBDepth / 2) / Cos(PI / 6) + dTemp, _
                 -(parBWidth / 2 + parInsulationThickness), (0)
    oEnPoint.Set parTransitionLength, _
                    -(parBWidth / 2 + parInsulationThickness), 0
    
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Complex String
    oStPoint.Set parTransitionLength, -(parBWidth / 2 + parInsulationThickness), _
                            0
    Set oCompStr = PlaceTrCString(oStPoint, oCollection)
    
    oVector.Set 0, 1, 0
    Set objInsGooseNeck = PlaceProjection(m_OutputColl, oCompStr, oVector, _
                            parBWidth + 2 * parInsulationThickness, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsGooseNeck
    Set objInsGooseNeck = Nothing
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    
    'Create Gooseneck 2
    oVector.Set 0, 1, 0
    Set objInsGooseNeck = PlaceProjection(m_OutputColl, oCompStr, oVector, _
                    parBWidth + 2 * parInsulationThickness, True)
    
    Dim oTransMat As AutoMath.DT4x4
    Set oTransMat = New DT4x4
    oTransMat.LoadIdentity
    oVector.Set 1, 0, 0
    oTransMat.Rotate PI, oVector
    objInsGooseNeck.Transform oTransMat
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsGooseNeck
    Set objInsGooseNeck = Nothing
    Set oCompStr = Nothing

    'Remove references
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oTransMat = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
