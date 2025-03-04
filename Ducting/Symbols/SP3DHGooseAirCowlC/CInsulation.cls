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
'   Creation Date:  Monday, Aug 6 2007
'   Description:
'    CR-124215 Goose Type Air Cowl Type C
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   06-Aug-2007     RUK     Created
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
    Dim parBWidth As Double             'b
    Dim parBDepth As Double             'a
    Dim parNeckWidth As Double          'B
    Dim parNeckDepth As Double          'A
    Dim parTransitionLength As Double   'D
    Dim parInnerRadius As Double        'r
    Dim parExpLen As Double             'd
    Dim parEndLen As Double             'e
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parBDepth = arrayOfInputs(4)
    parNeckWidth = arrayOfInputs(5)
    parNeckDepth = arrayOfInputs(6)
    parTransitionLength = arrayOfInputs(7)
    parInnerRadius = arrayOfInputs(8)
    parExpLen = arrayOfInputs(9)
    parEndLen = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)

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
    oStPoint.Set parTransitionLength, -(parNeckWidth / 2 + parInsulationThickness), _
                                (parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parNeckWidth / 2 + parInsulationThickness), _
                    (parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, (parNeckWidth / 2 + parInsulationThickness), _
                            (parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parNeckWidth / 2 + parInsulationThickness), _
                            -(parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength, (parNeckWidth / 2 + parInsulationThickness), _
                                -(parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, -(parNeckWidth / 2 + parInsulationThickness), _
                            -(parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength, -(parNeckWidth / 2 + parInsulationThickness), _
                                -(parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, -(parNeckWidth / 2 + parInsulationThickness), _
                        (parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set parTransitionLength, -(parNeckWidth / 2 + parInsulationThickness), _
                                (parNeckDepth + parInsulationThickness)
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
    Dim objInsGooseNeck1 As Object
    Dim oSweepCurve As IngrGeom3D.Arc3d
    
    Set oSweepCurve = New Arc3d
    Set oRectCompStr = New ComplexString3d
    Set oCollection = New Collection
    
    oCenter.Set parTransitionLength, 0, -(parNeckDepth + parInnerRadius)
    oStPoint.Set parTransitionLength, _
                            0, 0 + parInsulationThickness
    oEnPoint.Set parTransitionLength, 0, _
                    -(2 * parNeckDepth + 2 * parInnerRadius + parInsulationThickness)
    
    oVector.Set 0, 1, 0
    Set oSweepCurve = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                            oCenter.x, oCenter.y, oCenter.z, _
                                            oVector.x, oVector.y, oVector.z, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    'Line 1
    oStPoint.Set parTransitionLength, (parNeckWidth / 2 + parInsulationThickness), _
                            0 + parInsulationThickness
    oEnPoint.Set oStPoint.x, -(parNeckWidth / 2 + parInsulationThickness), _
                            0 + parInsulationThickness
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, -(parNeckWidth / 2 + parInsulationThickness), _
                            0 + parInsulationThickness
    oEnPoint.Set oStPoint.x, -(parNeckWidth / 2 + parInsulationThickness), _
                        -(parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength, -(parNeckWidth / 2 + parInsulationThickness), _
                        -(parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parNeckWidth / 2 + parInsulationThickness), _
                            -(parNeckDepth + parInsulationThickness)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength, (parNeckWidth / 2 + parInsulationThickness), _
                                    -(parNeckDepth + parInsulationThickness)
    oEnPoint.Set oStPoint.x, (parNeckWidth / 2 + parInsulationThickness), _
                                        0 + parInsulationThickness
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set parTransitionLength, (parNeckWidth / 2 + parInsulationThickness), _
                                    0 + parInsulationThickness
    Set oRectCompStr = PlaceTrCString(oStPoint, oCollection)
    
    Dim stnorm() As Double
    Dim ednorm() As Double
    
    Dim Surfset   As IngrGeom3D.IJElements
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oSweepCurve, oRectCompStr, _
                            CircularCorner, 0, stnorm, ednorm, False)
    
    'Set the Output
    For Each objInsGooseNeck1 In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsGooseNeck1
    Next objInsGooseNeck1
    
    Set objInsGooseNeck1 = Nothing
    Set Surfset = Nothing
    
    'Create Gooseneck 2
    Dim objInsGooseNeck2 As Object
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oSweepCurve, oRectCompStr, _
                            CircularCorner, 0, stnorm, ednorm, False)
    
    Dim oTransMat As AutoMath.DT4x4
    Set oTransMat = New DT4x4

    oTransMat.LoadIdentity
    oVector.Set 1, 0, 0
    oTransMat.Rotate PI, oVector
    For Each objInsGooseNeck2 In Surfset
        If Not objInsGooseNeck2 Is Nothing Then
            Exit For
        End If
    Next objInsGooseNeck2
    objInsGooseNeck2.Transform oTransMat

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsGooseNeck2
    Set objInsGooseNeck2 = Nothing
    Set Surfset = Nothing
    Set oRectCompStr = Nothing
    Set oSweepCurve = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oTransMat = Nothing
    
    'Create the Expansion 1
    Dim objInsExpansion As Object
    oCenter.Set (parTransitionLength - parExpLen), 0, _
                                (3 * parNeckDepth / 2) + 2 * parInnerRadius
    Set objInsExpansion = PlaceTrapezoid(m_OutputColl, oCenter, _
                            parBDepth + 2 * parInsulationThickness, _
                            parBWidth + 2 * parInsulationThickness, _
                            parNeckDepth + 2 * parInsulationThickness, _
                            parNeckWidth + 2 * parInsulationThickness, _
                            parExpLen, True, 0, PI / 2, 0)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsExpansion
    Set objInsExpansion = Nothing
    
    'Create Expansion 2
    oCenter.Set (parTransitionLength - parExpLen), 0, _
                                -((3 * parNeckDepth / 2) + 2 * parInnerRadius)
    Set objInsExpansion = PlaceTrapezoid(m_OutputColl, oCenter, _
                                parBDepth + 2 * parInsulationThickness, _
                                parBWidth + 2 * parInsulationThickness, _
                                parNeckDepth + 2 * parInsulationThickness, _
                                parNeckWidth + 2 * parInsulationThickness, _
                                parExpLen, True, 0, PI / 2, 0)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsExpansion
    Set objInsExpansion = Nothing
        
    'Create End 1
    Dim objInsEnd As Object
    
    oStPoint.Set (parTransitionLength - parExpLen), _
                    (parBWidth / 2 + parInsulationThickness), _
                    ((3 * parNeckDepth / 2) + 2 * parInnerRadius - parBDepth / 2 - parInsulationThickness)
    oEnPoint.Set (parTransitionLength - parExpLen - parEndLen), _
                    -(parBWidth / 2 + parInsulationThickness), _
                    oStPoint.z + parBDepth + 2 * parInsulationThickness
    Set objInsEnd = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsEnd
    Set objInsEnd = Nothing
    
    'Create End 2
    oStPoint.Set (parTransitionLength - parExpLen), _
                        parBWidth / 2 + parInsulationThickness, _
                -((3 * parNeckDepth / 2) + 2 * parInnerRadius - parBDepth / 2 - parInsulationThickness)
    oEnPoint.Set (parTransitionLength - parExpLen - parEndLen), _
                    -(parBWidth / 2 + parInsulationThickness), _
                    -((3 * parNeckDepth / 2) + 2 * parInnerRadius + parBDepth / 2 + parInsulationThickness)
    Set objInsEnd = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsEnd
    Set objInsEnd = Nothing
        
    'Remove references
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

