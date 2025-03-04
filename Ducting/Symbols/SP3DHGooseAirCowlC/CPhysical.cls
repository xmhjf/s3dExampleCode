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
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          RUK
'   Creation Date:  Monday, Aug 6 2007
'   Description:
'    CR-124215 Goose Type Air Cowl Type C.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   06-Aug-2007     RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim objTransition As Object
    
    
    'Circle
        'To get ruled surface between circle and rectangle circle should be created as a
        'complex string having 4 arcs
        
    Dim oCirCompStr As IngrGeom3D.ComplexString3d
    Dim oArc As Object

    Set oCirCompStr = New IngrGeom3D.ComplexString3d
    Set oCollection = New Collection

    'Arc 1
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, parWidth / 2, 0
    oEnPoint.Set 0, 0, -parWidth / 2
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing


    'Arc 2
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, 0, -parWidth / 2
    oEnPoint.Set 0, -parWidth / 2, 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    'Arc 3
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, -parWidth / 2, 0
    oEnPoint.Set 0, 0, parWidth / 2
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    'Arc 4
    oCenter.Set 0, 0, 0
    oStPoint.Set 0, 0, parWidth / 2
    oEnPoint.Set 0, parWidth / 2, 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set 0, parWidth / 2, 0
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
    oStPoint.Set parTransitionLength, -parNeckWidth / 2, parNeckDepth
    oEnPoint.Set oStPoint.x, parNeckWidth / 2, parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, parNeckWidth / 2, parNeckDepth
    oEnPoint.Set oStPoint.x, parNeckWidth / 2, -parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength, parNeckWidth / 2, -parNeckDepth
    oEnPoint.Set oStPoint.x, -parNeckWidth / 2, -parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength, -parNeckWidth / 2, -parNeckDepth
    oEnPoint.Set oStPoint.x, -parNeckWidth / 2, parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set parTransitionLength, -parNeckWidth / 2, parNeckDepth
    Set oRectCompStr = PlaceTrCString(oStPoint, oCollection)
   
    Set objTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                  oCirCompStr, oRectCompStr, True)
   
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransition
    Set objTransition = Nothing
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    Set oCirCompStr = Nothing
    Set oRectCompStr = Nothing
    
    'Create the Gooseneck 1
    Dim objGooseNeck1 As Object
    Dim oSweepCurve As IngrGeom3D.Arc3d
    
    Set oSweepCurve = New Arc3d
    Set oRectCompStr = New ComplexString3d
    Set oCollection = New Collection
    
    oCenter.Set parTransitionLength, 0, -(parNeckDepth + parInnerRadius)
    oStPoint.Set parTransitionLength, 0, 0
    oEnPoint.Set parTransitionLength, 0, -(2 * parNeckDepth + 2 * parInnerRadius)
    
    oVector.Set 0, 1, 0
    Set oSweepCurve = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                            oCenter.x, oCenter.y, oCenter.z, _
                                            oVector.x, oVector.y, oVector.z, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    'Line 1
    oStPoint.Set parTransitionLength, parNeckWidth / 2, 0
    oEnPoint.Set oStPoint.x, -parNeckWidth / 2, 0
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength, -parNeckWidth / 2, 0
    oEnPoint.Set oStPoint.x, -parNeckWidth / 2, -parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    oStPoint.Set parTransitionLength, -parNeckWidth / 2, -parNeckDepth
    oEnPoint.Set oStPoint.x, parNeckWidth / 2, -parNeckDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength, parNeckWidth / 2, -parNeckDepth
    oEnPoint.Set oStPoint.x, parNeckWidth / 2, 0
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set parTransitionLength, parNeckWidth / 2, 0
    Set oRectCompStr = PlaceTrCString(oStPoint, oCollection)
    
    Dim stnorm() As Double
    Dim ednorm() As Double
    
    Dim Surfset   As IngrGeom3D.IJElements
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oSweepCurve, oRectCompStr, _
                            CircularCorner, 0, stnorm, ednorm, False)
    
    'Set the Output
    For Each objGooseNeck1 In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objGooseNeck1
    Next objGooseNeck1
    
    Set objGooseNeck1 = Nothing
    Set Surfset = Nothing
    
    'Create Gooseneck 2
    Dim objGooseNeck2 As Object
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oSweepCurve, oRectCompStr, _
                            CircularCorner, 0, stnorm, ednorm, False)
    
    Dim oTransMat As AutoMath.DT4x4
    Set oTransMat = New DT4x4

    oTransMat.LoadIdentity
    oVector.Set 1, 0, 0
    oTransMat.Rotate PI, oVector
    For Each objGooseNeck2 In Surfset
        If Not objGooseNeck2 Is Nothing Then
            Exit For
        End If
    Next objGooseNeck2
    objGooseNeck2.Transform oTransMat

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objGooseNeck2
    Set objGooseNeck2 = Nothing
    Set Surfset = Nothing
    Set oRectCompStr = Nothing
    Set oSweepCurve = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oTransMat = Nothing
    
    'Create the Expansion 1
    Dim objExpansion As Object
    oCenter.Set (parTransitionLength - parExpLen), 0, _
                                (3 * parNeckDepth / 2) + 2 * parInnerRadius
    Set objExpansion = PlaceTrapezoid(m_OutputColl, oCenter, parBDepth, parBWidth, _
                            parNeckDepth, parNeckWidth, parExpLen, True, 0, PI / 2, 0)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objExpansion
    Set objExpansion = Nothing
    
    'Create Expansion 2
    oCenter.Set (parTransitionLength - parExpLen), 0, _
                                -((3 * parNeckDepth / 2) + 2 * parInnerRadius)
    Set objExpansion = PlaceTrapezoid(m_OutputColl, oCenter, parBDepth, parBWidth, _
                            parNeckDepth, parNeckWidth, parExpLen, True, 0, PI / 2, 0)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objExpansion
    Set objExpansion = Nothing
        
    'Create End 1
    Dim objEnd As Object
    
    oStPoint.Set (parTransitionLength - parExpLen), parBWidth / 2, _
                    (3 * parNeckDepth / 2) + 2 * parInnerRadius - parBDepth / 2
    oEnPoint.Set (parTransitionLength - parExpLen - parEndLen), -oStPoint.y, _
                    oStPoint.z + parBDepth
    Set objEnd = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEnd
    Set objEnd = Nothing
    
    'Create End 2
    oStPoint.Set (parTransitionLength - parExpLen), parBWidth / 2, _
                    -((3 * parNeckDepth / 2) + 2 * parInnerRadius - parBDepth / 2)
    oEnPoint.Set (parTransitionLength - parExpLen - parEndLen), -oStPoint.y, _
                    -((3 * parNeckDepth / 2) + 2 * parInnerRadius + parBDepth / 2)
    Set objEnd = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEnd
    Set objEnd = Nothing
        
    'Create Pad Eye 1
    Dim dPadWidth As Double
    Dim dPadHt1 As Double
    Dim dPadHt2 As Double
    Dim dPadThickness As Double
    
    'Assume the Following
    dPadWidth = 0.15
    dPadHt1 = 0.15
    dPadHt2 = 0.05
    dPadThickness = 0.05
    
    Dim objPadEye1 As Object
    Dim oPadCompStr As IngrGeom3D.ComplexString3d
    
    Set oPadCompStr = New ComplexString3d
    Set oCollection = New Collection
    
    'Line 1
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth, _
                -dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth, _
                 dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth, _
                 dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + dPadHt1 / 6, _
                 dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
        
    'Line 3
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + dPadHt1 / 6, _
                 dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + _
                                        dPadHt1 - (dPadHt2 - dPadHt2 * Sin(PI / 6)), _
                 dPadWidth / 2 - (dPadHt2 * Sin(PI / 6)), (parNeckDepth + parInnerRadius + dPadThickness / 2)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
        
    'Arc
    oCenter.Set parTransitionLength + parInnerRadius + parNeckDepth + dPadHt1 - dPadHt2, _
                0, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + _
                                        dPadHt1 - (dPadHt2 - dPadHt2 * Sin(PI / 6)), _
                 dPadWidth / 2 - (dPadHt2 * Sin(PI / 6)), (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + _
                                        dPadHt1 - (dPadHt2 - dPadHt2 * Sin(PI / 6)), _
                 -(dPadWidth / 2 - (dPadHt2 * Sin(PI / 6))), (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oVector.Set 0, 0, -1
    
    Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                            oCenter.x, oCenter.y, oCenter.z, _
                                            oVector.x, oVector.y, oVector.z, _
                                            oStPoint.x, oStPoint.y, oStPoint.z, _
                                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line 4
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + _
                                        dPadHt1 - (dPadHt2 - dPadHt2 * Sin(PI / 6)), _
                 -(dPadWidth / 2 - (dPadHt2 * Sin(PI / 6))), (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + dPadHt1 / 6, _
                 -dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
        
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
        
    'Line 5
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth + dPadHt1 / 6, _
                 -dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    oEnPoint.Set parTransitionLength + parInnerRadius + parNeckDepth, _
                 -dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    Set oLine = Nothing
       
    oStPoint.Set parTransitionLength + parInnerRadius + parNeckDepth, _
                -dPadWidth / 2, (parNeckDepth + parInnerRadius + dPadThickness / 2)
    Set oPadCompStr = PlaceTrCString(oStPoint, oCollection)
    
    oVector.Set 0, 0, -1
    Set objPadEye1 = PlaceProjection(m_OutputColl, oPadCompStr, oVector, _
                                dPadThickness, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPadEye1
    Set objPadEye1 = Nothing
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    'Create PadEye 2
    Dim objPadEye2 As Object
    oVector.Set 0, 0, -1
    Set objPadEye2 = PlaceProjection(m_OutputColl, oPadCompStr, oVector, _
                                    dPadThickness, True)
    Set oTransMat = New DT4x4
    oTransMat.LoadIdentity
    oVector.Set 0, 0, -2 * (parNeckDepth + parInnerRadius)
    oTransMat.Translate oVector
    objPadEye2.Transform oTransMat
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPadEye2
    Set objPadEye2 = Nothing
    Set oTransMat = Nothing
    
    'Create the Support Sheet of thickness 7.5 MM
    Dim objSheet As Object
    oStPoint.Set (parTransitionLength + parInnerRadius + parNeckDepth), dPadWidth / 2, _
                    (parNeckDepth + parInnerRadius)
    oEnPoint.Set (parTransitionLength + parInnerRadius + parNeckDepth - 0.0075), -dPadWidth / 2, _
                    -(parNeckDepth + parInnerRadius)
    Set objSheet = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSheet
    Set objSheet = Nothing
    Set oPadCompStr = Nothing

    'Remove references
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oGeomFactory = Nothing
        
    '-------------------------------------------------------------------------------------------------
    '==================================================
    'BUILD HVACNOZZLE1 at the begining of the Header
    '==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    Dim dNozzLength(1 To 3) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    dNozzLength(3) = 0.001
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                parWidth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    Dim x As Double, y As Double, z As Double
    x = 0
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

    '==================================================
    'BUILD HVACNOZZLE2 at the end of the Reducing Tee
    '==================================================
    iPortIndex = 2
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width2

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                    GSCADNozzleEntities.Rectangular, lEndPrep(2), _
                    dThickness(2), dFlangeWidth(2), lFlowDir(2), parBWidth, _
                    parBDepth, CornerRadius, DimBaseOuter, PortStatus, _
                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                    m_OutputColl.ResourceManager)

    
    'Position of the nozzle should be the conenct point of the nozzle
    x = parTransitionLength - parExpLen - parEndLen
    y = 0#
    z = (3 * parNeckDepth / 2) + 2 * parInnerRadius
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir
    
    oDir.Set 0, 0, -1
    oDistribPort.SetRadialOrient oDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    
    '============================================
    'BUILD HVACNOZZLE3 at the end of the Branch
    '============================================

    iPortIndex = 3
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Rectangular, lEndPrep(3), _
                                dThickness(3), dFlangeWidth(3), lFlowDir(3), parBWidth, _
                                parBDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz3", dPortDepth(3), dCptOffSet(3), False, _
                                m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    ' Y Should be positive for Rectangular and negitive for Round
    x = parTransitionLength - parExpLen - parEndLen
    y = 0#
    z = -((3 * parNeckDepth / 2) + 2 * parInnerRadius)
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 0, 1
    oDistribPort.SetRadialOrient oDir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
