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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MA
'   Creation Date:  Monday, Feb 18 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy      who               change description
'   -----------     -----              ------------------
'   18.Feb.2008      MA               CR-119767 Created the symbol.
'   06.JUN.2008      MA               CR-142643 Implemented part data basis for the vaules 990 and 991.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Dim PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    PI = Atn(1) * 4
End Sub

Public Sub run(ByVal m_outputcoll As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim objInsFlangeBody As Object
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
    Dim dFlangeInsLength As Double
    Dim dFlangeInsWidth As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHubDiameter As Double
    Dim parFlangeWidth As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parHubDiameter = arrayOfInputs(4)
'    parFlangeWidth = arrayOfInputs(5)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_outputcoll, pipeDiam, flangeThick, flangeDiam, _
                                                                    cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_outputcoll, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                    cptOffset2, depth2
    
    dFlangeInsLength = flangeDiam + 2 * parInsulationThickness
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis
    Case Is <= 1, 313, 312
        parFacetoFace = arrayOfInputs(2)
        parHubDiameter = arrayOfInputs(4)
    Case 990, 991
        parFlangeWidth = arrayOfInputs(5)
    Case Else
        GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
    
    Dim dBoltPatternOffset As Double
    Dim dBoltPatternLength As Double
    Dim dBoltPatternWidth As Double
    Dim dDrillingTemplatePattern As Long
    
    Set oPipeComponent = oPartFclt
    Call oPipeComponent.GetDrillingTemplatePatternData(1, dDrillingTemplatePattern, _
                            dBoltPatternLength, dBoltPatternWidth, dBoltPatternOffset)

    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    Dim objInsConicalHub As Object
    Dim dConeRadius As Double
    Dim dPoints() As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oLine As IngrGeom3D.Line3d
    Dim objCollection As Collection
    Dim objFlangeOutline As IngrGeom3D.ComplexString3d
    Dim oTransmat As DT4x4
    Dim dOffset As Double
    Dim objInsulatedPort2 As Object
    Dim dFlangeBodyThickness As Double
    
    Select Case lPartDataBasis
    Case Is <= 1, 313, 312
        dFlangeBodyThickness = flangeThick + parInsulationThickness
        'Thickness of insulated flange should not exceed face to face dimension.
        If CmpDblGreaterthan(dFlangeBodyThickness, parFacetoFace) Then _
                                                    dFlangeBodyThickness = parFacetoFace
        
        If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
            dConeRadius = flangeDiam2 / 2
        Else
            dConeRadius = pipeDiam2 / 2
        End If
        
        If CmpDblEqual(parHubDiameter, 0) Then parHubDiameter = 1.2 * pipeDiam2
        If CmpDblGreaterthan(1.2 * pipeDiam2, flangeDiam) Then _
                            parHubDiameter = (flangeDiam + pipeDiam2) / 2
        
        'Create insulation for conical hub
        oBaseCenPoint.Set flangeThick, 0, 0
        oTopCenPoint.Set parFacetoFace, 0, 0
    
        Set objInsConicalHub = PlaceCone(m_outputcoll, oBaseCenPoint, oTopCenPoint, _
                                        parHubDiameter / 2 + parInsulationThickness, _
                                        dConeRadius + parInsulationThickness, True)
        'Set the output
        m_outputcoll.AddOutput "ConicalHubIns", objInsConicalHub
        Set objInsConicalHub = Nothing
       
        'Weldneck flange, with circular drilling template.
        If (dDrillingTemplatePattern = 1) Or (dDrillingTemplatePattern = 2) Or _
                                    (dDrillingTemplatePattern = 40) Then
            oStPoint.Set 0, 0, 0
            oEndPoint.Set dFlangeBodyThickness, 0, 0
        
            Set objInsFlangeBody = PlaceCylinder(m_outputcoll, oStPoint, _
                                                oEndPoint, dFlangeInsLength, True)
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
     
        If (dDrillingTemplatePattern = 5) Then
            'Weldneck flange, with square drilling template.
            dFlangeInsWidth = dFlangeInsLength
            oStPoint.Set 0, -dFlangeInsLength / 2, -dFlangeInsWidth / 2
            oEndPoint.Set dFlangeBodyThickness, dFlangeInsLength / 2, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            dFlangeInsWidth = dFlangeInsLength
            oStPoint.Set 0, -dFlangeInsLength / 2, -dFlangeInsWidth / 2
            oEndPoint.Set dFlangeBodyThickness, dBoltPatternLength / 2 + dBoltPatternOffset _
                                            + parInsulationThickness, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 15) Then
            'Weldneck flange, with rectangular drilling template.
            dFlangeInsWidth = (1.5 * flangeDiam * (dBoltPatternWidth / dBoltPatternLength)) + 2 * parInsulationThickness
            
            oStPoint.Set 0, -dFlangeInsLength / 2, -dFlangeInsWidth / 2
            oEndPoint.Set dFlangeBodyThickness, dFlangeInsLength / 2, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            oStPoint.Set 0, ((-dBoltPatternWidth / 2) - (0.15 * (dBoltPatternOffset _
                        + dBoltPatternWidth))) - parInsulationThickness, -dFlangeInsLength / 2
            oEndPoint.Set dFlangeBodyThickness, dBoltPatternWidth / 2 + dBoltPatternOffset _
                                    + parInsulationThickness, dFlangeInsLength / 2
            
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 25) Then
            'Weldneck flange, with rectangular drilling template with split flange face.
            Dim dZCoord As Double
            dZCoord = (0.5 * dBoltPatternWidth) + 0.5 * (flangeDiam - dBoltPatternLength)
            
            Dim dYCoord As Double
            dYCoord = 0.5 * dBoltPatternLength
            
            dFlangeInsWidth = (1.5 * flangeDiam * (dBoltPatternWidth / dBoltPatternLength))
            
            Dim dRadius As Double
            dRadius = 0.5 * (flangeDiam - dBoltPatternLength)
            
            Dim dConst As Double
            dConst = (0.5 * dBoltPatternWidth) - Sqr(dRadius ^ 2 - _
                                ((0.4 * flangeDiam) - (0.5 * dBoltPatternLength)) ^ 2)
            
            'Line
            oStPoint.Set 0, 0.4 * flangeDiam, dConst
            oEndPoint.Set 0, 0.4 * flangeDiam, -dConst
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEndPoint.x, oEndPoint.y, oEndPoint.z)
            Set objCollection = New Collection
            objCollection.Add oLine
            
            'Arc
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0.5 * dBoltPatternLength, -0.5 * dBoltPatternWidth
            oStPoint.Set 0, 0.4 * flangeDiam, -dConst
            oEndPoint.Set 0, 0.5 * dBoltPatternLength, -dZCoord
            
            Dim oArc1 As IngrGeom3D.Arc3d
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
            
            'Elliptical Arc
            Dim oEllipseCenter As Object
            Set oEllipseCenter = New AutoMath.DPosition
            oEllipseCenter.Set 0, 0, -dZCoord
            
            Dim dMajorRadius As Double
            Dim dMinorRadius As Double
            dMajorRadius = dYCoord
            dMinorRadius = dFlangeInsWidth / 2 - dZCoord
            'place a quarter ellipse
            Dim oEllipse As Object
            Set oEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                Nothing, _
                oEllipseCenter.x, oEllipseCenter.y, oEllipseCenter.z, _
                -1, 0, 0, _
                0, dYCoord, 0, _
                dMinorRadius / dMajorRadius, _
                0, PI)
            
            objCollection.Add oEllipse
            Set oEllipse = Nothing
            
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.5 * dBoltPatternLength, -0.5 * dBoltPatternWidth
            oStPoint.Set 0, -0.5 * dBoltPatternLength, -dZCoord
            oEndPoint.Set 0, -0.4 * flangeDiam, -dConst
    
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
    
            'Line
            oStPoint.Set 0, -0.4 * flangeDiam, -dConst
            oEndPoint.Set 0, -0.4 * flangeDiam, dConst
    
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEndPoint.x, oEndPoint.y, oEndPoint.z)
            objCollection.Add oLine
    
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.5 * dBoltPatternLength, 0.5 * dBoltPatternWidth
            oStPoint.Set 0, -0.4 * flangeDiam, dConst
            oEndPoint.Set 0, -0.5 * dBoltPatternLength, dZCoord
            
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
            
            'Elliptical Arc
            oEllipseCenter.Set 0, 0, dZCoord
            'place a quarter ellipse
            Set oEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                Nothing, _
                oEllipseCenter.x, oEllipseCenter.y, oEllipseCenter.z, _
                -1, 0, 0, _
                0, -dYCoord, 0, _
                dMinorRadius / dMajorRadius, _
                0, PI)
            
            objCollection.Add oEllipse
            Set oEllipse = Nothing
            Set oEllipseCenter = Nothing
            
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0.5 * dBoltPatternLength, 0.5 * dBoltPatternWidth
            oStPoint.Set 0, 0.5 * dBoltPatternLength, dZCoord
            oEndPoint.Set 0, 0.4 * flangeDiam, dConst
    
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
    
            oStPoint.Set 0, 0.4 * flangeDiam, dConst
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
            
            Set oTransmat = New DT4x4
            oTransmat.LoadIdentity
            oTransmat.[Scale] ((dZCoord + parInsulationThickness) / dZCoord)
            objFlangeOutline.Transform oTransmat
            
            Set objInsFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                     objFlangeOutline, 1, 0, 0, dFlangeBodyThickness, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
            Set oLine = Nothing
            Set oArc1 = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
            Set oGeomFactory = Nothing
        End If
        
        If (dDrillingTemplatePattern = 30) Then
            'Weldneck flange, with triangular drilling template.
            'The center is at the centroid of the equilateral triangle
            'Side of the triangle is equal to Flange Diameter.
            
            ReDim dPoints(0 To 8)  'representing points in the Y-Z plane
            'Point 1
             dPoints(0) = 0      'X
            'The height of the traingular flange is flange dia x cos (30deg).
            'The centroid divides the total height of the triangle in a way that the height
            'from centroid to the base is one-third of the total height of the triangle.
            dPoints(1) = -dFlangeInsLength * Cos(PI / 6) / 3 'Y
            dPoints(2) = -dFlangeInsLength * 0.5        'Z
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = -dPoints(2)
            'Point 3
            dPoints(6) = dPoints(0)
            dPoints(7) = dFlangeInsLength * Cos(PI / 6) * 2 / 3
            dPoints(8) = 0
        
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                     dPoints(0), dPoints(1), dPoints(2), _
                                     dPoints(3), dPoints(4), dPoints(5))
            Set objCollection = New Collection
            objCollection.Add oLine
        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                     dPoints(3), dPoints(4), dPoints(5), _
                                     dPoints(6), dPoints(7), dPoints(8))
            objCollection.Add oLine
        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                     dPoints(6), dPoints(7), dPoints(8), _
                                     dPoints(0), dPoints(1), dPoints(2))
            objCollection.Add oLine
        
            oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
        
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
        
            'Project the Flange Outline along X axis with length Flange body Thickness.
            Set objInsFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                     objFlangeOutline, 1, 0, 0, dFlangeBodyThickness, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
            Set oLine = Nothing
            Set oGeomFactory = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
        End If
     
        If (dDrillingTemplatePattern = 35) Then
            'Weldneck flange, with oval drilling template.
            dFlangeInsWidth = 0.4 * flangeDiam  'Assumtion Flange Width is 40% of its height
            
            ReDim dPoints(0 To 17) 'representing points in the Y-Z plane
            'Point 1
            dPoints(0) = 0      'X
            dPoints(1) = 0      'Y
            dPoints(2) = -dFlangeInsWidth * 0.5        'Z
            
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dBoltPatternLength * 0.5
            dPoints(5) = -0.5 * (flangeDiam - dBoltPatternLength)
            
            'Point 3
            dPoints(6) = dPoints(0)
            dPoints(7) = dPoints(4)
            dPoints(8) = -dPoints(5)
            
            'Point 4
            dPoints(9) = dPoints(0)
            dPoints(10) = dPoints(1)
            dPoints(11) = -dPoints(2)
            
            'Point 5
            dPoints(12) = dPoints(6)
            dPoints(13) = -dPoints(7)
            dPoints(14) = dPoints(8)
            
            'Point 6
            dPoints(15) = dPoints(3)
            dPoints(16) = -dPoints(4)
            dPoints(17) = dPoints(5)
        
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dPoints(0), dPoints(1), dPoints(2), _
                        dPoints(3), dPoints(4), dPoints(5))
        
            Set objCollection = New Collection
            objCollection.Add oLine
        
            Dim oArc As IngrGeom3D.Arc3d
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, dBoltPatternLength / 2 + (0.5 * (flangeDiam - dBoltPatternLength)), 0
            oStPoint.Set dPoints(3), dPoints(4), dPoints(5)
            oEndPoint.Set dPoints(6), dPoints(7), dPoints(8)
    
            Set oArc = PlaceTrArcBy3Pts(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc
        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                         dPoints(6), dPoints(7), dPoints(8), _
                         dPoints(9), dPoints(10), dPoints(11))
            objCollection.Add oLine
        
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                         dPoints(9), dPoints(10), dPoints(11), _
                         dPoints(12), dPoints(13), dPoints(14))
            objCollection.Add oLine
        
            'This Arc is mirror image of the above arc.
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -(dBoltPatternLength / 2 + (0.5 * (flangeDiam - dBoltPatternLength))), 0
            oStPoint.Set dPoints(12), dPoints(13), dPoints(14)
            oEndPoint.Set dPoints(15), dPoints(16), dPoints(17)
    
            Set oArc = PlaceTrArcBy3Pts(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                         dPoints(15), dPoints(16), dPoints(17), _
                         dPoints(0), dPoints(1), dPoints(2))
            objCollection.Add oLine
        
            oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
            
            Set oTransmat = New DT4x4
            oTransmat.LoadIdentity
            oTransmat.[Scale] (((0.5 * dFlangeInsWidth) + parInsulationThickness) / (0.5 * dFlangeInsWidth))
            objFlangeOutline.Transform oTransmat
        
            'Project the Flange Upper Outline along X axis with length Flange body Thickness.
            Set objInsFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                    objFlangeOutline, 1, 0, 0, dFlangeBodyThickness, True)
           
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
            Set oLine = Nothing
            Set oArc = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
            Set oGeomFactory = Nothing
        
        End If
        
        'Place Insulation for Port 2
        If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace, 0, 0
            dFlangeBodyThickness = parFacetoFace - flangeThick2 - parInsulationThickness
            If CmpDblLessThan(dFlangeBodyThickness, 0) Then dFlangeBodyThickness = 0
            oEndPoint.Set dFlangeBodyThickness, 0, 0
    
            Set objInsulatedPort2 = PlaceCylinder(m_outputcoll, oStPoint, oEndPoint, _
                                            flangeDiam2 + 2 * parInsulationThickness, True)
    
            'Set the output
            m_outputcoll.AddOutput "InsPort2", objInsulatedPort2
            Set objInsulatedPort2 = Nothing
        End If
    
    Case 990
        If (dDrillingTemplatePattern = 5) Then
            'Weldneck flange, with square drilling template.
            dFlangeInsWidth = dFlangeInsLength
            oStPoint.Set 0, -dFlangeInsLength / 2, -dFlangeInsWidth / 2
            oEndPoint.Set flangeThick, dFlangeInsLength / 2, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.5 * dBoltPatternLength
            dFlangeInsWidth = parFlangeWidth + 2 * parInsulationThickness
            dOffset = (dBoltPatternLength / 2) + (flangeDiam - (dBoltPatternOffset + dBoltPatternLength))
            
            oStPoint.Set 0, -dOffset - parInsulationThickness, -dFlangeInsWidth / 2
            oEndPoint.Set flangeThick, dBoltPatternLength / 2 + dBoltPatternOffset _
                                            + parInsulationThickness, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 15) Then
            'Weldneck flange, with rectangular drilling template.
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeInsWidth = parFlangeWidth + 2 * parInsulationThickness
            
            oStPoint.Set 0, -dFlangeInsLength / 2, -dFlangeInsWidth / 2
            oEndPoint.Set flangeThick, dFlangeInsLength / 2, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeInsWidth = parFlangeWidth + 2 * parInsulationThickness
            dOffset = (dBoltPatternLength / 2) + (flangeDiam - (dBoltPatternOffset + dBoltPatternLength))
            
            oStPoint.Set 0, -dOffset - parInsulationThickness, -dFlangeInsWidth / 2
            oEndPoint.Set flangeThick, dBoltPatternLength / 2 + _
                            dBoltPatternOffset + parInsulationThickness, dFlangeInsWidth / 2
            
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        'Place Insulation for Port 2
        If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set flangeThick2, 0, 0
            dFlangeBodyThickness = flangeThick2 - NEGLIGIBLE_THICKNESS - parInsulationThickness
            If CmpDblLessThan(dFlangeBodyThickness, 0) Then dFlangeBodyThickness = 0
            oEndPoint.Set dFlangeBodyThickness, 0, 0

            Set objInsulatedPort2 = PlaceCylinder(m_outputcoll, oStPoint, oEndPoint, _
                                            flangeDiam2 + 2 * parInsulationThickness, True)

            'Set the output
            m_outputcoll.AddOutput "InsPort2", objInsulatedPort2
            Set objInsulatedPort2 = Nothing
        End If
    
    Case 991
        Dim dBlockLength As Double
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.5 * dBoltPatternLength
            dFlangeInsWidth = parFlangeWidth + 2 * parInsulationThickness
            dOffset = (dBoltPatternLength / 2) + (flangeDiam - (dBoltPatternOffset + dBoltPatternLength))
            
            'Putting a check to ensure that flange thickness or block length is always greater than or equal to
            'flange diameter of Port2
            dBlockLength = flangeThick
            If CmpDblLessThan(dBlockLength, flangeDiam2) Then dBlockLength = flangeDiam2
            
            oStPoint.Set -dBlockLength / 2, -dOffset - parInsulationThickness, -dFlangeInsWidth / 2
            oEndPoint.Set dBlockLength / 2, dBoltPatternLength / 2 + dBoltPatternOffset _
                                            + parInsulationThickness, dFlangeInsWidth / 2
        
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeInsWidth = parFlangeWidth + 2 * parInsulationThickness
            dOffset = (dBoltPatternLength / 2) + (flangeDiam - (dBoltPatternOffset + dBoltPatternLength))
            
            'Putting a check to ensure that flange thickness or block length is always greater than or equal to
            'flange diameter of Port2
            dBlockLength = flangeThick
            If CmpDblLessThan(dBlockLength, flangeDiam2) Then dBlockLength = flangeDiam2
                        
            oStPoint.Set -dBlockLength / 2, -dOffset - parInsulationThickness, -dFlangeInsWidth / 2
            oEndPoint.Set dBlockLength / 2, dBoltPatternLength / 2 + _
                            dBoltPatternOffset + parInsulationThickness, dFlangeInsWidth / 2
            
            Set objInsFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objInsFlangeBody
        End If
        
        'Place Insulation for Port 2
        If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set 0, (dBoltPatternLength / 2 + dBoltPatternOffset), 0
            dFlangeBodyThickness = (dBoltPatternLength / 2 + dBoltPatternOffset) - flangeThick2 - parInsulationThickness
            If CmpDblLessThan(dFlangeBodyThickness, 0) Then dFlangeBodyThickness = 0
            oEndPoint.Set 0, dFlangeBodyThickness, 0

            Set objInsulatedPort2 = PlaceCylinder(m_outputcoll, oStPoint, oEndPoint, _
                                            flangeDiam2 + 2 * parInsulationThickness, True)

            'Set the output
            m_outputcoll.AddOutput "InsPort2", objInsulatedPort2
            Set objInsulatedPort2 = Nothing
        End If
    
    Case Else
        GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
    
    Set objInsFlangeBody = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oPipeComponent = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing
    Set oTransmat = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
