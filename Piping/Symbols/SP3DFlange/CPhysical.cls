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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Monday, Feb 18 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                ------------------
'   18.Feb.2008      MA               CR-119767 Created the symbol.
'   06.JUN.2008      MA               CR-142643 Implemented part data basis for the vaules 990 and 991.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Dim PI As Double

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
    Dim objFlangeBody As Object
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
    Dim dFlangeLength As Double
    Dim dFlangeWidth As Double
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
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    'This symbol can be associated with the following Part data basis values
    '1. Flange, Length-through-hub dimension excludes flange face projection,
    '                              conical hub - 313 (The default case)
    '2. Flange, Length-through-hub dimension includes flange face projection,
    '                              conical hub - 312
    '3. Flange, block - 990
    '4. Flange, L-block - 991
    
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    Dim blnLengthThrHubInclFaceProj As Boolean
    Select Case lPartDataBasis 'Case where Length-through-hub dimension excludes flange face projection
    Case Is <= 1, 313
        parFacetoFace = arrayOfInputs(2)
        parHubDiameter = arrayOfInputs(4)
        blnLengthThrHubInclFaceProj = False
    Case 312  'Case where Length-through-hub dimension includes flange face projection
        parFacetoFace = arrayOfInputs(2)
        parHubDiameter = arrayOfInputs(4)
        blnLengthThrHubInclFaceProj = True
    Case 990, 991
        parFlangeWidth = arrayOfInputs(5)
    Case Else
        GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
    
    RetrieveParameters 1, oPartFclt, m_outputcoll, pipeDiam, flangeThick, flangeDiam, _
                                                        cptOffset, depth, blnLengthThrHubInclFaceProj
    RetrieveParameters 2, oPartFclt, m_outputcoll, pipeDiam2, flangeThick2, flangeDiam2, _
                                                        cptOffset2, depth2, blnLengthThrHubInclFaceProj
    dFlangeLength = flangeDiam
    Dim dBoltPatternOffset As Double
    Dim dBoltPatternLength As Double
    Dim dBoltPatternWidth As Double
    Dim dDrillingTemplatePattern As Long
    
    Call oPipeComponent.GetDrillingTemplatePatternData(1, dDrillingTemplatePattern, _
                            dBoltPatternLength, dBoltPatternWidth, dBoltPatternOffset)
     
    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    Dim objConicalHub As Object
    Dim dConeRadius As Double
    Dim dPoints() As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oLine As IngrGeom3D.Line3d
    Dim objCollection As Collection
    Dim objFlangeOutline As IngrGeom3D.ComplexString3d
    Dim dOffset As Double
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    Select Case lPartDataBasis
    Case Is <= 1, 313, 312
        'Assumption: Hub Diameter is an optional input
        If CmpDblEqual(parHubDiameter, 0) Then parHubDiameter = 1.2 * pipeDiam2
        If CmpDblGreaterthan(1.2 * pipeDiam2, flangeDiam) Then _
                            parHubDiameter = (flangeDiam + pipeDiam2) / 2
        
        If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
            dConeRadius = flangeDiam2 / 2
        Else
            dConeRadius = pipeDiam2 / 2
        End If
        
        'Create conical hub
        oBaseCenPoint.Set flangeThick, 0, 0
        oTopCenPoint.Set parFacetoFace, 0, 0
        
        Set objConicalHub = PlaceCone(m_outputcoll, oBaseCenPoint, oTopCenPoint, _
                                            parHubDiameter / 2, dConeRadius, True)
        'Set the output
        m_outputcoll.AddOutput "ConicalHub", objConicalHub
        Set objConicalHub = Nothing
        
        'Weldneck flange, with circular drilling template.
        If (dDrillingTemplatePattern = 1) Or (dDrillingTemplatePattern = 2) Or _
                                    (dDrillingTemplatePattern = 40) Then
            'Create circular flange face geometry
            oStPoint.Set 0, 0, 0
            oEndPoint.Set flangeThick, 0, 0
        
            Set objFlangeBody = PlaceCylinder(m_outputcoll, oStPoint, _
                                                oEndPoint, dFlangeLength, True)
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
         
        If (dDrillingTemplatePattern = 5) Then
            'Weldneck flange, with square drilling template.
            'Create square flange face geometry
            dFlangeWidth = dFlangeLength
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dFlangeLength / 2, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            'Create square with offset flange face geometry
            dFlangeWidth = dFlangeLength
            
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dBoltPatternLength / 2 + dBoltPatternOffset, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 15) Then
            'Weldneck flange, with rectangular drilling template.
            'Create rectangular flange face geometry
            dFlangeWidth = 1.5 * dFlangeLength * (dBoltPatternWidth / dBoltPatternLength)
            
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dFlangeLength / 2, dFlangeWidth / 2
        
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            'Create rectangular with offset flange face geometry
            oStPoint.Set 0, (-dBoltPatternWidth / 2) - (0.15 * (dBoltPatternOffset + dBoltPatternWidth)), -dFlangeLength / 2
            oEndPoint.Set flangeThick, dBoltPatternWidth / 2 + dBoltPatternOffset, dFlangeLength / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 25) Then
            'Weldneck flange, with rectangular drilling template with split flange face.
            Dim dZCoord As Double
            dZCoord = (0.5 * dBoltPatternWidth) + 0.5 * (dFlangeLength - dBoltPatternLength)
            
            Dim dYCoord As Double
            dYCoord = 0.5 * dBoltPatternLength
            
            dFlangeWidth = 1.5 * dFlangeLength * (dBoltPatternWidth / dBoltPatternLength)
            
            Dim dRadius As Double
            dRadius = 0.5 * (dFlangeLength - dBoltPatternLength)
            
            Dim dConst As Double
            dConst = (0.5 * dBoltPatternWidth) - Sqr(dRadius ^ 2 - _
                                ((0.4 * dFlangeLength) - (0.5 * dBoltPatternLength)) ^ 2)
            
            'Line
            oStPoint.Set 0, 0.4 * dFlangeLength, dConst
            oEndPoint.Set 0, 0.4 * dFlangeLength, -dConst
            
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
            oStPoint.Set 0, 0.4 * dFlangeLength, -dConst
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
            dMinorRadius = dFlangeWidth / 2 - dZCoord
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
            oEndPoint.Set 0, -0.4 * dFlangeLength, -dConst
    
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
    
            'Line
            oStPoint.Set 0, -0.4 * dFlangeLength, -dConst
            oEndPoint.Set 0, -0.4 * dFlangeLength, dConst
    
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEndPoint.x, oEndPoint.y, oEndPoint.z)
            objCollection.Add oLine
    
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.5 * dBoltPatternLength, 0.5 * dBoltPatternWidth
            oStPoint.Set 0, -0.4 * dFlangeLength, dConst
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
            oEndPoint.Set 0, 0.4 * dFlangeLength, dConst
    
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenPoint)
            objCollection.Add oArc1
    
            oStPoint.Set 0, 0.4 * dFlangeLength, dConst
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
    
            Set objFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                     objFlangeOutline, 1, 0, 0, flangeThick, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
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
            dPoints(1) = -dFlangeLength * Cos(PI / 6) / 3 'Y
            dPoints(2) = -dFlangeLength * 0.5        'Z
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = -dPoints(2)
            'Point 3
            dPoints(6) = dPoints(0)
            dPoints(7) = dFlangeLength * Cos(PI / 6) * 2 / 3
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
            Set objFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                     objFlangeOutline, 1, 0, 0, flangeThick, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
            Set oLine = Nothing
            Set oGeomFactory = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
        End If
         
        If (dDrillingTemplatePattern = 35) Then
            'Weldneck flange, with oval drilling template.
            dFlangeWidth = 0.4 * dFlangeLength  'Assumtion Flange Width is 40% of its height
            
            ReDim dPoints(0 To 17) 'representing points in the Y-Z plane
            'Point 1
            dPoints(0) = 0      'X
            dPoints(1) = 0      'Y
            dPoints(2) = -dFlangeWidth * 0.5        'Z
            
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dBoltPatternLength * 0.5
            dPoints(5) = -0.5 * (dFlangeLength - dBoltPatternLength)
            
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
            oCenPoint.Set 0, dBoltPatternLength / 2 + (0.5 * (dFlangeLength - dBoltPatternLength)), 0
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
            oCenPoint.Set 0, -(dBoltPatternLength / 2 + (0.5 * (dFlangeLength - dBoltPatternLength))), 0
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
        
            'Project the Flange Upper Outline along X axis with length Flange body Thickness.
            Set objFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                    objFlangeOutline, 1, 0, 0, flangeThick, True)
           
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
            Set oLine = Nothing
            Set oArc = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
            Set oGeomFactory = Nothing
        End If
        
        'Place Nozzle 1
        oDir.Set -1, 0, 0
        oPlacePoint.Set (depth - cptOffset), 0, 0
    
        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_outputcoll, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        oPlacePoint.Set parFacetoFace - (depth2 - cptOffset2), 0, 0
        oDir.Set 1, 0, 0
    
        Set objNozzle = CreateNozzle(2, oPartFclt, m_outputcoll, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
     
    Case 990
        If (dDrillingTemplatePattern = 5) Then
            'Weldneck flange, with square drilling template.
            'Create square flange face geometry
            dFlangeWidth = dFlangeLength
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dFlangeLength / 2, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            'Create square with offset flange face geometry
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.5 * dBoltPatternLength
            dFlangeWidth = parFlangeWidth
            dOffset = (dBoltPatternLength / 2) + (dFlangeLength - (dBoltPatternOffset + dBoltPatternLength))
            
            oStPoint.Set 0, -dOffset, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dBoltPatternLength / 2 + dBoltPatternOffset, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 15) Then
            'Weldneck flange, with rectangular drilling template.
            'Create rectangular flange face geometry
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeWidth = parFlangeWidth
            
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dFlangeLength / 2, dFlangeWidth / 2
        
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            'Create rectangular with offset flange face geometry
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeWidth = parFlangeWidth
            dOffset = (dBoltPatternLength / 2) + (dFlangeLength - (dBoltPatternOffset + dBoltPatternLength))
            
            oStPoint.Set 0, -dOffset, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dBoltPatternLength / 2 + dBoltPatternOffset, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        'Place Nozzle 1
        oDir.Set -1, 0, 0
        oPlacePoint.Set (depth - cptOffset), 0, 0
   
        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_outputcoll, oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        oPlacePoint.Set flangeThick - (depth2 - cptOffset2), 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_outputcoll, oDir, oPlacePoint)
 
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
    Case 991
        Dim dBlockLength As Double
        If (dDrillingTemplatePattern = 10) Then
            'Weldneck flange, with square drilling template with offset.
            'Create square with offset flange face geometry
            'Assumption: FlangeWidth is an optional input
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.5 * dBoltPatternLength
            dFlangeWidth = parFlangeWidth
            dOffset = (dBoltPatternLength / 2) + (dFlangeLength - (dBoltPatternOffset + dBoltPatternLength))
            
            'Putting a check to ensure that flange thickness or block length is always greater than or equal to
            'flange diameter of Port2
            dBlockLength = flangeThick
            If CmpDblLessThan(dBlockLength, flangeDiam2) Then dBlockLength = flangeDiam2
        
            oStPoint.Set -dBlockLength / 2, -dOffset, -dFlangeWidth / 2
            oEndPoint.Set dBlockLength / 2, dBoltPatternLength / 2 + dBoltPatternOffset, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)

            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        If (dDrillingTemplatePattern = 20) Then
            'Weldneck flange, with rectangular drilling template with offset.
            'Create rectangular with offset flange face geometry
            If CmpDblEqual(parFlangeWidth, 0) Then parFlangeWidth = 1.7 * dBoltPatternWidth
            dFlangeWidth = parFlangeWidth
            dOffset = (dBoltPatternLength / 2) + (dFlangeLength - (dBoltPatternOffset + dBoltPatternLength))
            
            'Putting a check to ensure that flange thickness or block length is always greater than or equal to
            'flange diameter of Port2
            dBlockLength = flangeThick
            If CmpDblLessThan(dBlockLength, flangeDiam2) Then dBlockLength = flangeDiam2
            
            oStPoint.Set -dBlockLength / 2, -dOffset, -dFlangeWidth / 2
            oEndPoint.Set dBlockLength / 2, dBoltPatternLength / 2 + dBoltPatternOffset, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_outputcoll, oStPoint, oEndPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), objFlangeBody
        End If
        
        'Place Nozzle 1
        oDir.Set -1, 0, 0
        oPlacePoint.Set -flangeDiam2 / 2 + depth - cptOffset, 0, 0

        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_outputcoll, oDir, oPlacePoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        oPlacePoint.Set 0, (dBoltPatternLength / 2 + dBoltPatternOffset) - depth2 + cptOffset2, 0
        oDir.Set 0, 1, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_outputcoll, oDir, oPlacePoint)
 
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    Case Else
        GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
        
    Set objFlangeBody = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oPipeComponent = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
