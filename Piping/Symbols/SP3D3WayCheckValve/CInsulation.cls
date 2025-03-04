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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         dkl/KKC
'   Creation Date:  Monday, Apr 11 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   04.05.2007      KKC     Created:  CR-117167  Create valve symbols for use in mining industry
'   27.Aug.2007     VRG      TR-124959  Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam3        As Double 'Port3 dNPD could be different
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim depth           As Double
    Dim flangeDiam      As Double
    Dim sptOffset3       As Double
    Dim depth3           As Double
    Dim flangeThick3     As Double
    Dim flangeDiam3      As Double
    
    Dim iOutput     As Integer
    Dim parPortSeparation As Double
    Dim parFace3toInletPort As Double
    Dim parAngle As Double
    Dim parValveWidth As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace3toInletPort = arrayOfInputs(2)   'B
    parPortSeparation = arrayOfInputs(3) 'A
    parAngle = arrayOfInputs(4)
    parValveWidth = arrayOfInputs(5) 'A
    parInsulationThickness = arrayOfInputs(6)
    Dim lPartDataBasis As Long
    
    iOutput = 0
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis <= 1) Or (lPartDataBasis = 265) Then
    ' The center is assumed at port 3 (The outlet port)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(parValveWidth, 0) Then parValveWidth = flangeDiam
        Dim dTravellingBallDia As Double
        dTravellingBallDia = parValveWidth
    
    '   3. Valve inlet adapter width is taken to be half of the dNPD.
        Dim dValInAdaWidth As Double
        dValInAdaWidth = parFace3toInletPort / 8
    
    '   4. The same value is taken for Valve outlet at port3
        Dim dNozzleLength3 As Double
        dNozzleLength3 = dValInAdaWidth
        
    ' Insert your code for output 7 (Body of Leg Port1 Side (joins port1 to port3)
        'Define half port separation value, Cosine and Sine
        Dim dHalfPortSeparation As Double
        Dim dCosA As Double
        Dim dSinA As Double
        
        dHalfPortSeparation = parPortSeparation / 2
        dCosA = Cos(parAngle / 2)
        dSinA = Sin(parAngle / 2)
    
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                            sptOffset3, depth3
        
        Dim oStPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        ' Face to Inlet Port dimension includes the flange thickness at port 3
        oStPoint.Set -parFace3toInletPort + dValInAdaWidth * dCosA, _
                                              dHalfPortSeparation - dValInAdaWidth * dSinA, 0
        
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        Dim oCircle1 As IngrGeom3D.Circle3d
        Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    -dCosA, dSinA, 0, _
                                                     dTravellingBallDia / 2 + parInsulationThickness)
        
        Dim oCircle3 As IngrGeom3D.Circle3d
        
        oStPoint.Set -dNozzleLength3, 0, 0
         
        Set oCircle3 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                     1, 0, 0, _
                                                    dTravellingBallDia / 2 + parInsulationThickness)
                           
        Dim objBodyLegPort1Side As Object
        Set objBodyLegPort1Side = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                oCircle1, oCircle3, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBodyLegPort1Side
        Set objBodyLegPort1Side = Nothing
        Set oCircle1 = Nothing
         
    ' Insert your code for output 8 (Body of Leg Port 2 Side (joins port2 to port3))
        Dim oCircle2 As IngrGeom3D.Circle3d
        
        oStPoint.Set -parFace3toInletPort + dValInAdaWidth * dCosA, _
                                        -dHalfPortSeparation + dValInAdaWidth * dSinA, 0
    
        Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    -dCosA, -dSinA, 0, _
                                                     dTravellingBallDia / 2 + parInsulationThickness)
    
        Dim objBodyLegPort2Side As Object
        Set objBodyLegPort2Side = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                oCircle2, oCircle3, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBodyLegPort2Side
        Set objBodyLegPort2Side = Nothing
        Set oCircle2 = Nothing
        Set oCircle3 = Nothing
    
    'Insert your code for output 9 (Triangular surfaces on either sides of the valve)
        Dim objTriangularSurfaces As Object
        Dim LineStrPoints(0 To 11) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        
        ' Co-rodinates of center of Circle 1 created above
        LineStrPoints(0) = -parFace3toInletPort + dValInAdaWidth * dCosA
        LineStrPoints(1) = dHalfPortSeparation - dValInAdaWidth * dSinA
        LineStrPoints(2) = -dTravellingBallDia / 2 - parInsulationThickness
        
        ' Co-rodinates of center of Circle 2 created above
        LineStrPoints(3) = LineStrPoints(0)
        LineStrPoints(4) = -dHalfPortSeparation + dValInAdaWidth * dSinA
        LineStrPoints(5) = LineStrPoints(2)
        
        LineStrPoints(6) = -dNozzleLength3
        LineStrPoints(7) = 0
        LineStrPoints(8) = LineStrPoints(2)
        
        LineStrPoints(9) = LineStrPoints(0)
        LineStrPoints(10) = LineStrPoints(1)
        LineStrPoints(11) = LineStrPoints(2)
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 0, 1
        Set objTriangularSurfaces = PlaceProjection(m_OutputColl, oLineString, oAxisVec, dTravellingBallDia + parInsulationThickness * 2, True)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTriangularSurfaces
        Set objTriangularSurfaces = Nothing
        Set oAxisVec = Nothing
        Set oGeomFactory = Nothing
        Set oLineString = Nothing
        
    ' Insert your code for output 13 (Insulated flange at inlet Port 1)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        Dim oEnPoint As AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
    
        oStPoint.Set -parFace3toInletPort, dHalfPortSeparation, 0
        oEnPoint.Set -parFace3toInletPort + dValInAdaWidth * dCosA, dHalfPortSeparation - dValInAdaWidth * dSinA, 0
    
        Dim objInsFlangePort1 As Object
        Dim dFlangeInsulationDiameter As Double
        
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
            dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
        End If
        
        Set objInsFlangePort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dFlangeInsulationDiameter, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort1
        Set objInsFlangePort1 = Nothing
        
    ' Insert your code for output 15 (Insulated flange at inlet Port 2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set -parFace3toInletPort, -dHalfPortSeparation, 0
        oEnPoint.Set oStPoint.x + dValInAdaWidth * dCosA, oStPoint.y + dValInAdaWidth * dSinA, 0
        
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
            dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
        End If
        
        Dim ObjInsInletFlangePort2 As Object
        Set ObjInsInletFlangePort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                    dFlangeInsulationDiameter, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsInletFlangePort2
        Set ObjInsInletFlangePort2 = Nothing
    
    ' Insert your code for output 17 (Insulated Flange at Port 3)
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                sptOffset3, depth3
        oStPoint.Set 0, 0, 0
        oEnPoint.Set oStPoint.x - dNozzleLength3, 0, 0
            
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then
            dFlangeInsulationDiameter = pipeDiam3 + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam3 + parInsulationThickness * 2
        End If
        
        Dim objInsFlangePort3 As Object
        Set objInsFlangePort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dFlangeInsulationDiameter, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort3
        Set objInsFlangePort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    ElseIf lPartDataBasis = 270 Then
    '   Center of the symbol is assumed at port 3
    '   Assumptions:
    '   1. If Valve Width is not provided then it is taken to be 1.1 times Pipe OD of port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(parValveWidth, 0) Then parValveWidth = 1.4 * flangeDiam
        
    '   2. Valve inlet width (the cylinder at inlets and outlet) is taken to be one-eighth of the parFace3toInletPort.
        dValInAdaWidth = parFace3toInletPort / 8
            
    '   Constructing Body of the valve
    '   Assumptions:
    '   1. Height from top of the elliptical arc to its ends is 50% of Face 3 to inlet port
        Dim dHeightofEllipticalArc As Double
        dHeightofEllipticalArc = parFace3toInletPort * 0.5
        
    '   2. Half-width of the elliptical arc is twice the Port Separation between inlet ports
        Dim dHalfWidthofEllipticalArc As Double
        dHalfWidthofEllipticalArc = parPortSeparation
    
    '   Computation Ends of the elliptical arc to botttom of the body
        Dim dEndofEllipArctoBotm As Double
        dEndofEllipArctoBotm = dHalfWidthofEllipticalArc * Tan(parAngle / 2)
        
        Dim objCurvesCollection  As Collection
        Set objCurvesCollection = New Collection
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
           
        Dim LineString(0 To 8) As Double
        ' Point 1
        LineString(0) = -(dValInAdaWidth + dHeightofEllipticalArc + dEndofEllipArctoBotm + parInsulationThickness)
        LineString(1) = 0
        LineString(2) = parValveWidth / 2 + parInsulationThickness
        ' Point 2
        LineString(3) = -(dValInAdaWidth + dHeightofEllipticalArc)
        LineString(4) = -dHalfWidthofEllipticalArc - parInsulationThickness / Cos(parAngle / 2)
        LineString(5) = LineString(2)
        ' Point 3
        LineString(6) = LineString(3)
        LineString(7) = -LineString(4)
        LineString(8) = LineString(2)
        
        Dim objLineRight As IngrGeom3D.Line3d
        Set objLineRight = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                LineString(0), LineString(1), LineString(2), _
                LineString(3), LineString(4), LineString(5))
        objCurvesCollection.Add objLineRight
      
        Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
        'Note major to minor Axis Ratio is dHeightofEllipticalArc / dHalfWidthofEllipticalArc
        Set objEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(m_OutputColl.ResourceManager, _
            LineString(3), 0, LineString(2), _
            0, 0, 1, _
            0, LineString(4), 0, _
        (dHeightofEllipticalArc / dHalfWidthofEllipticalArc), 0, PI)
        objCurvesCollection.Add objEllipticalArc
    
        Dim objLineLeft As IngrGeom3D.Line3d
        Set objLineLeft = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
            LineString(6), LineString(7), LineString(8), _
            LineString(0), LineString(1), LineString(2))
        objCurvesCollection.Add objLineLeft
          
        Dim objValveBodyOutline  As IngrGeom3D.ComplexString3d
        Set oStPoint = New AutoMath.DPosition
        
        oStPoint.Set LineString(0), 0, LineString(2)
        Set objValveBodyOutline = PlaceTrCString(oStPoint, objCurvesCollection)
    
        Dim objValveBody As Object
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 0, -1
        Set objValveBody = PlaceProjection(m_OutputColl, objValveBodyOutline, oAxisVec, _
                                        parValveWidth + 2 * parInsulationThickness, True)
        
    ' Set the output 3
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
        Set oAxisVec = Nothing
        Set objValveBodyOutline = Nothing
        Set objLineRight = Nothing
        Set objEllipticalArc = Nothing
        Set objLineLeft = Nothing
        Set geomFactory = Nothing
        
    ' Insert your code for output 10 (Insulated Valve inlet at Port 1)
    '    Define half port separation value, Cosine and Sine
        dHalfPortSeparation = parPortSeparation / 2
        dCosA = Cos(parAngle / 2)
        dSinA = Sin(parAngle / 2)

        Set oEnPoint = New AutoMath.DPosition
    
        oStPoint.Set -parFace3toInletPort, dHalfPortSeparation, 0
        oEnPoint.Set -parFace3toInletPort + parFace3toInletPort * 0.38 * dCosA, _
                                            dHalfPortSeparation - parFace3toInletPort * 0.38 * dSinA, 0
    
        Dim objInsValveInletPort1 As Object
        Set objInsValveInletPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                   pipeDiam + parInsulationThickness * 2, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveInletPort1
        Set objInsValveInletPort1 = Nothing
    
    ' Insert your code for output 11 (Insulated flange at inlet Port 1)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFace3toInletPort, dHalfPortSeparation, 0
        oEnPoint.Set oStPoint.x + (flangeThick + parInsulationThickness) * dCosA, oStPoint.y - (flangeThick + parInsulationThickness) * dSinA, 0
    
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
            dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
        End If
    
        Set objInsFlangePort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        dFlangeInsulationDiameter, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort1
        Set objInsFlangePort1 = Nothing
    
    ' Insert your code for output 12 (Insulated Valve inlet at Port 2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                            sptOffset, depth
        oStPoint.Set -parFace3toInletPort, -dHalfPortSeparation, 0
        oEnPoint.Set oStPoint.x + parFace3toInletPort * 0.38 * dCosA, oStPoint.y + parFace3toInletPort * 0.38 * dSinA, 0
    
        Dim objInsValveInletPort2 As Object
        Set objInsValveInletPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                    pipeDiam + parInsulationThickness * 2, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveInletPort2
        Set objInsValveInletPort2 = Nothing
    
    ' Insert your code for output 13 (Insulated Flange at Port 2)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFace3toInletPort, -dHalfPortSeparation, 0
        oEnPoint.Set oStPoint.x + (flangeThick + parInsulationThickness) * dCosA, oStPoint.y + (flangeThick + parInsulationThickness) * dSinA, 0
    
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
            dFlangeInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam + parInsulationThickness * 2
        End If
        
        Dim objInsFlangePort2 As Object
        Set objInsFlangePort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dFlangeInsulationDiameter, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort2
        Set objInsFlangePort2 = Nothing
       
    ' Insert your code for output 17 (Insulated Valve outlet at Port 3)
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
        oStPoint.Set 0, 0, 0
        oEnPoint.Set oStPoint.x - dValInAdaWidth * 2, 0, 0
    
        Dim objInsValveOutletPort3 As Object
        Set objInsValveOutletPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                    pipeDiam3 + parInsulationThickness * 2, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveOutletPort3
        Set objInsValveOutletPort3 = Nothing
        
    ' Insert your code for output 18 (Insulated Flange at Port 3)
        If CmpDblEqual(flangeThick3, 0) Then
            flangeThick3 = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set 0, 0, 0
        oEnPoint.Set -flangeThick3 - parInsulationThickness, 0, 0
    
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then
            dFlangeInsulationDiameter = pipeDiam3 + parInsulationThickness * 2
        Else
            dFlangeInsulationDiameter = flangeDiam3 + parInsulationThickness * 2
        End If
    
        Set objInsFlangePort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dFlangeInsulationDiameter, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsFlangePort3
        Set objInsFlangePort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    End If
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
