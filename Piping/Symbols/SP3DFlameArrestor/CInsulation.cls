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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         Sundar
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol constructed as per the Model 7628 in the Growth corporation's Catalog GrothC.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   27.Aug.2007     VRG     TR-124959 Insulation for ports was ensured for different end preparations
'   22.APR.2008     MA      CR-136577 Implemeted Part Data Basis for the values 911 to 920.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
 
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam1       As Double
    Dim flangeThick1    As Double
    Dim flangeDiam1     As Double
    Dim pipeDiam3       As Double
    Dim flangeThick3    As Double
    Dim flangeDiam3     As Double
    Dim pipeDiam4       As Double
    Dim flangeThick4    As Double
    Dim flangeDiam4     As Double
    Dim iOutput         As Integer

    Dim dInsulationDia   As Double

    Dim parFacetoFace As Double
    Dim parPipingSpecialtyHeight As Double
    Dim parInsulationThickness As Double
    Dim parBodyOutsideDiameter As Double
    Dim parPortHOffset As Double
    Dim parPortVOffset As Double
    Dim parDiameter As Double
    Dim parHoodHeight As Double
    Dim parElementFlangeOutsideDia As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parPipingSpecialtyHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parBodyOutsideDiameter = arrayOfInputs(5)
'    parPortHOffset = arrayOfInputs(6)
'    parPortVOffset = arrayOfInputs(7)
'    parDiameter = arrayOfInputs(8)
'    parHoodHeight = arrayOfInputs(9)
'    parElementFlangeOutsideDia = arrayOfInputs(10)
    
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim dLineStringPoints() As Double
    Dim oAxisVect As New AutoMath.DVector
    Dim oCenterPoint As New AutoMath.DPosition
    Set oAxisVect = New AutoMath.DVector
    Set oCenterPoint = New AutoMath.DPosition
    Dim oLineString As IngrGeom3D.LineString3d
    Dim objInsBody As Object
    Dim ObjConnector1 As Object
    Dim ObjConnector2 As Object
    Dim oPoints(1 To 8) As New AutoMath.DPosition
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Dim objInsulatedPort1 As Object
    Dim objInsulatedPort2 As Object
    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    Dim oCircle1 As IngrGeom3D.Circle3d
    Dim oCircle2 As IngrGeom3D.Circle3d
    Dim oEccentricCone As IngrGeom3D.RuledSurface3d
    Dim dConeDiameter As Double
    Dim dConst As Double
    Dim dAngle As Double
    Dim dConeRadius As Double
    Dim dNozzleLength As Double
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, _
                                                                        sptOffset, depth
    Dim dInsulationDiameter As Double
    If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
        dInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
    End If
    
    If lPartDataBasis = 1 Or lPartDataBasis = 911 Or lPartDataBasis = 912 Or _
            lPartDataBasis = 913 Or lPartDataBasis = 914 Or lPartDataBasis = 915 Or _
            lPartDataBasis = 916 Or lPartDataBasis = 917 Or lPartDataBasis = 918 Then

        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                            sptOffset, depth
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            dInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
    End If
    
    ' This symbol is based on the following three Part data basis values that govern its geometry,
    '   Flame arrester, in-line, horizontal - 911
    '   Flame arrester, in-line, horizontal with drain plugs - 912
    '   Flame arrester, in-line, horizontal with instrumentation taps - 913
    '   Flame arrester, in-line, horizontal, eccentric type 1 - 914
    '   Flame arrester, in-line, horizontal, eccentric type 2 - 915
    '   Flame arrester, in-line, vertical with drain plugs - 916
    '   Flame arrester, detonation - 917
    '   Flame arrester, detonation with drain port - 918
    '   Flame arrester, end-of-line, type 1 - 919
    '   Flame arrester, end-of-line, type 2 - 920
    
    Select Case lPartDataBasis
        
        Case Is <= 1
            parFacetoFace = arrayOfInputs(2)
            parPipingSpecialtyHeight = arrayOfInputs(3)
    
            'Insert your code for output 1(Insulation for Flame Arrestor  Body)
            Dim CenterPos As AutoMath.DPosition
            Set CenterPos = New AutoMath.DPosition
            Dim dCentertoBotOfBody  As Double
            Dim dFlameArrestorBodyLen  As Double
        
            'Assumptions:
            '1. The Body of Trap is taken to be cylinder
            '2. The cylinder length is taken at 40% of parFacetoFace
            '3. Diameter of Connector at Port1 and Port2 is taken to be 90% of flange diameter
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset, depth
            If CmpDblEqual(flangeThick1, 0) Then flangeThick1 = NEGLIGIBLE_THICKNESS
            If CmpDblLessThanOrEqualTo(flangeDiam1, pipeDiam1) Then flangeDiam1 = pipeDiam1
            dCentertoBotOfBody = flangeDiam1 / 2
            dFlameArrestorBodyLen = 0.4 * parFacetoFace
        
            oStPoint.Set -dFlameArrestorBodyLen / 2, _
                        (parPipingSpecialtyHeight / 2 - dCentertoBotOfBody), 0
            CenterPos.Set oStPoint.x, oStPoint.y, oStPoint.z
            oEnPoint.Set oStPoint.x + dFlameArrestorBodyLen, oStPoint.y, oStPoint.z
            dInsulationDia = parPipingSpecialtyHeight + 2 * parInsulationThickness
            
            Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                             dInsulationDia, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
        
            'Insert your code for output 2(Insulation for Connector Body at Port 1)
            Dim objCircle1       As IngrGeom3D.Circle3d
            Dim objCircle2       As IngrGeom3D.Circle3d
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set objCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                -parFacetoFace / 2 + flangeThick1, 0, 0, _
                                -1, 0, 0, 0.9 * flangeDiam1 / 2 + parInsulationThickness)
                                
            Set objCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                CenterPos.x, CenterPos.y, CenterPos.z, _
                                1, 0, 0, parPipingSpecialtyHeight / 2 + parInsulationThickness)
                                
            Set ObjConnector1 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                 objCircle1, objCircle2, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
        
            'Insert your code for output 3(Insulation for Connector Body at Port 2)
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            If CmpDblLessThanOrEqualTo(flangeDiam, pipeDiam) Then flangeDiam = pipeDiam
            If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
            Set objCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                parFacetoFace / 2 - flangeThick, 0, 0, _
                                1, 0, 0, 0.9 * flangeDiam / 2 + parInsulationThickness)
                                
            Set objCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                -CenterPos.x, CenterPos.y, CenterPos.z, _
                                -1, 0, 0, parPipingSpecialtyHeight / 2 + parInsulationThickness)
            
            Set ObjConnector2 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                 objCircle1, objCircle2, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
        Case 911
            parFacetoFace = arrayOfInputs(2)
            parElementFlangeOutsideDia = arrayOfInputs(10)
            
            'Insert your code for output 1(Insulation for Flame Arrestor  Body)
            'Point 1
            ReDim dLineStringPoints(0 To 23) As Double
            dLineStringPoints(0) = -0.15 * parFacetoFace
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = -0.15 * parFacetoFace
            dLineStringPoints(4) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = -0.05 * parFacetoFace
            dLineStringPoints(7) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = -0.05 * parFacetoFace
            dLineStringPoints(10) = 0.4 * parElementFlangeOutsideDia + parInsulationThickness
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = 0.05 * parFacetoFace
            dLineStringPoints(13) = 0.4 * parElementFlangeOutsideDia + parInsulationThickness
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = 0.05 * parFacetoFace
            dLineStringPoints(16) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(17) = 0
            
            'Point 7
            dLineStringPoints(18) = 0.15 * parFacetoFace
            dLineStringPoints(19) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(20) = 0
            
            'Point 8
            dLineStringPoints(21) = 0.15 * parFacetoFace
            dLineStringPoints(22) = 0
            dLineStringPoints(23) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, 0, 0
            Set objInsBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
            
            'Insert your code for output 2(Insulation for Connector Body at Port 1)
            oBaseCenPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oTopCenPoint.Set -0.15 * parFacetoFace, 0, 0
            
            Set ObjConnector1 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                    dInsulationDiameter / 2, 0.35 * parElementFlangeOutsideDia _
                                    + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
            
            'Insert your code for output 3(Insulation for Connector Body at Port 2)
            oBaseCenPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oTopCenPoint.Set 0.15 * parFacetoFace, 0, 0
            
            Set ObjConnector2 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                    dInsulationDiameter / 2, 0.35 * parElementFlangeOutsideDia _
                                    + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
    
        Case 912, 913
            parFacetoFace = arrayOfInputs(2)
            parPortHOffset = arrayOfInputs(6)
            parPortVOffset = arrayOfInputs(7)
            parElementFlangeOutsideDia = arrayOfInputs(10)
            
            'Insert your code for output 1(Insulation for Flame Arrestor  Body)
            'Point 1
            ReDim dLineStringPoints(0 To 23) As Double
            dLineStringPoints(0) = -0.15 * parFacetoFace
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = -0.15 * parFacetoFace
            dLineStringPoints(4) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = -0.05 * parFacetoFace
            dLineStringPoints(7) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = -0.05 * parFacetoFace
            dLineStringPoints(10) = 0.4 * parElementFlangeOutsideDia + parInsulationThickness
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = 0.05 * parFacetoFace
            dLineStringPoints(13) = 0.4 * parElementFlangeOutsideDia + parInsulationThickness
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = 0.05 * parFacetoFace
            dLineStringPoints(16) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(17) = 0
            
            'Point 7
            dLineStringPoints(18) = 0.15 * parFacetoFace
            dLineStringPoints(19) = parElementFlangeOutsideDia / 2 + parInsulationThickness
            dLineStringPoints(20) = 0
            
            'Point 8
            dLineStringPoints(21) = 0.15 * parFacetoFace
            dLineStringPoints(22) = 0
            dLineStringPoints(23) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, 0, 0
            Set objInsBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
        
            'Insert your code for output 7(Connector Body at Port 1)
            oBaseCenPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oTopCenPoint.Set -0.15 * parFacetoFace, 0, 0
            
            dConst = (0.5 * parElementFlangeOutsideDia - 0.5 * pipeDiam) / (0.5 * parFacetoFace - flangeThick)
            dAngle = Atn(dConst)
            
            dConeRadius = ((0.35 * parFacetoFace - flangeThick) * Tan(dAngle)) + 0.5 * pipeDiam
            
            Set ObjConnector1 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                        dInsulationDiameter / 2, dConeRadius + parInsulationThickness, True)
  
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing

            'Insert your code for output 8(Connector Body at Port 2)
            oBaseCenPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oTopCenPoint.Set 0.15 * parFacetoFace, 0, 0
            
            Set ObjConnector2 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                        dInsulationDiameter / 2, dConeRadius + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
        
        Case 914
            parFacetoFace = arrayOfInputs(2)
            parElementFlangeOutsideDia = arrayOfInputs(10)
            
            'Place Eccentric Cone 1
            Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -parFacetoFace / 2 + flangeThick, 0, 0, _
                            -1, 0, 0, pipeDiam / 2 + parInsulationThickness)

            dConeDiameter = (pipeDiam + parElementFlangeOutsideDia) / 2
            Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0, _
                            -1, 0, 0, dConeDiameter / 2 + parInsulationThickness)

            Set oEccentricCone = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                                m_OutputColl.ResourceManager, oCircle1, oCircle2, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEccentricCone
            Set oEccentricCone = Nothing
            
            'Place Eccentric Cone 2
            Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            parFacetoFace / 2 - flangeThick, 0, 0, _
                            1, 0, 0, pipeDiam / 2 + parInsulationThickness)
            
            Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0, _
                            1, 0, 0, dConeDiameter / 2 + parInsulationThickness)
            
            Set oEccentricCone = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                                m_OutputColl.ResourceManager, oCircle1, oCircle2, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEccentricCone
            Set oEccentricCone = Nothing
            
            'Place Cylinder 1
            oStPoint.Set -0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            oEnPoint.Set -0.05 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parElementFlangeOutsideDia + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
            'Place Cylinder 2
            oStPoint.Set 0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            oEnPoint.Set 0.05 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parElementFlangeOutsideDia + parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "cyl2", ObjConnector2
            Set ObjConnector2 = Nothing

            'Place Cylinder 3
            oStPoint.Set -0.05 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            oEnPoint.Set 0.05 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                0.8 * parElementFlangeOutsideDia + parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "cyl3", ObjConnector2
            Set ObjConnector2 = Nothing
        
        Case 915
            parFacetoFace = arrayOfInputs(2)
            parPipingSpecialtyHeight = arrayOfInputs(3)
            
            'Place Eccentric Cone 1
            Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -parFacetoFace / 2 + flangeThick, 0, 0, _
                            -1, 0, 0, pipeDiam / 2 + parInsulationThickness)
            
            dConeDiameter = (pipeDiam + parPipingSpecialtyHeight) / 2
            Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0, _
                            -1, 0, 0, dConeDiameter / 2 + parInsulationThickness)
            
            Set oEccentricCone = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                                m_OutputColl.ResourceManager, oCircle1, oCircle2, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEccentricCone
            Set oEccentricCone = Nothing
            
            'Place Eccentric Cone 2
            Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            parFacetoFace / 2 - flangeThick, 0, 0, _
                            1, 0, 0, pipeDiam / 2 + parInsulationThickness)
            
            Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0, _
                            1, 0, 0, dConeDiameter / 2 + parInsulationThickness)
            
            Set oEccentricCone = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                                m_OutputColl.ResourceManager, oCircle1, oCircle2, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEccentricCone
            Set oEccentricCone = Nothing
            
            'Place Cylinder 1
            oStPoint.Set -0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            oEnPoint.Set 0.15 * parFacetoFace, (dConeDiameter - pipeDiam) / 2, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parPipingSpecialtyHeight + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
        Case 916
            parFacetoFace = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(5)
            parPortHOffset = arrayOfInputs(6)
            parPortVOffset = arrayOfInputs(7)
        
            'Insert your code for output 6(Flame Arrestor Body)
            oStPoint.Set -0.15 * parFacetoFace, 0, 0
            oEnPoint.Set 0.15 * parFacetoFace, 0, 0
            
            Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parBodyOutsideDiameter + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
            
            'Insert your code for output 7(Connector Body at Port 1)
            oBaseCenPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oTopCenPoint.Set -0.15 * parFacetoFace, 0, 0
            
            dConst = (0.5 * parBodyOutsideDiameter - 0.5 * pipeDiam) / (0.5 * parFacetoFace - flangeThick)
            dAngle = Atn(dConst)
            
            dConeRadius = ((0.35 * parFacetoFace - flangeThick) * Tan(dAngle)) + 0.5 * pipeDiam
            
            Set ObjConnector1 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                        dInsulationDiameter / 2, dConeRadius + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
            
            'Insert your code for output 8(Connector Body at Port 2)
            oBaseCenPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oTopCenPoint.Set 0.15 * parFacetoFace, 0, 0
            
            Set ObjConnector2 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                        dInsulationDiameter / 2, dConeRadius + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
        
        Case 917, 918
            parFacetoFace = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(5)
            
            'Insert your code for output 1(Insulation for Flame Arrestor  Body)
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set -0.3 * parFacetoFace, pipeDiam / 2 + parInsulationThickness, 0
            oPoints(2).Set -0.28 * parFacetoFace, 0.4 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(3).Set -0.24 * parFacetoFace, 0.45 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(4).Set -0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(5).Set 0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(6).Set 0.24 * parFacetoFace, 0.45 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(7).Set 0.28 * parFacetoFace, 0.4 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(8).Set 0.3 * parFacetoFace, pipeDiam / 2 + parInsulationThickness, 0
            
            Set objBsplineCurve = PlaceTrBspline(7, oPoints)
            
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, 0, 0
        
            Set objInsBody = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVect, _
                                                        oCenterPoint, 2 * PI, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
            Set oCenterPoint = Nothing
            Set oAxisVect = Nothing
            Set objBsplineCurve = Nothing
        
            'Insert your code for output 2(Insulation for Connector Body at Port 1)
            oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oEnPoint.Set -0.3 * parFacetoFace, 0, 0
            
            Set ObjConnector1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                dInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
            
            'Insert your code for output 3(Insulation for Connector Body at Port 2)
            oStPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oEnPoint.Set 0.3 * parFacetoFace, 0, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                dInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
        
        Case 919
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parDiameter = arrayOfInputs(8)
      
            'Place Cone 1
            oBaseCenPoint.Set flangeThick, 0, 0
            oTopCenPoint.Set 0.3 * parPipingSpecialtyHeight, 0, 0
            
            Set ObjConnector1 = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, _
                                        dInsulationDiameter / 2, 0.25 * parDiameter _
                                        + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
            
            'Place Cylinder 1
            oStPoint.Set 0.3 * parPipingSpecialtyHeight, 0, 0
            oEnPoint.Set 0.55 * parPipingSpecialtyHeight, 0, 0
         
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        0.6 * parDiameter + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
            'Place Cylinder 2
            oStPoint.Set 0.55 * parPipingSpecialtyHeight, 0, 0
            oEnPoint.Set 0.8 * parPipingSpecialtyHeight, 0, 0
     
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        0.6 * parDiameter + parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
            'Place Cone 2
            'Point 1
            ReDim dLineStringPoints(0 To 11) As Double
            dLineStringPoints(0) = 0.8 * parPipingSpecialtyHeight - parInsulationThickness / 2
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = 0.8 * parPipingSpecialtyHeight - parInsulationThickness / 2
            dLineStringPoints(4) = parDiameter / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = parPipingSpecialtyHeight + parInsulationThickness
            dLineStringPoints(7) = 0.35 * parDiameter + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = parPipingSpecialtyHeight + parInsulationThickness
            dLineStringPoints(10) = 0
            dLineStringPoints(11) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, 0, 0
        
            Set objInsBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenterPoint, 2 * PI, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
    
        Case 920
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parDiameter = arrayOfInputs(8)
            parHoodHeight = arrayOfInputs(9)
        
            'Place Cylinder 1
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.3 * (parPipingSpecialtyHeight - parHoodHeight), 0, 0
            
            Set ObjConnector1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                dInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector1
            Set ObjConnector1 = Nothing
            
            'Place Cylinder 2
            oStPoint.Set 0.3 * (parPipingSpecialtyHeight - parHoodHeight), 0, 0
            oEnPoint.Set 0.7 * (parPipingSpecialtyHeight - parHoodHeight), 0, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                dInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
            'Place Cylinder 3
            oStPoint.Set 0.7 * (parPipingSpecialtyHeight - parHoodHeight), 0, 0
            oEnPoint.Set (parPipingSpecialtyHeight - parHoodHeight), 0, 0
            
            Set ObjConnector2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                dInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector2
            Set ObjConnector2 = Nothing
            
            'Place Cone 2
            'Point 1
            ReDim dLineStringPoints(0 To 11) As Double
            dLineStringPoints(0) = parPipingSpecialtyHeight - parHoodHeight - parInsulationThickness / 2
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = parPipingSpecialtyHeight - parHoodHeight - parInsulationThickness / 2
            dLineStringPoints(4) = parDiameter / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = parPipingSpecialtyHeight + parInsulationThickness
            dLineStringPoints(7) = 0.4 * parDiameter + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = parPipingSpecialtyHeight + parInsulationThickness
            dLineStringPoints(10) = 0
            dLineStringPoints(11) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, 0, 0
            Set objInsBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenterPoint, 2 * PI, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
            Set objInsBody = Nothing
        
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.

    End Select
            
    If lPartDataBasis = 1 Or lPartDataBasis = 911 Or lPartDataBasis = 912 Or _
            lPartDataBasis = 913 Or lPartDataBasis = 914 Or lPartDataBasis = 915 Or _
            lPartDataBasis = 916 Or lPartDataBasis = 917 Or lPartDataBasis = 918 Then
        'Insert your code for output 4(Insulation Port1 Side)
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick1 + parInsulationThickness, 0, 0
        dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam1, flangeDiam1) Then dInsulationDia = pipeDiam1 + 2 * parInsulationThickness
    
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dInsulationDia, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort1
        Set objInsulatedPort1 = Nothing
    
        'Insert your code for output 5(Insulation Port2 Side)
        oStPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
        Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dInsulationDia, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort2
        Set objInsulatedPort2 = Nothing
    End If
    
    If lPartDataBasis = 919 Or lPartDataBasis = 920 Then
        'Insert your code for output 4(Insulation Port1 Side)
        oStPoint.Set 0, 0, 0
        
        Dim dInsThickness As Double
        dInsThickness = flangeThick1 + parInsulationThickness
        If CmpDblGreaterthan(dInsThickness, parPipingSpecialtyHeight) Then _
                    dInsThickness = parPipingSpecialtyHeight
        
        oEnPoint.Set dInsThickness, 0, 0
        
        dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam1, flangeDiam1) Then dInsulationDia = pipeDiam1 + 2 * parInsulationThickness
        
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort1
        Set objInsulatedPort1 = Nothing
    End If
    
    If lPartDataBasis = 912 Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                        sptOffset, depth
        dNozzleLength = 0.0508
        oStPoint.Set -parPortHOffset, -parPortVOffset, 0
        oEnPoint.Set -parPortHOffset, -parPortVOffset + dNozzleLength, 0
        
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam3 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort1
        Set objInsulatedPort1 = Nothing
        
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, _
                                                                        sptOffset, depth
        oStPoint.Set parPortHOffset, -parPortVOffset, 0
        oEnPoint.Set parPortHOffset, -parPortVOffset + dNozzleLength, 0
        
        Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam4 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort2
        Set objInsulatedPort2 = Nothing
    End If
    
    If lPartDataBasis = 913 Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                        sptOffset, depth
        dNozzleLength = 0.0508
        oStPoint.Set -parPortHOffset, parPortVOffset, 0
        oEnPoint.Set -parPortHOffset, parPortVOffset - dNozzleLength, 0
        
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam3 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort1
        Set objInsulatedPort1 = Nothing
        
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, _
                                                                        sptOffset, depth
        oStPoint.Set parPortHOffset, parPortVOffset, 0
        oEnPoint.Set parPortHOffset, parPortVOffset - dNozzleLength, 0
        
        Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam4 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort2
        Set objInsulatedPort2 = Nothing
    End If
    
    If lPartDataBasis = 916 Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                        sptOffset, depth
        dNozzleLength = 0.0508
        dAngle = PI / 4
        oStPoint.Set -(-0.5 * parFacetoFace + parPortVOffset) * (-Cos(dAngle)), (-parPortHOffset) * Sin(dAngle), 0
        oEnPoint.Set -(-0.5 * parFacetoFace + parPortVOffset + dNozzleLength) * (-Cos(dAngle)), (-parPortHOffset + dNozzleLength) * Sin(dAngle), 0
        
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam3 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort1
        Set objInsulatedPort1 = Nothing
        
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, _
                                                                        sptOffset, depth
        dAngle = PI / 4
        oStPoint.Set (0.5 * parFacetoFace - parPortVOffset) * Cos(dAngle), (-parPortHOffset) * Sin(dAngle), 0
        oEnPoint.Set (0.5 * parFacetoFace - parPortVOffset - dNozzleLength) * Cos(dAngle), (-parPortHOffset + dNozzleLength) * Sin(dAngle), 0
        
        Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam4 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort2
        Set objInsulatedPort2 = Nothing
    End If
    
    If lPartDataBasis = 918 Then
        parPortHOffset = arrayOfInputs(6)
        parPortVOffset = arrayOfInputs(7)
    
        'Place insulation for Drain port 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                        sptOffset, depth
        dAngle = PI / 4
        dNozzleLength = 0.0508
        oStPoint.Set (0.5 * parFacetoFace - parPortHOffset) * Cos(dAngle), (-parPortVOffset) * Sin(dAngle), 0
        oEnPoint.Set (0.5 * parFacetoFace - parPortHOffset - dNozzleLength) * Cos(dAngle), (-parPortVOffset + dNozzleLength) * Sin(dAngle), 0
      
        Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam3 + parInsulationThickness * 2, True)
        'Set the output
        m_OutputColl.AddOutput "DrainPortIns", objInsulatedPort1
        Set objInsulatedPort1 = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxisVect = Nothing
    Set oCenterPoint = Nothing
    Set oLineString = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

