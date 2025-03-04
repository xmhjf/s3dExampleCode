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
'   Author:         MA
'   Creation Date:  Wednesday, Sep 12 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' References:
'1. Thermodynamic Steam Traps Product Brochure, Model T250,Angle; www.spiraxsarco.com/us
'2. Thermodynamic Steam Traps Product Brochure, Model T250,Straightway; www.spiraxsarco.com/us
'3. Thermodynamic Steam Traps Product Brochure, Model BPT21; www.spiraxsarco.com/us
'4. Thermodynamic Steam Traps Product Brochure, Model MST18; www.spiraxsarco.com/us
'5. Thermodynamic Steam Traps Product Brochure, Model UBP30; www.spiraxsarco.com/us
'6. www.armstrong-intl.com/common/allproductscatalog/seriestt.pdf
'7. RIFObi Thermo Controlled Condensate Steam Traps, Type 2102 Data Sheet;
'   Rifox-Hans Richter GmbH Spezialarmaturen; www.rifox.de
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
'''
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint1 As AutoMath.DPosition
    Dim oDir1        As AutoMath.DVector
    Dim oPlacePoint2 As AutoMath.DPosition
    Dim oDir2       As AutoMath.DVector
    Set oPlacePoint1 = New AutoMath.DPosition
    Set oDir1 = New AutoMath.DVector
    Set oPlacePoint2 = New AutoMath.DPosition
    Set oDir2 = New AutoMath.DVector
    
    Dim LineStrPoints(0 To 26)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oArcPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim objSteamTrapBdy As Object
    Dim objCyl1 As Object
    Dim objCyl2 As Object
    Dim objBox As Object
    Dim dCyldia As Double
    Dim dNozzleLength As Double
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oLine1   As IngrGeom3D.Line3d
    Dim oLine2   As IngrGeom3D.Line3d
    Dim oArc1  As IngrGeom3D.Arc3d
    Dim oCurve         As IngrGeom3D.ComplexString3d
    Dim oCurveCol      As Collection
    
    Dim parCenterlinetoCenterline As Double
    Dim parPipingSpecialtyHeight   As Double
    Dim parPipingSpecialtyDiameter As Double
    Dim parOutletCenterlineHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parFacetoCenter As Double
    Dim parCenterlineHeight As Double
    Dim parFacetoFace   As Double
    Dim parPipingSpecialtyDepth As Double
    Dim parPipingSpecialtyLength As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parInletCenterlineHeight As Double
    Dim parInletWidth As Double
    Dim parInlettoOutletCenterlineWidth As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
      
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parCenterlinetoCenterline = arrayOfInputs(2)
'    parPipingSpecialtyHeight = arrayOfInputs(3)
'    parPipingSpecialtyDiameter = arrayOfInputs(4)
'    parOutletCenterlineHeight = arrayOfInputs(5)
'    parBodyOutsideDiameter = arrayOfInputs(6)
'    parFacetoCenter = arrayOfInputs(7)
'    parCenterlineHeight = arrayOfInputs(8)
'    parFacetoFace = arrayOfInputs(9)
'    parPipingSpecialtyDepth = arrayOfInputs(10)
'    parPipingSpecialtyLength = arrayOfInputs(11)
'    parPipingSpecialtyWidth = arrayOfInputs(12)
'    parInletCenterlineHeight = arrayOfInputs(13)
'    parInletWidth = arrayOfInputs(14)
'    parInlettoOutletCenterlineWidth = arrayOfInputs(15)
'    parOffset = arrayOfInputs(16)
    parInsulationThickness = arrayOfInputs(17)

    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2
  

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
' This symbol is based on the following Part data basis values that govern its geometry
'  (Part data Basis value -405): Balanced pressure thermostatic steam trap, Type 1
'  (Part data Basis value -406): Balanced pressure thermostatic steam trap, Type 2
'  (Part data Basis value -407): Balanced pressure thermostatic steam trap, Type 3
'  (Part data Basis value -408): Balanced pressure thermostatic steam trap, Type 4
'  (Part data Basis value -409): Balanced pressure thermostatic steam trap, Type 5
'  (Part data Basis value -410): Balanced pressure thermostatic steam trap, Type 6
'  (Part data Basis value -411): Balanced pressure thermostatic steam trap, Type 7
    
    
    Select Case lPartDataBasis

        Case Is <= 1, 405

            parCenterlinetoCenterline = arrayOfInputs(2)
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parOutletCenterlineHeight = arrayOfInputs(5)
            parBodyOutsideDiameter = arrayOfInputs(6)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then _
                        parBodyOutsideDiameter = 0.55 * parPipingSpecialtyDiameter

            'Point 1
            LineStrPoints(0) = -parOutletCenterlineHeight / 2
            LineStrPoints(1) = 0
            LineStrPoints(2) = 0

            'Point 2
            LineStrPoints(3) = LineStrPoints(0)
            LineStrPoints(4) = parBodyOutsideDiameter / 2
            LineStrPoints(5) = LineStrPoints(2)

            'Point 3
            LineStrPoints(6) = 0.75 * (parPipingSpecialtyHeight - parOutletCenterlineHeight)
            LineStrPoints(7) = LineStrPoints(4)
            LineStrPoints(8) = LineStrPoints(2)

            'Point 4
            LineStrPoints(9) = LineStrPoints(6)
            LineStrPoints(10) = parPipingSpecialtyDiameter / 2
            LineStrPoints(11) = LineStrPoints(2)

            'Point 5
            LineStrPoints(12) = 0.9 * (parPipingSpecialtyHeight - parOutletCenterlineHeight)
            LineStrPoints(13) = LineStrPoints(10)
            LineStrPoints(14) = LineStrPoints(2)

            'Point 6
            LineStrPoints(15) = LineStrPoints(12)
            LineStrPoints(16) = parPipingSpecialtyDiameter / 4
            LineStrPoints(17) = LineStrPoints(2)

            'Point 7
            LineStrPoints(18) = (parPipingSpecialtyHeight - parOutletCenterlineHeight)
            LineStrPoints(19) = LineStrPoints(16)
            LineStrPoints(20) = LineStrPoints(2)
            
            'Point 8
            LineStrPoints(21) = LineStrPoints(18)
            LineStrPoints(22) = 0
            LineStrPoints(23) = LineStrPoints(2)
            
            'Point 9
            LineStrPoints(24) = LineStrPoints(0)
            LineStrPoints(25) = LineStrPoints(1)
            LineStrPoints(26) = LineStrPoints(2)

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    9, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            
            Set objSteamTrapBdy = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSteamTrapBdy
            Set objSteamTrapBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parOutletCenterlineHeight / 2, 0, 0
            oEnPoint.Set -0.7 * parOutletCenterlineHeight, 0, 0
            dCyldia = 0.2 * parPipingSpecialtyDiameter
            
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "Cyl1", objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -parCenterlinetoCenterline / 2, 0
            oEnPoint.Set 0, -0.7 * parCenterlinetoCenterline, 0
            dCyldia = 0.2 * parPipingSpecialtyDiameter
           
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.3 * parOutletCenterlineHeight
            oPlacePoint1.Set -parOutletCenterlineHeight - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.3 * parCenterlinetoCenterline
            oPlacePoint2.Set 0, -parCenterlinetoCenterline - sptOffset2 + depth2, 0
            oDir2.Set 0, -1, 0
           
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 406

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parBodyOutsideDiameter = arrayOfInputs(6)
            parFacetoCenter = arrayOfInputs(7)
            parCenterlineHeight = arrayOfInputs(8)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then _
                            parBodyOutsideDiameter = 0.6 * parPipingSpecialtyDiameter

            Dim LineStrPoints1(0 To 32)  As Double
            'Point 1
            LineStrPoints1(0) = 0
            LineStrPoints1(1) = (parPipingSpecialtyHeight - parCenterlineHeight)
            LineStrPoints1(2) = 0

            'Point 2
            LineStrPoints1(3) = -parPipingSpecialtyDiameter / 4
            LineStrPoints1(4) = LineStrPoints1(1)
            LineStrPoints1(5) = LineStrPoints1(2)

            'Point 3
            LineStrPoints1(6) = LineStrPoints1(3)
            LineStrPoints1(7) = 0.9 * (parPipingSpecialtyHeight - parCenterlineHeight)
            LineStrPoints1(8) = LineStrPoints1(2)

            'Point 4
            LineStrPoints1(9) = -parPipingSpecialtyDiameter / 2
            LineStrPoints1(10) = LineStrPoints1(7)
            LineStrPoints1(11) = LineStrPoints1(2)

            'Point 5
            LineStrPoints1(12) = LineStrPoints1(9)
            LineStrPoints1(13) = 0.75 * (parPipingSpecialtyHeight - parCenterlineHeight)
            LineStrPoints1(14) = LineStrPoints1(2)

            'Point 6
            LineStrPoints1(15) = -parBodyOutsideDiameter / 2
            LineStrPoints1(16) = LineStrPoints1(13)
            LineStrPoints1(17) = LineStrPoints1(2)

            'Point 7
            LineStrPoints1(18) = LineStrPoints1(15)
            LineStrPoints1(19) = -0.7 * parCenterlineHeight
            LineStrPoints1(20) = LineStrPoints1(2)

            'Point 8
            LineStrPoints1(21) = -parPipingSpecialtyDiameter / 8
            LineStrPoints1(22) = LineStrPoints1(19)
            LineStrPoints1(23) = LineStrPoints1(2)

            'Point 9
            LineStrPoints1(24) = LineStrPoints1(21)
            LineStrPoints1(25) = -parCenterlineHeight
            LineStrPoints1(26) = LineStrPoints1(2)

            'Point 10
            LineStrPoints1(27) = LineStrPoints1(0)
            LineStrPoints1(28) = LineStrPoints1(25)
            LineStrPoints1(29) = LineStrPoints1(2)

            'Point 11
            LineStrPoints1(30) = LineStrPoints1(0)
            LineStrPoints1(31) = LineStrPoints1(1)
            LineStrPoints1(32) = LineStrPoints1(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    11, LineStrPoints1)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)

            Set objSteamTrapBdy = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSteamTrapBdy
            Set objSteamTrapBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parBodyOutsideDiameter / 2, 0, 0
            oEnPoint.Set -0.75 * parFacetoCenter, 0, 0
            dCyldia = 0.2 * parPipingSpecialtyDiameter
            
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "Cyl1", objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set parBodyOutsideDiameter / 2, 0, 0
            oEnPoint.Set 0.75 * parFacetoCenter, 0, 0
            dCyldia = 0.2 * parPipingSpecialtyDiameter
           
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.25 * parFacetoCenter
            oPlacePoint1.Set -parFacetoCenter - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.25 * parFacetoCenter
            oPlacePoint2.Set parFacetoCenter + sptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
           
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 407

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parPipingSpecialtyDepth = arrayOfInputs(10)
            parPipingSpecialtyLength = arrayOfInputs(11)

            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.4 * parFacetoFace, 0, 0
            oEnPoint.Set 0.4 * parFacetoFace, 0, 0
            dCyldia = 0.4 * parPipingSpecialtyHeight
            
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Box1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.45 * parPipingSpecialtyLength, 0.5 * parPipingSpecialtyHeight, _
                                0.5 * parPipingSpecialtyDiameter
            oEnPoint.Set 0.45 * parPipingSpecialtyLength, 0.2 * parPipingSpecialtyHeight, _
                                -0.5 * parPipingSpecialtyDiameter
            
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
         
            'Set the output
            m_OutputColl.AddOutput "Box1", objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints2(0 To 14)  As Double
            'Point 1
            LineStrPoints2(0) = 0
            LineStrPoints2(1) = 0.5 * parPipingSpecialtyHeight
            LineStrPoints2(2) = 0

            'Point 2
            LineStrPoints2(3) = -0.5 * parPipingSpecialtyDiameter
            LineStrPoints2(4) = LineStrPoints2(1)
            LineStrPoints2(5) = LineStrPoints2(2)

            'Point 3
            LineStrPoints2(6) = -0.3 * parPipingSpecialtyDiameter
            LineStrPoints2(7) = parPipingSpecialtyHeight
            LineStrPoints2(8) = LineStrPoints2(2)

            'Point 4
            LineStrPoints2(9) = 0
            LineStrPoints2(10) = LineStrPoints2(7)
            LineStrPoints2(11) = LineStrPoints2(2)

            'Point 5
            LineStrPoints2(12) = LineStrPoints2(0)
            LineStrPoints2(13) = LineStrPoints2(1)
            LineStrPoints2(14) = LineStrPoints2(2)

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    5, LineStrPoints2)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)

            Set objSteamTrapBdy = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            m_OutputColl.AddOutput "SteamTrapBdy", objSteamTrapBdy
            Set objSteamTrapBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Box2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parPipingSpecialtyLength, -0.2 * parPipingSpecialtyHeight, _
                                0.1 * parPipingSpecialtyHeight
            oEnPoint.Set 0.5 * parPipingSpecialtyLength, -parPipingSpecialtyDepth, _
                                -0.1 * parPipingSpecialtyHeight
            
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       
            'Set the output
            m_OutputColl.AddOutput "Box2", objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.1 * parFacetoFace
            oPlacePoint1.Set -0.5 * parFacetoFace - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.1 * parFacetoFace
            oPlacePoint2.Set 0.5 * parFacetoFace + sptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
          
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 408

            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFacetoFace, 0, 0
            oEnPoint.Set 0.3 * parFacetoFace, 0, 0
            dCyldia = parPipingSpecialtyDiameter

            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.2 * parFacetoFace
            oPlacePoint1.Set -0.5 * parFacetoFace - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.2 * parFacetoFace
            oPlacePoint2.Set 0.5 * parFacetoFace + sptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
          
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 409

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parPipingSpecialtyWidth = arrayOfInputs(12)
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
            dCyldia = (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0, _
                        -0.45 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            oEnPoint.Set 0, 0, _
                        -0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            dCyldia = 0.8 * parFacetoFace
           
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Dim objCyl3 As Object
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0, _
                        -0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            oEnPoint.Set 0, 0, _
                        -(parPipingSpecialtyWidth + 0.5 * parPipingSpecialtyDiameter)
            dCyldia = 0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            
            Set objCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl3", objCyl3
            Set objCyl3 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, _
                    0.4 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter), _
                            -parPipingSpecialtyWidth
            oEnPoint.Set 0, _
                (0.08 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)) + _
                            0.8 * parPipingSpecialtyHeight, _
                                -parPipingSpecialtyWidth
            dCyldia = parPipingSpecialtyDiameter

            Set objCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl3
            Set objCyl3 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Create SemiEllipsoid
            Dim oQuarterEllipse As Object
            Dim oSemiEllipsoid As New IngrGeom3D.Revolution3d
            Dim oEllCen As AutoMath.DPosition
            Set oEllCen = New AutoMath.DPosition
            Dim oEllNormal As AutoMath.DVector
            Set oEllNormal = New AutoMath.DVector
            Dim oEllMajor As AutoMath.DPosition
            Set oEllMajor = New AutoMath.DPosition
            Dim dMMRatio As Double
            dMMRatio = 0.2 * (parPipingSpecialtyHeight - (0.4 * (parPipingSpecialtyWidth - _
                                                    parPipingSpecialtyDiameter / 2))) / _
                            (parPipingSpecialtyDiameter / 2)
            
            oEllCen.Set 0, parPipingSpecialtyHeight - 0.2 * (parPipingSpecialtyHeight - _
                                (0.4 * (parPipingSpecialtyWidth - _
                                            parPipingSpecialtyDiameter / 2))), _
                                            -parPipingSpecialtyWidth
            
            oEllNormal.Set 0, 0, -1
            oEllMajor.Set -0.5 * parPipingSpecialtyDiameter, 0, 0
            
            Set oQuarterEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                    Nothing, oEllCen.x, oEllCen.y, oEllCen.z, _
                                                    oEllNormal.x, oEllNormal.y, oEllNormal.z, _
                                                    oEllMajor.x, oEllMajor.y, oEllMajor.z, _
                                                    dMMRatio, 0, PI / 2)
            
            Dim oAxis As New AutoMath.DVector
            oAxis.Set 0, 1, 0
            
            Set oSemiEllipsoid = PlaceRevolution(m_OutputColl, oQuarterEllipse, oAxis, oEllCen, 2 * PI, True)
           
            'Set the Output
            m_OutputColl.AddOutput "BucketBody", oSemiEllipsoid
            Set oSemiEllipsoid = Nothing
            Set oQuarterEllipse = Nothing
            Set oEllCen = Nothing
            Set oEllNormal = Nothing
            Set oEllMajor = Nothing
            
            'Place Nozzle 1
            oPlacePoint1.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir1, oPlacePoint1)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            oPlacePoint2.Set parFacetoFace / 2 + sptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 410

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parInletCenterlineHeight = arrayOfInputs(13)
            parInletWidth = arrayOfInputs(14)
            parInlettoOutletCenterlineWidth = arrayOfInputs(15)
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth, 0, 0
            oEnPoint.Set -0.9 * parInlettoOutletCenterlineWidth, 0, 0
            dCyldia = 0.3 * parPipingSpecialtyHeight
            
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set 0, -0.85 * parInletCenterlineHeight, 0
            dCyldia = 0.3 * parPipingSpecialtyHeight
           
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Revolution
            'Line1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set -0.6 * parInlettoOutletCenterlineWidth, _
                                        -0.6 * parInletCenterlineHeight, 0
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Line2
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth, _
                                        -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set -0.6 * parInlettoOutletCenterlineWidth, _
                            (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            
            Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
          
            'Arc1
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth, _
                            (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            oEnPoint.Set 0, (parPipingSpecialtyHeight - parInletCenterlineHeight), 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            Set oCurveCol = New Collection
            oCurveCol.Add oLine1
            oCurveCol.Add oLine2
            oCurveCol.Add oArc1
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            
            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            PI = 4 * Atn(1)
            
            Set objSteamTrapBdy = PlaceRevolution(m_OutputColl, oCurve, oAxisVect, oCenPoint, 2 * PI, True)
         
            'Set the output
            m_OutputColl.AddOutput "SteamTrapBdy", objSteamTrapBdy
            Set objSteamTrapBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLine1 = Nothing
            Set oLine2 = Nothing
            Set oArc1 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            
            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.1 * parInlettoOutletCenterlineWidth

            oPlacePoint1.Set -parInlettoOutletCenterlineWidth - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.15 * parInletCenterlineHeight
            oPlacePoint2.Set 0, -parInletCenterlineHeight - sptOffset2 + depth2, 0
            oDir2.Set 0, -1, 0
           
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing
            
       Case 411

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parOffset = arrayOfInputs(16)
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.45 * parFacetoFace, 0, 0
            oEnPoint.Set 0.45 * parFacetoFace, 0, 0
            dCyldia = 0.4 * parPipingSpecialtyHeight
           
            Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1
            Set objCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Box 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFacetoFace, 0.2 * parPipingSpecialtyHeight, _
                                                    0.4 * parPipingSpecialtyHeight
            oEnPoint.Set 0.3 * parFacetoFace, 0.5 * parPipingSpecialtyHeight, _
                                                    -0.4 * parPipingSpecialtyHeight
           
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
         
            'Set the output
            m_OutputColl.AddOutput "Box1", objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0.5 * parPipingSpecialtyHeight, 0
            oEnPoint.Set 0, 0.9 * parPipingSpecialtyHeight, 0
            dCyldia = 0.8 * parPipingSpecialtyDiameter
           
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Create SemiEllipsoid
            Dim oQuarterEllipse2 As Object
            Dim oSemiEllipsoid2 As New IngrGeom3D.Revolution3d
            Dim oEllCen2 As AutoMath.DPosition
            Set oEllCen2 = New AutoMath.DPosition
            Dim oEllNormal2 As AutoMath.DVector
            Set oEllNormal2 = New AutoMath.DVector
            Dim oEllMajor2 As AutoMath.DPosition
            Set oEllMajor2 = New AutoMath.DPosition
            Dim dMMRatio2 As Double
            dMMRatio2 = (0.1 * parPipingSpecialtyHeight) / (0.4 * parPipingSpecialtyDiameter)
            oEllCen2.Set 0, 0.9 * parPipingSpecialtyHeight, 0
            oEllNormal2.Set 0, 0, 1
            oEllMajor2.Set 0.4 * parPipingSpecialtyDiameter, 0, 0
            
            Set oQuarterEllipse2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                    Nothing, oEllCen2.x, oEllCen2.y, oEllCen2.z, _
                                                    oEllNormal2.x, oEllNormal2.y, oEllNormal2.z, _
                                                    oEllMajor2.x, oEllMajor2.y, oEllMajor2.z, _
                                                    dMMRatio2, 0, PI / 2)
         
            Dim oAxis2 As New AutoMath.DVector
            oAxis2.Set 0, 1, 0

            Set oSemiEllipsoid2 = PlaceRevolution(m_OutputColl, oQuarterEllipse2, oAxis2, oEllCen2, 2 * PI, True)

            'Set the Output
            m_OutputColl.AddOutput "SteamTrapBdy", oSemiEllipsoid2
            Set oSemiEllipsoid2 = Nothing
            Set oQuarterEllipse2 = Nothing
            Set oEllCen2 = Nothing
            Set oEllNormal2 = Nothing
            Set oEllMajor2 = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parPipingSpecialtyDiameter * Cos(PI / 6), _
                                 -parOffset * Sin(PI / 6), 0
            oEnPoint.Set 0.4 * parPipingSpecialtyDiameter * Cos(PI / 6), _
                                -0.1 * parPipingSpecialtyHeight * Sin(PI / 6), 0
            dCyldia = 0.4 * parPipingSpecialtyHeight
          
            Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
         
            'Set the output
            m_OutputColl.AddOutput "Cyl2", objCyl2
            Set objCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.05 * parFacetoFace
            oPlacePoint1.Set -0.5 * parFacetoFace - sptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
         
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                    oDir1, oPlacePoint1, dNozzleLength)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
            Set objNozzle1 = Nothing
            Set oDir1 = Nothing
            Set oPlacePoint1 = Nothing

            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
            dNozzleLength = 0.05 * parFacetoFace
            oPlacePoint2.Set 0.5 * parFacetoFace + sptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
           
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                        oDir2, oPlacePoint2, dNozzleLength)
         
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

    Case Else
        GoTo ErrorLabel:

    End Select

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
