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
'   Creation Date:  Friday, Aug 31 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-121408
'   References:
'   1. www.armstrong-intl.com/common/allproductscatalog/series200.pdf
'   2. www.armstrong-intl.com/common/allproductscatalog/series300.pdf
'   3. www.spiraxsarco.com/us
'   4. http://www.armstrong-intl.com/common/allproductscatalog/series800-813.pdf
'   5. www.armstrong-intl.com/common/allproductscatalog/series411g-421.pdf
'   6. www.armstrong-intl.com/common/allproductscatalog/series880.pdf

'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'   12.10.2007          MA         CR-121408 Removed (iOutput = iOutput + 1) for variable outputs.
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
    
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oArcPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim objBucketBody As Object
    Dim objCyl As Object
    Dim dCyldia As Double
    Dim oProjBdy As Object
    Dim oCurve As IngrGeom3D.ComplexString3d
    Dim oCurveCol As Collection
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim parFacetoFace As Double
    Dim parDiameter   As Double
    Dim parBodyOutsideDiameter As Double
    Dim parOffset As Double
    Dim parHeight As Double
    Dim parCenterlineHeight As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
      
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parDiameter = arrayOfInputs(3)
'    parBodyOutsideDiameter = arrayOfInputs(4)
'    parOffset = arrayOfInputs(5)
'    parHeight = arrayOfInputs(6)
'    parCenterlineHeight = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)

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
'  (Part data Basis value -397): Inverted bucket steam trap,
'                                         vertical installation, Type 1
'  (Part data Basis value -398): Inverted bucket steam trap,
'                                         vertical installation with offset, Type 2
'  (Part data Basis value -399): Inverted bucket steam trap,
'                                         horizontal installation, Type 1
'  (Part data Basis value -400): Inverted bucket steam trap,
'                                         horizontal installation, Type 2
'  (Part data Basis value -401): Inverted bucket steam trap,
'                                         horizontal installation, Type 3
'  (Part data Basis value -402): Inverted bucket steam trap, horizontal
'                                         installation with integral strainer, Type 4
   
    Select Case lPartDataBasis
    
        Case Is <= 1, 397
        
            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            
            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.75 * parDiameter
            
            'Point 1
            LineStrPoints(0) = -parFacetoFace / 2
            LineStrPoints(1) = parBodyOutsideDiameter / 2
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = 0.25 * parFacetoFace
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = parDiameter / 2
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = 0.35 * parFacetoFace
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = LineStrPoints(9)
            LineStrPoints(13) = 0
            LineStrPoints(14) = LineStrPoints(2)
        
            'Point 6
            LineStrPoints(15) = LineStrPoints(0)
            LineStrPoints(16) = LineStrPoints(13)
            LineStrPoints(17) = LineStrPoints(2)
            
            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            
            Set objBucketBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBucketBody
            Set objBucketBody = Nothing
            
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.35 * parFacetoFace, -0.13 * parDiameter, 0
            oEnPoint.Set 0.5 * parFacetoFace, -0.13 * parDiameter, 0
            dCyldia = 0.5 * parDiameter
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
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
    
        Case 398

            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parOffset = arrayOfInputs(5)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.75 * parDiameter

            'Point 1
            LineStrPoints(0) = -parFacetoFace / 2
            LineStrPoints(1) = parBodyOutsideDiameter / 2
            LineStrPoints(2) = 0

            'Point 2
            LineStrPoints(3) = 0.25 * parFacetoFace
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)

            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = parDiameter / 2
            LineStrPoints(8) = LineStrPoints(2)

            'Point 4
            LineStrPoints(9) = 0.35 * parFacetoFace
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)

            'Point 5
            LineStrPoints(12) = LineStrPoints(9)
            LineStrPoints(13) = 0
            LineStrPoints(14) = LineStrPoints(2)

            'Point 6
            LineStrPoints(15) = LineStrPoints(0)
            LineStrPoints(16) = LineStrPoints(13)
            LineStrPoints(17) = LineStrPoints(2)

            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)

            Set objBucketBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBucketBody
            Set objBucketBody = Nothing

            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.35 * parFacetoFace, -parOffset, 0
            oEnPoint.Set 0.5 * parFacetoFace, -parOffset, 0
            dCyldia = 0.3 * parDiameter
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

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
            oPlacePoint2.Set parFacetoFace / 2 + sptOffset2 - depth2, -parOffset, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
            Set objNozzle2 = Nothing
            Set oDir2 = Nothing
            Set oPlacePoint2 = Nothing

        Case 399
            
            parFacetoFace = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(6)
            parCenterlineHeight = arrayOfInputs(7)
            
            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.6 * parFacetoFace
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -0.2 * parFacetoFace, 0, 0
            dCyldia = pipeDiam
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.2 * parFacetoFace, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam2
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.2 * parFacetoFace, -0.3 * (parHeight - parCenterlineHeight), 0
            oEnPoint.Set 0.2 * parFacetoFace, -0.3 * (parHeight - parCenterlineHeight), 0
            dCyldia = 1.75 * (parHeight - parCenterlineHeight)
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -pipeDiam / 2, 0
            oEnPoint.Set 0, -0.9 * parCenterlineHeight, 0
            dCyldia = parBodyOutsideDiameter

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
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
            dMMRatio = (0.1 * parCenterlineHeight) / (0.5 * parBodyOutsideDiameter)
            oEllCen.Set 0, -0.9 * parCenterlineHeight, 0
            oEllNormal.Set 0, 0, -1
            oEllMajor.Set parBodyOutsideDiameter / 2, 0, 0

            Set oQuarterEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                    Nothing, oEllCen.x, oEllCen.y, oEllCen.z, _
                                                    oEllNormal.x, oEllNormal.y, oEllNormal.z, _
                                                    oEllMajor.x, oEllMajor.y, oEllMajor.z, _
                                                    dMMRatio, 0, PI / 2)
            
            Dim oAxis As New AutoMath.DVector
            oAxis.Set 0, -1, 0
            
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

        Case 400

            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(6)
            parCenterlineHeight = arrayOfInputs(7)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.8 * parDiameter
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam
  
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints2(0 To 26)  As Double
            'Point 1
            LineStrPoints2(0) = 0
            LineStrPoints2(1) = (parHeight - parCenterlineHeight) / 2
            LineStrPoints2(2) = 0
            
            'Point 2
            LineStrPoints2(3) = -parDiameter / 2
            LineStrPoints2(4) = LineStrPoints2(1)
            LineStrPoints2(5) = LineStrPoints2(2)
            
            'Point 3
            LineStrPoints2(6) = LineStrPoints2(3)
            LineStrPoints2(7) = 0.3 * (parHeight - parCenterlineHeight)
            LineStrPoints2(8) = LineStrPoints2(2)
            
            'Point 4
            LineStrPoints2(9) = -0.5 * parBodyOutsideDiameter
            LineStrPoints2(10) = LineStrPoints2(7)
            LineStrPoints2(11) = LineStrPoints2(2)
            
            'Point 5
            LineStrPoints2(12) = LineStrPoints2(9)
            LineStrPoints2(13) = -0.8 * parCenterlineHeight
            LineStrPoints2(14) = LineStrPoints2(2)
        
            'Point 6
            LineStrPoints2(15) = -0.2 * parBodyOutsideDiameter
            LineStrPoints2(16) = LineStrPoints2(13)
            LineStrPoints2(17) = LineStrPoints2(2)
            
            'Point 7
            LineStrPoints2(18) = LineStrPoints2(15)
            LineStrPoints2(19) = -parCenterlineHeight
            LineStrPoints2(20) = LineStrPoints2(2)
            
            'Point 8
            LineStrPoints2(21) = LineStrPoints2(0)
            LineStrPoints2(22) = LineStrPoints2(19)
            LineStrPoints2(23) = LineStrPoints2(2)
            
            'Point 9
            LineStrPoints2(24) = LineStrPoints2(0)
            LineStrPoints2(25) = LineStrPoints2(1)
            LineStrPoints2(26) = LineStrPoints2(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    9, LineStrPoints2)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, -1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)
            
            Set objBucketBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBucketBody
            Set objBucketBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.125 * parDiameter, (parHeight - parCenterlineHeight) / 2, 0
            oEnPoint.Set 0.125 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            dCyldia = 0.5 * parDiameter

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            oEnPoint.Set 0.25 * parDiameter, (parHeight - parCenterlineHeight), 0
            dCyldia = 0.2 * parDiameter

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
           
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Projection
            Dim oLine1   As IngrGeom3D.Line3d
            Dim oLine2   As IngrGeom3D.Line3d
            Dim oArc1    As IngrGeom3D.Arc3d
            Dim oLine3   As IngrGeom3D.Line3d
            Dim oLine4   As IngrGeom3D.Line3d
            Dim oLine5   As IngrGeom3D.Line3d

            'Line1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition

            oStPoint.Set -0.5 * parBodyOutsideDiameter, _
                                0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            oEnPoint.Set -parDiameter / 2, _
                                0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Line2
            oStPoint.Set -parDiameter / 2, _
                                    0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            oEnPoint.Set -parDiameter / 2, -0.7 * parCenterlineHeight, pipeDiam / 4
            Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Arc1
            oStPoint.Set -parDiameter / 2, -0.7 * parCenterlineHeight, pipeDiam / 4
            oEnPoint.Set -0.2 * parBodyOutsideDiameter, -0.9 * parCenterlineHeight, _
                                            pipeDiam / 4
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set -0.3 * parDiameter, -0.9 * parCenterlineHeight, pipeDiam / 4
            
            Set oArc1 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)
            
            'Line3
            oStPoint.Set -0.2 * parBodyOutsideDiameter, -0.9 * parCenterlineHeight, _
                                                pipeDiam / 4
            oEnPoint.Set -0.2 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                                pipeDiam / 4
            Set oLine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
          
            'Line4
            oStPoint.Set -0.2 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                                pipeDiam / 4
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                                pipeDiam / 4
            Set oLine4 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
           
            'Line5
            oStPoint.Set -0.5 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                                pipeDiam / 4
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, _
                                    0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            Set oLine5 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
          
            Set oCurveCol = New Collection
            oCurveCol.Add oLine1
            oCurveCol.Add oLine2
            oCurveCol.Add oArc1
            oCurveCol.Add oLine3
            oCurveCol.Add oLine4
            oCurveCol.Add oLine5

            oStPoint.Set -0.5 * parBodyOutsideDiameter, _
                                    0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4

            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
          
            Set oProjBdy = PlaceProjection(m_OutputColl, oCurve, oAxisVect, pipeDiam / 2, True)
       
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdy

            Set oProjBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLine1 = Nothing
            Set oLine2 = Nothing
            Set oArc1 = Nothing
            Set oLine3 = Nothing
            Set oLine4 = Nothing
            Set oLine5 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing
          
            'Place Projection
            Dim oLineA   As IngrGeom3D.Line3d
            Dim oLineB   As IngrGeom3D.Line3d
            Dim oLineC   As IngrGeom3D.Line3d
            Dim oArc    As IngrGeom3D.Arc3d

            'LineA
            oStPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
            oEnPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, pipeDiam2 / 4

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineA = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineB
            oStPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, pipeDiam2 / 4
            oEnPoint.Set 0.45 * parFacetoFace, pipeDiam2 / 2, pipeDiam2 / 4
            Set oLineB = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineC
            oStPoint.Set 0.45 * parFacetoFace, pipeDiam2 / 2, pipeDiam2 / 4
            oEnPoint.Set 0.45 * parFacetoFace, 0.8 * (parHeight - parCenterlineHeight), _
                            pipeDiam2 / 4
            Set oLineC = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Arc
            oStPoint.Set 0.45 * parFacetoFace, 0.8 * (parHeight - parCenterlineHeight), _
                                pipeDiam2 / 4
            oEnPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set 0.4 * parFacetoFace, 0.96 * (parHeight - parCenterlineHeight), _
                                 pipeDiam2 / 4

            Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)

            Set oCurveCol = New Collection
            oCurveCol.Add oLineA
            oCurveCol.Add oLineB
            oCurveCol.Add oLineC
            oCurveCol.Add oArc

            oStPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
          
            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
          
            Set oProjBdy = PlaceProjection(m_OutputColl, oCurve, oAxisVect, pipeDiam2 / 2, True)
        
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdy

            Set oProjBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineA = Nothing
            Set oLineB = Nothing
            Set oArc = Nothing
            Set oLineC = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing
            
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

        Case 401

            parFacetoFace = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(6)
            parCenterlineHeight = arrayOfInputs(7)
         
            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.6 * parFacetoFace
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -0.075 * parFacetoFace, 0, 0
            dCyldia = pipeDiam
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.175 * parFacetoFace, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam2
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.05 * parFacetoFace, -pipeDiam / 2, 0
            oEnPoint.Set 0.05 * parFacetoFace, (parHeight - parCenterlineHeight), 0
            dCyldia = 0.25 * parFacetoFace
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints1(0 To 23)  As Double
            'Point 1
            LineStrPoints1(0) = 0
            LineStrPoints1(1) = -pipeDiam / 2
            LineStrPoints1(2) = 0
            
            'Point 2
            LineStrPoints1(3) = -0.4 * parFacetoFace
            LineStrPoints1(4) = LineStrPoints1(1)
            LineStrPoints1(5) = LineStrPoints1(2)
            
            'Point 3
            LineStrPoints1(6) = LineStrPoints1(3)
            LineStrPoints1(7) = -0.25 * parCenterlineHeight
            LineStrPoints1(8) = LineStrPoints1(2)
            
            'Point 4
            LineStrPoints1(9) = -parBodyOutsideDiameter / 2
            LineStrPoints1(10) = LineStrPoints1(7)
            LineStrPoints1(11) = LineStrPoints1(2)
            
            'Point 5
            LineStrPoints1(12) = LineStrPoints1(9)
            LineStrPoints1(13) = -0.8 * parCenterlineHeight
            LineStrPoints1(14) = LineStrPoints1(2)
        
            'Point 6
            LineStrPoints1(15) = -0.2 * parBodyOutsideDiameter
            LineStrPoints1(16) = -parCenterlineHeight
            LineStrPoints1(17) = LineStrPoints1(2)
            
            'Point 7
            LineStrPoints1(18) = LineStrPoints1(0)
            LineStrPoints1(19) = LineStrPoints1(16)
            LineStrPoints1(20) = LineStrPoints1(2)
            
            'Point 8
            LineStrPoints1(21) = LineStrPoints1(0)
            LineStrPoints1(22) = LineStrPoints1(1)
            LineStrPoints1(23) = LineStrPoints1(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    8, LineStrPoints1)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, -1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)
            
            Set objBucketBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            m_OutputColl.AddOutput "BucketBody", objBucketBody
            Set objBucketBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

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

        Case 402

            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parHeight = arrayOfInputs(6)
            parCenterlineHeight = arrayOfInputs(7)
         
            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.8 * parDiameter
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam
  
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Point 1
            LineStrPoints(0) = 0
            LineStrPoints(1) = (parHeight - parCenterlineHeight) / 2
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = -parDiameter / 2
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = 0.3 * (parHeight - parCenterlineHeight)
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = -0.5 * parBodyOutsideDiameter
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = LineStrPoints(9)
            LineStrPoints(13) = -0.8 * parCenterlineHeight
            LineStrPoints(14) = LineStrPoints(2)
        
            'Point 6
            LineStrPoints(15) = LineStrPoints(0)
            LineStrPoints(16) = LineStrPoints(13)
            LineStrPoints(17) = LineStrPoints(2)
            
            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, -1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)
            
            Set objBucketBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBucketBody
            Set objBucketBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.125 * parDiameter, (parHeight - parCenterlineHeight) / 2, 0
            oEnPoint.Set 0.125 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            dCyldia = 0.5 * parDiameter

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
         
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            oEnPoint.Set 0.25 * parDiameter, (parHeight - parCenterlineHeight), 0
            dCyldia = 0.2 * parDiameter

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
          
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Projection
            Dim oLine11   As IngrGeom3D.Line3d
            Dim oLine22   As IngrGeom3D.Line3d
            Dim oArc11    As IngrGeom3D.Arc3d
            Dim oLine33   As IngrGeom3D.Line3d

            'Line11
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition

            oStPoint.Set -0.5 * parBodyOutsideDiameter, _
                                    0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            oEnPoint.Set -parDiameter / 2, _
                                    0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine11 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
         
            'Line22
            oStPoint.Set -parDiameter / 2, _
                                0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            oEnPoint.Set -parDiameter / 2, -0.7 * parCenterlineHeight, pipeDiam / 4
            Set oLine22 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
           
            'Arc11
            oStPoint.Set -parDiameter / 2, -0.7 * parCenterlineHeight, pipeDiam / 4
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, -0.95 * parCenterlineHeight, _
                                            pipeDiam / 4
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set -0.45 * parDiameter, -0.9 * parCenterlineHeight, pipeDiam / 4
            
            Set oArc11 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)
           
            'Line33
            oStPoint.Set -0.5 * parBodyOutsideDiameter, -0.95 * parCenterlineHeight, _
                                                pipeDiam / 4
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, _
                                0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4
            
            Set oLine33 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
         
            Set oCurveCol = New Collection
            oCurveCol.Add oLine11
            oCurveCol.Add oLine22
            oCurveCol.Add oArc11
            oCurveCol.Add oLine33

            oStPoint.Set -0.5 * parBodyOutsideDiameter, _
                                0.3 * (parHeight - parCenterlineHeight), pipeDiam / 4

            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
          
            Set oProjBdy = PlaceProjection(m_OutputColl, oCurve, oAxisVect, pipeDiam / 2, True)
      
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdy

            Set oProjBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLine11 = Nothing
            Set oLine22 = Nothing
            Set oArc11 = Nothing
            Set oLine33 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing
            
            'Place Box 1
            Dim objBox1 As Object
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parBodyOutsideDiameter, -parCenterlineHeight, pipeDiam / 4
            oEnPoint.Set 0.45 * parDiameter, -0.8 * parCenterlineHeight, -pipeDiam / 4
            
            Set objBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            m_OutputColl.AddOutput "BucketBody", objBox1
            Set objBox1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
           
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.45 * parDiameter, -0.9 * parCenterlineHeight, 0
            oEnPoint.Set 0.5 * parDiameter, -0.9 * parCenterlineHeight, 0
            dCyldia = 0.3 * parCenterlineHeight

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Projection
            Dim oLineA2   As IngrGeom3D.Line3d
            Dim oLineB2   As IngrGeom3D.Line3d
            Dim oLineC2   As IngrGeom3D.Line3d
            Dim oArc2    As IngrGeom3D.Arc3d

            'LineA2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition

            oStPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
            oEnPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, pipeDiam2 / 4

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineA2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineB2
            oStPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, pipeDiam2 / 4
            oEnPoint.Set 0.45 * parFacetoFace, pipeDiam2 / 2, pipeDiam2 / 4
            Set oLineB2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineC2
            oStPoint.Set 0.45 * parFacetoFace, pipeDiam2 / 2, pipeDiam2 / 4
            oEnPoint.Set 0.45 * parFacetoFace, 0.8 * (parHeight - parCenterlineHeight), _
                            pipeDiam2 / 4
            Set oLineC2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Arc2
            oStPoint.Set 0.45 * parFacetoFace, 0.8 * (parHeight - parCenterlineHeight), _
                            pipeDiam2 / 4
            oEnPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set 0.4 * parFacetoFace, 0.93 * (parHeight - parCenterlineHeight), _
                                 pipeDiam2 / 4

            Set oArc2 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)

            Set oCurveCol = New Collection
            oCurveCol.Add oLineA2
            oCurveCol.Add oLineB2
            oCurveCol.Add oLineC2
            oCurveCol.Add oArc2

            oStPoint.Set 0.3 * parDiameter, (parHeight - parCenterlineHeight), pipeDiam2 / 4
          
            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
            
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
            
            Set oProjBdy = PlaceProjection(m_OutputColl, oCurve, oAxisVect, pipeDiam2 / 2, True)
          
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdy

            Set oProjBdy = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineA2 = Nothing
            Set oLineB2 = Nothing
            Set oArc2 = Nothing
            Set oLineC2 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing

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
        
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
