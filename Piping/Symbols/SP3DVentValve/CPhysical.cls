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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          KKC
'   Creation Date:  Thursday, April 5 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on Tech-Taylor Vaccum Breaker
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----   ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
'   25-Jul-07       RRK     TR-122066:Corrected the port 2 location
'   16.11.2007      MA      CR-128808 Implemented Part data basis for values 510 and 511.
'   04.02.2008      MA      CR-135704 Implemented Part data basis for values 508 and 509.
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 23) As Double
    
    Dim parFacetoFace As Double
    Dim parMajorBodyDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFacetoEnd As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim ObjBody As Object
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oPoints(1 To 5) As New AutoMath.DPosition
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Dim oAxisVec As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objMiddlePortion As Object
    Dim dCylDia As Double
    Dim objCyl As Object

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parMajorBodyDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFacetoEnd = arrayOfInputs(5)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -508): Vent valve, one port with cylindrical outlet cover
'   (Part data Basis value -509): Vent valve, one port with conical outlet cover
'   (Part data Basis value -510): Vent valve, one port
'   (Part data Basis value -511): Vent valve, one port with outlet cover
'   (Part data Basis value -512): Vent valve, two ports or (Default)
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis

        Case Is <= 1, 512
            
            parFacetoFace = arrayOfInputs(2)
            
            'Insert your code for output 2(Body of Vent Valve)
            'Point 1
            LineStringPoints(0) = -0.4 * parFacetoFace
            LineStringPoints(1) = 0
            LineStringPoints(2) = 0
            
            'Point 2
            LineStringPoints(3) = -0.4 * parFacetoFace
            LineStringPoints(4) = pipeDiam / 2
            LineStringPoints(5) = 0
            
            'Point 3
            LineStringPoints(6) = -0.1 * parFacetoFace
            LineStringPoints(7) = 0.35 * parMajorBodyDiameter
            LineStringPoints(8) = 0
            
            'Point 4
            LineStringPoints(9) = -0.1 * parFacetoFace
            LineStringPoints(10) = 0.5 * parMajorBodyDiameter
            LineStringPoints(11) = 0
            
            'Point 5
            LineStringPoints(12) = 0.1 * parFacetoFace
            LineStringPoints(13) = 0.5 * parMajorBodyDiameter
            LineStringPoints(14) = 0
            
            'Point 6
            LineStringPoints(15) = 0.1 * parFacetoFace
            LineStringPoints(16) = 0.35 * parMajorBodyDiameter
            LineStringPoints(17) = 0
            
            'Point 7
            LineStringPoints(18) = 0.6 * parFacetoFace
            LineStringPoints(19) = pipeDiam / 2
            LineStringPoints(20) = 0
            
            'Point 8
            LineStringPoints(21) = 0.6 * parFacetoFace
            LineStringPoints(22) = 0
            LineStringPoints(23) = 0
            
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 8, LineStringPoints)
            axisVect.Set 1, 0, 0
            centerPoint.Set 0, 0, 0
            Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
            Set ObjBody = Nothing
              
            'Place Nozzle 1
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
        
            oPlacePoint.Set -0.4 * parFacetoFace - sptOffset + depth, 0, 0
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
         
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            
            'Place Nozzle 2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
            oPlacePoint.Set 0.6 * parFacetoFace + sptOffset - depth, 0, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
         
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
            Set centerPoint = Nothing
            Set axisVect = Nothing
            Set oLineString = Nothing

        Case 510
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 10% of the Face to End and the
            'diameter is 70% of the Major Body Diameter
            oStPoint.Set 0.9 * parFacetoEnd, 0, 0
            oEnPoint.Set parFacetoEnd, 0, 0
            dCylDia = 0.7 * parMajorBodyDiameter
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDia, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Place Cylinder 2
            'Assumption : The height of the cylinder is 10% of the Face to End
            oStPoint.Set 0.8 * parFacetoEnd, 0, 0
            oEnPoint.Set 0.9 * parFacetoEnd, 0, 0
            dCylDia = parMajorBodyDiameter
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDia, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Insert your code for middle portion of the body of the vent valve
            'Assumption : 1) The height of the cylindrical portion is 45% of the Face to End
            '             2) The diameter of the cylindrical portion is 80% of the Major Body Diameter

            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.8 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(2).Set 0.45 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(3).Set 0.35 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(4).Set 0.25 * parFacetoEnd, 0.3 * parMajorBodyDiameter, 0
            oPoints(5).Set 0.2 * parFacetoEnd, pipeDiam / 2, 0
            
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
          
            Set objMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
      
            'Set the output
            m_OutputColl.AddOutput "Body", objMiddlePortion
            Set objMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
        
            'Place Cylinder 3
            'Assumption : The height of this cylinder is 20% of the Face to End
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.2 * parFacetoEnd, 0, 0

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        
            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
        
        Case 511
            
            parFacetoEnd = arrayOfInputs(5)
          
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 10% of the Face to End
            oStPoint.Set 0.6 * parFacetoEnd, 0, 0
            oEnPoint.Set 0.7 * parFacetoEnd, 0, 0
            dCylDia = parMajorBodyDiameter
          
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDia, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Insert your code for middle portion of the body of the vent valve
            'Assumption : 1) The height of the cylindrical portion is 30% of the Face to End
            '             2) The diameter of the cylindrical portion is 80% of the Major Body Diameter
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.6 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(2).Set 0.35 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(3).Set 0.3 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(4).Set 0.25 * parFacetoEnd, 0.35 * parMajorBodyDiameter, 0
            oPoints(5).Set 0.2 * parFacetoEnd, pipeDiam / 2, 0
            
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
        
            Set objMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMiddlePortion
            Set objMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
        
            'Place Cylinder 2
            'Assumption : The height of this cylinder is 20% of the Face to End
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.2 * parFacetoEnd, 0, 0

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl", objCyl
            Set objCyl = Nothing
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.7 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(2).Set 0.85 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oPoints(3).Set 0.9 * parFacetoEnd, 0.35 * parMajorBodyDiameter, 0
            oPoints(4).Set 0.98 * parFacetoEnd, 0.25 * parMajorBodyDiameter, 0
            oPoints(5).Set parFacetoEnd, 0, 0
            
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
      
            Set objMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
  
            'Set the output
            m_OutputColl.AddOutput "Body", objMiddlePortion
            Set objMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
            
            'Place Cylinder 3
            oStPoint.Set 0.7 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oEnPoint.Set 0.95 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parMajorBodyDiameter, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl3", objCyl
            Set objCyl = Nothing
            
            'Place Cylinder 4
            oStPoint.Set 0.7 * parFacetoEnd, -0.4 * parMajorBodyDiameter, 0
            oEnPoint.Set 0.95 * parFacetoEnd, -0.4 * parMajorBodyDiameter, 0

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parMajorBodyDiameter, True)

            'Set the output
            m_OutputColl.AddOutput "Cyl4", objCyl
            Set objCyl = Nothing
            
        Case 508
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 20% of the Face to End
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.2 * parFacetoEnd, 0, 0
          
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Place Cylinder 2
            'Assumption : The height of the cylinder is 60% of the Face to End
            oStPoint.Set 0.2 * parFacetoEnd, 0, 0
            oEnPoint.Set 0.8 * parFacetoEnd, 0, 0
          
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parMajorBodyDiameter, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
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
            dMMRatio = (0.2 * parFacetoEnd) / (0.5 * parMajorBodyDiameter)
            oEllCen.Set 0.8 * parFacetoEnd, 0, 0
            oEllNormal.Set 0, 0, -1
            oEllMajor.Set 0, parMajorBodyDiameter / 2, 0

            Set oQuarterEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                Nothing, oEllCen.x, oEllCen.y, oEllCen.z, _
                                                oEllNormal.x, oEllNormal.y, oEllNormal.z, _
                                                oEllMajor.x, oEllMajor.y, oEllMajor.z, _
                                                dMMRatio, 0, PI / 2)
 
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
   
            Set oSemiEllipsoid = PlaceRevolution(m_OutputColl, oQuarterEllipse, oAxisVec, oEllCen, 2 * PI, True)
 
            'Set the Output
            m_OutputColl.AddOutput "SemiEllipsoid", oSemiEllipsoid
            Set oSemiEllipsoid = Nothing
            Set oQuarterEllipse = Nothing
            Set oEllCen = Nothing
            Set oEllNormal = Nothing
            Set oEllMajor = Nothing
            Set oAxisVec = Nothing
            
        Case 509
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 20% of the Face to End
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.2 * parFacetoEnd, 0, 0
        
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Insert your code for output 2(Body of Vent Valve)
            Dim dLineStrPoints(0 To 20)  As Double
            'Point 1
            dLineStrPoints(0) = 0.2 * parFacetoEnd
            dLineStrPoints(1) = 0
            dLineStrPoints(2) = 0
            
            'Point 2
            dLineStrPoints(3) = 0.2 * parFacetoEnd
            dLineStrPoints(4) = parMajorBodyDiameter / 2
            dLineStrPoints(5) = 0
            
            'Point 3
            dLineStrPoints(6) = 0.9 * parFacetoEnd
            dLineStrPoints(7) = 0.25 * parMajorBodyDiameter
            dLineStrPoints(8) = 0
            
            'Point 4
            dLineStrPoints(9) = 0.9 * parFacetoEnd
            dLineStrPoints(10) = 0.125 * parMajorBodyDiameter
            dLineStrPoints(11) = 0
            
            'Point 5
            dLineStrPoints(12) = parFacetoEnd
            dLineStrPoints(13) = 0.125 * parMajorBodyDiameter
            dLineStrPoints(14) = 0
            
            'Point 6
            dLineStrPoints(15) = parFacetoEnd
            dLineStrPoints(16) = 0
            dLineStrPoints(17) = 0
            
            'Point 7
            dLineStrPoints(18) = 0.2 * parFacetoEnd
            dLineStrPoints(19) = 0
            dLineStrPoints(20) = 0
          
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenPoint, 2 * PI, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
            Set ObjBody = Nothing
            Set oLineString = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing

        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    If (lPartDataBasis = 508) Or (lPartDataBasis = 509) Or (lPartDataBasis = 510) Or _
                                            (lPartDataBasis = 511) Then
        'Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
    
        oPlacePoint.Set depth - sptOffset, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    End If
        
    Exit Sub
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set geomFactory = Nothing
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
