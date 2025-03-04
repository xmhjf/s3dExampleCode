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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          KKC
'   Creation Date:  Thursday, April 5 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on Tech-Taylor Vaccum Breaker
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
'   16.11.2007      MA      CR-128808 Implemented Part data basis for values 510 and 511.
'   04.02.2008      MA      CR-135704 Implemented Part data basis for values 508 and 509.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error message
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
    Dim ObjInsBody As Object
    
    Dim parFacetoFace As Double
    Dim parMajorBodyDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFacetoEnd As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 17) As Double
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oPoints(1 To 5) As New AutoMath.DPosition
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Dim dInsCylDia As Double
    Dim objInsCyl As Object
    Dim objInsMiddlePortion As Object
    Dim oAxisVec As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    
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
            LineStringPoints(4) = pipeDiam / 2 + parInsulationThickness
            LineStringPoints(5) = 0
        
            'Point 3
            LineStringPoints(6) = -0.1 * parFacetoFace - parInsulationThickness / 2
            LineStringPoints(7) = 0.5 * parMajorBodyDiameter + parInsulationThickness
            LineStringPoints(8) = 0
            
            'Point 4
            LineStringPoints(9) = 0.1 * parFacetoFace + parInsulationThickness / 2
            LineStringPoints(10) = 0.5 * parMajorBodyDiameter + parInsulationThickness
            LineStringPoints(11) = 0
        
            'Point 5
            LineStringPoints(12) = 0.6 * parFacetoFace
            LineStringPoints(13) = pipeDiam / 2 + parInsulationThickness
            LineStringPoints(14) = 0
            
            'Point 6
            LineStringPoints(15) = 0.6 * parFacetoFace
            LineStringPoints(16) = 0
            LineStringPoints(17) = 0
            
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
            axisVect.Set 1, 0, 0
            centerPoint.Set 0, 0, 0
            Set ObjInsBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
            Set ObjInsBody = Nothing
            
            'Insert the code for Insulation of the port
            Dim dInsulationDiameter As Double
            Dim ObjInsulatedPort1 As Object
            Dim ObjInsulatedPort2 As Object
            
            If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
            oStPoint.Set -0.4 * parFacetoFace, 0, 0
            oEnPoint.Set -0.4 * parFacetoFace + flangeThick + parInsulationThickness, 0, 0
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            End If
            
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
            
            'Insert the code for Insulation of the port2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
            
            oStPoint.Set 0.6 * parFacetoFace, 0, 0
            oEnPoint.Set 0.6 * parFacetoFace - flangeThick - parInsulationThickness, 0, 0
                
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
            
            Set centerPoint = Nothing
            Set axisVect = Nothing
            Set oLineString = Nothing
            
        Case 510
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 10% of the Face to End and the
            'diameter is 70% of the Major Body Diameter
            oStPoint.Set 0.9 * parFacetoEnd, 0, 0
            oEnPoint.Set parFacetoEnd + parInsulationThickness, 0, 0
            dInsCylDia = 0.7 * parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
            'Place Cylinder 2
            'Assumption : The height of the cylinder is 10% of the Face to End
            oStPoint.Set 0.8 * parFacetoEnd, 0, 0
            oEnPoint.Set 0.9 * parFacetoEnd, 0, 0
            dInsCylDia = parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
            'Insert your code for middle portion of the body of the vent valve
            'Assumption : 1) The height of the cylindrical portion is 45% of the Face to End
            '             2) The diameter of the cylindrical portion is 80% of the Major Body Diameter
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.8 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(2).Set 0.45 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(3).Set 0.35 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(4).Set 0.25 * parFacetoEnd, 0.3 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(5).Set 0.2 * parFacetoEnd, pipeDiam / 2 + parInsulationThickness, 0
            
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
          
            Set objInsMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsMiddlePortion
            Set objInsMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
        
            'Place Cylinder 3
            'Assumption : The height of this cylinder is 20% of the Face to End
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.2 * parFacetoEnd, 0, 0
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsCylDia = flangeDiam + parInsulationThickness * 2
            Else
                dInsCylDia = pipeDiam + parInsulationThickness * 2
            End If
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
        
            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
          
        Case 511

            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 1
            'Assumption : The height of the cylinder is 10% of the Face to End
            oStPoint.Set 0.6 * parFacetoEnd, 0, 0
            oEnPoint.Set 0.7 * parFacetoEnd, 0, 0
            dInsCylDia = parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
            'Insert your code for middle portion of the body of the vent valve
            'Assumption : 1) The height of the cylindrical portion is 30% of the Face to End
            '             2) The diameter of the cylindrical portion is 80% of the Major Body Diameter
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.6 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(2).Set 0.35 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(3).Set 0.3 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(4).Set 0.25 * parFacetoEnd, 0.35 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(5).Set 0.2 * parFacetoEnd, pipeDiam / 2 + parInsulationThickness, 0
            
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
        
            Set objInsMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsMiddlePortion
            Set objInsMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set 0.7 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(2).Set 0.85 * parFacetoEnd, 0.4 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(3).Set 0.9 * parFacetoEnd, 0.35 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(4).Set 0.98 * parFacetoEnd + parInsulationThickness, 0.25 * parMajorBodyDiameter + parInsulationThickness, 0
            oPoints(5).Set parFacetoEnd + parInsulationThickness, 0, 0
           
            Set objBsplineCurve = PlaceTrBspline(4, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
      
            Set objInsMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
  
            'Set the output
            m_OutputColl.AddOutput "InsBody", objInsMiddlePortion
            Set objInsMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
            
            'Place Cylinder 3
            oStPoint.Set 0.7 * parFacetoEnd, 0.4 * parMajorBodyDiameter, 0
            oEnPoint.Set 0.95 * parFacetoEnd + parInsulationThickness, 0.4 * parMajorBodyDiameter, 0
            dInsCylDia = 0.1 * parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl3", objInsCyl
            Set objInsCyl = Nothing
            
            'Place Cylinder 4
            oStPoint.Set 0.7 * parFacetoEnd, -0.4 * parMajorBodyDiameter, 0
            oEnPoint.Set 0.95 * parFacetoEnd + parInsulationThickness, -0.4 * parMajorBodyDiameter, 0
            dInsCylDia = 0.1 * parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl4", objInsCyl
            Set objInsCyl = Nothing
            
        Case 508
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Place Cylinder 2
            'Assumption : The height of the cylinder is 80% of the Face to End
            oStPoint.Set 0.2 * parFacetoEnd - parInsulationThickness / 2, 0, 0
            oEnPoint.Set parFacetoEnd + parInsulationThickness, 0, 0
            dInsCylDia = parMajorBodyDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
       
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
        Case 509
            
            parFacetoEnd = arrayOfInputs(5)
            
            'Insert your code for output 2(Body of Vent Valve)
            Dim dLineStrPoints(0 To 20)  As Double
            'Point 1
            dLineStrPoints(0) = 0.2 * parFacetoEnd - parInsulationThickness / 2
            dLineStrPoints(1) = 0
            dLineStrPoints(2) = 0
            
            'Point 2
            dLineStrPoints(3) = 0.2 * parFacetoEnd - parInsulationThickness / 2
            dLineStrPoints(4) = parMajorBodyDiameter / 2 + parInsulationThickness
            dLineStrPoints(5) = 0
            
            'Point 3
            dLineStrPoints(6) = 0.9 * parFacetoEnd + parInsulationThickness / 2
            dLineStrPoints(7) = 0.25 * parMajorBodyDiameter + parInsulationThickness
            dLineStrPoints(8) = 0
            
            'Point 4
            dLineStrPoints(9) = 0.9 * parFacetoEnd + parInsulationThickness / 2
            dLineStrPoints(10) = 0.125 * parMajorBodyDiameter + parInsulationThickness
            dLineStrPoints(11) = 0
            
            'Point 5
            dLineStrPoints(12) = parFacetoEnd + parInsulationThickness
            dLineStrPoints(13) = 0.125 * parMajorBodyDiameter + parInsulationThickness
            dLineStrPoints(14) = 0
            
            'Point 6
            dLineStrPoints(15) = parFacetoEnd + parInsulationThickness
            dLineStrPoints(16) = 0
            dLineStrPoints(17) = 0
            
            'Point 7
            dLineStrPoints(18) = 0.2 * parFacetoEnd - parInsulationThickness / 2
            dLineStrPoints(19) = 0
            dLineStrPoints(20) = 0

            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            Set ObjInsBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
            Set ObjInsBody = Nothing
            Set oLineString = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    If (lPartDataBasis = 508) Or (lPartDataBasis = 509) Or (lPartDataBasis = 511) Then _
    
       'Place Cylinder 1 (Nozzle Insulation)
       'Assumption : The height of this cylinder is 20% of the Face to End
       oStPoint.Set 0, 0, 0
       oEnPoint.Set 0.2 * parFacetoEnd, 0, 0
    
       If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
           dInsCylDia = flangeDiam + parInsulationThickness * 2
       Else
           dInsCylDia = pipeDiam + parInsulationThickness * 2
       End If
    
       Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
    
       'Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
       Set objInsCyl = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

