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
'   Author:         MA
'   Creation Date:  Monday, Nov 05 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   90 Degree Street Elbow.
'   This symbol is developed for CR-129350
'   Source: Hoke Precision Instrument Pipe Fittings Catalog; 79026 12/05;
'   Hoke Incorporated, www.hoke.com
'
'   Change History:
'   dd.mmm.yyyy     who               change description
'   -----------     -----             ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI           As Double
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Const NEGLIGIBLE_VALUE = 0.0001

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
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
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Integer
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim dNozzleLength As Double
    Dim ObjInsTangentPortion1 As Object
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                            flangeDiam, cptOffset, depth

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                            flangeDiam2, cptOffset2, depth2
 
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -490): 90° street elbow, hexagonal
'   (Part data Basis value -491): 90° street elbow, flat
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis

        Case Is <= 1, 490
            
            'Insert your code for output 1(Body)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            
            'Create complex string for Hexagonal body
            Dim dAngle As Double
            dAngle = PI / 3
            oStPoint.Set -parFace1toCenter, 0, 0
            Dim iCount As Integer
            Dim dLineStrPoints(0 To 20) As Double
            For iCount = 1 To 7
                dLineStrPoints(3 * iCount - 3) = oStPoint.x
                dLineStrPoints(3 * iCount - 2) = oStPoint.y + (pipeDiam / 2 + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle)
                dLineStrPoints(3 * iCount - 1) = oStPoint.z + (pipeDiam / 2 + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle)
            Next iCount
           
            Dim oHexagon As IngrGeom3D.LineString3d
            Set oHexagon = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)

            Dim oHorLine   As IngrGeom3D.Line3d
            Dim oArc  As IngrGeom3D.Arc3d
            Dim oVerLine   As IngrGeom3D.Line3d
        
            'Horizontal line
            oStPoint.Set -parFace1toCenter, 0, 0
            oEnPoint.Set -0.3 * parFace1toCenter, 0, 0
        
            Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                   oStPoint.x, oStPoint.y, oStPoint.z, _
                                   oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
            'Arc
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFace1toCenter, 0, 0
            oEnPoint.Set 0, -0.3 * parFace1toCenter, 0
            oCenPoint.Set -0.3 * parFace1toCenter, -0.3 * parFace1toCenter, 0
          
            Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
        
            'Vertical line
            oStPoint.Set 0, -0.3 * parFace1toCenter, 0
            oEnPoint.Set 0, -(parFace2toCenter - 0.4 * parFace2toCenter), 0
        
            Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                   oStPoint.x, oStPoint.y, oStPoint.z, _
                                   oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
            Dim oSweepCurve As IngrGeom3D.ComplexString3d
            Dim oSweepCurveCol As Collection
            Set oSweepCurveCol = New Collection
            oSweepCurveCol.Add oHorLine
            oSweepCurveCol.Add oArc
            oSweepCurveCol.Add oVerLine
        
            Dim oStartBC   As AutoMath.DPosition
            Set oStartBC = New AutoMath.DPosition
            oStartBC.Set -parFace1toCenter, 0, 0
        
            Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
        
            Dim Surfset   As IngrGeom3D.IJElements
            Dim stnorm() As Double
            Dim ednorm() As Double
        
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                              oSweepCurve, oHexagon, CircularCorner, 0, stnorm, ednorm, True)
        
            Dim ObjSurface As Object
            For Each ObjSurface In Surfset
               m_OutputColl.AddOutput "InsulatedBody_", ObjSurface
            Next ObjSurface
        
            Set oStartBC = Nothing
            Set oCenPoint = Nothing
            Set oHorLine = Nothing
            Set oArc = Nothing
            Set oVerLine = Nothing
            Set oHexagon = Nothing
            Set oSweepCurve = Nothing
            Set oSweepCurveCol = Nothing
            Set Surfset = Nothing
            Set oGeomFactory = Nothing
           
        Case 491
            
            Dim oLine11 As IngrGeom3D.Line3d
            Dim oLine22 As IngrGeom3D.Line3d
            Dim oLine33 As IngrGeom3D.Line3d
            Dim oLine44 As IngrGeom3D.Line3d
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition

            'Create a complex string for box cross-section
            'Line11
            oStPoint.Set -parFace1toCenter + flangeThick, _
                                0.5 * pipeDiam + parInsulationThickness, _
                                -0.5 * pipeDiam - parInsulationThickness
            oEnPoint.Set -parFace1toCenter + flangeThick, _
                                0.5 * pipeDiam + parInsulationThickness, _
                                0.5 * pipeDiam + parInsulationThickness

            Set oLine11 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Line22
            oStPoint.Set -parFace1toCenter + flangeThick, _
                                0.5 * pipeDiam + parInsulationThickness, _
                                0.5 * pipeDiam + parInsulationThickness
            oEnPoint.Set -parFace1toCenter + flangeThick, _
                                -0.5 * pipeDiam - parInsulationThickness, _
                                0.5 * pipeDiam + parInsulationThickness

            Set oLine22 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Line33
            oStPoint.Set -parFace1toCenter + flangeThick, _
                                -0.5 * pipeDiam - parInsulationThickness, _
                                0.5 * pipeDiam + parInsulationThickness
            oEnPoint.Set -parFace1toCenter + flangeThick, _
                                -0.5 * pipeDiam - parInsulationThickness, _
                                -0.5 * pipeDiam - parInsulationThickness

            Set oLine33 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Line44
            oStPoint.Set -parFace1toCenter + flangeThick, _
                                -0.5 * pipeDiam - parInsulationThickness, _
                                -0.5 * pipeDiam - parInsulationThickness
            oEnPoint.Set -parFace1toCenter + flangeThick, _
                                0.5 * pipeDiam + parInsulationThickness, _
                                -0.5 * pipeDiam - parInsulationThickness

            Set oLine44 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)

            Dim oCurvesColl2 As Collection
            Set oCurvesColl2 = New Collection
            oCurvesColl2.Add oLine11
            oCurvesColl2.Add oLine22
            oCurvesColl2.Add oLine33
            oCurvesColl2.Add oLine44

            Dim oBox As IngrGeom3D.ComplexString3d
            oStPoint.Set -parFace1toCenter + flangeThick, _
                                    0.5 * pipeDiam + parInsulationThickness, _
                                    -0.5 * pipeDiam - parInsulationThickness
            Set oBox = PlaceTrCString(oStPoint, oCurvesColl2)

            Dim oHorLine2   As IngrGeom3D.Line3d
            Dim oArc2  As IngrGeom3D.Arc3d
            Dim oVerLine2   As IngrGeom3D.Line3d

            'Horizontal line
            oStPoint.Set -parFace1toCenter + flangeThick, 0, 0
            oEnPoint.Set -0.3 * parFace1toCenter, 0, 0

            Set oHorLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                   oStPoint.x, oStPoint.y, oStPoint.z, _
                                   oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Arc
            Dim oCenPoint2 As AutoMath.DPosition
            Set oCenPoint2 = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFace1toCenter, 0, 0
            oEnPoint.Set 0, -0.3 * parFace1toCenter, 0
            oCenPoint2.Set -0.3 * parFace1toCenter, -0.3 * parFace1toCenter, 0

            Set oArc2 = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint2)

            'Vertical line
            oStPoint.Set 0, -0.3 * parFace1toCenter, 0
            oEnPoint.Set 0, -0.6 * parFace2toCenter, 0

            Set oVerLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                   oStPoint.x, oStPoint.y, oStPoint.z, _
                                   oEnPoint.x, oEnPoint.y, oEnPoint.z)

            Dim oSweepCurve2 As IngrGeom3D.ComplexString3d
            Dim oSweepCurveCol2 As Collection
            Set oSweepCurveCol2 = New Collection
            oSweepCurveCol2.Add oHorLine2
            oSweepCurveCol2.Add oArc2
            oSweepCurveCol2.Add oVerLine2

            Dim oStartBC2   As AutoMath.DPosition
            Set oStartBC2 = New AutoMath.DPosition
            oStartBC2.Set -parFace1toCenter + flangeThick, 0, 0

            Set oSweepCurve2 = PlaceTrCString(oStartBC2, oSweepCurveCol2)

            Dim Surfset2   As IngrGeom3D.IJElements
            Dim stnorm2() As Double
            Dim ednorm2() As Double

            Set Surfset2 = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                              oSweepCurve2, oBox, CircularCorner, 0, stnorm2, ednorm2, True)

            Dim ObjSurface2 As Object
            For Each ObjSurface2 In Surfset2
               m_OutputColl.AddOutput "InsulatedBody_", ObjSurface2
            Next ObjSurface2

            Set oLine11 = Nothing
            Set oLine22 = Nothing
            Set oLine33 = Nothing
            Set oLine44 = Nothing
            Set oBox = Nothing
            Set oCurvesColl2 = Nothing
            Set oStartBC2 = Nothing
            Set oCenPoint2 = Nothing
            Set oHorLine2 = Nothing
            Set oArc2 = Nothing
            Set oVerLine2 = Nothing
            Set oSweepCurve2 = Nothing
            Set oSweepCurveCol2 = Nothing
            Set Surfset2 = Nothing
            Set oGeomFactory = Nothing
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    'Construct Insulation for the tangential portion 1
    dNozzleLength = 0.4 * parFace2toCenter
    If CmpDblGreaterthan(dNozzleLength, LINEAR_TOLERANCE) Then
        oStPoint.Set 0, -parFace2toCenter, 0
        oEnPoint.Set 0, -parFace2toCenter + dNozzleLength, 0
 
        Set ObjInsTangentPortion1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam2 + parInsulationThickness * 2, True)
 
        'Set the output
        m_OutputColl.AddOutput "InsulatedBody_", ObjInsTangentPortion1
        Set ObjInsTangentPortion1 = Nothing
    End If
    
    'Place Insulation for Port 1
    If CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set -parFace1toCenter, 0, 0
        If CmpDblGreaterthanOrEqualTo(-parFace1toCenter + flangeThick + parInsulationThickness, 0) Then
            oEnPoint.Set 0, 0, 0
        Else
            oEnPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
        End If

        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        flangeDiam + 2 * parInsulationThickness, True)

        'Set the output
        m_OutputColl.AddOutput "InsulatedBody_", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
     End If

    'Place Insulation for Port 2
    If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
        If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
        oStPoint.Set 0, -parFace2toCenter, 0
        If CmpDblGreaterthanOrEqualTo(-parFace2toCenter + flangeThick2 + parInsulationThickness, 0) Then
            oEnPoint.Set 0, 0, 0
        Else
            oEnPoint.Set 0, -parFace2toCenter + flangeThick2 + parInsulationThickness, 0
        End If
       
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                flangeDiam2 + 2 * parInsulationThickness, True)

        'Set the output
        m_OutputColl.AddOutput "InsulatedBody_", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
     End If
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    End Sub
