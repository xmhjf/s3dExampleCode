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
'   Creation Date:  Friday, Aug 31 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-121408
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'   12.10.2007          MA         CR-121408 Removed (iOutput = iOutput + 1) for variable outputs.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim objInsulatedBody As Object

    Dim parFacetoFace As Double
    Dim parDiameter   As Double
    Dim parBodyOutsideDiameter As Double
    Dim parOffset As Double
    Dim parHeight As Double
    Dim parCenterlineHeight As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double

    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objInsCyl As Object
    Dim dCyldia As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oArcPoint As AutoMath.DPosition
    Dim oProjBdyIns As Object
    Dim oCurve As IngrGeom3D.ComplexString3d
    Dim oCurveCol As Collection
    Dim objInsulatedNozzle1 As Object
    Dim objInsulatedNozzle2 As Object
    Dim parInsulationDiameter As Double

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
    
    Select Case lPartDataBasis

        Case Is <= 1, 397
        
            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
        
            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.75 * parDiameter
            
            'Point 1
            LineStrPoints(0) = -parFacetoFace / 2
            LineStrPoints(1) = parBodyOutsideDiameter / 2 + parInsulationThickness
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = 0.25 * parFacetoFace - parInsulationThickness
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = parDiameter / 2 + parInsulationThickness
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = 0.35 * parFacetoFace + parInsulationThickness
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
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.35 * parFacetoFace, -0.13 * parDiameter, 0
            oEnPoint.Set 0.5 * parFacetoFace, -0.13 * parDiameter, 0
            dCyldia = 0.5 * parDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
          
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
       
        Case 398
            
            parFacetoFace = arrayOfInputs(2)
            parDiameter = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            parOffset = arrayOfInputs(5)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then parBodyOutsideDiameter = 0.75 * parDiameter

            'Point 1
            LineStrPoints(0) = -parFacetoFace / 2
            LineStrPoints(1) = parBodyOutsideDiameter / 2 + parInsulationThickness
            LineStrPoints(2) = 0

            'Point 2
            LineStrPoints(3) = 0.25 * parFacetoFace - parInsulationThickness
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)

            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = parDiameter / 2 + parInsulationThickness
            LineStrPoints(8) = LineStrPoints(2)

            'Point 4
            LineStrPoints(9) = 0.35 * parFacetoFace + parInsulationThickness
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

            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.35 * parFacetoFace, -parOffset, 0
            oEnPoint.Set 0.5 * parFacetoFace, -parOffset, 0
            dCyldia = 0.3 * parDiameter + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
            
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
            dCyldia = pipeDiam + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
    
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.2 * parFacetoFace, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam2 + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.2 * parFacetoFace, -0.3 * (parHeight - parCenterlineHeight) / 2, 0
            oEnPoint.Set 0.2 * parFacetoFace, -0.3 * (parHeight - parCenterlineHeight) / 2, 0
            
            dCyldia = 1.75 * (parHeight - parCenterlineHeight) + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -pipeDiam / 2, 0
            oEnPoint.Set 0, -parCenterlineHeight - parInsulationThickness, 0
            dCyldia = parBodyOutsideDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
            
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
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If
            
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If
            
            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dCyldia = dInsulationDia1
            Else
                dCyldia = dInsulationDia2
            End If
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints2(0 To 26)  As Double
            'Point 1
            LineStrPoints2(0) = 0
            LineStrPoints2(1) = (parHeight - parCenterlineHeight) / 2
            LineStrPoints2(2) = 0
            
            'Point 2
            LineStrPoints2(3) = -parDiameter / 2 - parInsulationThickness
            LineStrPoints2(4) = LineStrPoints2(1)
            LineStrPoints2(5) = LineStrPoints2(2)
            
            'Point 3
            LineStrPoints2(6) = LineStrPoints2(3)
            LineStrPoints2(7) = 0.3 * (parHeight - parCenterlineHeight)
            LineStrPoints2(8) = LineStrPoints2(2)
            
            'Point 4
            LineStrPoints2(9) = -0.5 * parBodyOutsideDiameter - parInsulationThickness
            LineStrPoints2(10) = LineStrPoints2(7)
            LineStrPoints2(11) = LineStrPoints2(2)
            
            'Point 5
            LineStrPoints2(12) = LineStrPoints2(9)
            LineStrPoints2(13) = -0.8 * parCenterlineHeight
            LineStrPoints2(14) = LineStrPoints2(2)
        
            'Point 6
            LineStrPoints2(15) = -0.2 * parBodyOutsideDiameter - parInsulationThickness
            LineStrPoints2(16) = LineStrPoints2(13)
            LineStrPoints2(17) = LineStrPoints2(2)
            
            'Point 7
            LineStrPoints2(18) = LineStrPoints2(15)
            LineStrPoints2(19) = -parCenterlineHeight - parInsulationThickness
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
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
        
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.125 * parDiameter, (parHeight - parCenterlineHeight) / 2, 0
            oEnPoint.Set 0.125 * parDiameter, _
                        0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, 0
            dCyldia = 0.5 * parDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            oEnPoint.Set 0.25 * parDiameter, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, 0
            dCyldia = 0.2 * parDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
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
                                    0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -parDiameter / 2 - parInsulationThickness, _
                                    0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Line2
            oStPoint.Set -parDiameter / 2 - parInsulationThickness, _
                                    0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -parDiameter / 2 - parInsulationThickness, _
                            -0.7 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Arc1
            oStPoint.Set -parDiameter / 2 - parInsulationThickness, _
                                -0.7 * parCenterlineHeight - parInsulationThickness, _
                                            pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.2 * parBodyOutsideDiameter - parInsulationThickness, _
                                -0.9 * parCenterlineHeight - parInsulationThickness, _
                                        pipeDiam / 4 + parInsulationThickness
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set -0.3 * parDiameter - parInsulationThickness, _
                                -0.9 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            
            Set oArc1 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)
            
            'Line3
            oStPoint.Set -0.2 * parBodyOutsideDiameter - parInsulationThickness, _
                                -0.9 * parCenterlineHeight - parInsulationThickness, _
                                            pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.2 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                            pipeDiam / 4 + parInsulationThickness
            Set oLine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
          
            'Line4
            oStPoint.Set -0.2 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                            pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                            pipeDiam / 4 + parInsulationThickness
            Set oLine4 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
           
            'Line5
            oStPoint.Set -0.5 * parBodyOutsideDiameter, -0.8 * parCenterlineHeight, _
                                            pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, _
                                    0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
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
                                0.3 * (parHeight - parCenterlineHeight), _
                                    pipeDiam / 4 + parInsulationThickness

            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
          
            Set oProjBdyIns = PlaceProjection(m_OutputColl, oCurve, oAxisVect, _
                                pipeDiam / 2 + 2 * parInsulationThickness, True)
          
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdyIns

            Set oProjBdyIns = Nothing
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
            oStPoint.Set 0.3 * parDiameter, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, _
                                    pipeDiam2 / 4 + parInsulationThickness

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineA = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineB
            oStPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, _
                                pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.45 * parFacetoFace + parInsulationThickness, pipeDiam2 / 2, _
                                        pipeDiam2 / 4 + parInsulationThickness
            Set oLineB = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineC
            oStPoint.Set 0.45 * parFacetoFace + parInsulationThickness, pipeDiam2 / 2, _
                                        pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                        0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                                pipeDiam2 / 4 + parInsulationThickness
            Set oLineC = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Arc
            oStPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                            0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                        pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.3 * parDiameter, _
                                (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                    pipeDiam2 / 4 + parInsulationThickness
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set 0.4 * parFacetoFace + parInsulationThickness, _
                            0.96 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                    pipeDiam2 / 4 + parInsulationThickness

            Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)

            Set oCurveCol = New Collection
            oCurveCol.Add oLineA
            oCurveCol.Add oLineB
            oCurveCol.Add oLineC
            oCurveCol.Add oArc

            oStPoint.Set 0.3 * parDiameter, _
                                (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                    pipeDiam2 / 4 + parInsulationThickness
          
            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
          
            Set oProjBdyIns = PlaceProjection(m_OutputColl, oCurve, oAxisVect, _
                                    pipeDiam2 / 2 + 2 * parInsulationThickness, True)
        
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdyIns

            Set oProjBdyIns = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineA = Nothing
            Set oLineB = Nothing
            Set oArc = Nothing
            Set oLineC = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing
            
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
                
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
            dCyldia = pipeDiam + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.175 * parFacetoFace, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam2 + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.05 * parFacetoFace, -pipeDiam / 2, 0
            oEnPoint.Set 0.05 * parFacetoFace, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, 0
            dCyldia = 0.25 * parFacetoFace + 2 * parInsulationThickness
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints1(0 To 23)  As Double
            'Point 1
            LineStrPoints1(0) = 0
            LineStrPoints1(1) = -pipeDiam / 2
            LineStrPoints1(2) = 0
            
            'Point 2
            LineStrPoints1(3) = -0.4 * parFacetoFace - parInsulationThickness
            LineStrPoints1(4) = LineStrPoints1(1)
            LineStrPoints1(5) = LineStrPoints1(2)
            
            'Point 3
            LineStrPoints1(6) = LineStrPoints1(3)
            LineStrPoints1(7) = -0.25 * parCenterlineHeight
            LineStrPoints1(8) = LineStrPoints1(2)
            
            'Point 4
            LineStrPoints1(9) = -parBodyOutsideDiameter / 2 - parInsulationThickness
            LineStrPoints1(10) = LineStrPoints1(7)
            LineStrPoints1(11) = LineStrPoints1(2)
            
            'Point 5
            LineStrPoints1(12) = LineStrPoints1(9)
            LineStrPoints1(13) = -0.8 * parCenterlineHeight - parInsulationThickness
            LineStrPoints1(14) = LineStrPoints1(2)
        
            'Point 6
            LineStrPoints1(15) = -0.2 * parBodyOutsideDiameter - parInsulationThickness
            LineStrPoints1(16) = -parCenterlineHeight - parInsulationThickness
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
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
            
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
  
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If
            
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If
            
            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dCyldia = dInsulationDia1
            Else
                dCyldia = dInsulationDia2
            End If
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Point 1
            LineStrPoints(0) = 0
            LineStrPoints(1) = (parHeight - parCenterlineHeight) / 2
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = -parDiameter / 2 - parInsulationThickness
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = 0.3 * (parHeight - parCenterlineHeight)
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = -0.5 * parBodyOutsideDiameter - parInsulationThickness
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
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.125 * parDiameter, (parHeight - parCenterlineHeight) / 2, 0
            oEnPoint.Set 0.125 * parDiameter, _
                        0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, 0
            dCyldia = 0.5 * parDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parDiameter, 0.8 * (parHeight - parCenterlineHeight), 0
            oEnPoint.Set 0.25 * parDiameter, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, 0
            dCyldia = 0.2 * parDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
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
                                0.3 * (parHeight - parCenterlineHeight), _
                                    pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -parDiameter / 2 - parInsulationThickness, _
                                    0.3 * (parHeight - parCenterlineHeight), _
                                            pipeDiam / 4 + parInsulationThickness
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine11 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Line22
            oStPoint.Set -parDiameter / 2 - parInsulationThickness, _
                                    0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -parDiameter / 2 - parInsulationThickness, _
                           -0.7 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            Set oLine22 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Arc11
            oStPoint.Set -parDiameter / 2 - parInsulationThickness, _
                           -0.7 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.5 * parBodyOutsideDiameter - parInsulationThickness, _
                        -0.95 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set -0.45 * parDiameter - parInsulationThickness, _
                            -0.9 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            
            Set oArc11 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)
          
            'Line33
            oStPoint.Set -0.5 * parBodyOutsideDiameter - parInsulationThickness, _
                        -0.95 * parCenterlineHeight - parInsulationThickness, _
                                    pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set -0.5 * parBodyOutsideDiameter, _
                                0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness
            
            Set oLine33 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
            Set oCurveCol = New Collection
            oCurveCol.Add oLine11
            oCurveCol.Add oLine22
            oCurveCol.Add oArc11
            oCurveCol.Add oLine33

            oStPoint.Set -0.5 * parBodyOutsideDiameter, _
                                0.3 * (parHeight - parCenterlineHeight), _
                                        pipeDiam / 4 + parInsulationThickness

            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
           
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1
       
            Set oProjBdyIns = PlaceProjection(m_OutputColl, oCurve, oAxisVect, _
                                        pipeDiam / 2 + 2 * parInsulationThickness, True)
       
            'Set the output
            m_OutputColl.AddOutput "BucketBody", oProjBdyIns

            Set oProjBdyIns = Nothing
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
            oStPoint.Set -0.5 * parBodyOutsideDiameter - parInsulationThickness, _
                            -parCenterlineHeight - parInsulationThickness, _
                             pipeDiam / 4 + parInsulationThickness
            oEnPoint.Set 0.45 * parDiameter + parInsulationThickness, _
                            -0.8 * parCenterlineHeight, _
                            -pipeDiam / 4 - parInsulationThickness
            
            Set objBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", objBox1
            Set objBox1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.45 * parDiameter, -0.9 * parCenterlineHeight, 0
            oEnPoint.Set 0.5 * parDiameter + parInsulationThickness, _
                                    -0.9 * parCenterlineHeight, 0
            dCyldia = 0.3 * parCenterlineHeight + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)
            
            'Set the output
            m_OutputColl.AddOutput "InsCyl", objInsCyl
            Set objInsCyl = Nothing
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

            oStPoint.Set 0.3 * parDiameter, _
                        (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                    pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, _
                                    pipeDiam2 / 4 + parInsulationThickness

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineA2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineB2
            oStPoint.Set 0.3 * parDiameter, pipeDiam2 / 2, _
                                    pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                                    pipeDiam2 / 2, pipeDiam2 / 4 + parInsulationThickness
            Set oLineB2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'LineC2
            oStPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                                    pipeDiam2 / 2, pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                        0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                            pipeDiam2 / 4 + parInsulationThickness
            Set oLineC2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)

            'Arc
            oStPoint.Set 0.45 * parFacetoFace + parInsulationThickness, _
                        0.8 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                        pipeDiam2 / 4 + parInsulationThickness
            oEnPoint.Set 0.3 * parDiameter, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                            pipeDiam2 / 4 + parInsulationThickness
            Set oArcPoint = New AutoMath.DPosition
            oArcPoint.Set 0.4 * parFacetoFace + parInsulationThickness, _
                            0.93 * (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                        pipeDiam2 / 4 + parInsulationThickness

            Set oArc2 = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oArcPoint)

            Set oCurveCol = New Collection
            oCurveCol.Add oLineA2
            oCurveCol.Add oLineB2
            oCurveCol.Add oLineC2
            oCurveCol.Add oArc2

            oStPoint.Set 0.3 * parDiameter, _
                            (parHeight - parCenterlineHeight) + parInsulationThickness, _
                                        pipeDiam2 / 4 + parInsulationThickness

            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)

            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 0, -1

            Set oProjBdyIns = PlaceProjection(m_OutputColl, oCurve, oAxisVect, _
                                            pipeDiam2 / 2 + 2 * parInsulationThickness, True)

            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", oProjBdyIns

            Set oProjBdyIns = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineA2 = Nothing
            Set oLineB2 = Nothing
            Set oArc2 = Nothing
            Set oLineC2 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
            Set oGeomFactory = Nothing
            
            'Place Insulation for Port 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
   
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
            Set objInsulatedNozzle1 = Nothing
            
            'Place Insulation for Port 2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
            Set objInsulatedNozzle2 = Nothing
        
    Case Else
        GoTo ErrorLabel:

    End Select

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
