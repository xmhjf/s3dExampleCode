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
'   Author:          KKC
'   Creation Date:  Wednesday, June 06 2007
'   Description:
'
'
'   Change History:
'   dd.mm.yyyy     who     change description
'   -----------    -----    ------------------
'   12.10.2007       KKC    CR-121408  Create additional symbols required by Intergraph's DIN piping specs
'   24.07.2008       PK     CR-145850  Made the attribute 'Instrument Width' optional for the partdatabasis 479
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    
    Dim iOutput     As Double
    Dim ObjCyl1 As Object
    Dim ObjCyl2 As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentWidth As Double
    Dim parSightDiameter As Double
    Dim parHousingDiameter As Double
    Dim parMajorBodyDiameter As Double
    Dim parViewingWindowWidth As Double
    Dim parViewingWindowHeight As Double
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim lPartDataBasis As Long
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oDir As AutoMath.DVector
    Dim parInsulationDiameter As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim Surfset   As IngrGeom3D.IJElements
    Dim ObjTopSupport As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(10)
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If lPartDataBasis = 479 Then
    '  Insert your code for output 1(Cylinder1)
        parFacetoFace = arrayOfInputs(2)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        Dim dInstrumentWidth As Double
        
         'Check to make the attribute 'Instrument Width' optional
        If CmpDblEqual(parInstrumentWidth, 0) Then
            'Setting the value of Instrument width to 2.5 times the pipe diameter 1 of the sight flow indicator
            dInstrumentWidth = 2.5 * pipeDiam
        Else
            dInstrumentWidth = parInstrumentWidth
        End If
        
        parFacetoCenter = parFacetoFace / 2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + parInsulationThickness, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
        Set oDir = New AutoMath.DVector
        oDir.Set 0, 1, 0
        
        Dim ObjCyl3 As IngrGeom3D.Revolution3d
        Set ObjCyl3 = New IngrGeom3D.Revolution3d
        Dim dlinestringPoints(0 To 14) As Double
        dlinestringPoints(0) = -parSightDiameter / 2
        dlinestringPoints(1) = -dInstrumentWidth / 2
        dlinestringPoints(2) = 0
        
        dlinestringPoints(3) = -1.3 * parSightDiameter / 2 - parInsulationThickness / 2
        dlinestringPoints(4) = -dInstrumentWidth / 2
        dlinestringPoints(5) = 0
        
        dlinestringPoints(6) = -1.3 * parSightDiameter / 2 - parInsulationThickness / 2
        dlinestringPoints(7) = dInstrumentWidth / 2
        dlinestringPoints(8) = 0
        
        dlinestringPoints(9) = -parSightDiameter / 2
        dlinestringPoints(10) = dInstrumentWidth / 2
        dlinestringPoints(11) = 0
         
        dlinestringPoints(12) = -parSightDiameter / 2
        dlinestringPoints(13) = -dInstrumentWidth / 2
        dlinestringPoints(14) = 0
        
        Dim oCenterPoint As AutoMath.DPosition
        Set oCenterPoint = New AutoMath.DPosition
        oCenterPoint.Set 0, 0, 0
        
        Dim oLineString As IngrGeom3D.LineString3d
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinestringPoints)
        Set ObjCyl3 = PlaceRevolution(m_OutputColl, oLineString, oDir, oCenterPoint, 2 * PI, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl3
        Set ObjCyl3 = Nothing
        
    ElseIf lPartDataBasis = 480 Then
        parFacetoFace = arrayOfInputs(2)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        parFacetoCenter = parFacetoFace / 2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + parInsulationThickness, True)
    ' Set the output
      
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing

        Dim dlinestringPoints1(0 To 20) As Double
        dlinestringPoints1(0) = -0.35 * parFacetoFace - parInsulationThickness
        dlinestringPoints1(1) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(2) = 0
        
        dlinestringPoints1(3) = -0.35 * parFacetoFace - parInsulationThickness
        dlinestringPoints1(4) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(5) = -0.35 * parFacetoFace - parInsulationThickness
        
        dlinestringPoints1(6) = 0.35 * parFacetoFace + parInsulationThickness
        dlinestringPoints1(7) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(8) = -0.35 * parFacetoFace - parInsulationThickness
        
        dlinestringPoints1(9) = 0.35 * parFacetoFace + parInsulationThickness
        dlinestringPoints1(10) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(11) = 0
         
        dlinestringPoints1(12) = parSightDiameter / 2
        dlinestringPoints1(13) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(14) = 0
        
        dlinestringPoints1(15) = 0
        dlinestringPoints1(16) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(17) = -parSightDiameter / 2
        
        dlinestringPoints1(18) = -parSightDiameter / 2
        dlinestringPoints1(19) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dlinestringPoints1(20) = 0
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Dim oProjCurve As Collection
        Set oProjCurve = New Collection
        
        Dim oBottomline1 As IngrGeom3D.Line3d
        Set oBottomline1 = New IngrGeom3D.Line3d
        Set oBottomline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dlinestringPoints1(18), dlinestringPoints1(19), dlinestringPoints1(20), dlinestringPoints1(0), dlinestringPoints1(1), dlinestringPoints1(2))
        
        oProjCurve.Add oBottomline1
        
        Dim oBottomline2 As IngrGeom3D.Line3d
        Set oBottomline2 = New IngrGeom3D.Line3d
        Set oBottomline2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dlinestringPoints1(0), dlinestringPoints1(1), dlinestringPoints1(2), dlinestringPoints1(3), dlinestringPoints1(4), dlinestringPoints1(5))
        
        oProjCurve.Add oBottomline2
        
        Dim oBottomline3 As IngrGeom3D.Line3d
        Set oBottomline3 = New IngrGeom3D.Line3d
        Set oBottomline3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dlinestringPoints1(3), dlinestringPoints1(4), dlinestringPoints1(5), dlinestringPoints1(6), dlinestringPoints1(7), dlinestringPoints1(8))
        
        oProjCurve.Add oBottomline3
        
        Dim oBottomline4 As IngrGeom3D.Line3d
        Set oBottomline4 = New IngrGeom3D.Line3d
        Set oBottomline4 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dlinestringPoints1(6), dlinestringPoints1(7), dlinestringPoints1(8), dlinestringPoints1(9), dlinestringPoints1(10), dlinestringPoints1(11))
        
        oProjCurve.Add oBottomline4
        
        Dim oBottomline5 As IngrGeom3D.Line3d
        Set oBottomline5 = New IngrGeom3D.Line3d
        Set oBottomline5 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dlinestringPoints1(9), dlinestringPoints1(10), dlinestringPoints1(11), dlinestringPoints1(12), dlinestringPoints1(13), dlinestringPoints1(14))
        
        
        oProjCurve.Add oBottomline5
        
        Dim oBottomArc As IngrGeom3D.Arc3d
        Set oBottomArc = New IngrGeom3D.Arc3d
        Set oBottomArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinestringPoints1(12), dlinestringPoints1(13), dlinestringPoints1(14), dlinestringPoints1(15), dlinestringPoints1(16), dlinestringPoints1(17), dlinestringPoints1(18), dlinestringPoints1(19), dlinestringPoints1(20))
        
        oProjCurve.Add oBottomArc
        
        Dim oProjCollection As IngrGeom3D.ComplexString3d
        Set oProjCollection = New IngrGeom3D.ComplexString3d
        oStPoint.Set dlinestringPoints1(18), dlinestringPoints1(19), dlinestringPoints1(20)
        Set oProjCollection = PlaceTrCString(oStPoint, oProjCurve)
        Set oDir = New AutoMath.DVector
        
        oDir.Set 0, 1, 0
        Dim OProj1 As Object
        
        Set OProj1 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
     ' Set the output
       
        m_OutputColl.AddOutput "ObjProj_", OProj1
        Set OProj1 = Nothing
        
        Dim oTransMatrix As AutoMath.DT4x4
        Set oTransMatrix = New AutoMath.DT4x4
        
        Dim OProj2 As Object
        
        Set OProj2 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI, oDir
        OProj2.Transform oTransMatrix
        
    ' Set the output
      
        m_OutputColl.AddOutput "ObjProj_", OProj2
        Set OProj2 = Nothing
        
        Dim OProj3 As Object
        
        Set OProj3 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        oDir.Set 0, -(parInstrumentWidth - 0.25 * pipeDiam), 0
        oTransMatrix.Translate oDir
        
        OProj3.Transform oTransMatrix
        
       
        m_OutputColl.AddOutput "ObjProj_", OProj3
        Set OProj3 = Nothing
        
        Dim OProj4 As Object
        oDir.Set 0, 1, 0
        Set OProj4 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        
        oTransMatrix.Rotate PI, oDir
        OProj4.Transform oTransMatrix
        
        oDir.Set 0, -(parInstrumentWidth - 0.25 * pipeDiam), 0
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oDir
        OProj4.Transform oTransMatrix
       
        m_OutputColl.AddOutput "ObjProj_", OProj4
        Set OProj4 = Nothing
 
    '  Insert your code for output 4(Sight Flow Indicator spacer 1)
    
        oStPoint.Set -0.3 * parFacetoFace, -parInstrumentWidth / 2 + 0.25 * pipeDiam, -0.3 * parFacetoFace
        oEnPoint.Set -0.3 * parFacetoFace, parInstrumentWidth / 2 - 0.25 * pipeDiam, -0.3 * parFacetoFace
        Dim ObjSightSpacer1 As Object
        Set ObjSightSpacer1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFacetoFace + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer1
        Set ObjSightSpacer1 = Nothing
    '  Insert your code for output 5(Sight Flow Indicator spacer 2)
        Dim ObjSightSpacer2 As Object
        Set ObjSightSpacer2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oDir.Set 0.6 * parFacetoFace, 0, 0
        oTransMatrix.Translate oDir
        ObjSightSpacer2.Transform oTransMatrix
'    ' Set the output
  
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer2
        Set ObjSightSpacer2 = Nothing
'    ' Insert your code for output 6(Sight Flow Indicator spacer 3)
        Dim ObjSightSpacer3 As Object
        Set ObjSightSpacer3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oDir.Set 0.6 * parFacetoFace, 0, 0.6 * parFacetoFace
        oTransMatrix.Translate oDir
        ObjSightSpacer3.Transform oTransMatrix
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer3
        Set ObjSightSpacer3 = Nothing
'
'    ' Insert your code for output 7(Sight Flow Indicator spacer 4
        Dim ObjSightSpacer4 As Object
        Set ObjSightSpacer4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oDir.Set 0, 0, 0.6 * parFacetoFace
        oTransMatrix.Translate oDir
        ObjSightSpacer4.Transform oTransMatrix
    ' Set the output
      
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer4
        Set ObjSightSpacer4 = Nothing

        Dim ObjSightCyl As Object

    ElseIf lPartDataBasis = 481 Then
        parFacetoFace = arrayOfInputs(2)
        parSightDiameter = arrayOfInputs(5)
        parFacetoCenter = parFacetoFace / 2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parSightDiameter + parInsulationThickness, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
    '  Insert your code for output 2(Sight Flow Indicator spacer 1)
    
        oStPoint.Set -parFacetoFace / 2 + flangeThick, -flangeDiam * 0.35, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, -flangeDiam * 0.35, 0
        
        Set ObjSightSpacer1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * pipeDiam + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer1
        Set ObjSightSpacer1 = Nothing
        
    '  Insert your code for output 3(Sight Flow Indicator spacer 2)
    
        oStPoint.Set -parFacetoFace / 2 + flangeThick, flangeDiam * 0.35, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, flangeDiam * 0.35, 0
        
        Set ObjSightSpacer2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * pipeDiam + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer2
        Set ObjSightSpacer2 = Nothing
        
    '  Insert your code for output 4(Sight Flow Indicator spacer 3)
    
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, flangeDiam * 0.35
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, flangeDiam * 0.35
        
        Set ObjSightSpacer3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * pipeDiam + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer3
        Set ObjSightSpacer3 = Nothing
        
    '  Insert your code for output 5(Sight Flow Indicator spacer 4)
    
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, -flangeDiam * 0.35
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, -flangeDiam * 0.35
        
        Set ObjSightSpacer4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.15 * pipeDiam + parInsulationThickness, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer4
        Set ObjSightSpacer4 = Nothing
        
    ElseIf lPartDataBasis = 482 Then
    
        parFacetoFace = arrayOfInputs(2)
        parViewingWindowWidth = arrayOfInputs(6)
        parViewingWindowHeight = arrayOfInputs(7)
        parMajorBodyDiameter = arrayOfInputs(8)
        parFacetoCenter = parFacetoFace / 2
        
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = 1.5 * pipeDiam
        If CmpDblEqual(parMajorBodyDiameter, 0) Then parMajorBodyDiameter = 0.8 * flangeDiam
        
    '  Insert your code for output 1(Cylinder1)
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set -parViewingWindowWidth / 2, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parMajorBodyDiameter + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
    ' Insert your code for Output 2(Cylinder 2)
    
        oStPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        oEnPoint.Set parViewingWindowWidth / 2, 0, 0
        
        Set ObjCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parMajorBodyDiameter + parInsulationThickness, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl2
        Set ObjCyl2 = Nothing
    ' Insert your code for Output 3(Viewing Window)
        Dim dViewingWindowDepth  As Double
        dViewingWindowDepth = Sqr((parMajorBodyDiameter / 2 + parInsulationThickness / 2) ^ 2 - (parViewingWindowHeight / 2) ^ 2)
        Dim dPoints1(0 To 8) As Double
        
        dPoints1(0) = -parViewingWindowWidth / 2
        dPoints1(1) = parViewingWindowHeight / 2
        dPoints1(2) = dViewingWindowDepth
        
        dPoints1(3) = -parViewingWindowWidth / 2
        dPoints1(4) = parMajorBodyDiameter / 2 + parInsulationThickness / 2
        dPoints1(5) = 0
        
        dPoints1(6) = -parViewingWindowWidth / 2
        dPoints1(7) = parViewingWindowHeight / 2
        dPoints1(8) = -dViewingWindowDepth
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        Set oProjCurve = New Collection
        Set oBottomline1 = New IngrGeom3D.Line3d
        Set oBottomline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                  dPoints1(6), dPoints1(7), dPoints1(8), dPoints1(0), dPoints1(1), dPoints1(2))
        
        oProjCurve.Add oBottomline1
        
        Set oBottomArc = New IngrGeom3D.Arc3d
        Set oBottomArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dPoints1(0), dPoints1(1), dPoints1(2), dPoints1(3), dPoints1(4), dPoints1(5), dPoints1(6), dPoints1(7), dPoints1(8))
        
        oProjCurve.Add oBottomArc
        
        Set oProjCollection = New IngrGeom3D.ComplexString3d
        oStPoint.Set dPoints1(6), dPoints1(7), dPoints1(8)
        Set oProjCollection = PlaceTrCString(oStPoint, oProjCurve)
        Set oDir = New AutoMath.DVector
        
        oDir.Set 1, 0, 0
        
        Set OProj1 = PlaceProjection(m_OutputColl, oProjCollection, oDir, parViewingWindowWidth, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjProj_", OProj1
        Set OProj1 = Nothing
        
        Set OProj2 = PlaceProjection(m_OutputColl, oProjCollection, oDir, parViewingWindowWidth, True)
    
        Set oTransMatrix = New AutoMath.DT4x4
        
        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI, oDir
        OProj2.Transform oTransMatrix
        
    ' Set the output
      
        m_OutputColl.AddOutput "ObjProj_", OProj2
        Set OProj2 = Nothing
    
     ElseIf lPartDataBasis = 483 Then
     '  Insert your code for output 1(Cylinder1)
        parFacetoFace = arrayOfInputs(2)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        parFacetoCenter = parFacetoFace / 2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
 
    '  Insert your code for output 3(Vertical Cylinder for flow indication)
        Set oDir = New AutoMath.DVector
        oDir.Set 0, 0, 1
    
        Set ObjCyl3 = New IngrGeom3D.Revolution3d

        dlinestringPoints(0) = -parSightDiameter / 2
        dlinestringPoints(1) = 0.4 * pipeDiam
        dlinestringPoints(2) = -parInstrumentWidth / 2
        
        dlinestringPoints(3) = -1.3 * parSightDiameter / 2 - parInsulationThickness / 2
        dlinestringPoints(4) = 0.4 * pipeDiam
        dlinestringPoints(5) = -parInstrumentWidth / 2
        
        dlinestringPoints(6) = -1.3 * parSightDiameter / 2 - parInsulationThickness / 2
        dlinestringPoints(7) = 0.4 * pipeDiam
        dlinestringPoints(8) = parInstrumentWidth / 2
        
        dlinestringPoints(9) = -parSightDiameter / 2
        dlinestringPoints(10) = 0.4 * pipeDiam
        dlinestringPoints(11) = parInstrumentWidth / 2
         
        dlinestringPoints(12) = -parSightDiameter / 2
        dlinestringPoints(13) = 0.4 * pipeDiam
        dlinestringPoints(14) = -parInstrumentWidth / 2
        
'        Dim oCenterPoint As AutoMath.DPosition
        Set oCenterPoint = New AutoMath.DPosition
        oCenterPoint.Set 0, 0.4 * pipeDiam, 0
        
'        Dim oLineString As IngrGeom3D.LineString3d
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinestringPoints)
        Set ObjCyl3 = PlaceRevolution(m_OutputColl, oLineString, oDir, oCenterPoint, 2 * PI, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl3
        Set ObjCyl3 = Nothing
        
    ElseIf lPartDataBasis = 484 Then
        parFacetoFace = arrayOfInputs(2)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        parHousingDiameter = arrayOfInputs(9)
        parFacetoCenter = parFacetoFace / 2
        If CmpDblEqual(parSightDiameter, 0) Then parSightDiameter = 0.75 * parHousingDiameter
        If CmpDblEqual(parInstrumentWidth, 0) Then parInstrumentWidth = 0.8 * pipeDiam + parSightDiameter / 2
    '  Insert your code for output 1(Cylinder1)
        oStPoint.Set -parFacetoFace / 2, -0.7 * pipeDiam - parInsulationThickness / 2, -0.7 * pipeDiam - parInsulationThickness / 2
        oEnPoint.Set parFacetoFace / 2, 0.7 * pipeDiam + parInsulationThickness / 2, 0.7 * pipeDiam + parInsulationThickness / 2
        
        Dim Objbox As Object
        Set Objbox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ' Set the output
      
        m_OutputColl.AddOutput "ObjBox_", Objbox
        Set Objbox = Nothing
        
    '  Insert your code for output 2(Vertical Cylinder for Housing)
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, parInstrumentWidth - parSightDiameter / 2, 0
        
        Set ObjCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parHousingDiameter + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl2
        Set ObjCyl2 = Nothing
    '  Insert your code for output 2(Vertical Cylinder for Housing)
        Dim ObjIndicator As Object
        oStPoint.Set 0, parInstrumentWidth - parSightDiameter / 2, 0
        Set ObjIndicator = PlaceSphere(m_OutputColl, oStPoint, parSightDiameter / 2)
     ' Set the output
       
        m_OutputColl.AddOutput "ObjSphere_", ObjIndicator
        Set ObjIndicator = Nothing
        
    ElseIf lPartDataBasis = 485 Then
        parFacetoFace = arrayOfInputs(2)
        parSightDiameter = arrayOfInputs(5)
        parFacetoCenter = parFacetoFace / 2
        
    '  Insert your code for output 1(Cylinder1)
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parSightDiameter, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
        
    ElseIf lPartDataBasis = 486 Then
        parFacetoFace = arrayOfInputs(2)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        parFacetoCenter = parFacetoFace / 2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing

        Dim dPoints(0 To 17) As Double
        dPoints(0) = -parSightDiameter - parInsulationThickness
        dPoints(1) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(2) = 0
        
        dPoints(3) = 0
        dPoints(4) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(5) = -parSightDiameter - parInsulationThickness
        
        dPoints(6) = parSightDiameter + parInsulationThickness
        dPoints(7) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(8) = 0
        
        dPoints(9) = parSightDiameter / 2
        dPoints(10) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(11) = 0
        
        dPoints(12) = 0
        dPoints(13) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(14) = -parSightDiameter / 2
         
        dPoints(15) = -parSightDiameter / 2
        dPoints(16) = parInstrumentWidth / 2 - 0.25 * pipeDiam
        dPoints(17) = 0
        
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        Set oProjCurve = New Collection
        
        
        Set oBottomline1 = New IngrGeom3D.Line3d
        Set oBottomline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dPoints(15), dPoints(16), dPoints(17), dPoints(0), dPoints(1), dPoints(2))
        
        oProjCurve.Add oBottomline1
        
        Dim oBottomArc1 As IngrGeom3D.Arc3d
        Set oBottomArc1 = New IngrGeom3D.Arc3d
        Set oBottomArc1 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                 dPoints(0), dPoints(1), dPoints(2), dPoints(3), dPoints(4), dPoints(5), dPoints(6), dPoints(7), dPoints(8))
        
        oProjCurve.Add oBottomArc1
        
        
        Set oBottomline2 = New IngrGeom3D.Line3d
        Set oBottomline2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                 dPoints(6), dPoints(7), dPoints(8), dPoints(9), dPoints(10), dPoints(11))
        
        oProjCurve.Add oBottomline2
        
        Dim oBottomArc2 As IngrGeom3D.Arc3d
        Set oBottomArc2 = New IngrGeom3D.Arc3d
        Set oBottomArc2 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                 dPoints(9), dPoints(10), dPoints(11), dPoints(12), dPoints(13), dPoints(14), dPoints(15), dPoints(16), dPoints(17))
        
        oProjCurve.Add oBottomArc2
        Set oProjCollection = New IngrGeom3D.ComplexString3d
        oStPoint.Set dPoints(15), dPoints(16), dPoints(17)
        Set oProjCollection = PlaceTrCString(oStPoint, oProjCurve)
        Set oDir = New AutoMath.DVector
        
        oDir.Set 0, 1, 0
        
        Set OProj1 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
     ' Set the output
       
        m_OutputColl.AddOutput "ObjProj_", OProj1
        Set OProj1 = Nothing
        Set oTransMatrix = New AutoMath.DT4x4
        Set OProj2 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI, oDir
        OProj2.Transform oTransMatrix
        
    ' Set the output
       
        m_OutputColl.AddOutput "ObjProj_", OProj2
        Set OProj2 = Nothing
        
        Set OProj3 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        oDir.Set 0, -(parInstrumentWidth - 0.25 * pipeDiam), 0
        oTransMatrix.Translate oDir
        
        OProj3.Transform oTransMatrix
        
        
        m_OutputColl.AddOutput "ObjProj_", OProj3
        Set OProj3 = Nothing
        
        oDir.Set 0, 1, 0
        Set OProj4 = PlaceProjection(m_OutputColl, oProjCollection, oDir, 0.25 * pipeDiam, True)
        
        oTransMatrix.LoadIdentity
        
        oTransMatrix.Rotate PI, oDir
        OProj4.Transform oTransMatrix
        
        oDir.Set 0, -(parInstrumentWidth - 0.25 * pipeDiam), 0
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oDir
        OProj4.Transform oTransMatrix
        
        m_OutputColl.AddOutput "ObjProj_", OProj4
        Set OProj4 = Nothing
    '  Insert your code for output 4(Sight Flow Indicator spacer 1)
    
        oStPoint.Set -0.75 * parSightDiameter, -parInstrumentWidth / 2 + 0.25 * pipeDiam, 0
        oEnPoint.Set -0.75 * parSightDiameter, parInstrumentWidth / 2 - 0.25 * pipeDiam, 0
        
        Set ObjSightSpacer1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parSightDiameter + parInsulationThickness, True)
    ' Set the output
     
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer1
        Set ObjSightSpacer1 = Nothing
    '  Insert your code for output 5(Sight Flow Indicator spacer 2)
        
        Set ObjSightSpacer2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parSightDiameter + parInsulationThickness, True)
        oTransMatrix.LoadIdentity
        oDir.Set 1.5 * parSightDiameter, 0, 0
        oTransMatrix.Translate oDir
        ObjSightSpacer2.Transform oTransMatrix
'    ' Set the output
      
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer2
        Set ObjSightSpacer2 = Nothing
'    ' Insert your code for output 6(Sight Flow Indicator spacer 3)

       
        Set ObjSightSpacer3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parSightDiameter + parInsulationThickness, True)
        oTransMatrix.LoadIdentity
        oDir.Set 0.75 * parSightDiameter, 0, 0.75 * parSightDiameter
        oTransMatrix.Translate oDir
        ObjSightSpacer3.Transform oTransMatrix
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer3
        Set ObjSightSpacer3 = Nothing
        
'    ' Insert your code for output 7(Sight Flow Indicator spacer 4
       
        Set ObjSightSpacer4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parSightDiameter + parInsulationThickness, True)
        oTransMatrix.LoadIdentity
        oDir.Set 0.75 * parSightDiameter, 0, -0.75 * parSightDiameter
        oTransMatrix.Translate oDir
        ObjSightSpacer4.Transform oTransMatrix
        
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjSightSpacer4
        Set ObjSightSpacer4 = Nothing
        
    ElseIf lPartDataBasis = 487 Then
        '  Insert your code for output 1(Cylinder1)
        parFacetoCenter = arrayOfInputs(3)
        parInstrumentWidth = arrayOfInputs(4)
        parSightDiameter = arrayOfInputs(5)
        parFacetoFace = 2 * parFacetoCenter
        
        If CmpDblEqual(parSightDiameter, 0) Then parSightDiameter = 0.4 * parFacetoCenter
        
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
        
        Set ObjCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)
    ' Set the output
        
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl1
        Set ObjCyl1 = Nothing
        
        Set oDir = New AutoMath.DVector
        oDir.Set 0, 1, 0
        
        Set ObjCyl3 = New IngrGeom3D.Revolution3d
        Dim dlinestringPoints2(0 To 26) As Double
        
        dlinestringPoints2(0) = -parSightDiameter / 2
        dlinestringPoints2(1) = -parInstrumentWidth
        dlinestringPoints2(2) = 0
        
        dlinestringPoints2(3) = -1.7 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(4) = -parInstrumentWidth
        dlinestringPoints2(5) = 0
        
        dlinestringPoints2(6) = -1.7 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(7) = -parInstrumentWidth + flangeThick
        dlinestringPoints2(8) = 0
        
        dlinestringPoints2(9) = -1.3 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(10) = -parInstrumentWidth + flangeThick
        dlinestringPoints2(11) = 0
        
        dlinestringPoints2(12) = -1.3 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(13) = parInstrumentWidth - flangeThick
        dlinestringPoints2(14) = 0
        
        dlinestringPoints2(15) = -1.7 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(16) = parInstrumentWidth - flangeThick
        dlinestringPoints2(17) = 0
         
        dlinestringPoints2(18) = -1.7 * parSightDiameter / 2 - parInsulationThickness
        dlinestringPoints2(19) = parInstrumentWidth
        dlinestringPoints2(20) = 0
        
        dlinestringPoints2(21) = -parSightDiameter / 2
        dlinestringPoints2(22) = parInstrumentWidth
        dlinestringPoints2(23) = 0
        
        dlinestringPoints2(24) = -parSightDiameter / 2
        dlinestringPoints2(25) = -parInstrumentWidth
        dlinestringPoints2(26) = 0
        
        Set oCenterPoint = New AutoMath.DPosition
        oCenterPoint.Set 0, 0, 0
    
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dlinestringPoints2)
        Set ObjCyl3 = PlaceRevolution(m_OutputColl, oLineString, oDir, oCenterPoint, 2 * PI, True)
    ' Set the output
       
        m_OutputColl.AddOutput "ObjCyl_", ObjCyl3
        Set ObjCyl3 = Nothing
    End If
'   Insulation for the ports
'   Insulation for Port 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort1 As Object
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'   Set the output
   
    m_OutputColl.AddOutput "InsPort_", ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Place Insulation for Port 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort2 As Object
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'   Set the output
   
    m_OutputColl.AddOutput "InsPort_", ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
