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
'   Author:         KKC
'   Creation Date:  Tuesday, Sep 18 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI                  As Double
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjHoriCylinder  As Object
        
    Dim parFacetoFace As Double
    Dim parTrapHeight As Double
    Dim parTrapBodyDiameter As Double
    Dim parTrapLength1 As Double
    Dim parTrapLenght2 As Double
    Dim parTrapHeight1 As Double
    Dim parTrapHeight2 As Double
    Dim parOffset As Double
    Dim parPipingSpecialtyHeight As Double
    Dim parPipingSpecialtyDepth As Double
    Dim parNozzleOffset2 As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parPortToPortDepOff As Double
    Dim parPortToPortHeiOff As Double
    Dim parTrapOffset1 As Double
    Dim parTrapOffset2 As Double
    Dim parTrapLength As Double
    Dim parDepth1 As Double
    Dim parDepth2 As Double
    Dim parNozzleOffset1 As Double
    Dim lPartdataBasis As Long
    Dim parInsulationThickness As Double
    Dim parInsulationDiameter As Double
    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    Dim dFacetoCenter As Double

    Dim oStPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartdataBasis = oPipeComponent.PartDataBasis
    
' Inputs
    
    
    If lPartdataBasis <= 1 Or lPartdataBasis = 390 Then
        parFacetoFace = arrayOfInputs(2)
        parTrapHeight = arrayOfInputs(3)
        parTrapBodyDiameter = arrayOfInputs(4)
        parTrapLength1 = arrayOfInputs(5)
        parTrapLenght2 = arrayOfInputs(6)
        parInsulationThickness = arrayOfInputs(22)

    ElseIf lPartdataBasis = 391 Then
        parFacetoFace = arrayOfInputs(2)
        parTrapBodyDiameter = arrayOfInputs(4)
        parTrapHeight1 = arrayOfInputs(3)
        parTrapHeight2 = arrayOfInputs(8)
        parOffset = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(22)
        
    ElseIf lPartdataBasis = 392 Then
        parPipingSpecialtyHeight = arrayOfInputs(14)  'B
        parPipingSpecialtyDepth = arrayOfInputs(16)  'C
        parNozzleOffset2 = arrayOfInputs(17)   'D
        parPipingSpecialtyWidth = arrayOfInputs(15)   'H
        parPortToPortDepOff = arrayOfInputs(18)  'K
        parPortToPortHeiOff = arrayOfInputs(19)   'M
        parNozzleOffset1 = arrayOfInputs(20)  'N
        parInsulationThickness = arrayOfInputs(22)
        
    ElseIf lPartdataBasis = 393 Then
        parFacetoFace = arrayOfInputs(2)
        parTrapBodyDiameter = arrayOfInputs(4)
        parOffset = arrayOfInputs(9)
        parTrapHeight = arrayOfInputs(21)
        parInsulationThickness = arrayOfInputs(22)
        
    ElseIf lPartdataBasis = 394 Then
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyWidth = arrayOfInputs(15)
        parTrapBodyDiameter = arrayOfInputs(4)
        parTrapOffset1 = arrayOfInputs(10)
        parTrapOffset2 = arrayOfInputs(11)
        parTrapHeight = arrayOfInputs(21)
        parInsulationThickness = arrayOfInputs(22)
        
    ElseIf lPartdataBasis = 395 Then
        parTrapHeight = arrayOfInputs(21)
        parPipingSpecialtyWidth = arrayOfInputs(15)
        parPortToPortHeiOff = arrayOfInputs(19)
        parOffset = arrayOfInputs(9)
        parDepth1 = arrayOfInputs(7)
        parDepth2 = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(22)
        
    End If
    iOutput = 0
    If lPartdataBasis <= 1 Or lPartdataBasis = 390 Then
    ' Insert your code for output 1(Horizontal Cylinder)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjHoriCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, (pipeDiam * 1.1 + 2 * parInsulationThickness), True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinder
        Set ObjHoriCylinder = Nothing
      
    ' Insert your code for output 2(Vertical Cylinder)
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0.5 * parTrapHeight, 0
        Dim ObjVertCylinder As Object
        Set ObjVertCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parTrapBodyDiameter + 2 * parInsulationThickness, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylinder
        Set ObjVertCylinder = Nothing
        
    ' Insert your code for the output 3(Trap Body)
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Dim Oline1 As IngrGeom3D.Line3d
        Dim oArc1 As Object
        Dim OLine2 As IngrGeom3D.Line3d
        Dim Oline3 As IngrGeom3D.Line3d
        Dim OCollection As Collection
        Set OCollection = New Collection
        oStPoint.Set 0, 0.5 * parTrapHeight - parInsulationThickness, 0
        oEnPoint.Set -parTrapBodyDiameter / 2, 0.5 * parTrapHeight - parInsulationThickness, 0
        
        Set Oline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, _
                               oStPoint.y, oStPoint.z, oEnPoint.x, oEnPoint.y, oEnPoint.z)
        OCollection.Add Oline1
        
        oCenPoint.Set -parTrapBodyDiameter / 2, 0.75 * parTrapHeight, 0
       
        Dim dMMRatio As Double
        dMMRatio = (parTrapLength1 - parTrapBodyDiameter / 2 + parInsulationThickness) / (0.25 * parTrapHeight + parInsulationThickness)
        Set oArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                            oCenPoint.x, oCenPoint.y, oCenPoint.z, 0, 0, -1, 0, -0.25 * parTrapHeight - parInsulationThickness, 0, dMMRatio, 0, PI)
                            
        OCollection.Add oArc1
        
        oStPoint.Set -parTrapBodyDiameter / 2, parTrapHeight + parInsulationThickness, 0
        oEnPoint.Set 0, parTrapHeight + parInsulationThickness, 0
    
        Set Oline3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, _
                               oStPoint.y, oStPoint.z, oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
        OCollection.Add Oline3
        Set Oline3 = Nothing
        oStPoint.Set 0, parTrapHeight + parInsulationThickness, 0
        oEnPoint.Set 0, 0.5 * parTrapHeight - parInsulationThickness, 0
    
        Set Oline3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oEnPoint.x, oEnPoint.y, oEnPoint.z)
        OCollection.Add Oline3
        
        Dim OTrCurve As IngrGeom3D.ComplexString3d
        Set OTrCurve = PlaceTrCString(oEnPoint, OCollection)
        
        Dim OTrapBody As Object
        Dim ORevVec As AutoMath.DVector
        Set ORevVec = New AutoMath.DVector
        ORevVec.Set 0, 1, 0
        oCenPoint.Set 0.001, 0.75 * parTrapHeight, 0
        
        Set OTrapBody = PlaceRevolution(m_OutputColl, OTrCurve, ORevVec, oCenPoint, PI * 2, True)
        
    ' Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OTrapBody
        Set OTrapBody = Nothing
        
    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    ElseIf lPartdataBasis = 391 Then
    ' Insert your code for output 1(Horizontal Cylinder)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjHoriCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinder
        Set ObjHoriCylinder = Nothing
      
    ' Insert your code for output 2(Top Cylinder)
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, parTrapHeight1 + parInsulationThickness, 0
        Dim ObjTopCylinder As Object
        If CmpDblEqual(parTrapBodyDiameter, 0) Then parTrapBodyDiameter = 0.6 * parFacetoFace
        Set ObjTopCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * parTrapBodyDiameter + 2 * parInsulationThickness, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylinder
        Set ObjTopCylinder = Nothing
    
    ' Insert your code for output 3(Bottom Cylinder)
        
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, -parTrapHeight2 - parInsulationThickness, 0
        Dim ObjBottomCylinder As Object
        Set ObjBottomCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parTrapBodyDiameter + 2 * parInsulationThickness, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomCylinder
        Set ObjBottomCylinder = Nothing
        
    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    ElseIf lPartdataBasis = 392 Then
      
    '   Assumption: Origin of the symbol is taken at intersection point of:
    '   horizontal line drawn at mid-point of 'M'
    '   Vertical line drawn at mid-point of 'C'
    
    ' Insert your code for output 6(Body of Trap)
 
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        Dim dProjStringPoints(0 To 17) As Double
        dProjStringPoints(0) = parPipingSpecialtyDepth / 2
        dProjStringPoints(1) = -parPortToPortHeiOff / 2 - parNozzleOffset2 - parInsulationThickness
        dProjStringPoints(2) = -parNozzleOffset1 - parInsulationThickness
        
        dProjStringPoints(3) = parPipingSpecialtyDepth / 2
        dProjStringPoints(4) = -parPortToPortHeiOff / 2 - parNozzleOffset2 - parInsulationThickness
        dProjStringPoints(5) = 0.3 * (parPipingSpecialtyWidth - parNozzleOffset1)
        
        dProjStringPoints(6) = parPipingSpecialtyDepth / 2
        dProjStringPoints(7) = parPipingSpecialtyHeight / 2 - (parPortToPortHeiOff / 2 + parNozzleOffset2)
        dProjStringPoints(8) = 0.3 * (parPipingSpecialtyWidth - parNozzleOffset1)
        
        dProjStringPoints(9) = 0
        dProjStringPoints(10) = -parPipingSpecialtyHeight / 2 - parInsulationThickness
        dProjStringPoints(11) = 0
        
        dProjStringPoints(12) = parPipingSpecialtyDepth / 2
        dProjStringPoints(13) = parPipingSpecialtyHeight - parPortToPortHeiOff / 2 - parNozzleOffset2 + parInsulationThickness
        dProjStringPoints(14) = -parNozzleOffset1 - parInsulationThickness
        
        dProjStringPoints(15) = parPipingSpecialtyDepth / 2
        dProjStringPoints(16) = parPipingSpecialtyHeight - parPortToPortHeiOff / 2 - parNozzleOffset2 + parInsulationThickness
        dProjStringPoints(17) = 0.3 * (parPipingSpecialtyWidth - parNozzleOffset1)
        
        Set Oline1 = New IngrGeom3D.Line3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set Oline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dProjStringPoints(0), dProjStringPoints(1), dProjStringPoints(2), _
                                             dProjStringPoints(3), dProjStringPoints(4), dProjStringPoints(5))
     
    
        Set OCollection = New Collection
        OCollection.Add Oline1
        Dim OEllArc As IngrGeom3D.EllipticalArc3d
        oCenPoint.Set dProjStringPoints(6), dProjStringPoints(7), dProjStringPoints(8)
        
        dMMRatio = (0.7 * (parPipingSpecialtyWidth - parNozzleOffset1) + parInsulationThickness) / (parPipingSpecialtyHeight / 2 + parInsulationThickness)
        Set OEllArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                       -1, 0, 0, dProjStringPoints(9), dProjStringPoints(10), dProjStringPoints(11), dMMRatio, 0, PI)
                                               
     
        OCollection.Add OEllArc

        Set OLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dProjStringPoints(15), dProjStringPoints(16), dProjStringPoints(17), _
                                             dProjStringPoints(12), dProjStringPoints(13), dProjStringPoints(14))
       
        OCollection.Add OLine2


        Set Oline3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dProjStringPoints(12), dProjStringPoints(13), dProjStringPoints(14), _
                                             dProjStringPoints(0), dProjStringPoints(1), dProjStringPoints(2))
       
        OCollection.Add Oline3

        oStPoint.Set dProjStringPoints(0), dProjStringPoints(1), dProjStringPoints(2)
        Dim oDir As AutoMath.DVector
        Set oDir = New AutoMath.DVector
        oDir.Set -1, 0, 0
        Set OTrCurve = PlaceTrCString(oStPoint, OCollection)
        Set OTrapBody = PlaceProjection(m_OutputColl, OTrCurve, oDir, parPipingSpecialtyDepth - 2 * parPortToPortDepOff, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OTrapBody
        Set OTrapBody = Nothing

    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        dFacetoCenter = parPipingSpecialtyDepth / 2 - parPortToPortDepOff
        
        oStPoint.Set -dFacetoCenter, parPortToPortHeiOff / 2, 0
        oEnPoint.Set -dFacetoCenter + parPortToPortDepOff, parPortToPortHeiOff / 2, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        dFacetoCenter = parPipingSpecialtyDepth / 2
        oStPoint.Set -dFacetoCenter, -parPortToPortHeiOff / 2, 0
        oEnPoint.Set -dFacetoCenter + 2 * parPortToPortDepOff, -parPortToPortHeiOff / 2, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    ElseIf lPartdataBasis = 393 Then
    ' Place Horizontal cylinder
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2, 0, 0
        Set ObjHoriCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinder
        Set ObjHoriCylinder = Nothing
    ' Place bottom cylinder
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set 0, -0.075 * parTrapHeight, 0
        oEnPoint.Set 0, -0.325 * parTrapHeight - parInsulationThickness, 0
        Set ObjVertCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parTrapBodyDiameter + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylinder
        Set ObjVertCylinder = Nothing
    'Place top cylinder
        oStPoint.Set 0, -0.075 * parTrapHeight, 0
        oEnPoint.Set 0, 0.175 * parTrapHeight, 0
        Dim OVertCylinder1 As Object
        Set OVertCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * parTrapBodyDiameter + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OVertCylinder1
        Set OVertCylinder1 = Nothing
    'Place Trap Body
        oCenPoint.Set 0, 0.175 * parTrapHeight, 0
        oEnPoint.Set 0, 0.5 * parTrapHeight + parInsulationThickness, 0
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oDir = New AutoMath.DVector
        dMMRatio = (0.4 * parFacetoFace + parInsulationThickness) / (0.5 * parTrapHeight + parInsulationThickness)
        Set oArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                       oCenPoint.x, oCenPoint.y, oCenPoint.z, 0, 0, 1, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z, dMMRatio, 0, PI / 2)
        oDir.Set 0, 1, 0
        Set OTrapBody = PlaceRevolution(m_OutputColl, oArc1, oDir, oCenPoint, 2 * PI, True)
        
    ' Set the output
        m_OutputColl.AddOutput "Trap_", OTrapBody
        Set OTrapBody = Nothing
    
    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    ElseIf lPartdataBasis = 394 Then
    ' Place Horizontal cylinder
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2, 0, 0
        Set ObjHoriCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinder
        Set ObjHoriCylinder = Nothing
    ' Place Bottom Box
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        
        oStPoint.Set -parFacetoFace / 2 + flangeThick, flangeDiam / 2 + parInsulationThickness, parTrapOffset2 + parInsulationThickness
        oEnPoint.Set parFacetoFace / 2 - flangeThick, -0.7 * parTrapOffset1 - parInsulationThickness, -(flangeDiam / 2 + 0.05 * parTrapHeight + parInsulationThickness)
        
        Dim OBottomBox As Object
        Set OBottomBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OBottomBox
        Set OBottomBox = Nothing
        
    ' Place the Inclined cylinder
        oStPoint.Set 1.2 * parFacetoFace / 2, -parTrapOffset1, 0
        oEnPoint.Set 0, 0, 0
        
        Dim OInclinedCylinder As Object
        Set OInclinedCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * flangeDiam + 2 * parInsulationThickness, True)
        
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OInclinedCylinder
        Set OInclinedCylinder = Nothing
        
    ' Place Top Cylinder
       
        oStPoint.Set 0, 0, -0.5 * parTrapHeight
        oEnPoint.Set 0, 0, 0
        
        Set OVertCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parTrapBodyDiameter + 2 * parInsulationThickness, True)
        
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OVertCylinder1
        Set OVertCylinder1 = Nothing
        
    ' Place the Trap Body
    
       Dim dRevCurPoints(0 To 23) As Double
       Dim dCylAxisHeight As Double
       dCylAxisHeight = (parPipingSpecialtyWidth - parTrapOffset1) - parTrapBodyDiameter / 2
       
       dRevCurPoints(0) = 0
       dRevCurPoints(1) = dCylAxisHeight + parTrapBodyDiameter / 2 + parInsulationThickness
       dRevCurPoints(2) = -0.4 * parTrapHeight
       
       dRevCurPoints(3) = 0
       dRevCurPoints(4) = dCylAxisHeight
       dRevCurPoints(5) = -0.4 * parTrapHeight
       
       dRevCurPoints(6) = 0
       dRevCurPoints(7) = dCylAxisHeight
       dRevCurPoints(8) = -0.9 * parTrapHeight
       
       dRevCurPoints(9) = 0
       dRevCurPoints(10) = dCylAxisHeight
       dRevCurPoints(11) = -parTrapHeight - parInsulationThickness
       
       dRevCurPoints(12) = 0
       dRevCurPoints(13) = dCylAxisHeight + parTrapBodyDiameter / 2 + parInsulationThickness
       dRevCurPoints(14) = -0.9 * parTrapHeight
       
       dRevCurPoints(15) = 0
       dRevCurPoints(16) = parTrapBodyDiameter / 2 + parInsulationThickness
       dRevCurPoints(17) = 0
       
       dRevCurPoints(18) = 0
       dRevCurPoints(19) = dCylAxisHeight
       dRevCurPoints(20) = -0.3 * parTrapHeight + parInsulationThickness
       
       dRevCurPoints(21) = 0
       dRevCurPoints(22) = dCylAxisHeight
       dRevCurPoints(23) = -parTrapHeight - parInsulationThickness
       
       dMMRatio = (0.1 * parTrapHeight) / (parTrapBodyDiameter / 2)
      
       Set OCollection = New Collection
       Set oGeomFactory = New IngrGeom3D.GeometryFactory
       Set oArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, dRevCurPoints(3), _
                                              dRevCurPoints(4), dRevCurPoints(5), 1, 0, 0, dRevCurPoints(15), dRevCurPoints(16), dRevCurPoints(17), dMMRatio, 0, PI / 2)
       Set oDir = New AutoMath.DVector
       oDir.Set 0, 0, -1
       oCenPoint.Set 0, dCylAxisHeight - 0.0001, -0.7 * parTrapHeight
       Set OTrapBody = PlaceRevolution(m_OutputColl, oArc1, oDir, oCenPoint, PI * 2, True)
       m_OutputColl.AddOutput "Strainer_", OTrapBody
       Set OTrapBody = Nothing
       Set oArc1 = Nothing
       
       Set Oline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevCurPoints(0), _
                        dRevCurPoints(1), dRevCurPoints(2), dRevCurPoints(12), dRevCurPoints(13), dRevCurPoints(14))
       Set OTrapBody = PlaceRevolution(m_OutputColl, Oline1, oDir, oCenPoint, PI * 2, True)
       m_OutputColl.AddOutput "Strainer_", OTrapBody
       Set OTrapBody = Nothing
       Set Oline1 = Nothing
       Dim oArc2 As IngrGeom3D.EllipticalArc3d
       Set oArc2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, dRevCurPoints(6), _
                                             dRevCurPoints(7), dRevCurPoints(8), -1, 0, 0, dRevCurPoints(15), dRevCurPoints(16), dRevCurPoints(17), dMMRatio, 0, PI / 2)
       Set OTrapBody = PlaceRevolution(m_OutputColl, oArc2, oDir, oCenPoint, PI * 2, True)
       m_OutputColl.AddOutput "Strainer_", OTrapBody
       Set OTrapBody = Nothing
       Set oArc2 = Nothing
    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    ElseIf lPartdataBasis = 395 Then
    ' Place Front Box
        Dim dOffset1 As Double
        dOffset1 = parTrapHeight - parPortToPortHeiOff - parOffset
        oStPoint.Set -parInsulationThickness, -(parPortToPortHeiOff / 2 + 0.7 * parOffset + parInsulationThickness), 0.3 * parPipingSpecialtyWidth + parInsulationThickness
        oEnPoint.Set 0.2 * parDepth1, parPortToPortHeiOff / 2 + 0.7 * dOffset1 + parInsulationThickness, -0.3 * parPipingSpecialtyWidth - parInsulationThickness
        Dim OFBox As Object
        Set OFBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OFBox
        Set OFBox = Nothing
    ' Place Flat Oval Body
        Dim OProjString1(0 To 23) As Double
        
        OProjString1(0) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(1) = parPortToPortHeiOff / 2 + dOffset1
        OProjString1(2) = 0
        
        OProjString1(3) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(4) = 0.8 * (parPortToPortHeiOff / 2 + dOffset1)
        OProjString1(5) = parPipingSpecialtyWidth / 2 + parInsulationThickness
        
        OProjString1(6) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(7) = -0.8 * (parPortToPortHeiOff / 2 + parOffset)
        OProjString1(8) = parPipingSpecialtyWidth / 2 + parInsulationThickness
        
        OProjString1(9) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(10) = -(parPortToPortHeiOff / 2 + parOffset)
        OProjString1(11) = 0
        
        OProjString1(12) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(13) = -0.8 * (parPortToPortHeiOff / 2 + parOffset)
        OProjString1(14) = -parPipingSpecialtyWidth / 2 - parInsulationThickness
        
        OProjString1(15) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(16) = 0.8 * (parPortToPortHeiOff / 2 + dOffset1)
        OProjString1(17) = -parPipingSpecialtyWidth / 2 - parInsulationThickness
        
        OProjString1(18) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(19) = 0.8 * (parPortToPortHeiOff / 2 + dOffset1)
        OProjString1(20) = 0
        
        OProjString1(21) = 0.2 * parDepth1 - parInsulationThickness
        OProjString1(22) = -0.8 * (parPortToPortHeiOff / 2 + parOffset)
        OProjString1(23) = 0
        
        dMMRatio = (0.2 * (parPortToPortHeiOff / 2 + dOffset1) + parInsulationThickness) / (parPipingSpecialtyWidth / 2 + parInsulationThickness)
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, OProjString1(18), _
                              OProjString1(19), OProjString1(20), 1, 0, 0, 0, 0, -parPipingSpecialtyWidth / 2 - parInsulationThickness, dMMRatio, 0, PI)
       
        Set OCollection = New Collection
        
        Set Oline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, OProjString1(3), _
                                   OProjString1(4), OProjString1(5), OProjString1(6), OProjString1(7), OProjString1(8))
                                            
        
        Set oArc2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, OProjString1(21), _
                              OProjString1(22), OProjString1(23), 1, 0, 0, 0, 0, parPipingSpecialtyWidth / 2 + parInsulationThickness, dMMRatio, 0, PI)
     
        Set OLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, OProjString1(12), _
                                   OProjString1(13), OProjString1(14), OProjString1(15), OProjString1(16), OProjString1(17))
                                   
                                   
        OCollection.Add oArc1
        OCollection.Add Oline1
        OCollection.Add oArc2
        OCollection.Add OLine2
        oStPoint.Set OProjString1(15), OProjString1(16), OProjString1(17)
        Set OTrCurve = PlaceTrCString(oStPoint, OCollection)
        Set oDir = New AutoMath.DVector
        oDir.Set 1, 0, 0
        Set OTrapBody = PlaceProjection(m_OutputColl, OTrCurve, oDir, 0.1 * parDepth1 + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), OTrapBody
        Set OTrapBody = Nothing
        
     ' Place Box
        oStPoint.Set 0.3 * parDepth1 - parInsulationThickness, -0.8 * (parPortToPortHeiOff / 2 + parOffset) - parInsulationThickness, 0.4 * parPipingSpecialtyWidth + parInsulationThickness
        oEnPoint.Set parDepth1 + parInsulationThickness, 0.8 * (parPortToPortHeiOff / 2 + dOffset1) + parInsulationThickness, -0.4 * parPipingSpecialtyWidth - parInsulationThickness
        
        Dim oTrapBox As Object
        Set oTrapBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTrapBox
        Set oTrapBox = Nothing
        
    ' Place Box
        oStPoint.Set parDepth1 - parInsulationThickness, 0.6 * (parPortToPortHeiOff / 2 + parOffset) + parInsulationThickness, 0.4 * parPipingSpecialtyWidth + parInsulationThickness
        oEnPoint.Set parDepth1 + parDepth2 + parInsulationThickness, -0.6 * (parPortToPortHeiOff / 2 + dOffset1) - parInsulationThickness, -0.4 * parPipingSpecialtyWidth - parInsulationThickness
        
        Dim oTrapBox1 As Object
        Set oTrapBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTrapBox1
        Set oTrapBox1 = Nothing

    ' Place Insulation for Port 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set 0, -parPortToPortHeiOff / 2, 0
        oEnPoint.Set flangeThick, -parPortToPortHeiOff / 2, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle1
        Set ObjInsulatedNozzle1 = Nothing
        
    ' Place Insulation for Port 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        oStPoint.Set 0, parPortToPortHeiOff / 2, 0
        oEnPoint.Set flangeThick, parPortToPortHeiOff / 2, 0
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    
    '   Set the output
       
        m_OutputColl.AddOutput "InsPort_", ObjInsulatedNozzle2
        Set ObjInsulatedNozzle2 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
         
    
    End If
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub


