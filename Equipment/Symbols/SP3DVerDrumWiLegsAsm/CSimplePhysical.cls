VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:50 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thusday, June 5 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-182.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Dim PI As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
    
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    
    
    Dim iOutput     As Double
    Dim parVesselDiameter   As Double
    Dim parVesselTantoTan   As Double
    Dim parCPtoSupport  As Double
    Dim parInsulationThickness  As Double
    
    Dim parSupportOrientation   As Double
    Dim parSupportLength    As Double
    Dim parSupportWidth As Double
    Dim parSupportHeight    As Double
    
    Dim parBackingPlateThickness    As Double
    Dim parBackingPlateDepth    As Double
    
    Dim parLengthBaseplate  As Double
    Dim parWidthBaseplate   As Double
    Dim parBasePlateThickness   As Double
    
    Dim parXboltHole    As Double
    Dim parYboltHole    As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '4
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parCPtoSupport = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    '6
    parSupportOrientation = arrayOfInputs(6)
    parSupportLength = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportHeight = arrayOfInputs(9)
    parBackingPlateThickness = arrayOfInputs(10)
    parBackingPlateDepth = arrayOfInputs(11)
    '3
    parLengthBaseplate = arrayOfInputs(12)
    parWidthBaseplate = arrayOfInputs(13)
    parBasePlateThickness = arrayOfInputs(14)
    '2
    parXboltHole = arrayOfInputs(15)
    parYboltHole = arrayOfInputs(16)
    

    Dim VesselRadius As Double
    Dim VesselOuterDiameter As Double
    Dim shellThickness As Double
    
    iOutput = 0
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    shellThickness = 0.0508
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness
    VesselRadius = VesselOuterDiameter / 2
    'Insert your cade Place Drum Body
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set 0, 0, 0
    endPt.Set 0, 0, parVesselTantoTan
    
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPt, endPt, VesselOuterDiameter

    'Insert you code for Head near Control Point
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As AutoMath.DVector
    Dim revCenPt As AutoMath.DPosition
    
    Set oAxisVect = New AutoMath.DVector
    Set revCenPt = New AutoMath.DPosition
    
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double
    Dim DomeMinorRadius As Double
    Dim oExchangerArc As IngrGeom3D.EllipticalArc3d
    Dim ObjFrontHead As Object
    
    DomeMinorRadius = 0.5 * (VesselOuterDiameter / 2)
    cenX = 0
    cenY = 0
    cenZ = 0

    MajorX = VesselOuterDiameter / 2
    MajorY = 0
    MajorZ = 0

    mMRatio = DomeMinorRadius / (VesselOuterDiameter / 2)
    StartAngle = PI / 2
    SweepAngle = PI / 2
    
    norX = 0
    norY = 1
    norZ = 0
    
    Set oExchangerArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oAxisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, 0
    'Revolve it about X-Axiz
    
    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oExchangerArc, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing
    Set oExchangerArc = Nothing
    
    'Insert you code for Head away from Control Point
    cenX = 0
    cenY = 0
    cenZ = parVesselTantoTan

    MajorX = VesselOuterDiameter / 2
    MajorY = 0
    MajorZ = 0

    mMRatio = DomeMinorRadius / (VesselOuterDiameter / 2)
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2
    
    norX = 0
    norY = 1
    norZ = 0
    
    Set oExchangerArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oAxisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, parVesselTantoTan
    'Revolve it about X-Axiz
    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oExchangerArc, oAxisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing
    Set revCenPt = Nothing
    
'Insert your code for backing plate
    Dim SupportTopZee As Double
    Dim BackingPlateProj As Double
    Dim dlinepts(0 To 14) As Double
    Dim SubtendedAng As Double
    Dim RadiustoEdge As Double
    Dim AdjSide As Double
    
    BackingPlateProj = 0.07      'as per the DOW document
    RadiustoEdge = VesselRadius + parBackingPlateThickness
    SupportTopZee = parSupportHeight + parBasePlateThickness - parCPtoSupport
    AdjSide = Sqr(Abs((VesselRadius ^ 2 - (parSupportLength / 2 + BackingPlateProj) ^ 2)))
    SubtendedAng = Atn((parSupportLength / 2 + BackingPlateProj) / AdjSide)
    
    'Point 1
    dlinepts(0) = VesselRadius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(1) = VesselRadius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(2) = SupportTopZee + BackingPlateProj
    
    'Point 2
    dlinepts(3) = RadiustoEdge * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(4) = RadiustoEdge * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(5) = dlinepts(2)
    
    'Point 3
    dlinepts(6) = RadiustoEdge * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(7) = RadiustoEdge * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(8) = dlinepts(2)
    
    'Point 4
    dlinepts(9) = VesselRadius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(10) = VesselRadius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(11) = dlinepts(2)
    
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim ObjBackPlate As IngrGeom3D.Projection3d
    Dim ProjVect As AutoMath.DVector
    Dim oLine1 As IngrGeom3D.Line3d
    Set ProjVect = New AutoMath.DVector
    
    'Line
    Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                                                            dlinepts(3), dlinepts(4), dlinepts(5))
    
    'Curve
    Dim oCurve1 As IngrGeom3D.Arc3d
    
    Set oCurve1 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, 0, 0, dlinepts(2), _
                                                        dlinepts(3), dlinepts(4), dlinepts(5), dlinepts(6), dlinepts(7), dlinepts(8))
                                                        
    'Line 2
    Dim oLine2 As IngrGeom3D.Line3d
    Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                                                            dlinepts(9), dlinepts(10), dlinepts(11))
    'Curve
    Dim oCurve2 As IngrGeom3D.Arc3d
    
    Set oCurve2 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, 0, 0, dlinepts(2), _
                                                        dlinepts(9), dlinepts(10), dlinepts(11), dlinepts(0), dlinepts(1), dlinepts(2))
    
    Dim oCompStr As IngrGeom3D.ComplexString3d
    Dim oElement As Collection
    Set oElement = New Collection
    
    oElement.Add oLine1
    oElement.Add oCurve1
    oElement.Add oLine2
    oElement.Add oCurve2
    startPt.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oCompStr = PlaceTrCString(startPt, oElement)
    
    ProjVect.Set 0, 0, -1
    Set ObjBackPlate = PlaceProjection(m_OutputColl, oCompStr, ProjVect, parBackingPlateDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackPlate
    Set ObjBackPlate = Nothing
    Set oLine1 = Nothing
    Set oLine2 = Nothing
    Set oCurve1 = Nothing
    Set oCurve2 = Nothing
    Set oElement = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
'Transform the Object and get other backing plates
    'Backing Plate 2
    Dim TransMatrix As IJDT4x4
    Dim pVector As AutoMath.DVector
    
    Set pVector = New AutoMath.DVector
    Set TransMatrix = New DT4x4
    
    pVector.Set 0, 0, 1
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI / 2, pVector
    oCompStr.Transform TransMatrix
    Set ObjBackPlate = PlaceProjection(m_OutputColl, oCompStr, ProjVect, parBackingPlateDepth, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackPlate
    Set ObjBackPlate = Nothing
    
'Insert your code for backing plate 3
    'Set TransMatrix = Nothing
    'Set TransMatrix = New DT4x4
    
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI, pVector
    oCompStr.Transform TransMatrix
    Set ObjBackPlate = PlaceProjection(m_OutputColl, oCompStr, ProjVect, parBackingPlateDepth, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackPlate
    Set ObjBackPlate = Nothing
    
'Insert your code for backing plate 4
    'Set TransMatrix = Nothing
    'Set TransMatrix = New DT4x4
    
    TransMatrix.LoadIdentity
    TransMatrix.Rotate 1.5 * PI, pVector
    oCompStr.Transform TransMatrix
    Set ObjBackPlate = PlaceProjection(m_OutputColl, oCompStr, ProjVect, parBackingPlateDepth, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackPlate
    Set ObjBackPlate = Nothing
    oCompStr.RemoveCurve True
    Set oCompStr = Nothing
    Set ProjVect = Nothing
    
 ' Insert your code for Support 1
    Dim Radius As Double
        
    SubtendedAng = Atn((parSupportLength / 2) / RadiustoEdge)
    Radius = Sqr((RadiustoEdge ^ 2) + (parSupportLength / 2) ^ 2)
                        
    'Point 1
    dlinepts(0) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(1) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(2) = SupportTopZee
    
    'Point 2
    dlinepts(3) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(4) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(5) = SupportTopZee
    
    'Point 3
    SubtendedAng = Atn((parSupportLength / 2) / (RadiustoEdge + parSupportWidth))
    Radius = Sqr((parSupportLength / 2) ^ 2 + (RadiustoEdge + parSupportWidth) ^ 2)
    
    dlinepts(6) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(7) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(8) = SupportTopZee - parSupportWidth
    
    'Point 4
    dlinepts(9) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(10) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(11) = SupportTopZee - parSupportWidth
    
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinepts)
    
    'Line String representing bottom of support
    SubtendedAng = Atn((parSupportLength / 2) / RadiustoEdge)
    Radius = Sqr((RadiustoEdge ^ 2) + (parSupportLength / 2) ^ 2)
    
    'Point 1
    dlinepts(0) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(1) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(2) = SupportTopZee - parSupportHeight
    
    'Point 2
    dlinepts(3) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(4) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(5) = SupportTopZee - parSupportHeight
    
    'Point 3
    SubtendedAng = Atn((parSupportLength / 2) / (RadiustoEdge + parSupportWidth))
    Radius = Sqr((parSupportLength / 2) ^ 2 + (RadiustoEdge + parSupportWidth) ^ 2)
    
    dlinepts(6) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(7) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(8) = SupportTopZee - parSupportHeight
    
    'Point 4
    dlinepts(9) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(10) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(11) = SupportTopZee - parSupportHeight
    
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    Dim oLineStr1 As IngrGeom3D.LineString3d
    Dim ObjSupport As IngrGeom3D.RuledSurface3d
    
    Set oLineStr1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinepts)
    Set ObjSupport = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oLineStr, oLineStr1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    
'Transform the Support to get other Support 2
    Set TransMatrix = Nothing
    Set TransMatrix = New DT4x4
    
    pVector.Set 0, 0, 1
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI / 2, pVector
    oLineStr.Transform TransMatrix
    oLineStr1.Transform TransMatrix
    Set ObjSupport = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oLineStr, oLineStr1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    
'Support 3
'    Set TransMatrix = Nothing
'    Set TransMatrix = New DT4x4
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI, pVector
    oLineStr.Transform TransMatrix
    oLineStr1.Transform TransMatrix
    Set ObjSupport = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oLineStr, oLineStr1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    
'Support 4
'    Set TransMatrix = Nothing
'    Set TransMatrix = New DT4x4

    TransMatrix.LoadIdentity
    TransMatrix.Rotate 1.5 * PI, pVector
    oLineStr.Transform TransMatrix
    oLineStr1.Transform TransMatrix
    Set ObjSupport = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oLineStr, oLineStr1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    Set oLineStr = Nothing
    Set oLineStr1 = Nothing
    
'Insert your code for Base Plate 1
    Dim ObjBaseSupport As IngrGeom3D.Projection3d
'    Dim BasePlateProjW As Double
'    Dim BasePlateProjL As Double
'
'    BasePlateProjW = parLengthBaseplate / 2 - parSupportLength / 2
'    BasePlateProjL = parWidthBaseplate / 2 - parSupportWidth / 2
    
    RadiustoEdge = RadiustoEdge + parSupportWidth / 2 - parWidthBaseplate / 2
    SubtendedAng = Atn((parLengthBaseplate / 2) / RadiustoEdge)
    Radius = Sqr((RadiustoEdge ^ 2) + (parLengthBaseplate / 2) ^ 2)
    
    'Point 1
    dlinepts(0) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(1) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(2) = SupportTopZee - parSupportHeight
    
    'Point 2
    dlinepts(3) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(4) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(5) = SupportTopZee - parSupportHeight
    
    'Point 3
    SubtendedAng = Atn((parLengthBaseplate / 2) / (RadiustoEdge + parWidthBaseplate))
    Radius = Sqr((parLengthBaseplate / 2) ^ 2 + (RadiustoEdge + parWidthBaseplate) ^ 2)
    
    dlinepts(6) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(7) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(8) = SupportTopZee - parSupportHeight
    
    'Point 4
    dlinepts(9) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(10) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(11) = SupportTopZee - parSupportHeight
    
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    Set oLineStr1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinepts)
    oAxisVect.Set 0, 0, -1
    Set ObjBaseSupport = PlaceProjection(m_OutputColl, oLineStr1, oAxisVect, parBasePlateThickness, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseSupport
    Set ObjBaseSupport = Nothing
    
'Base Plate 2
    Set TransMatrix = Nothing
    Set TransMatrix = New DT4x4
    
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI / 2, pVector
    oLineStr1.Transform TransMatrix
    Set ObjBaseSupport = PlaceProjection(m_OutputColl, oLineStr1, oAxisVect, parBasePlateThickness, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseSupport
    Set ObjBaseSupport = Nothing
'Base Plate 3
'    Set TransMatrix = Nothing
'    Set TransMatrix = New DT4x4
    
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI, pVector
    oLineStr1.Transform TransMatrix
    Set ObjBaseSupport = PlaceProjection(m_OutputColl, oLineStr1, oAxisVect, parBasePlateThickness, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseSupport
    Set ObjBaseSupport = Nothing
'Base Plate 4
'    Set TransMatrix = Nothing
'    Set TransMatrix = New DT4x4
    
    TransMatrix.LoadIdentity
    TransMatrix.Rotate 1.5 * PI, pVector
    oLineStr1.Transform TransMatrix
    Set ObjBaseSupport = PlaceProjection(m_OutputColl, oLineStr1, oAxisVect, parBasePlateThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseSupport
    Set ObjBaseSupport = Nothing
    Set oAxisVect = Nothing
    Set pVector = Nothing
    Set TransMatrix = Nothing
    Set oLine1 = Nothing
    
' Insert your code for output 9( Fundation Port)
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
          'Coordinate sysytem is at the bottom of support.
    '
    '               |-------|
    '   Y           |       |
    '   ^           |       |
    '   |           |       |
    '   |           |       |
    '   |           |       |
    '   -----> X    |       |
    '   Symbol CS   |       |
    '               |       |
    '               |   X   |
    '               |   ^   |
    '               |   |   |
    '               |   |   |
    '               |---|---|----> Y
    '                   Port CS
    
    
    Dim objEqpFoundationPort As IJEqpFoundationPort
'    Dim OriginPosition As Double
    Dim HoleSubAngle As Double
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort", _
                                                False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
'    Dim HoleDistFromOrigin As Double
    Dim hRadius As Double
    
'    HoleDistFromOrigin = parYboltHole / 2
'    OriginPosition = VesselRadius + parBackingPlateThickness
    
    Call objEqpFoundationPort.GetHoles(holes())
    HoleSubAngle = Atn(parXboltHole / (parYboltHole / 2))
    hRadius = Sqr((parYboltHole / 2) ^ 2 + parXboltHole ^ 2)
    
    holes(0, 1) = Cos(parSupportOrientation - HoleSubAngle) * hRadius
    holes(0, 2) = Sin(parSupportOrientation - HoleSubAngle) * hRadius
    holes(1, 1) = Cos(parSupportOrientation + HoleSubAngle) * hRadius
    holes(1, 2) = Sin(parSupportOrientation + HoleSubAngle) * hRadius

    Call objEqpFoundationPort.PutCS(0, _
                                                            0, _
                                                            SupportTopZee - parSupportHeight - parBasePlateThickness, 0, 1, 0, _
                                                            0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)
                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing
    
' Insert your code for output 9( Fundation Port for saddle 2)
    'Dim ObjFoundationPort1 As IJEqpFoundationPort
    
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort2", _
                                                False, m_OutputColl.ResourceManager)
                                                
    holes(0, 1) = Cos(PI / 2 + parSupportOrientation - HoleSubAngle) * hRadius
    holes(0, 2) = Sin(PI / 2 + parSupportOrientation - HoleSubAngle) * hRadius
    holes(1, 1) = Cos(PI / 2 + parSupportOrientation + HoleSubAngle) * hRadius
    holes(1, 2) = Sin(PI / 2 + parSupportOrientation + HoleSubAngle) * hRadius

    Call objEqpFoundationPort.PutCS(0, _
                                                            0, _
                                                            SupportTopZee - parSupportHeight - parBasePlateThickness, 0, 1, 0, _
                                                            0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)
                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing

'Insert your caode for Foundation port 3
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort3", _
                                                False, m_OutputColl.ResourceManager)
    holes(0, 1) = Cos(PI + parSupportOrientation - HoleSubAngle) * hRadius
    holes(0, 2) = Sin(PI + parSupportOrientation - HoleSubAngle) * hRadius
    holes(1, 1) = Cos(PI + parSupportOrientation + HoleSubAngle) * hRadius
    holes(1, 2) = Sin(PI + parSupportOrientation + HoleSubAngle) * hRadius

    Call objEqpFoundationPort.PutCS(0, _
                                                            0, _
                                                            SupportTopZee - parSupportHeight - parBasePlateThickness, 0, 1, 0, _
                                                            0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing
    
'Insert your caode for Foundation port 4
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort4", _
                                                False, m_OutputColl.ResourceManager)
    
    holes(0, 1) = Cos(1.5 * PI + parSupportOrientation - HoleSubAngle) * hRadius
    holes(0, 2) = Sin(1.5 * PI + parSupportOrientation - HoleSubAngle) * hRadius
    holes(1, 1) = Cos(1.5 * PI + parSupportOrientation + HoleSubAngle) * hRadius
    holes(1, 2) = Sin(1.5 * PI + parSupportOrientation + HoleSubAngle) * hRadius

    Call objEqpFoundationPort.PutCS(0, _
                                                            0, _
                                                            SupportTopZee - parSupportHeight - parBasePlateThickness, 0, 1, 0, _
                                                            0, 0, -1)
                                                            
    Call objEqpFoundationPort.SetHoles(holes)

                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing
    
    Set NozzlePHFactory = Nothing
    Set oGeomFactory = Nothing

    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

