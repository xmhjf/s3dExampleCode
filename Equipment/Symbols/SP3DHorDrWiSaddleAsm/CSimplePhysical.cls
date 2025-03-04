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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:17 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, May 26 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-162.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   28.Sep.2006    RRK                      TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
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

    'Saddle Outputs
    Dim ObjSaddlePlate As IngrGeom3D.Projection3d
    Dim ObjStiffnerPlate As IngrGeom3D.Projection3d
    Dim ObjBackingPlate1 As Object
    Dim ObjBackingPlate2  As Object
    Dim ObjBackingPlate3 As Object
    Dim ObjBackingPlate4 As Object
    Dim ObjBackingPlate5 As Object
    Dim ObjBackingPlate6 As Object
    Dim ObjFoundationPort As IJEqpFoundationPort

    Dim iOutput     As Double
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan     As Double
    Dim parInsulationThickness As Double

    'Saddle inputs
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation  As Double

    Dim parBasePlateThickness As Double
    Dim parBackingPlateThickness As Double
    Dim parBackingPlateWidth As Double
    Dim parStiffnerThickness As Double
    Dim parStiffnerThickness1 As Double
    Dim parStiffnerSpacing As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '5 for Drum
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parFirstSupportLocation = arrayOfInputs(4)
    parSecondSupportLocation = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    '11 Saddle inputs
    parSupportLength = arrayOfInputs(7)
    parSupportHeight = arrayOfInputs(8)
    parSupportWidth = arrayOfInputs(9)
    parBasePlateThickness = arrayOfInputs(10)
    parBackingPlateThickness = arrayOfInputs(11)
    parBackingPlateWidth = arrayOfInputs(12)
    parStiffnerThickness = arrayOfInputs(13)
    parStiffnerThickness1 = arrayOfInputs(14)
    parStiffnerSpacing = arrayOfInputs(15)
    parXboltHole = arrayOfInputs(16)
    parYboltHole = arrayOfInputs(17)

    
    Dim VesselRadius As Double

    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
    Dim VesselOuterDiameter     As Double
    Dim shellThickness As Double

    shellThickness = 0.0508         'Assumed value for Shell thickness
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness
    VesselRadius = VesselOuterDiameter / 2

    'Insert your code Place Drum Body
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition

    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition

    startPt.Set 0, 0, 0
    endPt.Set parVesselTantoTan, 0, 0

    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPt, endPt, VesselOuterDiameter

    'Insert you code for Head near Control Point
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
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

    MajorX = 0
    MajorY = 0
    MajorZ = VesselOuterDiameter / 2

    mMRatio = DomeMinorRadius / (VesselOuterDiameter / 2)
    StartAngle = PI
    SweepAngle = PI / 2

    norX = 0
    norY = 1
    norZ = 0

    Set oExchangerArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)

    oAxisVect.Set 1, 0, 0
    revCenPt.Set 0, 0, 0
    'Revolve it about X-Axiz

    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oExchangerArc, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing
    Set oExchangerArc = Nothing

    'Insert you code for Head away from Control Point
    cenX = parVesselTantoTan
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = VesselOuterDiameter / 2

    mMRatio = DomeMinorRadius / (VesselOuterDiameter / 2)
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2

    norX = 0
    norY = -1
    norZ = 0

    Set oExchangerArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)

    oAxisVect.Set 1, 0, 0
    revCenPt.Set parVesselTantoTan, 0, 0
    'Revolve it about X-Axiz
    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oExchangerArc, oAxisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing

'Saddle
 ' Insert your code for output 1(SaddlePlate Body just below the Drum along the length of Drum)
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection

    Dim dBackingPlateAngle As Double
    Dim dSupportPlateAngle As Double
    Dim dBackingPlateArcLength As Double
    Dim dLength As Double
    Dim dBackingPlateProjFromSupportPlate As Double  ''Parameter S in the sketch

    dBackingPlateProjFromSupportPlate = 0.14  ''5 1/2 inch

    ''dBackingPlateProjFromSupportPlate is an extra Projection of saddle from the Support plate( S )

    dLength = Sqr(Abs((VesselRadius + parBackingPlateThickness) ^ 2 - (parSupportLength / 2) ^ 2))
    dSupportPlateAngle = 2 * Atn((parSupportLength / 2) / dLength)
    dBackingPlateArcLength = (VesselRadius) * dSupportPlateAngle + 2 * dBackingPlateProjFromSupportPlate
    dBackingPlateAngle = dBackingPlateArcLength / (VesselRadius)

    Dim dlinepts() As Double
    ReDim dlinepts(0 To 17) As Double

    ''SaddlePlate Points
    dlinepts(0) = parFirstSupportLocation + parBackingPlateWidth / 2
    dlinepts(1) = (VesselRadius) * Sin(dBackingPlateAngle / 2)
    dlinepts(2) = -(VesselRadius) * Cos(dBackingPlateAngle / 2)

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = 0
    dlinepts(5) = -(VesselRadius)

    dlinepts(6) = dlinepts(0)
    dlinepts(7) = -dlinepts(1)
    dlinepts(8) = dlinepts(2)

    dlinepts(9) = dlinepts(0)
    dlinepts(10) = -(VesselRadius + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
    dlinepts(11) = -(VesselRadius + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)

    dlinepts(12) = dlinepts(0)
    dlinepts(13) = 0
    dlinepts(14) = -(VesselRadius + parBackingPlateThickness)

    dlinepts(15) = dlinepts(0)
    dlinepts(16) = (VesselRadius + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
    dlinepts(17) = -(VesselRadius + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                            dlinepts(3), dlinepts(4), dlinepts(5), _
                                                            dlinepts(6), dlinepts(7), dlinepts(8))


    oEleCollection.Add oArc
    Set oArc = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(15), dlinepts(16), dlinepts(17))


    oEleCollection.Add oArc
    Set oArc = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                      dlinepts(0), dlinepts(1), dlinepts(2))

    oEleCollection.Add oLine
    Set oLine = Nothing


    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set -1, 0, 0

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjSaddlePlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parBackingPlateWidth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddlePlate
    Set ObjSaddlePlate = Nothing

    Set oComplexStr = Nothing
    Dim count As Integer
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

    Set oEleCollection = Nothing

'Insert your code for Saddle 2 Plate
    ''SaddlePlate Points
    dlinepts(0) = parSecondSupportLocation + parBackingPlateWidth / 2
'    dlinepts(1) = VesselRadius * Sin(dBackingPlateAngle / 2)
'    dlinepts(2) = -VesselRadius * Cos(dBackingPlateAngle / 2)

    dlinepts(3) = dlinepts(0)
'    dlinepts(4) = 0
'    dlinepts(5) = -VesselRadius

    dlinepts(6) = dlinepts(0)
'    dlinepts(7) = -dlinepts(1)
'    dlinepts(8) = dlinepts(2)

    dlinepts(9) = dlinepts(0)
'    dlinepts(10) = -(VesselRadius + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
'    dlinepts(11) = -(VesselRadius + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)

    dlinepts(12) = dlinepts(0)
'    dlinepts(13) = 0
'    dlinepts(14) = -(VesselRadius + parBackingPlateThickness)

    dlinepts(15) = dlinepts(0)
'    dlinepts(16) = (VesselRadius + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
'    dlinepts(17) = -(VesselRadius + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                            dlinepts(3), dlinepts(4), dlinepts(5), _
                                                            dlinepts(6), dlinepts(7), dlinepts(8))

    Set oEleCollection = New Collection
    oEleCollection.Add oArc
    Set oArc = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(15), dlinepts(16), dlinepts(17))


    oEleCollection.Add oArc
    Set oArc = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                      dlinepts(0), dlinepts(1), dlinepts(2))

    oEleCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set -1, 0, 0

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjSaddlePlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parBackingPlateWidth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddlePlate
    Set ObjSaddlePlate = Nothing

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

'Insert your code for output 2(Ends Support Plate in Y-Z plane)
    ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = parFirstSupportLocation + parStiffnerThickness / 2
    dlinepts(1) = (VesselRadius + parBackingPlateThickness) * Sin(dSupportPlateAngle / 2)
    dlinepts(2) = -(VesselRadius + parBackingPlateThickness) * Cos(dSupportPlateAngle / 2)

    'Point 2
    dlinepts(3) = dlinepts(0)
    dlinepts(4) = parSupportLength / 2
    dlinepts(5) = -parSupportHeight + parBasePlateThickness

    'Point 3
    dlinepts(6) = dlinepts(0)
    dlinepts(7) = -dlinepts(4)
    dlinepts(8) = dlinepts(5)

    'Point 4
    dlinepts(9) = dlinepts(0)
    dlinepts(10) = -dlinepts(1)
    dlinepts(11) = dlinepts(2)

     'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = 0
    dlinepts(14) = -(VesselRadius + parBackingPlateThickness)

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set -1, 0, 0

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjStiffnerPlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parStiffnerThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffnerPlate
    Set ObjStiffnerPlate = Nothing

    'Clean up
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    Set oEleCollection = Nothing

'Saddle 2 Stiffener Plate set the output for other Saddle
    dlinepts(0) = parSecondSupportLocation + parStiffnerThickness / 2
'    dlinepts(1) = (VesselRadius + parBackingPlateThickness) * Sin(dSupportPlateAngle / 2)
'    dlinepts(2) = -(VesselRadius + parBackingPlateThickness) * Cos(dSupportPlateAngle / 2)

    dlinepts(3) = dlinepts(0)
'    dlinepts(4) = parSupportLength / 2
'    dlinepts(5) = -parSupportHeight + parBasePlateThickness

    dlinepts(6) = dlinepts(0)
'    dlinepts(7) = -dlinepts(4)
'    dlinepts(8) = dlinepts(5)

    dlinepts(9) = dlinepts(0)
'    dlinepts(10) = -dlinepts(1)
'    dlinepts(11) = dlinepts(2)

    dlinepts(12) = dlinepts(0)
'    dlinepts(13) = 0
'    dlinepts(14) = -(VesselRadius + parBackingPlateThickness)

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    Set oEleCollection = New Collection

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set -1, 0, 0

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjStiffnerPlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parStiffnerThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffnerPlate
    Set ObjStiffnerPlate = Nothing

    'Clean up
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

''Note:The same symbol is used for both 8ft vesseldia and 10ft vesseldia. Number of Support plates depends
      ''on Vessel Dia. So support spacing is calculte based on the Vessel dia.

    Dim dSpacingAngle As Double ''Variable for angle between Spacings

    ' Insert your code for output 3 ( Center Support Plate1 )

       dSpacingAngle = Atn((parStiffnerSpacing / 2) / (Sqr((VesselRadius + parBackingPlateThickness) ^ 2 - (parStiffnerSpacing / 2) ^ 2)))

       oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, -parStiffnerSpacing / 2 + parStiffnerThickness1, _
                                        -parSupportHeight + parBasePlateThickness
       oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, -parStiffnerSpacing / 2, _
                                        -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)
       Set ObjBackingPlate1 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

    ' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate1
        Set ObjBackingPlate1 = Nothing

'Saddle 2 Stiffener Plate set the output for other Saddle
        oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, -parStiffnerSpacing / 2 + parStiffnerThickness1, _
                    -parSupportHeight + parBasePlateThickness
       oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, -parStiffnerSpacing / 2, _
                     -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)
       Set ObjBackingPlate1 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate1
    Set ObjBackingPlate1 = Nothing

' Insert your code for output 4( Middle Support Plate )

        oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, -parStiffnerThickness1 / 2, _
                    -parSupportHeight + parBasePlateThickness

        oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, parStiffnerThickness1 / 2, _
                     -(VesselRadius + parBackingPlateThickness)

        Set ObjBackingPlate2 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

    ' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate2
        Set ObjBackingPlate2 = Nothing

'Saddle 2 Stiffener Plate set the output for other Saddle
        oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, -parStiffnerThickness1 / 2, _
                    -parSupportHeight + parBasePlateThickness

        oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, parStiffnerThickness1 / 2, _
                     -(VesselRadius + parBackingPlateThickness)

        Set ObjBackingPlate2 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate2
    Set ObjBackingPlate2 = Nothing

    ' Insert your code for output 5 ( Center Support Plate2 )

       oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, parStiffnerSpacing / 2 - parStiffnerThickness1, _
                    -parSupportHeight + parBasePlateThickness

       oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, parStiffnerSpacing / 2, _
                     -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

       Set ObjBackingPlate3 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

        ' Set the output
        iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate3
        Set ObjBackingPlate3 = Nothing

        'Saddle 2 Stiffener Plate set the output for other Saddle
        oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, parStiffnerSpacing / 2 - parStiffnerThickness1, _
                    -parSupportHeight + parBasePlateThickness

       oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, parStiffnerSpacing / 2, _
                     -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

       Set ObjBackingPlate3 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate3
        Set ObjBackingPlate3 = Nothing

If (CmpDblGreaterthanOrEqualTo(parVesselDiameter, 0.9) And CmpDblLessThanOrEqualTo(parVesselDiameter, 1.8)) Then
        ''If vessel dia is between 1.2 mts to 1.8 mts then skipping four outputs.
        iOutput = iOutput + 4
Else ''this is for vessel dia more than 1.8mts it has Two additional Stiffner plates(two more outputs)
         ' Insert your code for output 6( Center Support Plate 1 )

            dSpacingAngle = Atn((parStiffnerSpacing) / (Sqr(Abs((VesselRadius + parBackingPlateThickness) ^ 2 - (parStiffnerSpacing) ^ 2))))

            oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, -parStiffnerSpacing - parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, -parStiffnerSpacing, _
                          -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate4 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

         ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate4
            Set ObjBackingPlate4 = Nothing

            'Saddle 2 Stiffener Plate set the output for other Saddle
            oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, -parStiffnerSpacing - parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, -parStiffnerSpacing, _
                          -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate4 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate4
            Set ObjBackingPlate4 = Nothing

         ' Insert your code for output 7( Center Support Plate4 )

            oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, parStiffnerSpacing - parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, parStiffnerSpacing, _
                          -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate5 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

         ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate5
            Set ObjBackingPlate5 = Nothing

            'Saddle 2 Stiffener Plate set the output for other Saddle
            oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, parStiffnerSpacing - parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, parStiffnerSpacing, _
                          -(VesselRadius + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate5 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate5
            Set ObjBackingPlate5 = Nothing

End If
' Insert your code for output 8( Bottom Support Plate )

   oStPoint.Set parFirstSupportLocation + parSupportWidth / 2, parSupportLength / 2, _
                -parSupportHeight

   oEndPoint.Set parFirstSupportLocation - parSupportWidth / 2, -parSupportLength / 2, _
                 -parSupportHeight + parBasePlateThickness

   Set ObjBackingPlate6 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

' Set the output
   iOutput = iOutput + 1
   m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate6
    Set ObjBackingPlate6 = Nothing

'Saddle 2 Stiffener Plate set the output for other Saddle
    oStPoint.Set parSecondSupportLocation + parSupportWidth / 2, parSupportLength / 2, _
                -parSupportHeight

   oEndPoint.Set parSecondSupportLocation - parSupportWidth / 2, -parSupportLength / 2, _
                 -parSupportHeight + parBasePlateThickness

   Set ObjBackingPlate6 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate6
    Set ObjBackingPlate6 = Nothing

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
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort", _
                                                False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = (parSupportLength - parYboltHole) / 2
    holes(0, 2) = parXboltHole
    holes(1, 1) = (parSupportLength + parYboltHole) / 2
    holes(1, 2) = parXboltHole
    holes(2, 1) = (parSupportLength + parYboltHole) / 2
    holes(2, 2) = -parXboltHole
    holes(3, 1) = (parSupportLength - parYboltHole) / 2
    holes(3, 2) = -parXboltHole

    Call objEqpFoundationPort.PutCS(parFirstSupportLocation, -parSupportLength / 2, -parSupportHeight, 0, 1, 0, 0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing

' Insert your code for output 9( Fundation Port for saddle 2)
    Dim ObjFoundationPort1 As IJEqpFoundationPort

    Set ObjFoundationPort1 = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort1", _
                                                False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort1.PutCS(parSecondSupportLocation, -parSupportLength / 2, -parSupportHeight, 0, 1, 0, 0, 0, -1)
    Call ObjFoundationPort1.SetHoles(holes)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort1
    Set ObjFoundationPort1 = Nothing

    Set NozzlePHFactory = Nothing

    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oEleCollection = Nothing
    Set oGeomFactory = Nothing


    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

