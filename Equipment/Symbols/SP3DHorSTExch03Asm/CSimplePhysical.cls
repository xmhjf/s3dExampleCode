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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:54 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Apr 25, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This is a Shell and Tube Exchanger(E-166).
'   In this Symbol there are total 21 outputs.19 outputs are physical ,out of which 5 are pipe Nozzles.
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   28.Sep.2006     RRK                     TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Option Explicit
    Private PI As Double
    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
    Private Sub Class_Terminate()
        Set m_oSymGeomHelper = Nothing
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

        Dim parVesselDiameter As Double
        Dim parVesselTantoTan As Double
        Dim parVesselLength As Double
        Dim parVesselStartPoint As Double
        Dim parFaceDiameter As Double
        Dim parFaceThickness As Double
        Dim parSupportWidth As Double
        Dim parSupportLength As Double
        Dim parSupportHeight  As Double
        Dim parFirstSupportLocation  As Double
        Dim parSecondSupportLocation As Double
        Dim parBasePlateThickness As Double
        Dim parBackingPlateThickness As Double
        Dim parBackingPlateWidth As Double
        Dim parStiffnerThickness As Double
        Dim parStiffnerThickness1 As Double
        Dim parNozzlePosition As Double
        Dim parNozzlePosition1 As Double
        Dim parNozzlePosition2 As Double
        Dim parNozzlePosition3 As Double
        Dim parNozzlePosition4 As Double
        Dim parRadialPosition As Double
        Dim parRadialPosition1 As Double
        Dim parRadialPosition2 As Double
        Dim parRadialPosition3 As Double
        Dim parRadialPosition4 As Double
        Dim parNozzleOrientation As Double
        Dim parNozzleOrientation1 As Double
        Dim parNozzleOrientation2 As Double
        Dim parNozzleOrientation3 As Double
        Dim parNozzleOrientation4 As Double

        Dim iOutput     As Double

        Dim ObjFrontHeadBonnet As Object
        Dim ObjRearHeadBonnet As Object
        Dim ObjSupport1SaddlePlate As Object
        Dim ObjSupport1StiffnerPlate As Object
        Dim ObjSupport1BackingPlate1 As Object
        Dim ObjSupport1BackingPlate2  As Object
        Dim ObjSupport2SaddlePlate As Object
        Dim ObjSupport2StiffnerPlate As Object
        Dim ObjSupport2BackingPlate1 As Object
        Dim ObjSupport2BackingPlate2  As Object

    ' Inputs
        Set oPartFclt = arrayOfInputs(1)
        parVesselDiameter = arrayOfInputs(2)
        parVesselTantoTan = arrayOfInputs(3)
        parVesselLength = arrayOfInputs(4)
        parVesselStartPoint = arrayOfInputs(5)
        parFaceDiameter = arrayOfInputs(6)
        parFaceThickness = arrayOfInputs(7)
        parSupportWidth = arrayOfInputs(8)
        parSupportLength = arrayOfInputs(9)
        parSupportHeight = arrayOfInputs(10)
        parFirstSupportLocation = arrayOfInputs(11)
        parSecondSupportLocation = arrayOfInputs(12)
        parBasePlateThickness = arrayOfInputs(13)
        parBackingPlateThickness = arrayOfInputs(14)
        parBackingPlateWidth = arrayOfInputs(15)
        parStiffnerThickness = arrayOfInputs(16)
        parStiffnerThickness1 = arrayOfInputs(17)
        parNozzlePosition = arrayOfInputs(18)
        parNozzlePosition1 = arrayOfInputs(19)
        parNozzlePosition2 = arrayOfInputs(20)
        parNozzlePosition3 = arrayOfInputs(21)
        parNozzlePosition4 = arrayOfInputs(22)
        parRadialPosition = arrayOfInputs(23)
        parRadialPosition1 = arrayOfInputs(24)
        parRadialPosition2 = arrayOfInputs(25)
        parRadialPosition3 = arrayOfInputs(26)
        parRadialPosition4 = arrayOfInputs(27)
        parNozzleOrientation = arrayOfInputs(28)
        parNozzleOrientation1 = arrayOfInputs(29)
        parNozzleOrientation2 = arrayOfInputs(30)
        parNozzleOrientation3 = arrayOfInputs(31)
        parNozzleOrientation4 = arrayOfInputs(32)

        m_oSymGeomHelper.OutputCollection = m_OutputColl

        Dim oStPoint As New AutoMath.DPosition
        Dim oEnPoint As New AutoMath.DPosition

        Dim dShellLength As Double
        Dim dDomeHeight As Double

        dShellLength = parVesselTantoTan - parVesselLength
        dDomeHeight = parVesselDiameter / 4

    ' Insert your code for outputs 1(ObjFrontHeadBonnet)
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
        Dim MMRatio As Double
        Dim StartAngle As Double
        Dim SweepAngle As Double
        Dim norX As Double
        Dim norY As Double
        Dim norZ As Double
        Dim Linepts(0 To 5) As Double

        Dim oExchangerArc1 As IngrGeom3D.EllipticalArc3d
        Dim oExchangerArc2 As IngrGeom3D.EllipticalArc3d

        'First Curve
        cenX = -parVesselStartPoint
        cenY = 0
        cenZ = 0

        MajorX = 0
        MajorY = 0
        MajorZ = parVesselDiameter / 2

        MMRatio = dDomeHeight / (parVesselDiameter / 2)
        StartAngle = PI
        SweepAngle = PI / 2

        norX = 0
        norY = 1
        norZ = 0

        Set oExchangerArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                                                StartAngle, SweepAngle)

        oAxisVect.Set 1, 0, 0
        revCenPt.Set 0, 0, 0
        'Revolve it about X-Axiz
        Set ObjFrontHeadBonnet = PlaceRevolution(m_OutputColl, oExchangerArc1, oAxisVect, revCenPt, PI * 2, True)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHeadBonnet
        Set ObjFrontHeadBonnet = Nothing

        Dim Objcurves As IJDObject
        Set Objcurves = oExchangerArc1
        Objcurves.Remove
        Set Objcurves = Nothing

    ' Insert your code for output 2(ObjFrontHeadBodyCyl)As a Cylinder
        oStPoint.Set -parVesselStartPoint, 0, 0
        oEnPoint.Set parVesselLength - parVesselStartPoint - parFaceThickness, 0, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
        iOutput = iOutput + 1
         m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parVesselDiameter

    ' Insert your code for output 3(ObjFrontHeadFlangeTop)As a Cylinder
        oStPoint.Set parVesselLength - parVesselStartPoint - parFaceThickness, 0, 0
        oEnPoint.Set oStPoint.x + parFaceThickness, 0, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter

    ' Insert your code for output 4(ObjFrontHeadFlangeBottom)As a Cylinder
        oStPoint.Set parVesselLength - parVesselStartPoint, 0, 0
        oEnPoint.Set oStPoint.x + parFaceThickness, 0, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter

    ' Insert your code for output 5(ObjShellBodyCyl)As a Cylinder
        oStPoint.Set parVesselLength - parVesselStartPoint + parFaceThickness, 0, 0
        oEnPoint.Set parVesselLength - parVesselStartPoint + dShellLength, 0, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parVesselDiameter

    ' Insert your code for outputs 6(ObjRearHeadBonnet)
        cenX = parVesselLength - parVesselStartPoint + dShellLength
        cenY = 0
        cenZ = 0

        MajorX = 0
        MajorY = 0
        MajorZ = parVesselDiameter / 2

        MMRatio = dDomeHeight / (parVesselDiameter / 2)
        StartAngle = PI / 2
        SweepAngle = PI / 2

        norX = 0
        norY = 1
        norZ = 0

        Set oExchangerArc2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                                                StartAngle, SweepAngle)

        oAxisVect.Set 1, 0, 0
        revCenPt.Set 0, 0, 0

        'Revolve it about X-Axiz
        Set ObjRearHeadBonnet = PlaceRevolution(m_OutputColl, oExchangerArc2, oAxisVect, revCenPt, PI * 2, True)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadBonnet
        Set ObjRearHeadBonnet = Nothing

        Set Objcurves = oExchangerArc2
        Objcurves.Remove

        Set Objcurves = Nothing
        Set revCenPt = Nothing

    'Insert your code for Output 7(ObjSupport1SaddlePlate)
        Dim oComplexSaddleStr As IngrGeom3D.ComplexString3d
        Dim oArc As IngrGeom3D.Arc3d
        Dim oLine As IngrGeom3D.Line3d
        Dim oEleCollection           As Collection
        Dim dlinepts() As Double
        Dim dBackingPlateAngle As Double
        'Assumption
        ' The Total Backing Plate spacing angle is taken as 150 Deg. Hence dBackingPlateAngle is 30/2 deg i.e. PI/12
        dBackingPlateAngle = PI / 12

        ReDim dlinepts(0 To 17) As Double
         ''Support Points
        dlinepts(0) = parFirstSupportLocation - parBackingPlateWidth / 2
        dlinepts(1) = -Cos(dBackingPlateAngle) * (parVesselDiameter / 2)
        dlinepts(2) = -Sin(dBackingPlateAngle) * (parVesselDiameter / 2)

        dlinepts(3) = dlinepts(0)
        dlinepts(4) = -Cos(dBackingPlateAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)
        dlinepts(5) = -Sin(dBackingPlateAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)

        dlinepts(6) = dlinepts(0)
        dlinepts(7) = 0
        dlinepts(8) = -(parVesselDiameter / 2 + parBackingPlateThickness)

        dlinepts(9) = dlinepts(0)
        dlinepts(10) = Cos(dBackingPlateAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)
        dlinepts(11) = -Sin(dBackingPlateAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)

        dlinepts(12) = dlinepts(0)
        dlinepts(13) = Cos(dBackingPlateAngle) * (parVesselDiameter / 2)
        dlinepts(14) = -Sin(dBackingPlateAngle) * (parVesselDiameter / 2)

        dlinepts(15) = dlinepts(0)
        dlinepts(16) = 0
        dlinepts(17) = -parVesselDiameter / 2

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                          dlinepts(3), dlinepts(4), dlinepts(5))

        Set oEleCollection = New Collection
        oEleCollection.Add oLine
        Set oLine = Nothing

        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                                dlinepts(6), dlinepts(7), dlinepts(8), _
                                                                dlinepts(9), dlinepts(10), dlinepts(11))
        oEleCollection.Add oArc
        Set oArc = Nothing

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                          dlinepts(12), dlinepts(13), dlinepts(14))

        oEleCollection.Add oLine
        Set oLine = Nothing

        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                                dlinepts(15), dlinepts(16), dlinepts(17), _
                                                                dlinepts(0), dlinepts(1), dlinepts(2))
        oEleCollection.Add oArc
        Set oArc = Nothing

        oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
        oAxisVect.Set 1, 0, 0

        Set oComplexSaddleStr = PlaceTrCString(oStPoint, oEleCollection)

        Set ObjSupport1SaddlePlate = PlaceProjection(m_OutputColl, oComplexSaddleStr, oAxisVect, parBackingPlateWidth, True)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1SaddlePlate
        Set ObjSupport1SaddlePlate = Nothing

        Dim Count As Integer
        For Count = 1 To oEleCollection.Count
            oEleCollection.Remove 1
        Next Count

    ' Insert your code for output 8(ObjSupport1StiffnerPlate)
        Dim oComplexStiffnerStr As IngrGeom3D.ComplexString3d
        Dim dTurnPoint As Double
        Dim dSupportAngle As Double
        'Assumption
        'The Support Spacing angle is taken as 120 Deg. Hence the dSupportAngle is 60/2 i.e. PI / 6.
        dSupportAngle = PI / 6
        If CmpDblGreaterthanOrEqualTo(parVesselDiameter, parSupportLength) Then
            parSupportLength = parVesselDiameter * (19 / 12)    'We will maintain the Default Ratio of Support Length
        End If                                                                                       'to VesselDiameter
        
        dTurnPoint = Tan(dSupportAngle) * parSupportLength / 2
        ReDim dlinepts(0 To 20) As Double

        ''Support Points
        dlinepts(0) = parFirstSupportLocation - parStiffnerThickness / 2
        dlinepts(1) = -Cos(dSupportAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)
        dlinepts(2) = -Sin(dSupportAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)

        dlinepts(3) = dlinepts(0)
        dlinepts(4) = -parSupportLength / 2
        dlinepts(5) = -dTurnPoint

        dlinepts(6) = dlinepts(0)
        dlinepts(7) = -parSupportLength / 2
        dlinepts(8) = -parSupportHeight + parBasePlateThickness

        dlinepts(9) = dlinepts(0)
        dlinepts(10) = parSupportLength / 2
        dlinepts(11) = dlinepts(8)

        dlinepts(12) = dlinepts(0)
        dlinepts(13) = parSupportLength / 2
        dlinepts(14) = -dTurnPoint

        dlinepts(15) = dlinepts(0)
        dlinepts(16) = Cos(dSupportAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)
        dlinepts(17) = -Sin(dSupportAngle) * (parVesselDiameter / 2 + parBackingPlateThickness)

        dlinepts(18) = dlinepts(0)
        dlinepts(19) = 0
        dlinepts(20) = -(parVesselDiameter / 2 + parBackingPlateThickness)

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

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                          dlinepts(12), dlinepts(13), dlinepts(14))

        oEleCollection.Add oLine
        Set oLine = Nothing

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                          dlinepts(15), dlinepts(16), dlinepts(17))

        oEleCollection.Add oLine
        Set oLine = Nothing


        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                                dlinepts(18), dlinepts(19), dlinepts(20), _
                                                                dlinepts(0), dlinepts(1), dlinepts(2))
        oEleCollection.Add oArc
        Set oArc = Nothing

        oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
        Set oComplexStiffnerStr = PlaceTrCString(oStPoint, oEleCollection)

        oAxisVect.Set 1, 0, 0
        Set ObjSupport1StiffnerPlate = PlaceProjection(m_OutputColl, oComplexStiffnerStr, oAxisVect, parStiffnerThickness, True)
     ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1StiffnerPlate
        Set ObjSupport1StiffnerPlate = Nothing

        For Count = 1 To oEleCollection.Count
            oEleCollection.Remove 1
        Next Count

' Insert your code for output 9( Middle Support Plate )
        oStPoint.Set parFirstSupportLocation - parSupportWidth / 2, -parStiffnerThickness1 / 2, _
                    -parSupportHeight + parBasePlateThickness

        oEnPoint.Set parFirstSupportLocation + parSupportWidth / 2, parStiffnerThickness1 / 2, _
                     -(parVesselDiameter / 2 + parBackingPlateThickness)

        Set ObjSupport1BackingPlate1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1BackingPlate1
       Set ObjSupport1BackingPlate1 = Nothing

' Insert your code for output 10( Bottom Support Plate )
       oStPoint.Set parFirstSupportLocation - parSupportWidth / 2, parSupportLength / 2, -parSupportHeight
       oEnPoint.Set parFirstSupportLocation + parSupportWidth / 2, -parSupportLength / 2, -parSupportHeight + parBasePlateThickness

       Set ObjSupport1BackingPlate2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1BackingPlate2
       Set ObjSupport1BackingPlate2 = Nothing

    ' Insert your code for output 11(ObjSupport2SaddlePlate)
    ' The complex saddle string created for the Support1 is translated along X-Dir to use it for projection for 2nd Support.
        Dim translateMat     As New AutoMath.DT4x4   'For complexstring translation
        Dim xTransVec As New AutoMath.DVector
        xTransVec.Set 1, 0, 0
        xTransVec.Length = parSecondSupportLocation - parFirstSupportLocation                      'C.L. to C.L. of supports = 19 feet
        translateMat.LoadIdentity
        translateMat.Translate xTransVec

        'Transform the complex string to the new location
        oComplexSaddleStr.Transform translateMat

        Set ObjSupport2SaddlePlate = PlaceProjection(m_OutputColl, oComplexSaddleStr, oAxisVect, parBackingPlateWidth, True)
         ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2SaddlePlate
        Set ObjSupport2SaddlePlate = Nothing

        oComplexSaddleStr.RemoveCurve True
        Set oComplexSaddleStr = Nothing

' Insert your code for output 12(ObjSupport2StiffnerPlate)
    ' The complex Stiffner string created for the Support1 is translated along X-Dir to use it for projection for 2nd Support.

         'Transform the complex string to the new location
        oComplexStiffnerStr.Transform translateMat

        Set ObjSupport2StiffnerPlate = PlaceProjection(m_OutputColl, oComplexStiffnerStr, oAxisVect, parStiffnerThickness, True)
         ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2StiffnerPlate
        Set ObjSupport2StiffnerPlate = Nothing

        oComplexStiffnerStr.RemoveCurve True
        Set oComplexStiffnerStr = Nothing
        Set translateMat = Nothing
        Set xTransVec = Nothing

' Insert your code for output 13( Middle Support Plate )
        oStPoint.Set parSecondSupportLocation - parSupportWidth / 2, -parStiffnerThickness1 / 2, _
                    -parSupportHeight + parBasePlateThickness

        oEnPoint.Set parSecondSupportLocation + parSupportWidth / 2, parStiffnerThickness1 / 2, _
                     -(parVesselDiameter / 2 + parBackingPlateThickness)

        Set ObjSupport2BackingPlate1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2BackingPlate1
       Set ObjSupport2BackingPlate1 = Nothing

' Insert your code for output 8( Bottom Support Plate )
       oStPoint.Set parSecondSupportLocation - parSupportWidth / 2, parSupportLength / 2, -parSupportHeight
       oEnPoint.Set parSecondSupportLocation + parSupportWidth / 2, -parSupportLength / 2, -parSupportHeight + parBasePlateThickness

       Set ObjSupport2BackingPlate2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2BackingPlate2
       Set ObjSupport2BackingPlate2 = Nothing


    ' Place Nozzle 1 (A)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        Dim oPlacePoint As AutoMath.DPosition
        Dim oDir        As AutoMath.DVector
        Dim objNozzle   As IJDNozzle
        Dim dNozzleLength As Double
        Dim dMeetPoint As Double

        dMeetPoint = Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2))

        dNozzleLength = parRadialPosition - dMeetPoint

         If CmpDblLessThanOrEqualTo(dNozzleLength, 0.1) Then
            dNozzleLength = flangeThick
        End If
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set parNozzlePosition, _
                                -(parRadialPosition + sptOffset - depth) * Sin(parNozzleOrientation), _
                                (parRadialPosition + sptOffset - depth) * Cos(parNozzleOrientation)
        oDir.Set 0, -Sin(parNozzleOrientation), Cos(parNozzleOrientation)
        Set objNozzle = CreateNozzlePHWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    ' Place Nozzle 2 (B1)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        dMeetPoint = Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2))
        dNozzleLength = parRadialPosition1 - dMeetPoint
        If CmpDblLessThanOrEqualTo(dNozzleLength, 0.1) Then
            dNozzleLength = flangeThick
        End If

       oPlacePoint.Set parNozzlePosition1, _
                                    -(parRadialPosition1 + sptOffset - depth) * Sin(parNozzleOrientation1), _
                                    (parRadialPosition1 + sptOffset - depth) * Cos(parNozzleOrientation1)

        oDir.Set 0, -Sin(parNozzleOrientation1), Cos(parNozzleOrientation1)
        Set objNozzle = CreateNozzlePHWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    ' Place Nozzle 3 (B2)
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        dMeetPoint = Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2))
        dNozzleLength = parRadialPosition2 - dMeetPoint

        If CmpDblLessThanOrEqualTo(dNozzleLength, 0.1) Then
            dNozzleLength = flangeThick
        End If

       oPlacePoint.Set parNozzlePosition2, _
                                -(parRadialPosition2 + sptOffset - depth) * Sin(parNozzleOrientation2), _
                                (parRadialPosition2 + sptOffset - depth) * Cos(parNozzleOrientation2)

        oDir.Set 0, -Sin(parNozzleOrientation2), Cos(parNozzleOrientation2)
        Set objNozzle = CreateNozzlePHWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    ' Place Nozzle 4 (C)
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        dMeetPoint = Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2))
        dNozzleLength = parRadialPosition3 - dMeetPoint

        If CmpDblLessThanOrEqualTo(dNozzleLength, 0.1) Then
            dNozzleLength = flangeThick
        End If

       oPlacePoint.Set parNozzlePosition3, _
                                -(parRadialPosition3 + sptOffset - depth) * Sin(parNozzleOrientation3), _
                                (parRadialPosition3 + sptOffset - depth) * Cos(parNozzleOrientation3)

        oDir.Set 0, -Sin(parNozzleOrientation3), Cos(parNozzleOrientation3)
        Set objNozzle = CreateNozzlePHWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    ' Place Nozzle 5 (D)
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        dMeetPoint = Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2))
        dNozzleLength = parRadialPosition4 - dMeetPoint

        If CmpDblLessThanOrEqualTo(dNozzleLength, 0.1) Then
            dNozzleLength = flangeThick
        End If

        oPlacePoint.Set parNozzlePosition4, _
                                -(parRadialPosition4 + sptOffset - depth) * Sin(parNozzleOrientation4), _
                                (parRadialPosition4 + sptOffset - depth) * Cos(parNozzleOrientation4)

        oDir.Set 0, -Sin(parNozzleOrientation4), Cos(parNozzleOrientation4)

        Set objNozzle = CreateNozzlePHWithLength(5, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        Set oPlacePoint = Nothing
        Set oDir = Nothing

        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxisVect = Nothing
        Set oEleCollection = Nothing
        Set oGeomFactory = Nothing
        
        Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    End Sub
