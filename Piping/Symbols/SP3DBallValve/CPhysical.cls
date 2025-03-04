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
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Tuesday, Jan 8 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav            Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)   Removed the small value (0.0000001) that was added to the
'                                           end point for cones in symbol code (Which was a temporary fix).
'   02.Jan.2006     KKK                 CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)   Modified to accept default value of lPartDataBasis to be
'                                           less than or equal to 1.
'   25.sep.06       KKC                 TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                           Modified symbol code to use Face to Face as a input parameter.
'   16.May.2007     MA                  CR-113431 Implemented Part data basis for values 5,10,13,25,35,349,351 and 353.
'   14.July.2008    RUK                 CR-57547  User would like to see a more detailed symbol represetation on valves
'                                       Reference: "E-128 Ball Valves" of the Design Document
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:"    'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400    'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401    'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private PI As Double

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
    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim sptOffset2 As Double
    Dim flangeDiam2 As Double
    Dim depth2 As Double

    Dim ObjValveBody As Object
    Dim ObjLeftCone As Object
    Dim ObjRightCone As Object
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim objNozzle1 As GSCADNozzleEntities.IJDNozzle
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parSeattoSeat As Double
    Dim parNippleLength As Double
    Dim parValveFacetoFace As Double
    Dim parValveFacetoCenter As Double
    Dim parValCenLineToBot As Double
    Dim parOffsetFrmValCen As Double
    Dim parOperatorHeight As Double
    Dim parHwOffsetFrmValCen As Double
    Dim parValCenToHWTop As Double
    Dim parValCenToHWCen As Double
    Dim parExStemLenFrmValCen As Double

    Dim iOutput As Long

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
    'parFacetoCenter = arrayOfInputs(5)
    'parFace1toCenter = arrayOfInputs(6)
    'parFace2toCenter = arrayOfInputs(7)
    'parSeat1toCenter = arrayOfInputs(8)
    'parSeat2toCenter = arrayOfInputs(9)
    'parSeattoSeat = arrayOfInputs(10)
    'parNippleLength = arrayOfInputs(11)
    'parValveFacetoFace = arrayOfInputs(12)
    'parValveFacetoCenter = arrayOfInputs(13)

    If UBound(arrayOfInputs) > 13 Then
        parValCenLineToBot = arrayOfInputs(14)
        parOffsetFrmValCen = arrayOfInputs(15)
        parOperatorHeight = arrayOfInputs(16)
        parHwOffsetFrmValCen = arrayOfInputs(17)
        parValCenToHWTop = arrayOfInputs(18)
        parValCenToHWCen = arrayOfInputs(19)
        parExStemLenFrmValCen = arrayOfInputs(20)
    End If

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                       sptOffset2, depth2

    Dim dRadiusofCone As Double
    dRadiusofCone = pipeDiam / 2

    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dNozzleLength As Double    'Used in case where nipples are created
    dNozzleLength = flangeThick

    Dim oPlacePointPort1 As AutoMath.DPosition
    Set oPlacePointPort1 = New AutoMath.DPosition
    Dim oPlacePointPort2 As AutoMath.DPosition
    Set oPlacePointPort2 = New AutoMath.DPosition

    'This symbol is based on the following Part data basis values that govern its geometry
        '(Part data Basis value -5): FacetoFace
        '(Part data Basis value -10): FacetoCenter
        '(Part data Basis value -13): Face1toCenter and Face2toCenter
        '(Part data Basis value -25): SeattoSeat
        '(Part data Basis value -35): Seat1toCenter and Seat2toCenter
        '(Part data Basis value -349): NippleLength and ValveFacetoFace
        '(Part data Basis value -351): FacetoFace and ValveFacetoFace
        '(Part data Basis value -353): NippleLength and ValveFacetoCenter

    'Compute the values of dFace1toCenter, dFace2toCenter and port postions for respective PDB values

    Select Case lPartDataBasis

    Case Is <= 1, 5, 1002
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0

    Case 10, 1004
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0

    Case 13, 1006
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0

    Case 25, 1008
        parSeattoSeat = arrayOfInputs(10)

        dFace1toCenter = (parSeattoSeat / 2) + depth - sptOffset
        'Assumption: If depth is not provided, length of the half of the valve body
        'is 1.2 times half of the seat-to-seat dimension.
        If CmpDblLessThanOrEqualTo(depth, 0) Then _
           dFace1toCenter = (parSeattoSeat / 2) * 1.2

        dFace2toCenter = (parSeattoSeat / 2) + depth2 - sptOffset2
        If CmpDblLessThanOrEqualTo(depth2, 0) Then _
           dFace2toCenter = (parSeattoSeat / 2) * 1.2

        If CmpDblGreaterthan(flangeDiam, flangeDiam2) Then
            dRadiusofCone = flangeDiam / 2
        Else
            dRadiusofCone = flangeDiam2 / 2
        End If
        If CmpDblLessThanOrEqualTo(dRadiusofCone, 0) Then dRadiusofCone = pipeDiam / 2

        oPlacePointPort1.Set -parSeattoSeat / 2, 0, 0
        oPlacePointPort2.Set parSeattoSeat / 2, 0, 0

    Case 35, 1010
        parSeat1toCenter = arrayOfInputs(8)
        parSeat2toCenter = arrayOfInputs(9)

        dFace1toCenter = parSeat1toCenter + depth - sptOffset
        'Assumption: If depth is not provided, length of the half of the valve body
        'is 1.2 times half of the seat-to-seat dimension.
        If CmpDblLessThanOrEqualTo(depth, 0) Then _
           dFace1toCenter = parSeat1toCenter * 1.2

        dFace2toCenter = parSeat2toCenter + depth2 - sptOffset2
        If CmpDblLessThanOrEqualTo(depth2, 0) Then _
           dFace2toCenter = parSeat2toCenter * 1.2

        If CmpDblGreaterthan(flangeDiam, flangeDiam2) Then
            dRadiusofCone = flangeDiam / 2
        Else
            dRadiusofCone = flangeDiam2 / 2
        End If

        If CmpDblLessThanOrEqualTo(dRadiusofCone, 0) Then dRadiusofCone = pipeDiam / 2

        oPlacePointPort1.Set -parSeat1toCenter, 0, 0
        oPlacePointPort2.Set parSeat2toCenter, 0, 0

    Case 349
        parNippleLength = arrayOfInputs(11)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = parNippleLength
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick

        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0

    Case 351
        parFacetoFace = arrayOfInputs(2)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = (parFacetoFace - parValveFacetoFace) / 2
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick

        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0

    Case 353
        parNippleLength = arrayOfInputs(11)
        parValveFacetoCenter = arrayOfInputs(13)
        dFace1toCenter = parValveFacetoCenter
        dFace2toCenter = parValveFacetoCenter
        dNozzleLength = parNippleLength
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick

        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0

    Case Else
        GoTo ErrorLabel:

    End Select

    'Create the Body of the Ball valve
    Select Case lPartDataBasis
    Case 1, 5, 10, 13, 25, 35, 349, 351, 353
        'Insert your code for output 1(Cone on Left)
        Dim oBaseCenPoint As AutoMath.DPosition
        Dim oTopCenPoint As AutoMath.DPosition
        Set oBaseCenPoint = New AutoMath.DPosition
        Set oTopCenPoint = New AutoMath.DPosition
        oBaseCenPoint.Set -dFace1toCenter, 0, 0
        oTopCenPoint.Set 0, 0, 0

        Set ObjLeftCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, dRadiusofCone, 0, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
        Set ObjLeftCone = Nothing

        'Insert your code for output 2(Sphere)
        Dim oStartPos As AutoMath.DPosition
        Set oStartPos = New AutoMath.DPosition
        oStartPos.Set 0, 0, 0

        Set ObjValveBody = PlaceSphere(m_OutputColl, oStartPos, (dFace1toCenter + dFace2toCenter) / 8)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        Set oStartPos = Nothing

        'Insert your code for output 3(Cone on Right)
        oBaseCenPoint.Set dFace2toCenter, 0, 0
        oTopCenPoint.Set 0, 0, 0

        Set ObjRightCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, dRadiusofCone, 0, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
        Set ObjRightCone = Nothing
        Set oBaseCenPoint = Nothing
        Set oTopCenPoint = Nothing

    Case 1002, 1004, 1006, 1008, 1010
        Dim dValCenLineToBot As Double
        Dim dBallDiameter As Double
        Dim dMaxDia As Double   'pipe diameter Or flange diameter of Port 1 or Port 2, whcih ever is greater
        Dim dTempAngle As Double
        
        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oCenter As AutoMath.DPosition
        Dim oVector As AutoMath.DVector
        Dim oCompStr As IngrGeom3D.ComplexString3d
        Dim oLineStr As IngrGeom3D.LineString3d
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.Arc3d
        Dim oCollection As Collection
        
        Dim objBody As Object
        
        Set oGeomFact = New GeometryFactory
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oCenter = New DPosition
        Set oVector = New DVector
        Set oCompStr = New ComplexString3d
        Set oLineStr = New LineString3d
        Set oLine = New Line3d
        Set oArc = New Arc3d
        Set oCollection = New Collection
        
        'Get the value of pipe diameter Or flange diameter of Port 1 or Port 2, whcih ever is greater
        dMaxDia = IIf(CmpDblGreaterthan(BodyOD(pipeDiam, flangeDiam), BodyOD(pipeDiam2, flangeDiam2)), _
                        BodyOD(pipeDiam, flangeDiam), BodyOD(pipeDiam2, flangeDiam2))
                                
        'Caliculate "Valve Centerline to Bottom" such that if it is not provided,
        'it will be 30% more than the pipe radius or flange radius which ever is greater
        dValCenLineToBot = IIf(Not CmpDblEqual(parValCenLineToBot, 0), parValCenLineToBot, _
                                1.3 * dMaxDia / 2)
                                
        'Compute Ball diameter such that it should be equal to
        'greater flange diameter in case of flanged ends and
        'in case of other ends it should be more than greater pipe diameter and less than
        '"Valve Center Line to Bottom"
        dBallDiameter = IIf((CmpDblEqual(dMaxDia, pipeDiam) Or CmpDblEqual(dMaxDia, pipeDiam2)), _
                            dMaxDia + 0.5 * (dValCenLineToBot - dMaxDia / 2), dMaxDia)
                
        'Create a complex string and revolve it
        'Line 1
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set oStPoint.x, oStPoint.y + pipeDiam / 2, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        'Line 2
        dTempAngle = Atn(pipeDiam / Sqr(dBallDiameter ^ 2 - pipeDiam ^ 2))
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set -(dBallDiameter / 2) * Cos(dTempAngle), oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        'Arc 1
        dTempAngle = Atn(pipeDiam2 / Sqr(dBallDiameter ^ 2 - pipeDiam2 ^ 2))
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set (dBallDiameter / 2) * Cos(dTempAngle), pipeDiam2 / 2, oStPoint.z
        oCenter.Set 0, 0, 0
        oVector.Set 0, 0, -1
        Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenter, oVector)
        oCollection.Add oArc
        
        'Line 3
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set dFace2toCenter, oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        'Line 4
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, 0, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        oStPoint.Set -dFace1toCenter, 0, 0
        Set oCompStr = PlaceTrCString(oStPoint, oCollection)
        
        oVector.Set 1, 0, 0
        oCenter.Set 0, 0, 0
        Set objBody = oGeomFact.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                        oCompStr, oVector.x, oVector.y, oVector.z, _
                                        oCenter.x, oCenter.y, oCenter.z, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
        
        'Remove the objects from 'Collection' object
        Dim iCount As Long
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        
        'Create the Bottom projection
        Dim objBotProjection As Object
        Dim dProjectionDia As Double
        Dim dPoints() As Double
        
        ReDim dPoints(0 To 8) As Double

        dProjectionDia = 0.15 * (dFace1toCenter + dFace2toCenter)
        dTempAngle = Atn(dProjectionDia / Sqr(dBallDiameter ^ 2 - dProjectionDia ^ 2))
        
        'Point 1
        dPoints(0) = -dProjectionDia / 2
        dPoints(1) = -(dBallDiameter / 2) * Cos(dTempAngle)
        dPoints(2) = 0
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = -dValCenLineToBot
        dPoints(5) = dPoints(2)
        
        'Point 3
        dPoints(6) = 0
        dPoints(7) = dPoints(4)
        dPoints(8) = dPoints(2)
        
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 3, dPoints)
        
        oVector.Set 0, -1, 0
        oCenter.Set 0, 0, 0
        Set objBotProjection = PlaceRevolution(m_OutputColl, oLineStr, oVector, oCenter, _
                                                                2 * PI, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBotProjection
        Set objBotProjection = Nothing
        
        'Create the Top projection
        Dim objTopProjection As Object
        Dim dOffsetFrmValCen As Double
        
        ReDim dPoints(0 To 14) As Double
        
        dProjectionDia = 0.3 * (dFace1toCenter + dFace2toCenter)
        dOffsetFrmValCen = IIf(Not CmpDblEqual(parOffsetFrmValCen, LINEAR_TOLERANCE), _
                            parOffsetFrmValCen, 1.6 * dMaxDia / 2)
        dTempAngle = Atn(dProjectionDia / Sqr(dBallDiameter ^ 2 - dProjectionDia ^ 2))
        
        'Point 1
        dPoints(0) = -0.5 * dProjectionDia
        dPoints(1) = (dBallDiameter / 2) * Cos(dTempAngle)
        dPoints(2) = 0
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.9 * dOffsetFrmValCen
        dPoints(5) = dPoints(2)
        
        'Point 3
        dPoints(6) = -0.7 * dProjectionDia
        dPoints(7) = dPoints(4)
        dPoints(8) = dPoints(2)
        
        'Point 4
        dPoints(9) = dPoints(6)
        dPoints(10) = dOffsetFrmValCen
        dPoints(11) = dPoints(2)
        
        'Point 5
        dPoints(12) = 0
        dPoints(13) = dPoints(10)
        dPoints(14) = dPoints(2)
        
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
        
        oVector.Set 0, -1, 0
        oCenter.Set 0, dBallDiameter / 2, 0
        Set objTopProjection = PlaceRevolution(m_OutputColl, oLineStr, oVector, oCenter, 2 * PI, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopProjection
        Set objTopProjection = Nothing
        
        'Remove the References
        Set oGeomFact = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oCenter = Nothing
        Set oVector = Nothing
        Set oCompStr = Nothing
        Set oLineStr = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oCollection = Nothing
    
    End Select

    'Place Nozzle 1 for Output 4 (Nozzle1)
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                    oPlacePointPort1, dNozzleLength)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePointPort1 = Nothing

    'Place Nozzle 2 for output 5 (Nozzle2)
    oDir.Set 1, 0, 0
    Set objNozzle1 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                        oPlacePointPort2, dNozzleLength)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    Set oPlacePointPort2 = Nothing
    Set oDir = Nothing

    '****************************************************************************************************
                                        'Operator
    '****************************************************************************************************
    'Insert your code for output 7 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc As IJPartOcc
    
    If Not oPipeComponent Is Nothing Then
        'Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
        '1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
        '2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
        'symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, _
                                                OpOrigin, oDirX, oDirY, oDirZ)
            Set OpOrigin = Nothing
            'Update Occurrence Attributes of Operator.
            Dim oIJDInputsArg As IJDInputsArg
            Dim oIJDEditJDArgument As IJDEditJDArgument
            
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
            
            'Update the Occurence Operator Attributes
            'Offset From Valve Centerline
            UpdateOperatorAttributes oOperatorOcc, "IJUAOffsetFrmValCen", "OffsetFrmValCen", _
                                        dOffsetFrmValCen, oIJDEditJDArgument
            
            'Operator Height
            If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight", _
                                            parOperatorHeight, oIJDEditJDArgument
            End If
            
            'Handwheel Offset from Valve Centerline
            If CmpDblGreaterthan(parHwOffsetFrmValCen, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAHwOffsetFrmValCen", "HwOffsetFrmValCen", _
                                            parHwOffsetFrmValCen, oIJDEditJDArgument
            End If
            
            'Valve Centerline to Top of Handwheel
            If CmpDblGreaterthan(parValCenToHWTop, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValCenToHwTop", "ValCenToHWTop", _
                                            parValCenToHWTop, oIJDEditJDArgument
            End If
                            
            'Valve Centerline to Handwheel Centerline
            If CmpDblGreaterthan(parValCenToHWCen, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValCenToHWCen", "ValCenToHWCen", _
                                            parValCenToHWCen, oIJDEditJDArgument
            End If
                            
            'Extension Stem Length from Valve Centerline
            If CmpDblGreaterthan(parExStemLenFrmValCen, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAExStemLenFrmValCen", "ExStemLenFrmValCen", _
                                            parExStemLenFrmValCen, oIJDEditJDArgument
            End If
                
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
            Set oIJDEditJDArgument = Nothing
            Set oIJDInputsArg = Nothing
            
        End If
    End If
    
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    
HandleOperatorCodeLabel:
    '   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                            vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)
    
    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel
    
    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
'    Dim vNewValue As Variant
    Dim lInputIndex As Long
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long
    
    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1
    
    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub
        
    ' Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
'    vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue
    
    ' Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, MODULE & METHOD
   
End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                    oIJDEditJDArgument As IJDEditJDArgument)
                    
    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel
    
    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument
                
    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC
    
    oIJDEditJDArgument.SetArg oArgument
    
    Set oArgument = Nothing
    Set oPC = Nothing
    
    Exit Sub
    
ErrorLabel:
   Err.Raise Err.Number, MODULE & METHOD
        
End Sub

Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg
    
    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument
    
    Exit Function
    
ErrorLabel:
   Err.Raise Err.Number, MODULE & METHOD
End Function

