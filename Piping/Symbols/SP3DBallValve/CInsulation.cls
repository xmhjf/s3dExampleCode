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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Tuesday, Jan 8 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'Insulation Diameter is greater of Flange and Pipe Dia of both Pipes mating with the Valve
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav            Defined new input parameter HandwheelAngle.
'   02.Jan.2006     KKK                 CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)   Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC                 TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                           Modified symbol code to use Face to Face as a input parameter.
'   16.May.2007     MA                  CR-113431 Implemented Part data basis for values 5,10,13,25,35,349,351 and 353.
'   14.July.2008    RUK                 CR-57547  User would like to see a more detailed symbol represetation on valves
'                                       Reference: "E-128 Ball Valves" of the Design Document
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:"    'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim ObjBallInsulation As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim dInsulationDia As Double
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

    ' Insert your code for output 6(As a Cylinder)
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    Dim oStartInsulationPos As AutoMath.DPosition
    Set oStartInsulationPos = New AutoMath.DPosition
    Dim oEndInsulationPos As AutoMath.DPosition
    Set oEndInsulationPos = New AutoMath.DPosition

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
    dNozzleLength = 0

    Select Case lPartDataBasis

    Case Is <= 1, 5, 1002  'FacetoFace
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

    Case 10, 1004   'FacetoCenter
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter

    Case 13, 1006  'Face1toCenter and Face2toCenter
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter

    Case 25, 1008  'SeattoSeat
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

    Case 35, 1010    'Seat1toCenter and Seat2toCenter
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

    Case 349  'NippleLength and ValveFacetoFace
        parNippleLength = arrayOfInputs(11)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = parNippleLength

    Case 351  'FacetoFace and ValveFacetoFace
        parFacetoFace = arrayOfInputs(2)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = (parFacetoFace - parValveFacetoFace) / 2

    Case 353  'NippleLength and ValveFacetoCenter
        parNippleLength = arrayOfInputs(11)
        parValveFacetoCenter = arrayOfInputs(13)
        dFace1toCenter = parValveFacetoCenter
        dFace2toCenter = parValveFacetoCenter
        dNozzleLength = parNippleLength

    Case Else
        GoTo ErrorLabel:
    End Select

    Select Case lPartDataBasis
    Case 1, 5, 10, 13, 25, 35, 349, 351, 353
        oStartInsulationPos.Set -dFace1toCenter - dNozzleLength, 0, 0
        oEndInsulationPos.Set dFace2toCenter + dNozzleLength, 0, 0

        'In cases 25 and 35, the radius of the cone is greater than the pipe radius.
        'This radius is considered in the computation of insulation diameter below.
        If CmpDblGreaterthan(flangeDiam, dRadiusofCone * 2) Then
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
            dInsulationDia = dInsulationDia1
        Else
            dInsulationDia = dInsulationDia2
        End If
        Set ObjBallInsulation = PlaceCylinder(m_OutputColl, oStartInsulationPos, oEndInsulationPos, _
                                              dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBallInsulation
        Set ObjBallInsulation = Nothing
        Set oStartInsulationPos = Nothing
        Set oEndInsulationPos = Nothing

    Case 1002, 1004, 1006, 1008, 1010
        Dim dValCenLineToBot As Double
        Dim dInsBallDiameter As Double
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
        dInsBallDiameter = IIf((CmpDblEqual(dMaxDia, pipeDiam) Or CmpDblEqual(dMaxDia, pipeDiam2)), _
                            dMaxDia + 0.5 * (dValCenLineToBot - dMaxDia / 2), dMaxDia) _
                            + 2 * parInsulationThickness
                
        'Create a complex string and revolve it
        'Line 1
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set oStPoint.x, oStPoint.y + pipeDiam / 2 + parInsulationThickness, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        'Line 2
        dTempAngle = Atn((pipeDiam / 2 + parInsulationThickness) / _
                        Sqr((dInsBallDiameter / 2) ^ 2 - (pipeDiam / 2 + parInsulationThickness) ^ 2))
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set -(dInsBallDiameter / 2) * Cos(dTempAngle), oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        'Arc 1
        dTempAngle = Atn((pipeDiam2 / 2 + parInsulationThickness) / _
                        Sqr((dInsBallDiameter / 2) ^ 2 - (pipeDiam2 / 2 + parInsulationThickness) ^ 2))
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set (dInsBallDiameter / 2) * Cos(dTempAngle), _
                            pipeDiam2 / 2 + parInsulationThickness, oStPoint.z
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
        Dim dInsProjectionDia As Double
        Dim dPoints() As Double
        
        ReDim dPoints(0 To 8) As Double

        dInsProjectionDia = 0.15 * (dFace1toCenter + dFace2toCenter) + 2 * parInsulationThickness
        dTempAngle = Atn(dInsProjectionDia / Sqr(dInsBallDiameter ^ 2 - dInsProjectionDia ^ 2))
        
        'Point 1
        dPoints(0) = -dInsProjectionDia / 2
        dPoints(1) = -(dInsBallDiameter / 2) * Cos(dTempAngle)
        dPoints(2) = 0
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = -(dValCenLineToBot + parInsulationThickness)
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
        m_OutputColl.AddOutput "InsulatedBottomProjection", objBotProjection
        Set objBotProjection = Nothing
        
        'Create the Top projection
        Dim objTopProjection As Object
        Dim dOffsetFrmValCen As Double
        
        ReDim dPoints(0 To 14) As Double
        
        dInsProjectionDia = 0.3 * (dFace1toCenter + dFace2toCenter) + 2 * parInsulationThickness
        dOffsetFrmValCen = IIf(Not CmpDblEqual(parOffsetFrmValCen, LINEAR_TOLERANCE), _
                            parOffsetFrmValCen, 1.6 * dMaxDia / 2)
        dTempAngle = Atn(dInsProjectionDia / Sqr(dInsBallDiameter ^ 2 - dInsProjectionDia ^ 2))
        
        'Point 1
        dPoints(0) = -0.5 * dInsProjectionDia
        dPoints(1) = (dInsBallDiameter / 2) * Cos(dTempAngle)
        dPoints(2) = 0
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = 0.9 * dOffsetFrmValCen
        dPoints(5) = dPoints(2)
        
        'Point 3
        dPoints(6) = -0.7 * dInsProjectionDia
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
        oCenter.Set 0, dInsBallDiameter / 2, 0
        Set objTopProjection = PlaceRevolution(m_OutputColl, oLineStr, oVector, oCenter, 2 * PI, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedTopProjection", objTopProjection
        Set objTopProjection = Nothing
        
        'Create the Insulation for Ports
        Dim dInsThick1 As Double    'Length of Insulation need apply on Port 1
        Dim dInsThick2 As Double    'Length of Insulation need apply on Port 2
        Dim objInsPort As Object
        
        'Calculate the insulation Diameter on Both sides of valve as per the ports
        dInsulationDia1 = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        dInsulationDia2 = BodyOD(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness
                    
        'Insulation for port 1
        'Calculate the length of insulation for Port 1 and it should not be greater than Face 1 to Center
        dInsThick1 = flangeThick + IIf(CmpDblEqual(flangeThick, 0), _
                                NEGLIGIBLE_THICKNESS, parInsulationThickness)
        If CmpDblGreaterthan(dInsThick1, dFace1toCenter) Then dInsThick1 = dFace1toCenter
        
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -(dFace1toCenter - dInsThick1), oStPoint.y, oStPoint.z
        Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia1, True)
        
        'Set the Output
        m_OutputColl.AddOutput "InsulatedPort1", objInsPort
        Set objInsPort = Nothing
        
        'Insulation for port 2
        'Calculate the length of insulation for Port 2 and it should not be greater than Face 2 to Center
        dInsThick2 = flangeThick2 + IIf(CmpDblEqual(flangeThick2, 0), _
                                NEGLIGIBLE_THICKNESS, parInsulationThickness)
        If CmpDblGreaterthan(dInsThick2, dFace2toCenter) Then dInsThick2 = dFace2toCenter
        
        oStPoint.Set dFace2toCenter, 0, 0
        oEnPoint.Set (dFace2toCenter - dInsThick2), oStPoint.y, oStPoint.z
        Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia2, True)
        
        'Set the Output
        m_OutputColl.AddOutput "InsulatedPort2", objInsPort
        Set objInsPort = Nothing
        
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

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
