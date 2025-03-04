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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Thursday, June 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This valve consists of 7 outputs. This symbol has 2 Inputs FacetoFace and InsulationThickness
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----        ------------------

'   03.Dec.2004     svsmylav            Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   01.Aug.2006     KKK             TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   22.May.2007     MA             CR-113431 Implemented Part data basis for values 5,10,13,25,35,349,351 and 353.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oCenterPos       As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    oCenterPos.Set 0, 0, 0
    
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

    Dim ObjEllipsoid As Object
    Dim ObjRightCone As Object
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Dim ObjLeftCone As Object
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim iOutput     As Double
    Dim parHandwheelAngle   As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parSeattoSeat As Double
    Dim parNippleLength As Double
    Dim parValveFacetoFace As Double
    Dim parValveFacetoCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFacetoCenter = arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
'    parSeat1toCenter = arrayOfInputs(8)
'    parSeat2toCenter = arrayOfInputs(9)
'    parSeattoSeat = arrayOfInputs(10)
'    parNippleLength = arrayOfInputs(11)
'    parValveFacetoFace = arrayOfInputs(12)
'    parValveFacetoCenter = arrayOfInputs(13)

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
    Dim dNozzleLength As Double 'Used in case where nipples are created
    dNozzleLength = flangeThick
    
    Dim oPlacePointPort1 As AutoMath.DPosition
    Set oPlacePointPort1 = New AutoMath.DPosition
    Dim oPlacePointPort2 As AutoMath.DPosition
    Set oPlacePointPort2 = New AutoMath.DPosition
    Dim dEllipseMajorX   As Double

' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): FacetoFace
'   (Part data Basis value -10): FacetoCenter
'   (Part data Basis value -13): Face1toCenter and Face2toCenter
'   (Part data Basis value -25): SeattoSeat
'   (Part data Basis value -35): Seat1toCenter and Seat2toCenter
'   (Part data Basis value -349): NippleLength and ValveFacetoFace
'   (Part data Basis value -351): FacetoFace and ValveFacetoFace
'   (Part data Basis value -353): NippleLength and ValveFacetoCenter
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.25 * parFacetoFace)
    
    Case 10
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.25 * 2 * parFacetoCenter)
    
    Case 13
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        oPlacePointPort1.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.25 * (parFace1toCenter + parFace2toCenter))
    
    Case 25
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
        dEllipseMajorX = 0.6 * (0.25 * parSeattoSeat)
    
    Case 35
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
        dEllipseMajorX = 0.6 * (0.25 * (parSeat1toCenter + parSeat2toCenter))
    
    Case 349
        parNippleLength = arrayOfInputs(11)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = parNippleLength
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
        
        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.125 * (parValveFacetoFace + (2 * parNippleLength)))
    
    Case 351
        parFacetoFace = arrayOfInputs(2)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = (parFacetoFace - parValveFacetoFace) / 2
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.125 * parFacetoFace)
    
    Case 353
        parNippleLength = arrayOfInputs(11)
        parValveFacetoCenter = arrayOfInputs(13)
        dFace1toCenter = parValveFacetoCenter
        dFace2toCenter = parValveFacetoCenter
        dNozzleLength = parNippleLength
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
        oPlacePointPort1.Set -dFace1toCenter - dNozzleLength - sptOffset + depth, 0, 0
        oPlacePointPort2.Set dFace2toCenter + dNozzleLength + sptOffset2 - depth2, 0, 0
        dEllipseMajorX = 0.6 * (0.125 * ((2 * parValveFacetoCenter) + (2 * parNippleLength)))
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
' Insert your code for output 1(Cone on Left)
    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    oBaseCenPoint.Set -dFace1toCenter, 0, 0
    oTopCenPoint.Set 0, 0, 0
    Set ObjLeftCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, dRadiusofCone, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing

' Insert your code for output 2(Ellipsoid)
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxis            As AutoMath.DVector
    Set oAxis = New AutoMath.DVector
    Dim oCentPoint       As AutoMath.DPosition
    Set oCentPoint = New AutoMath.DPosition
    Dim dEllipseMajorY   As Double
    Dim dEllipseMajorZ   As Double
    Dim dEllipseMMRatio  As Double
    
'   Consider Ellipse Major Axis is 0.6 times the larger face to face distance /4 in all the cases
'   below except in the cases where nipplelength is used and
'   Minor axis as 0.4 times the same center distance.
    dEllipseMMRatio = 0.4 / 0.6
    dEllipseMajorY = oCenterPos.y
    dEllipseMajorZ = oCenterPos.z
  
    Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
    Set ObjEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, oCenterPos.x, oCenterPos.y, oCenterPos.z, _
                                            0, 0, 1, dEllipseMajorX, dEllipseMajorY, dEllipseMajorZ, _
                                            dEllipseMMRatio, 0, PI)
    oCentPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oAxis.Set 1, 0, 0
    Set ObjEllipsoid = PlaceRevolution(m_OutputColl, ObjEllipticalArc, oAxis, _
                        oCentPoint, 2 * PI, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipsoid
    Set ObjEllipsoid = Nothing
    Set oCenterPos = Nothing
    Set geomFactory = Nothing
    Set oAxis = Nothing
    Set oCentPoint = Nothing
    
'   remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = ObjEllipticalArc
    Set ObjEllipticalArc = Nothing
    sp3dElem1.Remove

' Insert your code for output 3(Cone on Right)
    oBaseCenPoint.Set dFace2toCenter, 0, 0
    oTopCenPoint.Set 0, 0, 0
    Set ObjRightCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, dRadiusofCone, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing
    
' Place Nozzle 1 for Output 4 (Nozzle1)
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePointPort1, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePointPort1 = Nothing
    
' Place Nozzle 2 for output 5 (Nozzle2)
    oDir.Set 1, 0, 0
    Set objNozzle1 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePointPort2, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    Set oPlacePointPort2 = Nothing
    Set oDir = Nothing
   
' Insert your code for output 7 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    
    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    
'    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
    On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
        If Err.Number = E_NOOPERATORPARTFOUND Then   'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
