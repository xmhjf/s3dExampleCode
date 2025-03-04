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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Monday, Jan 22 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy      who                    change description
'   -----------      -----                  ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   20.Jan.05       svsmylav                TR-59461: Condition "Not PressureRating" is replaced with "PressureRating = 0".
'   08.SEP.2006     KKC                     DI-95670  Replace names with initials in all revision history sheets and symbols
'   15.Jun.2007     KKC                     CR-102105 Added Valve Operator code to support operator on a check valve
'   19.Sep.2007     VRG                     CR-127655 Enhanced the existing check valve to support sanitary vertical and horizontal check valves
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput As Integer
    Dim oBody As Object
    Dim parFacetoFace As Double
    Dim parBodyDiameter As Double
    Dim parPortVOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    'parHandwheelAngle = arrayOfInputs(4) (optional parameter)
    'parBodyDiameter = arrayOfInputs(5) (optional parameter)
    'parPortVOffset = arrayOfInputs(6)  (optional parameter)
    If UBound(arrayOfInputs) >= 4 Then
        parHandwheelAngle = arrayOfInputs(4)
    End If

    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    'This symbol uses the following Part data basis values
    '1.Check valve, horizontal, specified by face-to-face dimension and offset - 455
    '2.Check valve, vertical, specified by face-to-face dimension - 456

    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    If (lPartDataBasis = 455 Or lPartDataBasis = 456) Then
        parBodyDiameter = arrayOfInputs(5)
        parPortVOffset = arrayOfInputs(6)
        If (lPartDataBasis = 456) Then parPortVOffset = 0
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory

        'Insert your code for RightHalfValve(Output 1)
        Dim oRightHalf As Object
        Dim oAxisVec As AutoMath.DVector
        Dim oCenter As AutoMath.DPosition
        Set oAxisVec = New AutoMath.DVector
        Set oCenter = New AutoMath.DPosition
        Dim dLineStrPoints(0 To 23) As Double
        Dim oLineString As IngrGeom3D.LineString3d

        dLineStrPoints(0) = -0.1 * parFacetoFace / 2
        dLineStrPoints(1) = 0
        dLineStrPoints(2) = 0
        dLineStrPoints(3) = -0.1 * parFacetoFace / 2
        dLineStrPoints(4) = 0.4 * parBodyDiameter
        dLineStrPoints(5) = 0
        dLineStrPoints(6) = 0
        dLineStrPoints(7) = 0.3 * parBodyDiameter
        dLineStrPoints(8) = 0
        dLineStrPoints(9) = 0
        dLineStrPoints(10) = 0.4 * parBodyDiameter
        dLineStrPoints(11) = 0
        dLineStrPoints(12) = 0.2 * parFacetoFace
        dLineStrPoints(13) = 0.4 * parBodyDiameter
        dLineStrPoints(14) = 0
        dLineStrPoints(15) = 0.2 * parFacetoFace
        dLineStrPoints(16) = pipeDiam / 2
        dLineStrPoints(17) = 0
        dLineStrPoints(18) = parFacetoFace / 2
        dLineStrPoints(19) = pipeDiam / 2
        dLineStrPoints(20) = 0
        dLineStrPoints(21) = parFacetoFace / 2
        dLineStrPoints(22) = 0
        dLineStrPoints(23) = 0

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStrPoints)
        oAxisVec.Set 1, 0, 0
        oCenter.Set 0, 0, 0
        Set oRightHalf = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRightHalf
        Set oRightHalf = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
        Set oCenter = Nothing

        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        
        'Insert your code for FlangeCylinder(Output)
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set -0.1 * parFacetoFace / 2, -0.1 * parBodyDiameter, 0
        oEnPoint.Set -0.5 * parFacetoFace / 2, -0.1 * parBodyDiameter, 0
        Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBodyDiameter, True)
        'Set the output
        m_OutputColl.AddOutput "Body_", oBody
        Set oBody = Nothing
        Dim oNormal As AutoMath.DVector
        Set oCenter = New AutoMath.DPosition
        Set oNormal = New AutoMath.DVector
        Dim oCircle1 As IngrGeom3D.Circle3d
        Dim oCircle2 As IngrGeom3D.Circle3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        oCenter.Set -0.7 * parFacetoFace / 2, -parPortVOffset, 0
        oNormal.Set -1, 0, 0
        Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                               oNormal.x, oNormal.y, oNormal.z, pipeDiam / 2)
                                               
      
        oCenter.Set -0.5 * parFacetoFace / 2, -0.1 * parBodyDiameter, 0
        oNormal.Set -1, 0, 0
        Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                               oNormal.x, oNormal.y, oNormal.z, 0.5 * parBodyDiameter)
                                               
                                               
        'code for making transition element between two circular planes
        Dim oTransition As Object
        Set oTransition = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                        oCircle1, oCircle2, True)
    
        'Set the output
        m_OutputColl.AddOutput "Body_", oTransition
        Set oTransition = Nothing
        Set oGeomFactory = Nothing
        Set oCenter = Nothing
        Set oNormal = Nothing
        Set oCircle1 = Nothing
        Set oCircle2 = Nothing
        'Insert your code for LeftHalfValveBody(Output)
        oStPoint.Set -parFacetoFace / 2, -parPortVOffset, 0
        oEnPoint.Set -0.7 * parFacetoFace / 2, -parPortVOffset, 0
        Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        'Set the output
        m_OutputColl.AddOutput "Body_", oBody
        Set oBody = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    ElseIf (lPartDataBasis <= 1) Then
        'Insert your code for output 1(Check Valve Body)
        Dim PressureRating As Long
        RetrieveParameterswithPressureRating 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth, PressureRating
        Dim BodyRadius As Double
        Dim flangeThick2 As Double
        BodyRadius = pipeDiam / 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, cptOffset, depth

        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
        'If Not PressureRating And flangeThick = 0 Then
        If PressureRating = 0 And flangeThick = 0 Then
            Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        Else
            Set oBody = PlaceCone(m_OutputColl, oStPoint, oEnPoint, BodyRadius, 0)
        End If
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    End If
    Set oBody = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, 0, 0
    If (lPartDataBasis = 455) Then oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, -parPortVOffset, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Insert your code for output 5 (Valve Operator)
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

    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
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
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
    'The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oOperatorOcc = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
