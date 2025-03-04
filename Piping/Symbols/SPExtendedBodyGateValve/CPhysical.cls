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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Friday, May 16 2008
'   Description:
'   Source: For PDB value 988: Forged Steel Valves Catalog, Bonney Forge, www.bonneyforge.com
'   For PDB value 13: Vogt Catalog
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.May.2008     MA      CR-141770 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CPhysical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private PI As Double

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
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    
    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFacetoFace = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
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
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim objCyl As Object
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim dLineStringPoints() As Double
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim objBody As Object
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    ' (Part data Basis value -988): Extended body gate valve, specified by face-to-face, face-to-center 1
    ' (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 988
        parFace1toCenter = arrayOfInputs(2)
        parFacetoFace = arrayOfInputs(4)
        
        Dim dWallThick As Double
        RetrievePipeWallThick 1, oPartFclt, dWallThick
        Dim dInnerDia1 As Double
        dInnerDia1 = pipeDiam - (2 * dWallThick)
        
        RetrievePipeWallThick 2, oPartFclt, dWallThick
        Dim dInnerDia2 As Double
        dInnerDia2 = pipeDiam2 - (2 * dWallThick)
       
        'Insert your code for output 1(Body of Valve)
        ReDim dLineStringPoints(23) As Double
        'Point 1
        dLineStringPoints(0) = 0
        dLineStringPoints(1) = dInnerDia1 / 2
        dLineStringPoints(2) = 0
        
        'Point 2
        dLineStringPoints(3) = -dInnerDia2 / 2
        dLineStringPoints(4) = dInnerDia1 / 2
        dLineStringPoints(5) = 0
        
        'Point 3
        dLineStringPoints(6) = -dInnerDia2 / 2
        dLineStringPoints(7) = dInnerDia1 / 2 + 0.1 * parFace1toCenter
        dLineStringPoints(8) = 0
        
        'Point 4
        dLineStringPoints(9) = -0.5 * pipeDiam
        dLineStringPoints(10) = dInnerDia1 / 2 + 0.2 * parFace1toCenter
        dLineStringPoints(11) = 0
        
        'Point 5
        dLineStringPoints(12) = -0.4 * pipeDiam
        dLineStringPoints(13) = dInnerDia1 / 2 + 0.4 * parFace1toCenter
        dLineStringPoints(14) = 0
        
        'Point 6
        dLineStringPoints(15) = -0.4 * pipeDiam
        dLineStringPoints(16) = dInnerDia1 / 2 + parFace1toCenter
        dLineStringPoints(17) = 0
        
        'Point 7
        dLineStringPoints(18) = 0
        dLineStringPoints(19) = dInnerDia1 / 2 + parFace1toCenter
        dLineStringPoints(20) = 0
        
        'Point 8
        dLineStringPoints(21) = 0
        dLineStringPoints(22) = dInnerDia1 / 2
        dLineStringPoints(23) = 0
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStringPoints)
        oAxisVect.Set 0, 1, 0
        oCenterPoint.Set 0.00001, 0, 0
        Set objBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
 
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
        
        'Place Ellipsoid
        Dim dEllipseMajorX   As Double
        Dim dEllipseMajorY   As Double
        Dim dEllipseMajorZ   As Double
        Dim dEllipseMMRatio  As Double
        Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
        Dim objEllipsoid As Object
        
        dEllipseMMRatio = (0.5 * pipeDiam) / (0.5 * (parFacetoFace - parFace1toCenter))
        dEllipseMajorX = 0
        dEllipseMajorY = 0.5 * (parFacetoFace - parFace1toCenter)
        dEllipseMajorZ = 0
    
        Set objEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, 0, dInnerDia1 / 2 + parFace1toCenter, 0, _
                                            0, 0, 1, dEllipseMajorX, dEllipseMajorY, dEllipseMajorZ, _
                                            dEllipseMMRatio, 0, PI)
        oAxisVect.Set 0, 1, 0
        oCenterPoint.Set 0, dInnerDia1 / 2 + parFace1toCenter, 0
        Set objEllipsoid = PlaceRevolution(m_OutputColl, objEllipticalArc, oAxisVect, _
                                            oCenterPoint, 2 * PI, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllipsoid
        Set objEllipsoid = Nothing
        
        Dim sp3dElem1 As IJDObject
        Set sp3dElem1 = objEllipticalArc
        Set objEllipticalArc = Nothing
        sp3dElem1.Remove
        Set sp3dElem1 = Nothing
        
        'Place Cylinder
        oStPoint.Set 0, dInnerDia1 / 2 + parFace1toCenter, 0
        oEnPoint.Set 0, dInnerDia1 / 2 + parFacetoFace, 0
        Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
        Set objCyl = Nothing
        
        'Place Nozzle 1
        oPlacePoint.Set 0, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        oPlacePoint.Set 0, dInnerDia1 / 2 + parFacetoFace + sptOffset2 - depth2, 0
        oDir.Set 0, 1, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
  
    Case 13
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        
        'Insert your code for output 1(Body of Valve)
        ReDim dLineStringPoints(29) As Double
        'Point 1
        dLineStringPoints(0) = -parFace1toCenter
        dLineStringPoints(1) = 0
        dLineStringPoints(2) = 0
        
        'Point 2
        dLineStringPoints(3) = -parFace1toCenter
        dLineStringPoints(4) = pipeDiam / 2
        dLineStringPoints(5) = 0
        
        'Point 3
        dLineStringPoints(6) = -parFace2toCenter - 0.9 * (parFace1toCenter - parFace2toCenter)
        dLineStringPoints(7) = pipeDiam / 2
        dLineStringPoints(8) = 0
        
        'Point 4
        dLineStringPoints(9) = -parFace2toCenter - 0.8 * (parFace1toCenter - parFace2toCenter)
        dLineStringPoints(10) = 0.6 * pipeDiam
        dLineStringPoints(11) = 0
        
        'Point 5
        dLineStringPoints(12) = -parFace2toCenter - 0.5 * (parFace1toCenter - parFace2toCenter)
        dLineStringPoints(13) = 0.6 * pipeDiam
        dLineStringPoints(14) = 0
        
        'Point 6
        dLineStringPoints(15) = -parFace2toCenter - 0.4 * (parFace1toCenter - parFace2toCenter)
        dLineStringPoints(16) = 0.4 * pipeDiam
        dLineStringPoints(17) = 0
        
        'Point 7
        dLineStringPoints(18) = -parFace2toCenter - 0.1 * (parFace1toCenter - parFace2toCenter)
        dLineStringPoints(19) = 0.4 * pipeDiam
        dLineStringPoints(20) = 0
        
        'Point 8
        dLineStringPoints(21) = -parFace2toCenter
        dLineStringPoints(22) = pipeDiam2 / 2
        dLineStringPoints(23) = 0
        
        'Point 9
        dLineStringPoints(24) = -parFace2toCenter
        dLineStringPoints(25) = 0
        dLineStringPoints(26) = 0
        
        'Point 10
        dLineStringPoints(27) = -parFace1toCenter
        dLineStringPoints(28) = 0
        dLineStringPoints(29) = 0
   
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dLineStringPoints)
        oAxisVect.Set -1, 0, 0
        oCenterPoint.Set 0, -0.00001, 0
        Set objBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set objBody = Nothing
        
        'Place Cylinder 1
        oStPoint.Set -parFace2toCenter, 0, 0
        oEnPoint.Set 0, 0, 0
        Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
        Set objCyl = Nothing
        
        'Place Cylinder 2
        oStPoint.Set parFace2toCenter, 0, 0
        oEnPoint.Set 0, 0, 0
        Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
        Set objCyl = Nothing
        
        'Insert your code for output 1(Nozzle1)
        oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
     
        'Insert your code for output 5(Nozzle2)
        oPlacePoint.Set parFace2toCenter + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenterPoint = Nothing
    Set oAxisVect = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oGeomFactory = Nothing

    'Insert your code for output 7 (Valve Operator)
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

    If lPartDataBasis = 1 Or lPartDataBasis = 988 Then
        oDirX.Set 0, Cos(parHandwheelAngle), Sin(parHandwheelAngle)
        oDirY.Set 1, 0, 0
        oDirZ.Set 0, Sin(parHandwheelAngle), -Cos(parHandwheelAngle)
    ElseIf lPartDataBasis = 13 Then
        oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
        oDirY.Set 0, 1, 0
        oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    End If
        
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
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
            If lPartDataBasis = 1 Or lPartDataBasis = 988 Then
                OpOrigin.Set 0, dInnerDia1 / 2 + parFace1toCenter, 0
            ElseIf lPartDataBasis = 13 Then
                OpOrigin.Set 0, 0, 0
            End If
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
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    Set OpOrigin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
