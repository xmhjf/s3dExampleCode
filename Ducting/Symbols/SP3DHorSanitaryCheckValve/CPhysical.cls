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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          VRG
'   Creation Date:  Wednesday, Sep 19 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
PI = 4 * Atn(1)
Set m_oSymGeomHelper = New SymbolServices
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Integer
    Dim parFacetoFace As Double
    Dim parBodyDiameter As Double
    Dim parPortVOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parBodyDiameter = arrayOfInputs(3)
    parPortVOffset = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
 
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
  'Insert your code for RightHalfValve(Output 1)
    Dim oRightHalf As Object
    Dim oAxisVec As AutoMath.DVector
    Dim oCenter As AutoMath.DPosition
    Set oAxisVec = New AutoMath.DVector
    Set oCenter = New AutoMath.DPosition
    Dim dLineStrPoints(0 To 35) As Double
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
    dLineStrPoints(18) = 0.5 * parFacetoFace / 2
    dLineStrPoints(19) = pipeDiam / 2
    dLineStrPoints(20) = 0
    dLineStrPoints(21) = 0.3 * parFacetoFace
    dLineStrPoints(22) = 0.4 * pipeDiam
    dLineStrPoints(23) = 0
    dLineStrPoints(24) = 0.4 * parFacetoFace
    dLineStrPoints(25) = 0.4 * pipeDiam
    dLineStrPoints(26) = 0
    dLineStrPoints(27) = 0.9 * parFacetoFace / 2
    dLineStrPoints(28) = pipeDiam / 2
    dLineStrPoints(29) = 0
    dLineStrPoints(30) = parFacetoFace / 2
    dLineStrPoints(31) = pipeDiam / 2
    dLineStrPoints(32) = 0
    dLineStrPoints(33) = parFacetoFace / 2
    dLineStrPoints(34) = 0
    dLineStrPoints(35) = 0
 
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 12, dLineStrPoints)
    oAxisVec.Set 1, 0, 0
    oCenter.Set 0, 0, 0
    Set oRightHalf = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "RightHalfValve", oRightHalf
    Set oRightHalf = Nothing
    Set oLineString = Nothing
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
 'Insert your code for LeftHalfValve(Output 2)
    Dim oLeftHalf As Object
    Dim dLineStrPoints2(0 To 20) As Double
    
    dLineStrPoints2(0) = -parFacetoFace / 2
    dLineStrPoints2(1) = -parPortVOffset
    dLineStrPoints2(2) = 0
    dLineStrPoints2(3) = -parFacetoFace / 2
    dLineStrPoints2(4) = pipeDiam / 2 - parPortVOffset
    dLineStrPoints2(5) = 0
    dLineStrPoints2(6) = -0.9 * parFacetoFace / 2
    dLineStrPoints2(7) = pipeDiam / 2 - parPortVOffset
    dLineStrPoints2(8) = 0
    dLineStrPoints2(9) = -0.8 * parFacetoFace / 2
    dLineStrPoints2(10) = 0.4 * pipeDiam - parPortVOffset
    dLineStrPoints2(11) = 0
    dLineStrPoints2(12) = -0.7 * parFacetoFace / 2
    dLineStrPoints2(13) = 0.4 * pipeDiam - parPortVOffset
    dLineStrPoints2(14) = 0
    dLineStrPoints2(15) = -0.5 * parFacetoFace / 2
    dLineStrPoints2(16) = 0
    dLineStrPoints2(17) = 0
    dLineStrPoints2(18) = -0.5 * parFacetoFace / 2
    dLineStrPoints2(19) = -parPortVOffset
    dLineStrPoints2(20) = 0
    
 
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints2)
    oAxisVec.Set -1, 0, 0
    oCenter.Set -parFacetoFace / 2, -parPortVOffset, 0
    Set oLeftHalf = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "LeftHalfValve", oLeftHalf
    Set oLeftHalf = Nothing
    Set oLineString = Nothing

   'Insert your code for FlangeCylinder(Output 3)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oFlangeCylinder As Object
    oStPoint.Set -0.1 * parFacetoFace / 2, -0.1 * parBodyDiameter, 0
    oEnPoint.Set -0.5 * parFacetoFace / 2, -0.1 * parBodyDiameter, 0
    Set oFlangeCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBodyDiameter, True)
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "FlangeCylinder", oFlangeCylinder
    Set oFlangeCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing


' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, -parPortVOffset, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
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

    oDirX.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
    oDirY.Set 0, 0, 1
    oDirZ.Set -Sin(parHandwheelAngle), Cos(parHandwheelAngle), 0

  Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    
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
            
            OpOrigin.Set -0.1 * parFacetoFace / 2, 0.4 * parBodyDiameter, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
    
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


