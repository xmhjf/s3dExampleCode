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
'   Author:         KKC
'   Creation Date:  Tuesday, Jun 10 2008
'   Description:
'    The symbol is prepared based on 120 deg angle valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-144044  Provide a 120 deg angle valve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

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
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Integer
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim parFace1toFace2 As Double
    Dim parFace2toCenter As Double
    Dim parValveHeight As Double
    Dim parValveDiamter As Double
    Dim ObjBodyCylinder1 As Object
    Dim ObjBodyCylinder2 As Object
    
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toFace2 = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parValveHeight = arrayOfInputs(4)
    parValveDiamter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
 
    iOutput = 0

    'Insert your code for output 1(Cylinder Port 2 of Body)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'Anti-Clock wise angle measurement is taken to be positive
    Dim dValveAng As Double
    dValveAng = 4 * PI / 3
    Dim dPort1XOffset As Double
    Dim dPort2XOffset As Double
    
    dPort1XOffset = parFace1toFace2 - parFace2toCenter * Tan(PI / 6)
    dPort2XOffset = parFace2toCenter * Tan(PI / 6)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set dPort2XOffset, -parFace2toCenter, 0
    
    Set ObjBodyCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder1
    Set ObjBodyCylinder1 = Nothing
    
    'Insert your code for output 2(Cylinder Port 1 of Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    
    oStPoint.Set -dPort1XOffset, 0, 0
    oEnPoint.Set (parValveHeight - dPort1XOffset), 0, 0
    
    Set ObjBodyCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parValveDiamter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder2
    Set ObjBodyCylinder2 = Nothing

    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -dPort1XOffset + depth - sptOffset, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    oPlacePoint.Set dPort2XOffset - (depth - sptOffset) * Sin(PI / 6), _
                            -parFace2toCenter + (depth - sptOffset) * Cos(PI / 6), 0
    oDir.Set Cos(PI / 3), -Sin(PI / 3), 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
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
    
    'Prepare transformation matrix and get operator direction
    Dim oTransMatrix As IJDT4x4
    Dim oVector As AutoMath.DVector

    Set oVector = New AutoMath.DVector
    Set oTransMatrix = New DT4x4

    oTransMatrix.LoadIdentity

    oVector.Set 1, 0, 0
    oTransMatrix.Rotate parHandwheelAngle, oVector
    oVector.Set 0, 0, 1
    oTransMatrix.Rotate -PI / 2, oVector

    oDirX.Set oTransMatrix.IndexValue(0), oTransMatrix.IndexValue(1), oTransMatrix.IndexValue(2)
    oDirY.Set oTransMatrix.IndexValue(4), oTransMatrix.IndexValue(5), oTransMatrix.IndexValue(6)
    oDirZ.Set oTransMatrix.IndexValue(8), oTransMatrix.IndexValue(9), oTransMatrix.IndexValue(10)
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
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
            OpOrigin.Set -dPort1XOffset, 0, 0
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
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    Set OpOrigin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
