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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Jun 9 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is prepared based on PDS symbol MC=TKDR, SN=V62.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.Feb.2008     MA   CR-135752 Implemented Part Data Basis for the value 13.
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
    Dim CenterPos       As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parHandwheelAngle   As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
    
    iOutput = 0

    'Insert your code for output 1(Cone Port1 of Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'This symbol is based on the following Part data basis values that govern its geometry
    '(Part data Basis value -10): Face-to-center dimension basis
    '(Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 10
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
   
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    'Anti-Clock wise angle measurement is taken to be positive
    Dim dValveAng As Double
    dValveAng = 5 * PI / 4

    stPoint.Set CenterPos.x + (dFace1toCenter - flangeThick) * Cos(dValveAng), _
                CenterPos.y + (dFace1toCenter - flangeThick) * Sin(dValveAng), _
                CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
    'Insert your code for output 2(Cone Port2 of Body)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    
    stPoint.Set CenterPos.x + dFace2toCenter - flangeThick, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set CenterPos.x + (dFace1toCenter + sptOffset - depth) * Cos(dValveAng), _
                    CenterPos.y + (dFace1toCenter + sptOffset - depth) * Sin(dValveAng), _
                    CenterPos.z
    oDir.Set Cos(dValveAng), Sin(dValveAng), 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    oPlacePoint.Set CenterPos.x + dFace2toCenter + sptOffset2 - depth2, CenterPos.y, _
                    CenterPos.z
    oDir.Set 1, 0, 0
    
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
    
    oDirX.Set Cos(parHandwheelAngle) * Cos(PI / 4), -Sin(PI / 4), Sin(parHandwheelAngle)
    oDirY.Set Sin(PI / 4), Cos(PI / 4), 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    
    'Prepare transformation matrix and get operator direction
    Dim oTransMatrix As IJDT4x4
    Dim oVector As AutoMath.DVector

    Set oVector = New AutoMath.DVector
    Set oTransMatrix = New DT4x4

    oTransMatrix.LoadIdentity

    oVector.Set Sin(PI / 4), Cos(PI / 4), 0
    oTransMatrix.Rotate parHandwheelAngle, oVector
    oVector.Set 0, 0, 1
    oTransMatrix.Rotate -PI / 4, oVector

    oDirX.Set oTransMatrix.IndexValue(0), oTransMatrix.IndexValue(1), oTransMatrix.IndexValue(2)
    oDirY.Set oTransMatrix.IndexValue(4), oTransMatrix.IndexValue(5), oTransMatrix.IndexValue(6)
    oDirZ.Set oTransMatrix.IndexValue(8), oTransMatrix.IndexValue(9), oTransMatrix.IndexValue(10)

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
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set CenterPos = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
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
