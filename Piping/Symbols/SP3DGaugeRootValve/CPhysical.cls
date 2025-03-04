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
'   Author:          MA
'   Creation Date:  Tuesday, May 2007
'   Description:
'    The following Part data basis cases are addressed for the parameters specified:
'    Case A (Part data Basis value -343): (Gauge Root Valve with Single Outlet)
'                                          FacetoFace,ValveWidth and Offset
'    Case B (Part data Basis value -346): (Gauge Root Valve with Multiple Outlet)
'                                 FacetoFace,ValveWidth,Offset,Port3Offset and Port4Offset
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   24.May.2007    MA       CR-113431: Implemented Part data basis for values 343 and 346.
'   25.Sep.2007    MA       CR-121926 Enhanced Gauge Root Valve to support four ports case.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
        
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parFacetoFace As Double
    Dim parValveWidth As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parValveWidth = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
'    parPort3Offset = arrayOfInputs(7)
'    parPort4Offset = arrayOfInputs(8)
    
    iOutput = 0
        
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                   sptOffset, depth

' Insert your code for output 2(Cone)
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim objCone As Object
    
'   Assumption: 30% Face-to-Face is taken to be start point and 5% is taken to be its length.
    oStPoint.Set -0.3 * parFacetoFace, 0, 0
    oEnPoint.Set -0.25 * parFacetoFace, 0, 0

    Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, _
                                                                parValveWidth / 2)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    
    ' Insert your code for output 3(Box)
    Dim objBox As Object
    
    oStPoint.Set -0.25 * parFacetoFace, -0.5 * parValveWidth, 0.5 * parValveWidth
    oEnPoint.Set 0.5 * parFacetoFace, 0.5 * parValveWidth, -0.5 * parValveWidth
    
    Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set objBox = Nothing

' Place Nozzle for output 1 (Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    
    oDir.Set -1, 0, 0
    oPlacePoint.Set -0.5 * parFacetoFace + depth - sptOffset, 0, 0
    
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Dim dNozzleLength As Double
    dNozzleLength = 0.2 * parFacetoFace
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                    oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing

' Place Nozzle for output 4 (Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
    oPlacePoint.Set 0.5 * parFacetoFace - depth + sptOffset, 0, 0
    oDir.Set 1, 0, 0
    
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    
'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    Select Case lPartDataBasis

        Case Is <= 1, 343

        Case 346
            Dim parPort3Offset As Double
            Dim parPort4Offset As Double
            parPort3Offset = arrayOfInputs(7)
            parPort4Offset = arrayOfInputs(8)

            Dim objNozzle3  As GSCADNozzleEntities.IJDNozzle
            Dim objNozzle4  As GSCADNozzleEntities.IJDNozzle

            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            ' Place Nozzle for output 1 (Nozzle3)
            oDir.Set 0, 0, 1
            oPlacePoint.Set (0.5 * parFacetoFace) - parPort3Offset, 0, _
                                                    0.5 * parValveWidth - depth + sptOffset

            Set objNozzle3 = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            m_OutputColl.AddOutput "Nozzle3", objNozzle3
            Set objNozzle3 = Nothing

            ' Place Nozzle for output 1 (Nozzle4)
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            oDir.Set 0, 0, -1
            oPlacePoint.Set (0.5 * parFacetoFace) - parPort4Offset, 0, _
                                -0.5 * parValveWidth + depth - sptOffset

            Set objNozzle4 = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            m_OutputColl.AddOutput "Nozzle4", objNozzle4
            Set objNozzle4 = Nothing

        Case Else
            GoTo ErrorLabel:

    End Select
    
    Set oPlacePoint = Nothing
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
            OpOrigin.Set (parOffset - 0.5 * parFacetoFace), 0.5 * parValveWidth, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
        If Err.Number = E_NOOPERATORPARTFOUND Then    'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub


