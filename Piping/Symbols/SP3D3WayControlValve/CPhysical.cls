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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 9 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC     DI-95670        Replace names with initials in all revision history sheets and symbols
'   21.FEB.2008     PK      CR-135702       Enhanced the existing 3 way control valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()
    ''
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput             As Double
    Dim ObjBodyCone1        As Object
    Dim ObjBodyCone2        As Object
    Dim ObjBodyCone3        As Object
    Dim ObjInsulatedCone1   As Object
    Dim ObjInsulatedPort1   As Object
    Dim ObjInsulatedCone2   As Object
    Dim ObjInsulatedPort2   As Object
    Dim ObjInsulatedCone3   As Object
    Dim ObjInsulatedPort3   As Object
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parFace3toCenter        As Double
    Dim parAngle                As Double
    Dim parInsulationThickness  As Double
    
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double
    
    'Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parHandwheelAngle   As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parFace3toCenter = arrayOfInputs(4)
    parAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    'parFacetoFace = arrayOfInputs(8)
    'parFacetoCenter = arrayOfInputs(9)
    
    iOutput = 0
    
    Dim oPipeComponent      As IJDPipeComponent
    Dim lPartdatabasis      As Long
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartdatabasis
    
        Case Is <= 1, 13                        'default case and Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parFace3toCenter = arrayOfInputs(4)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            If CmpDblEqual(dFace1toCenter, 0) Then
                dFace1toCenter = dFace2toCenter
            ElseIf CmpDblEqual(dFace2toCenter, 0) Then
                dFace2toCenter = dFace1toCenter
            End If
                
        Case 10                                 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(9)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
            
        Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(8)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
             
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    'Insert your code for output 1(Body Cone1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    stPoint.Set CenterPos.x - (dFace1toCenter - flangeThick), CenterPos.y, CenterPos.z
    'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    'value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
    'Insert your code for output 2(Body Cone2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    stPoint.Set CenterPos.x + dFace2toCenter - flangeThick, CenterPos.y, CenterPos.z
    'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    'value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
    'Insert your code for output 3(Body Cone3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    stPoint.Set CenterPos.x + (dFace3toCenter - flangeThick) * Cos(parAngle), _
                CenterPos.y + (dFace3toCenter - flangeThick) * Sin(parAngle), _
                CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Set ObjBodyCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone3
    Set ObjBodyCone3 = Nothing
    
    'Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set CenterPos.x - dFace1toCenter - sptOffset + depth, _
                    CenterPos.y, _
                    CenterPos.z
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set CenterPos.x + dFace2toCenter + sptOffset - depth, _
                    CenterPos.y, _
                    CenterPos.z
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set CenterPos.x + (dFace3toCenter + sptOffset - depth) * Cos(parAngle), _
                    CenterPos.y + (dFace3toCenter + sptOffset - depth) * Sin(parAngle), _
                    CenterPos.z
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Insert your code for output 13 (Valve Operator)
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

    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart   As IJDPart
    Dim oOperatorOcc    As IJPartOcc
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
            Set OpOrigin = Nothing
        End If
    End If
HandleOperatorCodeLabel:
    'The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
       
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
