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
'   Author:          svsmylav
'   Creation Date:  Thursday, Jan 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   26.Jul.2006     dkl                     CR-28369 Added implementation of PartDataBasis 5
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  04.05.2007      KKC         CR-117167  Implemented Part data basis for 290, 295 and 300
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Dim PI           As Double
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjValveBodyBox As Object
    Dim ObjInsulatedCyl As Object
    Dim ObjInsulatedBox As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parValveHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parFrameHeight As Double
    Dim parFrameWidth As Double
    Dim parFrameDepth As Double
    Dim parValveWidth As Double
    Dim dFrameDepth1 As Double
    Dim dFrameDepth2 As Double
    
'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)    'Set only if PartDataBasis=1 or 10
'    parFace2toCenter = arrayOfInputs(3)    'Set only if PartDataBasis=1 or 10
    parValveHeight = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
'    parFacetoFace = arrayOfInputs(7)       'Set only if PartDataBasis=5
'    parFrameHeight = arrayOfInputs(8)
'    parFrameWidth = arrayOfInputs(9)
'    parFrameDepth = arrayOfInputs(10)
'    parValveWidth = arrayOfInputs(11)

    iOutput = 0
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Select Case lPartDataBasis
    Case 5, 295  'Applicable for 5 "Face-to-Face dimension basis" and Knife Gate Valve Type 2
        parFacetoFace = arrayOfInputs(7)
        parFace1toCenter = parFacetoFace / 2  'parFace1toCenter and parFace2toCenter: These dimensions are calculated and used in computation.
        parFace2toCenter = parFacetoFace / 2
        
        If UBound(arrayOfInputs) >= 10 Then
            parFrameDepth = arrayOfInputs(10)
        End If
        
        If CmpDblEqual(parFrameDepth, 0) Then
           dFrameDepth1 = parFacetoFace / 2
           dFrameDepth2 = parFacetoFace / 2
        Else
           dFrameDepth1 = parFrameDepth / 2
           dFrameDepth2 = parFrameDepth / 2
        End If
           
    Case Is <= 1, 10  'Applicable for  1 "Default" & 10 -"Face-to-center dimension basis"
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        
        If UBound(arrayOfInputs) >= 10 Then
            parFrameDepth = arrayOfInputs(10)
        End If
        
        If CmpDblEqual(parFrameDepth, 0) Then
           dFrameDepth1 = parFace1toCenter
           dFrameDepth2 = parFace2toCenter
        Else
           dFrameDepth1 = parFrameDepth / 2
           dFrameDepth2 = parFrameDepth / 2
        End If
    Case 290 'Knife Gate Valve Type 1
        parFacetoFace = arrayOfInputs(7)
        parFrameHeight = arrayOfInputs(8)
        parFrameWidth = arrayOfInputs(9)
        parFrameDepth = arrayOfInputs(10)
        parValveWidth = arrayOfInputs(11)
        If CmpDblEqual(parFrameHeight, 0) Then parFrameHeight = parValveHeight * 0.6
        If CmpDblEqual(parFrameWidth, 0) Then parFrameWidth = parValveWidth / 2
        If CmpDblEqual(parFrameDepth, 0) Then parFrameDepth = parFacetoFace / 2
    Case 300 'Knife Gate Valve Type 4
        parFacetoFace = arrayOfInputs(7)
        parFrameWidth = arrayOfInputs(9)
        parFrameDepth = arrayOfInputs(10)
        If CmpDblEqual(parFrameWidth, 0) Then parFrameWidth = parValveWidth * 0.6
        If CmpDblEqual(parFrameDepth, 0) Then parFrameDepth = parFacetoFace / 2
    Case Else
        GoTo ErrorLabel
    End Select
    
    
' Insert your code for output 1(Valve Body Box)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If (lPartDataBasis <= 1) Or (lPartDataBasis = 5) Or (lPartDataBasis = 295) Or (lPartDataBasis = 10) Then
        Dim dValWidth As Double
        If CmpDblEqual(flangeDiam, 0) Then
            dValWidth = pipeDiam
        Else
            dValWidth = flangeDiam
        End If
        stPoint.Set CenterPos.x - dFrameDepth1 + flangeThick, _
                    CenterPos.y, _
                    CenterPos.z - dValWidth / 2
        enPoint.Set CenterPos.x + dFrameDepth2 - flangeThick, _
                    CenterPos.y + parValveHeight, _
                    CenterPos.z + dValWidth / 2
        Set ObjValveBodyBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyBox
        Set ObjValveBodyBox = Nothing
    
    ElseIf (lPartDataBasis = 290) Then
        stPoint.Set CenterPos.x - parFacetoFace / 3 + flangeThick, _
                    CenterPos.y - parValveHeight / 2, _
                    CenterPos.z - parValveWidth / 2
        enPoint.Set CenterPos.x + parFacetoFace / 3 - flangeThick, _
                    CenterPos.y + parValveHeight / 2, _
                    CenterPos.z + parValveWidth / 2
        Set ObjValveBodyBox = PlaceBox(m_OutputColl, stPoint, enPoint)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyBox
        Set ObjValveBodyBox = Nothing
    
    ' Insert your code for output 2(Frame for Operator support)
        Dim objFrame As Object
        Dim oBaseCenter As AutoMath.DPosition
        Set oBaseCenter = New AutoMath.DPosition
        oBaseCenter.Set 0, parValveHeight / 2, 0
        
        Set objFrame = PlaceTrapezoid(m_OutputColl, oBaseCenter, parValveWidth, _
                           parFrameDepth, parFrameWidth, parFrameDepth, _
                                        parFrameHeight, True, -PI / 2, 0, PI / 2)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
        Set objFrame = Nothing
    ElseIf lPartDataBasis = 300 Then
        stPoint.Set CenterPos.x - parFrameDepth / 2 + flangeThick, _
                    CenterPos.y, _
                    CenterPos.z - parFrameWidth / 2
        enPoint.Set CenterPos.x + parFrameDepth / 2 - flangeThick, _
                    CenterPos.y + parValveHeight, _
                    CenterPos.z + parFrameWidth / 2
        Set ObjValveBodyBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyBox
        Set ObjValveBodyBox = Nothing
    End If
    
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    If lPartDataBasis = 5 Or lPartDataBasis = 290 Or _
                    lPartDataBasis = 295 Or lPartDataBasis = 300 Then
        parFace1toCenter = parFacetoFace / 2
        parFace2toCenter = parFacetoFace / 2
    End If

    oPlacePoint.Set CenterPos.x - parFace1toCenter - sptOffset + depth, _
                    CenterPos.y, _
                    CenterPos.z
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace1toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set CenterPos.x + parFace2toCenter + sptOffset - depth, _
                    CenterPos.y, _
                    CenterPos.z
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleWithLength( _
                        2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace2toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 6 (Valve Operator)
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

'    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt

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

