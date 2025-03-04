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
'   Author:         MS
'   Creation Date:  Wednesday, 17 Mar 2004
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Orifice Flange with one Tap. FacetoTap is the position of the tap along the flow centerline measured from the bolted face of the flange.
'       The Tap position along the Y-axis is the tap female generic depth and offset measured from the flange outer radius.
'       This symbol is prepared based on the existing SP3D1OrificeFlange.C1OrificeFlange symbol, with insulation aspect added.
'       Note: Face3toCenter is to be used with appropriate part data basis value.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'  08.SEP.2006     KKC           DI-95670  Replace names with initials in all revision history sheets and symbols
'  11.Jan.2007     svsmylav      CR-112604 Added two optional inputs (i) TapOrientation and (ii) Face3toCenter
'                                Provided implementation for a) two taps b) Resistoflex orifice flange using part data basis value.
'                                Also, corrected port 1 to be created with length for all cases.
'  29.Jun.2007     svsmylav/dkl  Checked Ubound of 'arrays of inputs' and assigned 'parTapOrientation = arrayOfInputs(5)'. Added 'parFace3toCenter = arrayOfInputs(6)' in the part data basis = 75 'If' block.
'  1.NOV.2007      RRK           CR-123952 Updated the symbol to support part data basis values of 15, 20 and 76
'  20.12.2007      RUK           CR-CP·120874  Enhance orifice flange symbol to support 3, 4, and 5 pressure taps.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const CLV_SINGLE_TAP = 1
Private Const CLV_TAPS_90DEGREES_APART = 5
Private Const CLV_TAPS_180DEGREES_APART = 10
Private Const CLV_TAPS_THREE_45_180_45 = 15         'The angle are w.r.t the +ve Y-Axis (Up).
Private Const CLV_TAPS_FOUR_0_45_180_45 = 20        'The angle are w.r.t the +ve Y-Axis (Up).
Private Const CLV_TAPS_FIVE_0_45_180_45_90 = 25     'The angle are w.r.t the +ve Y-Axis (Up).
Private Const CLV_TAPS_FIVE_180_30_30_60_60 = 30    'The angle are w.r.t the +ve Y-Axis (Up).

Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parFacetoTap As Double
    Dim parInsulationThickness As Double
    Dim parTapOrientation As Long
    Dim parFace3toCenter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFacetoTap = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parTapOrientation = arrayOfInputs(5)
'    parFace3toCenter = arrayOfInputs(6)
    
    If UBound(arrayOfInputs) >= 5 Then
        parTapOrientation = arrayOfInputs(5)    'V7 instances will use this
    Else
        parTapOrientation = 1 'V6.1 instances will use this
    End If
    
    iOutput = 0
    
 'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    Dim blnLengthThrHubInclFaceProj As Boolean
    If lPartDataBasis <= 1 Or lPartDataBasis = 20 Or lPartDataBasis = 75 Then 'Cases where Length-through-hub dimension includes flange face projection
        blnLengthThrHubInclFaceProj = True
    ElseIf lPartDataBasis = 15 Or lPartDataBasis = 76 Then 'Cases where Length-through-hub dimension excludes flange face projection
        blnLengthThrHubInclFaceProj = False
    End If
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj

    If CmpDblEqual(parFace3toCenter, LINEAR_TOLERANCE) And CmpDblEqual(parFacetoTap, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(flangeThick, NEGLIGIBLE_THICKNESS) Then
            parFacetoTap = flangeThick / 2
        Else
            parFacetoTap = parFacetoFace / 2
        End If
        parFace3toCenter = flangeDiam / 2
    End If
    
    If (lPartDataBasis = 75 Or lPartDataBasis = 76) Then 'Consider Face3-to-center and do not consider Face-to-tap
        parFace3toCenter = arrayOfInputs(6)
        If CmpDblGreaterthan(flangeThick, NEGLIGIBLE_THICKNESS) Then
            parFacetoTap = flangeThick / 2
        Else
            parFacetoTap = parFacetoFace / 2
        End If
    Else
        parFace3toCenter = flangeDiam / 2
    End If

' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Dim oPlacePoint       As New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    If CmpDblLessThan(parFacetoFace, flangeThick) Then parFacetoFace = flangeThick

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                oPlacePoint, parFacetoFace)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set CenterPos = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj

    oPlacePoint.Set parFacetoFace - (depth - cptOffset), 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing

' Place 3,4 Taps
    Dim oTapPosition       As New AutoMath.DPosition
    Dim iNozzIndex As Long
    iNozzIndex = 3

'   Check value of parTapOrientation
    If parTapOrientation <= CLV_SINGLE_TAP Or parTapOrientation = CLV_TAPS_180DEGREES_APART Or _
                                        parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        'Create tap at zero degrees from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
        
        oTapPosition.Set parFacetoTap, parFace3toCenter - (depth - cptOffset), 0
        oDir.Set 0, 1, 0
        'Set the output
        Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                    iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
        iNozzIndex = iNozzIndex + 1
    End If
    
    If parTapOrientation = CLV_TAPS_90DEGREES_APART Or _
                        parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        'Create tap at 45 degrees in CW from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
        
        oTapPosition.Set parFacetoTap, (parFace3toCenter - (depth - cptOffset)) * Cos(PI / 4), _
            (parFace3toCenter - (depth - cptOffset)) * Sin(PI / 4)
        oDir.Set 0, Cos(PI / 4), Sin(PI / 4)
        
        'Set the output
        Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                    iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
        iNozzIndex = iNozzIndex + 1
    End If
    
' Place 5,6 Taps
    If parTapOrientation = CLV_TAPS_180DEGREES_APART Or _
                        parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Or _
                        parTapOrientation = CLV_TAPS_FIVE_180_30_30_60_60 Then
        'Create tap at 180 degrees from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
                    
        oTapPosition.Set parFacetoTap, -(parFace3toCenter + cptOffset - depth), 0
        oDir.Set 0, -1, 0
        
        'Set the output
        Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                    iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
        iNozzIndex = iNozzIndex + 1
    End If
    
    If parTapOrientation = CLV_TAPS_90DEGREES_APART Or _
                        parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        'Create tap at 45 degrees in CCW from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
        
        oTapPosition.Set parFacetoTap, (parFace3toCenter - (depth - cptOffset)) * Cos(PI / 4), _
        -(parFace3toCenter - (depth - cptOffset)) * Sin(PI / 4)
        oDir.Set 0, Cos(PI / 4), -Sin(PI / 4)
        
        'Set the output
        Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                    iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
        iNozzIndex = iNozzIndex + 1
    End If
    
    'Place Tap 5
    If parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                            cptOffset, depth, blnLengthThrHubInclFaceProj
        'Create tap at 90 degrees in CCW from +ve Y-axis
        oTapPosition.Set parFacetoTap, (parFace3toCenter - (depth - cptOffset)) * Cos(PI / 2), _
                        -(parFace3toCenter - (depth - cptOffset)) * Sin(PI / 2)
        oDir.Set 0, Cos(PI / 2), -Sin(PI / 2)
        
        'Set the output
        Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                    iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
        iNozzIndex = iNozzIndex + 1
    End If
    
    'Create the taps for parTapOrientation = 30, 30 ,60 ,60
    If parTapOrientation = CLV_TAPS_FIVE_180_30_30_60_60 Then
        Dim iCount As Integer
        Dim dTapAngle As Double
        dTapAngle = PI / 6
        
        For iCount = 1 To 4
            RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                cptOffset, depth, blnLengthThrHubInclFaceProj
            
            oTapPosition.Set parFacetoTap, (parFace3toCenter - (depth - cptOffset)) * Cos(dTapAngle), _
                        ((-1) ^ iCount) * (parFace3toCenter - (depth - cptOffset)) * Sin(dTapAngle)
            oDir.Set 0, Cos(dTapAngle), ((-1) ^ iCount) * Sin(dTapAngle)
            If (iCount Mod 2) = 0 Then
                dTapAngle = PI / 3
            End If
            
            'Set the output
            Call CreNozzAddToOutput(oPartFclt, m_OutputColl, oDir, oTapPosition, _
                                        iNozzIndex, lPartDataBasis, parFace3toCenter, flangeThick)
            iNozzIndex = iNozzIndex + 1
        Next iCount
    End If
    
    Set oDir = Nothing
    Set oTapPosition = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub CreNozzAddToOutput(oPartFclt As PartFacelets.IJDPart, OutputColl As Object, _
                    oDir As IJDVector, oTapPosition As IJDPosition, _
                    iNozzIndex As Long, lPartDataBasis As Long, dNozzLen As Double, _
                    dflangeThick As Double)
    Const METHOD = "CreNozzAddToOutput"
    On Error GoTo ErrorLabel
    
    Dim objNozzle As Object
    If (lPartDataBasis = 75 Or lPartDataBasis = 76) Then 'Consider Face3-to-center
        If CmpDblLessThan(dNozzLen, dflangeThick) Then dNozzLen = dflangeThick
        Set objNozzle = CreateNozzleWithLength(CInt(iNozzIndex), oPartFclt, OutputColl, oDir, _
                                                oTapPosition, dNozzLen)
    Else
        Set objNozzle = CreateNozzle(CInt(iNozzIndex), oPartFclt, OutputColl, oDir, oTapPosition)
    End If
     
    'Set the Output
    OutputColl.AddOutput "PNoz" & iNozzIndex, objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext

End Sub

