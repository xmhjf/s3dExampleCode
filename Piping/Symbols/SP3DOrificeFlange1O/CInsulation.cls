VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
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
'  29.Jun.2007     svsmylav/dkl  Checked Ubound of 'arrays of inputs' and assigned 'parTapOrientation = arrayOfInputs(5)'. Added 'parFace3toCenter = arrayOfInputs(6)' in the part data basis = 75 'If' block.
'  1.NOV.2007      RRK           CR-123952 Updated the symbol to support part data basis values of 15, 20 and 76
'  20.12.2007      RUK           CR-CP·120874  Enhance orifice flange symbol to support 3, 4, and 5 pressure taps.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
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
    Dim flangeDiam1      As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parFacetoTap As Double
    Dim parInsulationThickness As Double
    Dim parTapOrientation As Long
    Dim parFace3toCenter As Double

    Dim ObjBodyInsulation As Object
    Dim ObjNozz3Insulation As Object
    
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
        
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, cptOffset, depth

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
        parFace3toCenter = flangeDiam1 / 2
    End If

 ' Insert your code for output 3(Orifice Body Insulation)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia1 As Double
    Dim InsulationDia2 As Double
    Dim InsulationDia As Double

    If CmpDblGreaterthan(flangeDiam1, pipeDiam) Then
        InsulationDia1 = flangeDiam1 + parInsulationThickness * 2
    Else
        InsulationDia1 = pipeDiam + parInsulationThickness * 2
    End If
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDia2 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia2 = pipeDiam + parInsulationThickness * 2
    End If
    
    If CmpDblGreaterthan(InsulationDia1, InsulationDia2) Then
        InsulationDia = InsulationDia1
    Else
        InsulationDia = InsulationDia2
    End If
    
    stPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    enPoint.Set parFacetoFace, CenterPos.y, CenterPos.z
    
    Set ObjBodyInsulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                    InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsulation
    Set ObjBodyInsulation = Nothing
      
 ' Insert your code for output 5(Port3 Body Insulation)
    Dim iNozzIndex As Long
    Dim iOutputIndex As Long
    Dim objInsNozz As Object
    iNozzIndex = 3
    iOutputIndex = 3
    
'   Check value of parTapOrientation
    If parTapOrientation <= CLV_SINGLE_TAP Or parTapOrientation = CLV_TAPS_180DEGREES_APART Or _
                                            parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
    
        'Create insulation for tap at zero degrees from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                    flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
        
        InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                pipeDiam, flangeDiam, parFacetoTap)
        
        'For default (single tap) or 'Taps 180 degrees apart' cases
        stPoint.Set parFacetoTap, parFace3toCenter - flangeThick - parInsulationThickness, 0
        enPoint.Set parFacetoTap, parFace3toCenter, 0
        
        'Set the Output
        Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
        Set objInsNozz = Nothing
        iNozzIndex = iNozzIndex + 1
        iOutputIndex = iOutputIndex + 1
    End If
    
'   Create insulation for Port 3 body if it is Resistoflex Orifice flange
    If (lPartDataBasis = 75 Or lPartDataBasis = 76) Then
        stPoint.Set parFacetoTap, parFace3toCenter - flangeThick, 0
        enPoint.Set parFacetoTap, flangeDiam1 / 2, 0
        InsulationDia = pipeDiam + 2 * parInsulationThickness
        
        If CmpDblGreaterthan(InsulationDia, parFacetoTap) Then InsulationDia = 2 * parFacetoTap

        Dim ObjPort3BodyIns As Object
        Set ObjPort3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

        ' Set the output
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", ObjPort3BodyIns
        iOutputIndex = iOutputIndex + 1
        Set ObjPort3BodyIns = Nothing
    End If
    
    If parTapOrientation = CLV_TAPS_90DEGREES_APART Or _
                        parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                        parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        
        'Create insulation for tap at 45 degrees in CW from +ve Y-axis
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                    flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
        
        InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                pipeDiam, flangeDiam, parFacetoTap)
        
        stPoint.Set parFacetoTap, _
            (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 4), _
            (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 4)
        enPoint.Set parFacetoTap, parFace3toCenter * Cos(PI / 4), parFace3toCenter * Sin(PI / 4)
        
        'Set the Output
        Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
        Set objInsNozz = Nothing
        iNozzIndex = iNozzIndex + 1
        iOutputIndex = iOutputIndex + 1
    End If

    If parTapOrientation = CLV_TAPS_180DEGREES_APART Or _
                    parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                    parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                    parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Or _
                    parTapOrientation = CLV_TAPS_FIVE_180_30_30_60_60 Then
        
        'Create insulation for tap at 180 degrees from +ve Y-axis
        
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                    flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
        
        InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                pipeDiam, flangeDiam, parFacetoTap)
        
        stPoint.Set parFacetoTap, _
            -(parFace3toCenter - flangeThick - parInsulationThickness), 0
        enPoint.Set parFacetoTap, -parFace3toCenter, 0
        
        'Set the Output
        Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
        Set objInsNozz = Nothing
        iNozzIndex = iNozzIndex + 1
        iOutputIndex = iOutputIndex + 1
    End If
    
    If parTapOrientation = CLV_TAPS_90DEGREES_APART Or _
                    parTapOrientation = CLV_TAPS_THREE_45_180_45 Or _
                    parTapOrientation = CLV_TAPS_FOUR_0_45_180_45 Or _
                    parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        
        'Create insulation for tap at 45 degrees in CCW from +ve Y-axis
        
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                    flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
        
        InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                pipeDiam, flangeDiam, parFacetoTap)
        
        stPoint.Set parFacetoTap, _
                        (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 4), _
                        -(parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 4)
        enPoint.Set parFacetoTap, _
                        parFace3toCenter * Cos(PI / 4), -parFace3toCenter * Sin(PI / 4)
                        
        'Set the Output
        Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
        Set objInsNozz = Nothing
        iNozzIndex = iNozzIndex + 1
        iOutputIndex = iOutputIndex + 1
    End If

    'Place Tap 5
    If parTapOrientation = CLV_TAPS_FIVE_0_45_180_45_90 Then
        'Create insulation for tap at 90 degrees in CCW from +ve Y-axis
        
        RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                    flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
        
        InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                pipeDiam, flangeDiam, parFacetoTap)
        
        stPoint.Set parFacetoTap, _
                        (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 2), _
                        -(parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 2)
        enPoint.Set parFacetoTap, _
                        parFace3toCenter * Cos(PI / 2), -parFace3toCenter * Sin(PI / 2)
        
        'Set the output
        Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
        m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
        Set objInsNozz = Nothing
        iNozzIndex = iNozzIndex + 1
        iOutputIndex = iOutputIndex + 1
    End If

    'Create the insualtion for taps at parTapOrientation = 30, 30 ,60 ,60
    If parTapOrientation = CLV_TAPS_FIVE_180_30_30_60_60 Then
        Dim iCount As Integer
        Dim dTapAngle As Double
        dTapAngle = PI / 6
        
        For iCount = 1 To 4
            RetrieveParameters iNozzIndex, oPartFclt, m_OutputColl, pipeDiam, _
                                                        flangeThick, flangeDiam, cptOffset, depth
            If CmpDblLessThan(flangeThick, NEGLIGIBLE_THICKNESS) Then flangeThick = NEGLIGIBLE_THICKNESS
            
            InsulationDia = GetInsulationDia(lPartDataBasis, parInsulationThickness, _
                                    pipeDiam, flangeDiam, parFacetoTap)
            
            stPoint.Set parFacetoTap, _
                            (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(dTapAngle), _
                            ((-1) ^ iCount) * (parFace3toCenter - flangeThick - parInsulationThickness) _
                                                                                        * Sin(dTapAngle)
            enPoint.Set parFacetoTap, _
                            parFace3toCenter * Cos(dTapAngle), _
                            ((-1) ^ iCount) * parFace3toCenter * Sin(dTapAngle)
            
            If (iCount Mod 2) = 0 Then
                dTapAngle = PI / 3
            End If
            
            'Set the output
            Set objInsNozz = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
            m_OutputColl.AddOutput "Nozz" & iOutputIndex & "Insulation", objInsNozz
            Set objInsNozz = Nothing
            iNozzIndex = iNozzIndex + 1
            iOutputIndex = iOutputIndex + 1
        Next iCount
    End If

    Set CenterPos = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub

Private Function GetInsulationDia(ByVal lPartDataBasis As Double, ByVal dInsulationThickness As Double, _
                            ByVal dpipeDiam As Double, ByVal dflangeDiam As Double, _
                            ByVal parFacetoTap As Double) As Double
    Const METHOD = "run"
    On Error GoTo ErrorLabel
                            
    Dim dInsulationDia As Double
    
    If CmpDblGreaterthan(dflangeDiam, dpipeDiam) Then
        dInsulationDia = dflangeDiam + dInsulationThickness * 2
    Else
        dInsulationDia = dpipeDiam + dInsulationThickness * 2
    End If

'   Ensure that the InsulationDia does not exceed 'parFacetoTap' otherwise it protrudes
'   into orifice plate
    If (lPartDataBasis <> 75 And lPartDataBasis <> 76) Then
        If CmpDblGreaterthan(dInsulationDia, parFacetoTap) Then dInsulationDia = 2 * parFacetoTap
    End If
    
    GetInsulationDia = dInsulationDia
    
    Exit Function
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
                            
End Function

