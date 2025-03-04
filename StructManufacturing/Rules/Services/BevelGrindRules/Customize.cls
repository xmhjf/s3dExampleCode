VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''***********************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : StrMfgBevelGrind
''  File        : Customize.cls
''  Description : This class has the functionality to Start the commnad,listening to events of Ribbonbars
''  Author      : Seth Eden Hollingsead
''  History     :
''                28Jan2011     Seth Eden Hollingsead
''                                  - Creation
''***********************************************************************************
Option Explicit

Implements IJMfgBevelGrind
Private m_oUnitsofMeasure As IJUomVBInterface

Private Const MODULE = "Customize"

Private Sub Class_Initialize()
    Set m_oUnitsofMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
End Sub

Private Sub Class_Terminate()
    Set m_oUnitsofMeasure = Nothing
End Sub

Private Sub SetBevelParams(ByRef pParams As GSCADStructMfgGlobals.MfgBevelParams, _
                            Optional dAngle1_M As Double = 0, _
                            Optional dAngle1_UM As Double = 0, _
                            Optional dAngle2_M As Double = 0, _
                            Optional dAngle2_UM As Double = 0, _
                            Optional dAttachmentAngle As Double = 0, _
                            Optional dChamferAngle_M As Double = 0, _
                            Optional dChamferAngle_UM As Double = 0, _
                            Optional dChamferDepth_M As Double = 0, _
                            Optional dChamferDepth_UM As Double = 0, _
                            Optional dDepth1_M As Double = 0, _
                            Optional dDepth1_UM As Double = 0, _
                            Optional dDepth2_M As Double = 0, _
                            Optional dDepth2_UM As Double = 0, _
                            Optional eeBevelType As tagStrMfgBevelType = BEVEL_FET, _
                            Optional eeGrindShape As tagStrMfgGrindShape = FLAT, _
                            Optional lMarkingSide As Long = 0, _
                            Optional lMethod1_M As Long = 0, _
                            Optional lMethod1_UM As Long = 0, _
                            Optional lMethod2_M As Long = 0, _
                            Optional lMethod2_UM As Long = 0, _
                            Optional dNose As Double = 0, _
                            Optional dNoseAngle As Double = 0, _
                            Optional lNoseMethod As Long = 0, Optional dRadius As Double = 0, _
                            Optional dRootGap As Double = 0, Optional lSymmetrical As Long = 0)
    
    pParams.Angle1_M = dAngle1_M
    pParams.Angle1_UM = dAngle1_UM
    pParams.Angle2_M = dAngle2_M
    pParams.Angle2_UM = dAngle2_UM
    pParams.AttachmentAngle = dAttachmentAngle
    pParams.ChamferAngle_M = dChamferAngle_M
    pParams.ChamferAngle_UM = dChamferAngle_UM
    pParams.ChamferDepth_M = dChamferDepth_M
    pParams.ChamferDepth_UM = dChamferDepth_UM
    pParams.Depth1_M = dDepth1_M
    pParams.Depth1_UM = dDepth1_UM
    pParams.Depth2_M = dDepth2_M
    pParams.Depth2_UM = dDepth2_UM
    pParams.eBevelType = eeBevelType
    pParams.eGrindShape = eeGrindShape
    pParams.MarkingSide = lMarkingSide
    pParams.Method1_M = lMethod1_M
    pParams.Method1_UM = lMethod1_UM
    pParams.Method2_M = lMethod2_M
    pParams.Method2_UM = lMethod2_UM
    pParams.Nose = dNose
    pParams.NoseAngle = dNoseAngle
    pParams.NoseMethod = lNoseMethod
    pParams.Radius = dRadius
    pParams.RootGap = dRootGap
    pParams.Symmetrical = lSymmetrical
End Sub

Private Sub IJMfgBevelGrind_GetBevelAttributes(ByVal oPart As Object, ByVal PhysConn As Object, ByVal thickness As Double, ByVal BevelType As String, ByVal BevelSubType As String, pParams As GSCADStructMfgGlobals.MfgBevelParams)
'    Dim oPhysicalConnection As IJStructPhysicalConnection
    Dim eWeldType As SmartClassSubType
    Dim mmThickness As Double
    'Dim bUsePCChamfer As Boolean
    Dim vPCChamferValue As Variant
    Dim ResultType As EResultType
    
'    Set oPhysicalConnection = PhysConn
    'eWeldType = oPhysicalConnection.ConnectionSubType
    
    mmThickness = thickness
    'bUsePCChamfer = True
    If pParams.ChamferDepth_M > 0 Then
        mmThickness = mmThickness - pParams.ChamferDepth_M
        'bUsePCChamfer = False
    End If
    If pParams.ChamferDepth_UM > 0 Then
        mmThickness = mmThickness - pParams.ChamferDepth_UM
        'bUsePCChamfer = False
    End If
    
    Select Case BevelType
        Case "ButtI"
            pParams.Nose = mmThickness
        Case "ButtK"
            Select Case BevelSubType
                Case "auto"
                    If mmThickness > 20 And mmThickness <= 28 Then
                        SetBevelParams pParams, 30, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, (mmThickness - 8) / 2, _
                            (mmThickness - 8) / 2, , , , , , , , , , 8
                    ElseIf mmThickness > 28 And mmThickness <= 31 Then
                        SetBevelParams pParams, 30, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 18, _
                            10, , , , , , , , , , 8
                    Else
                        SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM
                    End If
                Case "K-BevelsEqual"
                    SetBevelParams pParams, 30, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, (mmThickness - 8) / 2, _
                        (mmThickness - 8) / 2, , , , , , , , , , 8
                Case "K-AntiRefBevel10mm"
                    SetBevelParams pParams, 30, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 18, _
                        10, , , , , , , , , , 8
            End Select
   
        Case "ButtV"
            Select Case BevelSubType
                Case "auto"
                    If mmThickness >= 8 And mmThickness < 10 Then
                        SetBevelParams pParams, 20, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness, , , , , , , , , , , , , , , 2
                    ElseIf mmThickness >= 10 And mmThickness < 20 Then
                        SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness, , , , , , , , , , , , , , , 3
                    ElseIf mmThickness > 20 And mmThickness <= 25 Then
                        SetBevelParams pParams, 20, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness, , , , , , , , , , , , , , , 3
                    Else
                        SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, pParams.ChamferDepth_UM
                    End If
                Case "V20-RG_2mm"
                    SetBevelParams pParams, 20, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness, , , , , , , , , , , , , , , 2
                Case "V25-RG_3mm"
                    SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness, , , , , , , , , , , , , , , 3
                Case "V20-RG_3mm"
                    SetBevelParams pParams, 20, , , , , , , , , _
                        mmThickness, , , , , , , , , , , , , , , 3
            End Select
            
        Case "ButtX"
            If mmThickness > 6 Then
                SetBevelParams pParams, 25, 30, , , , pParams.ChamferAngle_M, _
                    pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                    pParams.ChamferDepth_UM, mmThickness * 2 / 3, mmThickness * 1 / 3
            Else
                SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                    pParams.ChamferAngle_UM, pParams.ChamferDepth_M, pParams.ChamferDepth_UM
            End If
                    
        Case "ButtY"
            Select Case BevelSubType
                Case "auto"
                    If mmThickness >= 8 And mmThickness <= 15 Then
                        SetBevelParams pParams, 30, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 2, , , , , , , , , , , 2
                    ElseIf mmThickness > 15 And mmThickness <= 23 Then
                        SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 3, , , , , , , , , , , 3
                    ElseIf mmThickness > 23 And mmThickness <= 25 Then
                        SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 4, , , , , , , , , , , 4
                    ElseIf mmThickness > 25 And mmThickness <= 30 Then
                        SetBevelParams pParams, 22.5, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 5, , , , , , , , , , , 5
                    ElseIf mmThickness > 30 Then
                        SetBevelParams pParams, 22.5, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, mmThickness - 6, , , , , , , , , , , 6
                    Else
                        SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, pParams.ChamferDepth_UM
                    End If
                Case "Y20-Nose2"
                    SetBevelParams pParams, 30, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 2, , , , , , , , , , , 2
                Case "Y25-Nose3"
                    SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 3, , , , , , , , , , , 3
                Case "Y25-Nose4"
                    SetBevelParams pParams, 25, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 4, , , , , , , , , , , 4
                Case "Y22.5-Nose5"
                    SetBevelParams pParams, 22.5, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 5, , , , , , , , , , , 5
                Case "Y22.5-Nose6"
                    SetBevelParams pParams, 22.5, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, mmThickness - 6, , , , , , , , , , , 6
            End Select
        Case "TeeFillet1"
        Case "TeeFillet2"
        Case "TeeY"
            Select Case BevelSubType
                Case "auto"
                    If thickness <= 25 Then
                        SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness - 3, , , , , , , , , , , 3
                    ElseIf thickness > 25 And thickness <= 38 Then
                        SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness - 5, , , , , , , , , , , 5
                    ElseIf thickness > 38 Then
                        SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, pParams.ChamferDepth_UM, thickness / 2, , , , , , , , , , , thickness / 2
                    End If
                Case "Y-Nose3"
                    SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness - 3, , , , , , , , , , , 3
                Case "Y-Nose5"
                    SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness - 5, , , , , , , , , , , 5
                Case "Y-NoseThick/2"
                    SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness / 2, , , , , , , , , , , thickness / 2
            End Select
        Case "TeeK"
            Select Case BevelSubType
                Case "auto"
                    If thickness <= 25 Then
                        SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, (thickness - 3) / 2, (thickness - 3) / 2, _
                            , , , , , , , , , 3
                    Else
                        SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness / 3, thickness / 3, _
                            , , , , , , , , , thickness / 3
                    End If
                Case "K-Nose3"
                    SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, (thickness - 3) / 2, _
                        (thickness - 3) / 2, , , , , , , , , , 3
                Case "K-NoseThick/3"
                    SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness / 3, thickness / 3, _
                        , , , , , , , , , thickness / 3
            End Select
        Case "TeeV"
            Select Case BevelSubType
                Case "auto"
                    If thickness <= 25 Then
                        SetBevelParams pParams, 45, , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness
                    Else
                        SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM
                    End If
            End Select
        Case "TeeX"
            Select Case BevelSubType
                Case "auto"
                    If thickness > 25 And thickness <= 38 Then
                        SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness / 3, thickness * 2 / 3
                    ElseIf thickness > 38 Then
                        SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM, thickness * 2 / 5, thickness * 3 / 5
                    Else
                        SetBevelParams pParams, , , , , , pParams.ChamferAngle_M, _
                            pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                            pParams.ChamferDepth_UM
                    End If
                Case "X-Bevels1/3_2/3"
                    SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness / 3, thickness * 2 / 3
                Case "X-Bevels2/5_3/5"
                    SetBevelParams pParams, 45, 45, , , , pParams.ChamferAngle_M, _
                        pParams.ChamferAngle_UM, pParams.ChamferDepth_M, _
                        pParams.ChamferDepth_UM, thickness * 2 / 5, thickness * 3 / 5
            End Select
        Case Else
            SetBevelParams pParams
    End Select
End Sub
