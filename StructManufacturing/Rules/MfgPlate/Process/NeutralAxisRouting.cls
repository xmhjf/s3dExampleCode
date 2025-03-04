VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Routing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: Routing(NeutralAxis)
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2011.01.03    Siva     Created
'*******************************************************************************
Option Explicit

Private Const Module = "MfgPlateProcess.Routing"

Implements IJDMfgPlateNeutralAxisRule

Private Function IJDMfgPlateNeutralAxisRule_GetNeutralAxis(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "NeutralAxisRouting: IJDMfgPlateNeutralAxisRule_GetNeutralAxis"
    On Error GoTo ErrorHandler
    
    ' Use the plate part as input
    Dim dNARatio As Double
    Dim eInnerSurfSide As enumPlateSide
    Dim oSDPlatePart As New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPlatePart
    
    ' ******** SKDY specific code starts **********
    Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
    Set oMfgPlateWrapper.object = pPlatePart
    
    On Error Resume Next
    ' Calling CurvatureType function takes long time so do not call it if not really necessary.
    Dim ePlateCurvature As PlateCurvature
    ePlateCurvature = oMfgPlateWrapper.CurvatureType
    On Error GoTo ErrorHandler
    
    ' Check if it is a single curved/developable plate
    If ePlateCurvature = PLATE_CURVATURE_Knuckled Or _
        ePlateCurvature = PLATE_CURVATURE_SingleCurvature Or _
        ePlateCurvature = PLATE_CURVATURE_SingleCurvature_Knuckled Then
        
        Dim oPlnProdRouting As PlanningObjects.PlnProdRouting
        Set oPlnProdRouting = New PlanningObjects.PlnProdRouting
        Set oPlnProdRouting.object = pPlatePart
        
        ' Get the Bending machine info from Planning Production Routing
        Dim strMachine As String, strCode As String
        oPlnProdRouting.GetActionInfo "Bending", strMachine, strCode
        
        ' 300T Press Bender     -   1/2 Thickness
        ' 500T Press Bender     -   1/2 Thickness
        ' 1000T Press Bender    -   1/2 Thickness
        ' Bending Roller        -   1/2 Thickness
        ' 1750T Ship Bender     -   1/2 Thickness
        ' 3000T Press Bender    -   1/2 Thickness - 2mm
        
        If strMachine = "3000T Press Bender" Then
            Dim dHalfPlateThickness As Double
            dHalfPlateThickness = (0.5 * oSDPlatePart.PlateThickness * 1000#)
            
            IJDMfgPlateNeutralAxisRule_GetNeutralAxis = ((dHalfPlateThickness - 2#) / dHalfPlateThickness) * 0.5
            Exit Function
        Else
            IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
            Exit Function
        End If
    Else
        IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
        Exit Function
    End If
    
    ' ******** SKDY specific code ends **********
    
CleanUp:
    Set oSDPlatePart = Nothing
    Set oMfgPlateWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1095, , "RULES")
    GoTo CleanUp
End Function
 

