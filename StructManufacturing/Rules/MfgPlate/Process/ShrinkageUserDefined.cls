VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShrinkageUserDefined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module:
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const MODULE = "MfgPlateProcess.ShrinkageUserDefined"
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "ShrinkageUserDefined: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
        
    Dim oPlatePart As Object
    Dim oMfgPlatePart As IJMfgPlatePart
    
    Set oMfgPlatePart = oMfgPart
    oMfgPlatePart.GetDetailedPart oPlatePart
       
    Dim oPlatePartHlpr  As MfgRuleHelpers.PlatePartHlpr
    Set oPlatePartHlpr = New MfgRuleHelpers.PlatePartHlpr
    
    Set oPlatePartHlpr.object = oPlatePart
    
    '''' "Assembly Part Rule1" -- 51
    oPlatePartHlpr.UpdateScalingShrinkage 2, 51, False
    
CleanUp:
    Set oPlatePart = Nothing
    Set oMfgPlatePart = Nothing
    Set oPlatePartHlpr = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1098, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "ShrinkageUserDefined: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error GoTo ErrorHandler

CleanUp:
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1098, , "RULES")
End Sub
