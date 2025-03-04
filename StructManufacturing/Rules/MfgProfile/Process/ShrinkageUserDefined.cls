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
Const MODULE = "MfgPlateProcess.ShrinkageAssPart"
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)
    Const METHOD = "ShrinkageAssPart: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2068, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "ShrinkageAssPart: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error GoTo ErrorHandler
    
    Dim oProfilePart As Object
    Dim oMfgPartParent As IJMfgChild

    Set oMfgPartParent = oMfgPart
    Set oProfilePart = oMfgPartParent.GetParent

    Dim oMfgParent As IJMfgParent
    Set oMfgParent = oProfilePart
    
    Dim oProfilePartHelper  As MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHelper = New MfgRuleHelpers.ProfilePartHlpr
    
    Set oProfilePartHelper.object = oProfilePart
    
    oProfilePartHelper.UpdateScalingShrinkage 2, 51, False

CleanUp:
    Set oProfilePart = Nothing
    Set oMfgPartParent = Nothing
    Set oMfgParent = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2068, , "RULES")
End Sub
