VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginCalculated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module:
'
' Description:  Determines the process settings for the mfg plate
' The rule handles margin creation for Can plates, flange brackets
' Author:
'
' Comments:
' 2010.06.22    Siva    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const Module = "MfgMemberProcess"
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "MarginCalculated: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error Resume Next

    If oMfgPart Is Nothing Then
        GoTo ErrorHandler
    End If
    
    ' Get the profilepart object
    Dim oMfgChild As IJMfgChild
    Set oMfgChild = oMfgPart
    
    Dim oMemberPart As Object
    Set oMemberPart = oMfgChild.GetParent
    
    Dim oProfileWrapper             As MfgRuleHelpers.ProfilePartHlpr

    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oMemberPart
    
    oProfileWrapper.UpdateConstantMargin PROFILE_OFFSET, 1, MARGIN_VALUE, 2, 1, 0, "Calculated Margin"

Cleanup:
    Set oProfileWrapper = Nothing
    Set oMemberPart = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1060, , "RULES")
    GoTo Cleanup
End Sub

Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)
    Const METHOD = "MarginCalculated: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error Resume Next

Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1060, , "RULES")
    GoTo Cleanup
End Sub
