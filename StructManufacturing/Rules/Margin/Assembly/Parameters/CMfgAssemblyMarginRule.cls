VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMfgAssemblyMarginRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: GSCADSTRMFGRULES
' Module: CMfgAssemblyMarginRule
'
' Description:  Determines the proposed settings for the assembly margin command
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Const Module = "CMfgAssemblyMarginRule:"

Implements IJDAssemblyMarginRule

Private Sub IJDAssemblyMarginRule_GetAssemblyMarginParameters(ByVal pDispAssbly As Object, PlateMarginVal As Double, ProfileMarginVal As Double)
Const METHOD = "IJDAssemblyMarginRule_GetAssemblyMarginParameters"
On Error GoTo ErrorHandler

Dim oWorkcenter As String
Dim oMfgRuleHelper As MfgRuleHelpers.MfgAssemblyHelper
Set oMfgRuleHelper = New MfgAssemblyHelper

Set oMfgRuleHelper.object = pDispAssbly

oWorkcenter = oMfgRuleHelper.AssemblyWorkCenter

If oWorkcenter = "Dock1" Then
    PlateMarginVal = 30
    ProfileMarginVal = 50
Else
    PlateMarginVal = 20
    ProfileMarginVal = 30
End If

Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 6003, , "RULES")
End Sub
 
