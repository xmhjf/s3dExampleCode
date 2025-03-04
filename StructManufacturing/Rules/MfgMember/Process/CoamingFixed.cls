VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoamingFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: CoamingFixed
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule
Const MODULE = "MfgMemberProcess"

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "CoamingFixed: IJDMfgProcessRule_ExecuteRule"
    On Error GoTo ErrorHandler
    
    Exit Sub
    
    Dim oSDProfilePart As New StructDetailObjects.ProfilePart
    Set oSDProfilePart.object = Part

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

'    If oSDProfilePart.ProfileType <> sptEdgeReinforcement Then
'        oProfileWrapper.RemoveMargins 8 ' CoamingMargin
'        Exit Sub
'    End If

    oProfileWrapper.UpdateConstantMargin PROFILE_BASE, 8, 0.2  '8 - Coaming margin , 0.2 is margin value
    oProfileWrapper.UpdateConstantMargin PROFILE_OFFSET, 8, 0.2 '8 - Coaming margin , 0.2 is margin value

    Exit Sub

ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2024, , "RULES")

End Sub
 
