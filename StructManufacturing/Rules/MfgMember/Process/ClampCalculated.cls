VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClampCalculated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: ClampCalculated
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
    Const METHOD = "ClampCalculated: IJDMfgProcessRule_ExecuteRule"
    On Error Resume Next
    
    Exit Sub
    
    ' Don't we want to have the clampmargin value dependent on the amount of bend at the ends of the profile ?
    Dim oAssyChild As IJAssemblyChild
    Set oAssyChild = Part
    If (oAssyChild Is Nothing) Then Exit Sub
        
    ' determine if part is in assembly.  If so, then get needed information information
    Dim oAssyParent As IJAssembly
    Set oAssyParent = oAssyChild.Parent
    If (oAssyParent Is Nothing) Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim oAssemblyHelper As New MfgAssemblyHelper
    Set oAssemblyHelper.object = oAssyParent
    
    Dim oWorkCenter As String
    oWorkCenter = oAssemblyHelper.WorkCenterFromPart(Part)
   
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

    Dim dMargin As Double
    Select Case oWorkCenter
    Case "BendingDefault"
        dMargin = 0.5
    Case "FenderBender"
        dMargin = 0.35
    Case Else
        'Do nothing
        GoTo CleanUp
    End Select

    oProfileWrapper.UpdateConstantMargin PROFILE_BASE, 7, dMargin ' 7 - ClampMargin
    oProfileWrapper.UpdateConstantMargin PROFILE_OFFSET, 7, dMargin '7 - ClampMargin

CleanUp:
    Set oAssemblyHelper = Nothing
    Set oAssyChild = Nothing
    Set oAssyParent = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2019, , "RULES")
    GoTo CleanUp
End Sub
 
