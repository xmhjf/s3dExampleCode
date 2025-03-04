VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoamingCalculated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule

Const MODULE = "MfgProfileProcess"

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "CoamingCalculated: IJDMfgProcessRule_ExecuteRule"
    On Error GoTo ErrorHandler
     
    ' performance check - exit if not edge reinforcement
    Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart
    If TypeOf Part Is IJStiffenerPart Then
        Set oSDProfileWrapper.object = Part
        If oSDProfileWrapper.ProfileType <> sptEdgeReinforcement Then Exit Sub
    ElseIf TypeOf Part Is IJBeamPart Then
        Exit Sub
    End If
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    ' Don't we want to have the clampmargin value dependent on the amount of bend at the ends of the profile ?
    On Error Resume Next
    Dim oAssyChild As IJAssemblyChild
    Set oAssyChild = Part
    If (oAssyChild Is Nothing) Then
        oProfileWrapper.RemoveMargins 8 'CoamingMargin
        Exit Sub
    End If
        
    ' determine if part is in assembly.  If so, then get needed information information
    Dim oAssyParent As IJAssembly
    Set oAssyParent = oAssyChild.Parent
    If (oAssyParent Is Nothing) Then
        oProfileWrapper.RemoveMargins 8 'CoamingMargin
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Dim oAssemblyHelper As New MfgAssemblyHelper
    Set oAssemblyHelper.object = oAssyParent
    
    Dim oWorkCenter As String
    oWorkCenter = oAssemblyHelper.WorkCenterFromPart(Part)
    
    Dim dMargin As Double
    If oSDProfileWrapper.ProfileType = sptEdgeReinforcement Then
        Select Case oWorkCenter
        Case "BendingDefault"
            dMargin = 0.5
        Case "FenderBender"
            dMargin = 0.35
        Case Else
            dMargin = 0.2
        End Select
    End If
    
    oProfileWrapper.UpdateConstantMargin PROFILE_BASE, 8, dMargin ' 8 - CoamingMargin
    oProfileWrapper.UpdateConstantMargin PROFILE_OFFSET, 8, dMargin ' 8 - CoamingMargin

    Exit Sub
ErrorHandler:
    
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2023, , "RULES")

End Sub
 
