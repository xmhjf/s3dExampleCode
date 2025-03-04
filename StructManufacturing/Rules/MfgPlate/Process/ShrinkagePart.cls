VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShrinkagePart"
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
' Comments:
' 2010.07.14    StructMfg    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const MODULE = "MfgPlateProcess.ShrinkageYes"
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)
    Const METHOD = "ShrinkageYes: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1081, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "ShrinkageYes: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error GoTo ErrorHandler
    
    ''''create new part shrinkage
    
'    Dim oScalingshrinkageBOfactory As New ScalingShrFactory
'    Dim oPlatePart As Object
'    Dim oMfgPartParent As IJMfgChild
'
'    Set oMfgPartParent = oMfgPart
'    Set oPlatePart = oMfgPartParent.GetParent
'
'    Dim oMfgParent As IJMfgParent
'    Set oMfgParent = oPlatePart
'
'    Dim oPlatePartHlpr  As MfgRuleHelpers.PlatePartHlpr
'    Set oPlatePartHlpr = New MfgRuleHelpers.PlatePartHlpr
'
'    Set oPlatePartHlpr.object = oPlatePart
'
'    oPlatePartHlpr.UpdateScalingShrinkage 1, 1, True

CleanUp:
'    Set oPlatePart = Nothing
'    Set oMfgPartParent = Nothing
'    Set oMfgParent = Nothing

    Exit Sub
    
ErrorHandler:
MsgBox Err.Description
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1081, , "RULES")
End Sub
