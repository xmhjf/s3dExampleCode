VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeatureTabYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE As String = "MfgPlateProcess.FeatureTabYes"

Private Const FEATURETAB_MARKINGNAME = "FEATURETAB_MARKING"
Implements IJDMfgProcessRule
Implements IJDMfgPostUnfoldRule


Private Sub IJDMfgPostUnfoldRule_ExecuteRule(ByVal Part As Object, ByVal UnfoldedGeomColl As Object, ByVal oTabCandidateColl As IJElements, NewGeomColl As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "FeaturTabYes: IJDMfgProcessRule_ExecuteRule"
    On Error GoTo ErrorHandler

    ' Do not check for feature tabs to be placed on the geometry.
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1048, , "RULES")
End Sub

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    
End Sub
