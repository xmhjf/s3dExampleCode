VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompUD
'
' Description:  Apply Weld Compensation with User Defined Parameters
'
' Comments:
'
' 8 Feb, 2008    Creation
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkRule
 
Private m_oWCFactory As IJMfgWeldCompensationFactory
Private m_oWC As IJMfgWeldCompensation

Const MODULE = "MfgMemberProcess.WeldCompUD"

Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2037, , "RULES")
End Function
 
