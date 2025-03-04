VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgFrameSysRule
' Module: MfgPlate
'
' Author: Anand Hariharan
'
' Comments:
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgFrameSysRule.MfgPlate"

Implements IJMfgRuleForFrameSystem

Private Function IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity(ByVal pEntityDsp As Object) As IMSCoreCollections.IJElements

    Const METHOD = "IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity"
    On Error GoTo ErrorHandler

    Dim MfgQuery As IJDMfgFrameSysUtil
    Set MfgQuery = New CMfgCoordSys

    Set IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity = MfgQuery.FrameSystemsInRange(pEntityDsp)

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3002, , "RULES")
End Function
