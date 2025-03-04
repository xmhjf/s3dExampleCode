VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateMarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: ValidateMarking
'
' Description:  In this function, user can control Validate marks.
'  History:
'       StructMfg         October 6th, 2010   created
'
'*******************************************************************************
Implements IJMfgMarkingValidate
Const Module = "MfgMemberProcess.ValidateMarking"
Private Sub IJMfgMarkingValidate_ValidateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal MfgGeomCol3d As GSCADMfgRulesDefinitions.IJMfgGeomCol3d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateBeforeUnfold"
    On Error GoTo ErrorHandler
    
Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 2058, , "RULES")
    GoTo Cleanup
End Sub
Private Sub IJMfgMarkingValidate_ValidateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateAfterUnfold"
    On Error GoTo ErrorHandler
    
Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 2058, , "RULES")
    GoTo Cleanup
End Sub

