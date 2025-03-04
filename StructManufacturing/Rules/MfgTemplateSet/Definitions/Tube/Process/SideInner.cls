VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
' Module: SideInner
'
' Description:  Determines the process settings for the mfg template
'
' History:
'
'   10th Oct 2011    Siva     Creation
'*******************************************************************************
Option Explicit

Private Const MODULE = "StrMfgTemplateProcessTube.SideInner"

Implements IJDMfgTubeTemplateSideRule

Private Function IJDMfgTubeTemplateSideRule_GetTemplateSide(ByVal pDispTube As Object) As GSCADMfgRulesDefinitions.Mfg_TubeEnd
    Const METHOD = "IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    IJDMfgTubeTemplateSideRule_GetTemplateSide = TUBE_INNER
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
End Function

 
