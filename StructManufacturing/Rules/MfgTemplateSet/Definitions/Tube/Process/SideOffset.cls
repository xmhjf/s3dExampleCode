VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
' Module: SideOffset
'
' Description:  Determines the process settings for the mfg template
'
' History:
'
'   10th Oct 2011    Siva     Creation
'*******************************************************************************
Option Explicit

Private Const MODULE = "StrMfgTemplateProcessTube.SideOffset"

Implements IJDMfgTubeTemplateSideEndRule

Private Function IJDMfgTubeTemplateSideEndRule_GetTemplateSideEnd(ByVal pDispTube As Object) As GSCADMfgRulesDefinitions.Mfg_TubeEnd
    Const METHOD = "IJDMfgTubeTemplateSideRule_GetTemplateSideEnd"
    On Error GoTo ErrorHandler
    
    IJDMfgTubeTemplateSideEndRule_GetTemplateSideEnd = TUBE_OFFSET
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
End Function
