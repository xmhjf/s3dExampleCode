VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideWebLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2012, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessEdge
' Module: SideWebLeft
'
' Description:  Determines the process settings for the mfg Template
'
' History:
'
'   20th April 2012    Siva     Creation
'*******************************************************************************
Option Explicit

Private Const MODULE = "StrMfgTemplateProcessEdge.SideWebLeft"
Implements IJDMfgEdgeTemplateSideRule

Private Function IJDMfgEdgeTemplateSideRule_GetTemplateSide(ByVal pDispPart As Object) As Long
    Const METHOD = "IJDMfgEdgeTemplateSideRule_GetTemplateSide"
    On Error GoTo ErrorHandler

    'JXSEC_BOTTOM = 513 (&H201)
    'JXSEC_TOP = 514 (&H202)
    'JXSEC_WEB_LEFT = 257 (&H101)
    'JXSEC_WEB_RIGHT = 258 (&H102)
    
    IJDMfgEdgeTemplateSideRule_GetTemplateSide = 257

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
End Function

