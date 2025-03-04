VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtnOffsetFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2012, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessEdge
' Module:  ExtnOffsetFix
'
' Description:  Determines the process settings for the Edge Template extension, which will be applied to both
'               the base and offset side of the Edge template
'               Properties:
'                   dExtn_Base: - The Extension value to be applied on base side
'                   dExtn_Offset: - The Extension value to be applied on offset side
'
' History:
'   20th April 2012    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessEdge.ExtnOffsetFix"

Implements IJMfgEdgeTemplateExtensionRule

' Defines the Extension value to be used for the template beyond the seam defintion
Private Sub IJMfgEdgeTemplateExtensionRule_GetTemplateExtension(ByVal pDispPart As Object, dExtn_Base As Double, dExtn_Offset As Double)
Const METHOD = "IJMfgEdgeTemplateExtensionRule_GetTemplateExtension"
On Error GoTo ErrorHandler

    dExtn_Base = 0
    dExtn_Offset = 0.1
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9999, , "RULES")
End Sub

