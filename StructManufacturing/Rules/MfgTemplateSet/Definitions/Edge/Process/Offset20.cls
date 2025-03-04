VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Offset20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2012, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessEdge
' Module:  Offset20
'
' Description:  Determines the process settings for the Edge Template offset, which will be applied to find location of bottom curve
'
' History:
'   20th April 2012    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessEdge.Offset20"

Implements IJDMfgEdgeTemplateOffsetRule

' Defines the Offset value to be used for computing the template bottom line location
Private Function IJDMfgEdgeTemplateOffsetRule_GetTemplateOffset(ByVal pDispPart As Object, Optional ByVal lSectionID As Long) As Double
Const METHOD = "IJMfgEdgeTemplateExtensionRule_GetTemplateExtension"
On Error GoTo ErrorHandler

    IJDMfgEdgeTemplateOffsetRule_GetTemplateOffset = 0.02 ' 20 mm
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9999, , "RULES")
End Function

