VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MaxHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
' Module:
'
' Description:  Determines the process settings for the Template
'
' History:
'
'   10th Oct 2011    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessTube.MaxHeight"
Implements IJMfgTubeTemplateMaxHeightRule

Private Function IJMfgTubeTemplateMaxHeightRule_GetMaxHeight(ByVal pDispTube As Object) As Double
' Defines the Maximum Height applied for templates
' Currently only returns a fixed value but this could be
' in the future dependend on the workcenter assignment
    Const METHOD = "IJMfgTubeTemplateMaxHeightRule_GetMaxHeight"
    On Error GoTo ErrorHandler

    Dim MaxHeight As Double
    MaxHeight = 1

    ' Return value
    IJMfgTubeTemplateMaxHeightRule_GetMaxHeight = MaxHeight
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3007, , "RULES")
End Function
 

