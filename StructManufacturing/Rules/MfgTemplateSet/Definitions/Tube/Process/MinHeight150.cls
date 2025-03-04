VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinHeight150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
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
Private Const MODULE = "StrMfgTemplateProcessTube.HeightOneFifty"
Implements IJMfgTubeTemplateMinHeightRule

Private Function IJMfgTubeTemplateMinHeightRule_GetMinHeight(ByVal pDispTube As Object) As Double
    Const METHOD = "IJMfgTubeTemplateMinHeightRule_GetMinHeight"
    On Error GoTo ErrorHandler

    ' Return value
    IJMfgTubeTemplateMinHeightRule_GetMinHeight = 0.15
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3008, , "RULES")
End Function


