VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinHeightFifty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module:
'
' Description:  Determines the process settings for the Template
'
' Author:
'
' Comments:
' 2009.11.03    Siva     Created
'*******************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcess.HeightFifty"
Implements IJMfgTemplateMinHeightRule

Private Function IJMfgTemplateMinHeightRule_GetMinHeight(ByVal pPartDisp As Object) As Double
    Const METHOD = "IJMfgTemplateMinHeightRule_GetMinHeight"
    On Error GoTo ErrorHandler

    ' Return value
    IJMfgTemplateMinHeightRule_GetMinHeight = 0.05
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3008, , "RULES")
End Function
