VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module:
'
' Description:  Determines the process settings for the Template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcess.Frame"
Implements IJMfgTemplateTypeRule

Private Sub IJMfgTemplateTypeRule_GetTemplateType(ByVal pPartDisp As Object, TemplateType As String, Optional dInterval As Double)
    Const METHOD = "IJMfgTemplateTypeRule_GetTemplateType"
    On Error GoTo ErrorHandler

        TemplateType = "Frame"
        dInterval = 0.2

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3004, , "RULES")
End Sub
 
