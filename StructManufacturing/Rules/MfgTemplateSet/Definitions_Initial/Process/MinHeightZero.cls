VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinHeightZero"
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
Private Const MODULE = "MfgTemplateProcess.MinHeightZero"
Implements IJMfgTemplateMinHeightRule
Private Function IJMfgTemplateMinHeightRule_GetMinHeight(ByVal pPartDisp As Object) As Double
' Minimum Template Height
' Currently only returns a fixed value but this could be
' in the future dependend on the workcenter assignment
     Const METHOD = "IJMfgTemplateMinHeightRule_GetMinHeight"
On Error GoTo ErrorHandler

    Dim sERROR As String
    sERROR = ""

    Dim MinimumHeight As Double

    ' Set Minimum template Height value as zero
    MinimumHeight = 0#

    ' Return value
    IJMfgTemplateMinHeightRule_GetMinHeight = MinimumHeight
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3008, , "RULES")
End Function
 
