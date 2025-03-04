VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtensionOffSet"
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
' 2009.09.23    Triveni CH     Created
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateProcess.ExtensionOffSet"

Implements IJMfgTemplateExtensionRule2

Private Sub IJMfgTemplateExtensionRule2_GetExtension(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart, dExtension As Double, dOffset As Double, dMinExtn As Double)
' Defines the Extension value to be used for the template ebyond the seam defintion
Const METHOD = "IJMfgTemplateExtensionRule2_GetExtension"
On Error GoTo ErrorHandler

    dExtension = 0#
    dOffset = 0.25
    dMinExtn = 0.025

Exit Sub
  
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub


