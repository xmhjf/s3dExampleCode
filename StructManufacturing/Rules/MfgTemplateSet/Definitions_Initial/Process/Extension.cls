VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extension"
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
Private Const MODULE = "MfgTemplateProcess.Extension"
Implements IJMfgTemplateExtensionRule2

'Private Function IJMfgTemplateExtensionRule_GetExtension(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
'' Defines the Extension value to be used for the template ebyond the seam defintion
'    Const METHOD = "IJMfgTemplateExtensionRule_GetExtension"
'On Error GoTo ErrorHandler
'
'    Dim sERROR As String
'    sERROR = ""
'
'    Dim Extension As Double
'
'    ' Set Extension value
'    Extension = 0.05
'
'    ' Return Value
'    IJMfgTemplateExtensionRule_GetExtension = Extension
'  Exit Function
'
'ErrorHandler:
'    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
'End Function

'Note:While IJMfgTemplateExtensionRule is still supported, suggest to use IJMfgTemplateExtensionRule2
Private Sub IJMfgTemplateExtensionRule2_GetExtension(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart, dExtension As Double, dOffset As Double, dMinExtn As Double)
' Defines the Extension value to be used for the template ebyond the seam defintion
Const METHOD = "IJMfgTemplateExtensionRule2_GetExtension"
On Error GoTo ErrorHandler

    ' Set Extension value
    dExtension = 0.05
    
    'not needed
    dOffset = 0
    dMinExtn = 0

Exit Sub
  
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub
