VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtnNone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   No Extn Extension Process rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcess.Extension"
Implements IJMfgTemplateExtensionRule2


'Note:While IJMfgTemplateExtensionRule is still supported, suggest to use IJMfgTemplateExtensionRule2
Private Sub IJMfgTemplateExtensionRule2_GetExtension(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart, dExtension As Double, dOffset As Double, dMinExtn As Double)
' Defines the Extension value to be used for the template ebyond the seam defintion
Const METHOD = "IJMfgTemplateExtensionRule2_GetExtension"
On Error GoTo ErrorHandler

    'Constant extension of 0mm(i.e no extension needed)
    dExtension = 0#
    
    'offset not to be used
    dOffset = 0
    dMinExtn = 0

Exit Sub
  
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub
