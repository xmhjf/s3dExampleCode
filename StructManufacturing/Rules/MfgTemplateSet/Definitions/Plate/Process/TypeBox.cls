VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Box Type Process rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcess.Box"
Implements IJMfgTemplateTypeRule

Private Sub IJMfgTemplateTypeRule_GetTemplateType(ByVal pPartDisp As Object, TemplateType As String, Optional dInterval As Double)
    Const METHOD = "IJMfgTemplateTypeRule_GetTemplateType"
    On Error GoTo ErrorHandler

        TemplateType = "Box"
        dInterval = 0.2

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3004, , "RULES")
End Sub
 

