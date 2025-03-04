VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinHt150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   150 As Minimum Height Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************


Option Explicit
Private Const MODULE = "MfgTemplateProcessProfile.MinHt150"
Implements IJMfgTemplateMinHeightRule

Private Function IJMfgTemplateMinHeightRule_GetMinHeight(ByVal pPartDisp As Object) As Double
    Const METHOD = "IJMfgTemplateMinHeightRule_GetMinHeight"
    On Error GoTo ErrorHandler

    ' Return value
    IJMfgTemplateMinHeightRule_GetMinHeight = 0.15
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3008, , "RULES")
End Function


