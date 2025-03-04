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
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Fixed Maximum Height Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcessProfile.MaxHeight"
Implements IJMfgTemplateMaxHeightRule
Private Function IJMfgTemplateMaxHeightRule_GetMaxHeight(ByVal pPartDisp As Object) As Double
' Defines the Maximum Height applied for templates
' Currently only returns a fixed value but this could be
' in the future dependend on the workcenter assignment
     Const METHOD = "IJMfgTemplateMaxHeightRule_GetMaxHeight"
     
    On Error GoTo ErrorHandler

    Dim sERROR As String
    sERROR = ""

    Dim MaxHeight As Double

    ' Set Maxheight value
    MaxHeight = 0.6

    ' Return value
    IJMfgTemplateMaxHeightRule_GetMaxHeight = MaxHeight
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3007, , "RULES")
End Function
 
