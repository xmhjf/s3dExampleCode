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
'Copyright (C) 2012, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Fixed Maximum Height Process rule
'
'Description:
'History :
'   Siva        08/10/2012      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "StrMfgTemplateProcessEdge.MaxHeight"
Implements IJMfgEdgeTemplateMaxHeightRule

Private Function IJMfgEdgeTemplateMaxHeightRule_GetMaxHeight(ByVal pPartDisp As Object) As Double
    Const METHOD = "IJMfgEdgeTemplateMaxHeightRule_GetMaxHeight"
    On Error GoTo ErrorHandler

    ' Return Maximum template Height value as 1 m
    IJMfgEdgeTemplateMaxHeightRule_GetMaxHeight = 1#

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3007, , "RULES")
End Function
 


