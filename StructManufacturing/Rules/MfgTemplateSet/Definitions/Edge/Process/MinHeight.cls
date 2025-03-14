VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2012, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Fixed Minimum Height Process rule
'
'Description:
'History :
'   Siva    08/10/2012      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "StrMfgTemplateProcessEdge.MinHeight"
Implements IJMfgEdgeTemplateMinHeightRule

Private Function IJMfgEdgeTemplateMinHeightRule_GetMinHeight(ByVal pPartDisp As Object) As Double
    Const METHOD = "IJMfgEdgeTemplateMinHeightRule_GetMinHeight"
    On Error GoTo ErrorHandler
    
    ' Return default Minimum template Height value
    IJMfgEdgeTemplateMinHeightRule_GetMinHeight = 0.35
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3008, , "RULES")
End Function
