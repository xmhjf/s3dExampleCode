VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtnAlongEdge50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
' Module:  ExtnAlongEdge50
'
' Description:  Determines the process settings for the Template extension, which will be applied to both
'               the left and right side of the tube template
'               Properties:
'                   method: - How the extension is measured;
'                           0 : Linear extension
'                           1 : Perpendicular
'                           2 : Along edge
'                   measure: - how to measure the extension offset to the tube edge.
'                           0 : Along girth
'                           1 : Along tube axis
'                   value1: - The offset value to be applied
'                   value2: - reserved distance value for new method/offset type
'
' History:
'   10th Oct 2011    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessTube.ExtnAlongEdge50"

Implements IJMfgTubeTemplateExtensionRule

' Defines the Extension value to be used for the template ebyond the seam defintion
Private Sub IJMfgTubeTemplateExtensionRule_GetExtension(ByVal pDispTube As Object, lMethod_Left As Long, lMethod_Right As Long, lMeasure_Left As Long, _
                                lMeasure_Right As Long, dValue1_Left As Double, dValue2_Left As Double, dValue1_Right As Double, dValue2_Right As Double)
Const METHOD = "IJMfgTubeTemplateExtensionRule_GetExtension"
On Error GoTo ErrorHandler

    lMethod_Left = 2
    lMethod_Right = 2
    
    'offset not to be used
    lMeasure_Left = 0
    lMeasure_Right = 0
    
    dValue1_Left = 0.05
    dValue1_Right = 0.05
    
    dValue2_Left = 0
    dValue2_Right = 0
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9999, , "RULES")
End Sub
