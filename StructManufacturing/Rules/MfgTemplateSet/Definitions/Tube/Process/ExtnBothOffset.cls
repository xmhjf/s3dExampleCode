VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtBothOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
' Module:
'
' Description:  Determines the process settings for the Template extension, which will be applied to both
'               the left and right side of the tube template
'               Properties:
'                   method: - How the extension is measured;
'                           0 : linear extension
'                           1 : perpendicular
'                           2 : along edge
'                   measure: - how to measure the extension offset to the tube edge.
'                           0 : along girth
'                           1 : perpendicular to tube axis
'                           2 : along tube axis
'                   value1: - The offset value to be applied
'                   value2: - reserved distance value for new method/offset type
'
'
' Author:
'
' Comments:
' 10th Oct 2011    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessTube.ExtnBothOffset"
Implements IJMfgTubeTemplateExtensionRule

' Defines the Extension value to be used for the template ebyond the seam defintion
Private Sub IJMfgTubeTemplateExtensionRule_GetExtension(ByVal pDispTube As Object, dLeftExtension As Double, dLeftOffset As Double, dMinLeftExtn As Double, dRightExtension As Double, dRightOffset As Double, dMinRightExtn As Double)
Const METHOD = "IJMfgTubeTemplateExtensionRule_GetExtension"
On Error GoTo ErrorHandler

    lMethod_Left = 0        '
    lMethod_Right = 0
        
    'offset not to be used
    lMeasure_Left = 0
    lMeasure_Right = 0
    
    dValue1_Left = 0.2
    dValue2_Left = 0.2
    dValue1_Right = 0.2
    dValue2_Right = 0.2
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub
