VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtnRightfix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessTube
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
Private Const MODULE = "StrMfgTemplateProcessTube.ExtnRightFix"

Implements IJMfgTubeTemplateExtensionRule

' Defines the Extension value to be used for the template ebyond the seam defintion
Private Sub IJMfgTubeTemplateExtensionRule_GetExtension(ByVal pDispTube As Object, dLeftExtension As Double, dLeftOffset As Double, dMinLeftExtn As Double, dRightExtension As Double, dRightOffset As Double, dMinRightExtn As Double)
Const METHOD = "IJMfgTubeTemplateExtensionRule_GetExtension"
On Error GoTo ErrorHandler

    'Constant extension of 0mm(i.e no extension needed)
    dLeftExtension = 0#
    dRightExtension = 0.1
        
    'offset not to be used
    dLeftOffset = 0#
    dMinLeftExtn = 0#
    
    dRightOffset = 0#
    dMinRightExtn = 0#

    Exit Sub
  
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub



