VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SightLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateMarking
' Module: SightLine
'
' Description:  Determines the sighline for a template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateMarking.SightLine"

Implements IJMfgTemplateSightLineRule
Private Function IJMfgTemplateSightLineRule_GetSightLineOffset(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
' Defines the Sightline Offset value, default is 0
Const METHOD = "IJMfgTemplateSightLineRule_GetSightLineOffset"
On Error GoTo ErrorHandler

'    Dim sERROR As String
'    sERROR = ""

    Dim SightLine As Double
'    Dim oPlate As IMSPLateEntity.IJPlate

    'Set default value
    SightLine = 0

    ' Should be replaced by the TemplateType,
    ' pending on the type we can then modify the
    ' sightline offset
'    If oPlate.NamingCategory = "Shell" Then
'        SightLine = 100
'    Else
'        SightLine = 0
'    End If

    ' Return Value
    IJMfgTemplateSightLineRule_GetSightLineOffset = SightLine
   Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3013, , "RULES")
End Function
 
