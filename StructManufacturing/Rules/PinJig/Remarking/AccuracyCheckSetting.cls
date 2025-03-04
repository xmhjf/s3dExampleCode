VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccuracyCheckSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: AccuracyCheckSetting
'
' Description:
'
' Author: Suma Mallena
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigRule.AccuracyCheckSetting"

Implements IJDPinJigAccuracyRuleSetting

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigAccuracyRuleSetting_InitialAccuracyRuleSetting
' Purpose   : Update the value of each attribute,each related to Accuracy Check rule,
'---------------------------------------------------------------------------------------

Private Sub IJDPinJigAccuracyRuleSetting_InitialAccuracyRuleSetting(ByVal PartialPinJig As Object, ByVal AttributeColl As Object)
    Const METHOD = "IJDPinJigAccuracyRuleSetting_InitialAccuracyRuleSetting"
    On Error GoTo ErrorHandler

    Dim oAttribute As IJDAttribute
    Dim oAttrInfo As IJDAttributeInfo

    For Each oAttribute In AttributeColl
        Set oAttrInfo = oAttribute.AttributeInfo
         
        Select Case oAttrInfo.Name
       
            Case "AccuracyCheckRuleSet"
                oAttribute.Value = 151 'MfgPinJigAccuracyRuleSet_Default
            Case "BendPoints"
                oAttribute.Value = 1 'Apply
            Case "AftBoundary"
                oAttribute.Value = 0
            Case "LowerBoundary"
                oAttribute.Value = 0
            Case "ForeBoundary"
                oAttribute.Value = 0
            Case "UpperBoundary"
                oAttribute.Value = 0
            Case "Custom"
                oAttribute.Value = 0 'Ignore
        End Select
        Set oAttrInfo = Nothing
        Set oAttribute = Nothing
    Next
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5028, , "RULES")
End Sub
