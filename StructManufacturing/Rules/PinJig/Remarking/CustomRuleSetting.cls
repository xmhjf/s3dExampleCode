VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomRuleSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: CustomRuleSetting
'
' Description:
'
' Author: Raman
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigRule.CustomRuleSetting"

Implements IJDPinJigRemarkingRulesSetting

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigRemarkingRulesSetting_ProvideRemarkingRulesSetting
' Purpose   : Update the value of each attribute,each related to a remarking rule,
'               based on whether corresponding remarking lines need to be shown or not
'---------------------------------------------------------------------------------------
Private Sub IJDPinJigRemarkingRulesSetting_ProvideRemarkingRulesSetting(ByVal PartialPinJig As Object, ByVal AttributeColl As Object)
    Const METHOD = "IJDPinJigRemarkingRulesSetting_ProvideRemarkingRulesSetting"
    On Error GoTo ErrorHandler

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oAttribute As IJDAttribute
    Dim oAttrInfo As IJDAttributeInfo

    For Each oAttribute In AttributeColl
        Set oAttrInfo = oAttribute.AttributeInfo
        Select Case oAttrInfo.Name
            Case "RemarkRuleSet"
                oAttribute.Value = 1 'MfgPinJigRemarkingRuleSet_Default
            Case "SeamRemark"
                If oPinJig.Type = PinJigType_Assembly Then
                    oAttribute.Value = 1
                Else
                    oAttribute.Value = 0
                End If
            Case "PlateRemark"
                oAttribute.Value = 1
            Case "ProfileRemark"
                oAttribute.Value = 1
            Case "GridLineX_Remark"
                oAttribute.Value = 1
            Case "GridLineY_Remark"
                oAttribute.Value = 0
            Case "GridLineZ_Remark"
                oAttribute.Value = 0
            Case "RefCurveRemark"
                oAttribute.Value = 1
            Case "UserRemark"
                oAttribute.Value = 1
            Case "UserExtend"
                oAttribute.Value = 1
            Case "SeamControlRemark"  ' Meaningful only if suitable bulkload is done
                oAttribute.Value = True ' Set to a non-zero value corresponding to 'Type' in CustomInterface
        End Select
        Set oAttrInfo = Nothing
        Set oAttribute = Nothing
    Next

    ' Setting below property to True gets the bed of pins to adjust to the support plates
    ' based on the user's interaction of move/rotate.  The remarking graphics is unchanged.
    '
    ' oPinJig.AdjustPinBedUponMoveRotate = True
    '
    ' Set above value to false if you want To mimic pre V7 SP4 behaviour
    ' (i.e., a move/rotate modifies the remarking graphics,
    ' but the bed of pins remain at the supported plates' location)


    Set oPinJig = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5025, , "RULES")
End Sub


