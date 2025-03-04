VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: Profile Custom
'
' Description:  Add custom attributes for manufacturing profile in this module.
'
' Author:
'
' Comments:
' 2010.07.14    StructMfg    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const MODULE = "MfgPlateProcess"
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgMemberPart As Object)
    
    On Error GoTo ErrorHandler
    
    If oMfgMemberPart Is Nothing Then
        GoTo Cleanup
    End If
    
    ' Get the profile cross section object
    Dim oMfgChild As IJMfgChild
    Set oMfgChild = oMfgMemberPart
    
    Dim oMemberPart As Object
    Set oMemberPart = oMfgChild.GetParent
    
    ' Get the production info related custom attributes
    Dim varCustomA As Variant, varCustomB As Variant, varCustomC As Variant, varCustomD As Variant
    GetMemberCustomAttributes oMemberPart, varCustomA, varCustomB, varCustomC, varCustomD
    
    ' Get the attribute coll of the custon interface "IJMfgPartCustomAttr"
    Dim oAttribute As IJDAttribute
    Dim oAttrCol As IJDAttributesCol
    Set oAttrCol = GetAttributeColl(oMfgMemberPart, "IJMfgPartCustomAttr")
    
    ' Convert the Custom(A,B,C and D) values into mm
    Set oAttribute = oAttrCol.Item("Custom_A")
    oAttribute.Value = CStr(CDbl(varCustomA) * 1000)
    Set oAttribute = Nothing
    
    Set oAttribute = oAttrCol.Item("Custom_B")
    oAttribute.Value = CStr(CDbl(varCustomB) * 1000)
    Set oAttribute = Nothing
    
    Set oAttribute = oAttrCol.Item("Custom_C")
    oAttribute.Value = CStr(CDbl(varCustomC) * 1000)
    Set oAttribute = Nothing
    
    Set oAttribute = oAttrCol.Item("Custom_D")
    oAttribute.Value = CStr(CDbl(varCustomD) * 1000)
    Set oAttribute = Nothing
    
    ' Get the attribute coll of the custon interface "IJMfgProfileInfo"
    Set oAttrCol = Nothing
    Set oAttrCol = GetAttributeColl(oMfgMemberPart, "IJMfgProfileInfo")
    
    Set oAttribute = oAttrCol.Item("NC_Machine")
    If Not oAttribute Is Nothing Then
        oAttribute.Value = 11   'Default
    End If
    
    Set oAttribute = oAttrCol.Item("Primer")
    If Not oAttribute Is Nothing Then
        oAttribute.Value = 1  'Undefined
    End If
    
    
    Set oAttribute = oAttrCol.Item("Grinder")
    If Not oAttribute Is Nothing Then
        oAttribute.Value = 1  'Undefined
    End If
    
 
    '*** Start of Code For Routing Attribute****'
    
'    Dim oAttributes As IJDAttributes
'    Set oAttributes = oMfgMemberPart
'    On Error Resume Next        ' This is needed to protect the code when the catalog missing SKDYMfgProfilePartInfo_GUID
'    Set oAttrCol = oAttributes.CollectionOfAttributes(SKDYMfgProfilePartInfo_GUID)
'    On Error GoTo ErrorHandler
'    If Not oAttrCol Is Nothing Then
'        Set oAttribute = oAttrCol.Item("Routing")
'        strTable = "StrMfgStageAndBendingCodes"
'        If Not strRouting = vbNullString And Not strRouting = "None" Then
'            lCodeList = GetCodelistNumberFromShortValue(strRouting, strTable)
'            oAttribute.Value = lCodeList
'        End If
'    End If
    '******* End of Code For Grinder Attributes *********'
    
Cleanup:
    Set oMfgChild = Nothing
    Set oMemberPart = Nothing
    Set oAttrCol = Nothing
    
    Exit Sub
''''    Set oPlnProdRouting = Nothing
ErrorHandler:

End Sub
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)
End Sub


