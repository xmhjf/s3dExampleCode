VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WBSItemGroupNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'  File:    WBSItemGroupNameRule.cls
'
'  Abstract: The file contains a sample implementation of a naming rule which
'           could be used in conjunction with the Create WBS Items from Piping Parts
'           command.
'
'  Author: Mike Furno
'
'  23-May-2006  Created
'
'******************************************************************


Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "WBSItemNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const E_FAIL = -2147467259

Dim m_oErrors As IJEditErrors

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the IJWBSItemGroup.GroupName
' and IJSequence.Id proeperties.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    
    Dim strNamedParentsString As String
    Dim oNamedItem As IJNamedItem
    Dim oAttributeHelper As IJDAttributes
    Dim oAttributeCol As IJDAttributesCol
    Dim strLabel As String
    Dim strSeqId As String
    Dim strNewName As String
    
    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oAttributeHelper = pEntity
    
    On Error Resume Next
    Set oAttributeCol = oAttributeHelper.CollectionOfAttributes("IJWBSItemGroup")
    On Error GoTo label
    
    If (oAttributeCol Is Nothing) Then
        '  Object is not a WBSItemGroup user type.  Just give a dummy name
        strNewName = "WBS Item"
    Else
        strLabel = oAttributeCol.Item("GroupName").Value
        strSeqId = oAttributeHelper.CollectionOfAttributes("IJSequence").Item("Id").Value
        strNewName = strLabel + "-" + strSeqId
    End If
    
    'Check if New  name string constructed and old  name string existing are same
    If (strNewName <> pActiveEntity.NamingParentsString) Then
        pActiveEntity.NamingParentsString = strNewName
        Set oNamedItem = pEntity
        oNamedItem.Name = strNewName
    End If
        
    Exit Sub

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.

'  In this case there are no naming parents since the properties which are used to construct the name
'  are directly on the WBSItem object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo label

Exit Function

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
