VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WBSItemNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'  File:    WBSItemNameRule.cls
'
'  Abstract: The file contains a sample implementation of a naming rule
'       If the object is a WBSItem object, then the name rule is
'       "WBS Project parent Name" + "Unique counter"
'
'  Author: David Kelley
'
'  24-Apr-2006  Created
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
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "WBSProject Name" + Unique counter.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim oNameCounter As IJNameCounter
    Dim oChildNamedItem As IJNamedItem
    Dim oParentNamedItem As IJNamedItem
    Dim oWBSItem As IJWBSItem
    Dim oWBSChild As IJWBSChild
    
    Dim strChildName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strLocation As String

    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oChildNamedItem = pEntity
    Set oWBSChild = pEntity
    If Not oWBSChild Is Nothing Then
        Set oParentNamedItem = oWBSChild.GetParent
    End If
    If Not oParentNamedItem Is Nothing Then
        strChildName = oParentNamedItem.Name
    Else
        strChildName = oChildNamedItem.TypeString
    End If
    
    strNamedParentsString = pActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If (strChildName) <> strNamedParentsString Then
            pActiveEntity.NamingParentsString = strChildName
        
            strLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName, strLocation)
            
            If strLocation <> vbNullString Then
                strChildName = strChildName & "-" & "-" & strLocation & "-" & Format(nCount, strCountFormat)
            Else
                strChildName = strChildName & "-" & "-" & Format(nCount, strCountFormat)
            End If
            oChildNamedItem.Name = strChildName & nCount
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo label

    Dim oWBSParent As IJWBSParent
    Dim oWBSChild As IJWBSChild

    On Error Resume Next
    Set oWBSChild = pEntity
    
    If Not (oWBSChild Is Nothing) Then
        Set oWBSParent = oWBSChild.GetParent
        
        On Error GoTo label
    
        Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
        If Not (oWBSChild Is Nothing) Then
            Call IJNameRule_GetNamingParents.Add(oWBSParent)
        End If
    End If


    Set oWBSParent = Nothing
    Set oWBSChild = Nothing
Exit Function

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
