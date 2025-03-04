VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADHoleNamingRules
'
'  Abstract: The file contains a implementation of a SystemNameRule for Hole Management
'
'  Author: Jang Choi
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "CSysNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "Base Name" + "Object Name" + Index.
' "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName( _
        ByVal oHoleTraceAE As Object, _
        ByVal oNamingParents As IJElements, _
        ByVal oActiveEntity As Object)

    Const Method = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
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
    Dim oHoleTraceNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim strBaseName As String
    Dim strObjectType As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strHoleName As String
    Dim strLocationID As String

    If oHoleTraceAE Is Nothing Then
        GoTo ErrorHandler
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oHoleTraceNamedItem = oHoleTraceAE
    strObjectType = oHoleTraceNamedItem.TypeString

    ' Use Name for basename, remove blanks:
    strObjectType = Join(Split(strObjectType), "")

    If oNamingParents.Count > 0 Then
        Set oParent = oNamingParents.Item(1)
        
        On Error GoTo ErrorHandler
        
        strBaseName = GetBaseName(oParent)
        
        strNamedParentsString = oActiveEntity.NamingParentsString
        If (strBaseName + strObjectType) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strObjectType
            
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strBaseName, strLocationID)
            
            If Not strLocationID = vbNullString Then
                strHoleName = strBaseName & "_" & strObjectType & "_" & strLocationID & "_" & Format(nCount, strCountFormat)
            Else
                strHoleName = strBaseName & "_" & strObjectType & "_" & Format(nCount, strCountFormat)
            End If
            
            oHoleTraceNamedItem.Name = strHoleName
        End If
    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strObjectType <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strObjectType

            nCount = oNameCounter.GetCount(oModelResourceMgr, strObjectType)
            
            If Not strLocationID = vbNullString Then
                strHoleName = strObjectType + "-" + strLocationID & "_" & Format(nCount, strCountFormat)
            Else
                strHoleName = strObjectType + "-" + Format(nCount, strCountFormat)
            End If

            oHoleTraceNamedItem.Name = strHoleName
        End If
    End If
    
       
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oHoleTraceNamedItem = Nothing
    Set oParent = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub

ErrorHandler:
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oHoleTraceAE As Object) As IJElements

    Const Method = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild

    On Error Resume Next
    Set oSysChild = oHoleTraceAE
    Set oSysParent = oSysChild.GetParent
    
    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oSysParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oSysParent)
    End If


    Set oSysParent = Nothing
    Set oSysChild = Nothing
Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description

End Function

Private Function GetBaseName(ByVal oParent As Object) As String
    
    Const Method = "GetBaseName"
    
    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild
    Dim oNamedItem As IJNamedItem
    Dim strBaseName As String
    
    On Error Resume Next
    
    Set oSysChild = oParent
    Set oNamedItem = oSysChild
    
    On Error GoTo ErrorHandler
    
    If oSysChild Is Nothing Or oNamedItem Is Nothing Then
        GetBaseName = "ProjectRoot"
        GoTo Cleanup
    End If
    
    strBaseName = oNamedItem.Name
    
    Set oSysParent = oSysChild.GetParent
    Set oNamedItem = oSysParent
    Do Until InStr(oNamedItem.TypeString, "ProjectRoot") <> 0
        strBaseName = oNamedItem.Name & "_" & strBaseName
        
        On Error Resume Next
        
        Set oSysChild = oSysParent
        Set oSysParent = oSysChild.GetParent
        Set oNamedItem = oSysParent
        
        On Error GoTo ErrorHandler
    Loop

    GetBaseName = strBaseName
    
Cleanup:
    Set oNamedItem = Nothing
    Set oSysParent = Nothing
    Set oSysChild = Nothing
    
    Exit Function
    
ErrorHandler:
    GetBaseName = ""
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    GoTo Cleanup
End Function
 