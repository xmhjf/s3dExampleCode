VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsmSysNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADHoleNamingRules
'
'  Abstract: The file contains a implementation of a AsmSystemNameRule for Hole Management
'
'  Author: Jang Choi
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "CAsmSysNameRule: "
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
    Dim oAsmNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim oAsmParent As IJAssembly
    Dim oAsmChild As IJAssemblyChild
    
    Dim strSystemBaseName As String
    Dim strAssemblyBaseName As String
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

'''    MsgBox "Naming Parents Count : " & oNamingParents.Count, , METHOD
    
    If oNamingParents.Count > 0 Then
        For Each oParent In oNamingParents
            If TypeOf oParent Is IJAssembly Then
                Set oAsmParent = oParent
                strAssemblyBaseName = GetAssemblyBaseName(oParent)
            ElseIf TypeOf oParent Is IJSystem Then
                strSystemBaseName = GetSystemBaseName(oParent)
            End If
        Next oParent
        
        
        On Error GoTo ErrorHandler

        strNamedParentsString = oActiveEntity.NamingParentsString
        If (strSystemBaseName + strObjectType) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strObjectType
            
            If Len(Trim(strSystemBaseName)) <> 0 Then
                nCount = oNameCounter.GetCountEx(oModelResourceMgr, strSystemBaseName, strLocationID)
                
                If Not strLocationID = vbNullString Then
                    strHoleName = strSystemBaseName & "_" & strObjectType & "_" & strLocationID & "_" & Format(nCount, strCountFormat)
                Else
                    strHoleName = strSystemBaseName & "_" & strObjectType & "_" & Format(nCount, strCountFormat)
                End If
                
                oHoleTraceNamedItem.Name = strHoleName
            End If

''' HoleTraceAE wasn't added in Assembly Tab of WorkSpaceNavigator.
''' If HoleTraceAE is added in Assembly Tab of WorkSpaceNavigator, follow code will be tested.
'''
'''            If Len(Trim(strAssemblyBaseName)) <> 0 Then
'''                nCount = oNameCounter.GetCount(oModelResourceMgr, strAssemblyBaseName)
'''
'''                strHoleName = strAssemblyBaseName & "_" & strObjectType & Format(nCount, strCountFormat)
'''
'''                MsgBox strHoleName, , METHOD
'''
'''                On Error Resume Next
'''                Set oAsmChild = oHoleTraceAE
'''                Set oAsmNamedItem = oAsmChild
'''                oAsmNamedItem.Name = strHoleName
'''                oAsmParent.AddChild oAsmChild
'''            End If
        End If
    Else
        ' no parents - use just the strObjectName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strObjectType <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strObjectType

            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strObjectType, strLocationID)
            
            If Not strLocationID = vbNullString Then
                strHoleName = strObjectType & "_" & strLocationID & "_" & Format(nCount, strCountFormat)
            Else
                strHoleName = strObjectType & "_" & Format(nCount, strCountFormat)
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

    Dim oAsmParent As Object
    Dim oAsmChild As IJAssemblyChild

    Dim oNamedItem As IJNamedItem
    Dim oAsmNamedItem As IJNamedItem
    Dim oColObject As IJDTargetObjectCol
    Dim obj As Object
    Dim index As Long
    Dim strSysName As String
    Dim strAsmName As String
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Set oSysChild = oHoleTraceAE
    Set oSysParent = oSysChild.GetParent
    
    
    Set oNamedItem = oSysParent
    strSysName = oNamedItem.Name
    Set oColObject = oSysParent.GetChildren()
    If Not oColObject Is Nothing Then
        For index = 1 To oColObject.Count
            Set obj = oColObject.Item(index)
            Set oNamedItem = obj
            
            If oNamedItem.Name = strSysName & "/1" Then
                Set oAsmChild = obj
                Set oAsmParent = oAsmChild.Parent
                Set oAsmNamedItem = oAsmParent
                
                If Not oAsmNamedItem Is Nothing Then
                    strAsmName = Trim(oAsmNamedItem.Name)
                    If Len(strAsmName) <> 0 Then
                        Call IJNameRule_GetNamingParents.Add(oAsmParent)
                        Exit For
                    End If
                End If
            End If
        Next index
    End If

    Call IJNameRule_GetNamingParents.Add(oSysParent)


    Set oSysParent = Nothing
    Set oSysChild = Nothing
    Set oAsmParent = Nothing
    Set oAsmChild = Nothing
    Set oNamedItem = Nothing
Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError Module, Method
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description

End Function

Private Function GetSystemBaseName(ByVal oParent As Object) As String
    
    Const Method = "GetSystemBaseName"
    
    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild
    Dim oNamedItem As IJNamedItem
    Dim strBaseName As String
    
    On Error Resume Next
    
    Set oSysChild = oParent
    Set oNamedItem = oSysChild
    
    On Error GoTo ErrorHandler
    
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

    GetSystemBaseName = strBaseName
    
    Set oNamedItem = Nothing
    Set oSysParent = Nothing
    Set oSysChild = Nothing

    Exit Function
    
ErrorHandler:
    GetSystemBaseName = ""
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description

End Function

Private Function GetAssemblyBaseName(ByVal oParent As Object) As String
    
    Const Method = "GetAssemblyBaseName"
    
    Dim oAsmParent As IJAssembly
    Dim oAsmChild As IJAssemblyChild
    Dim oNamedItem As IJNamedItem
    Dim strBaseName As String
    
    On Error Resume Next
    
    Set oAsmChild = oParent
    Set oNamedItem = oAsmChild
    
    On Error GoTo ErrorHandler
    
    strBaseName = oNamedItem.Name
    
    Set oAsmParent = oAsmChild.Parent
    Set oNamedItem = oAsmParent
    Do Until InStr(oNamedItem.TypeString, "ProjectRoot") <> 0
        strBaseName = oNamedItem.Name & "_" & strBaseName
        
        On Error Resume Next
        
        Set oAsmChild = oAsmParent
        Set oAsmParent = oAsmChild.Parent
        Set oNamedItem = oAsmParent
        
        On Error GoTo ErrorHandler
    Loop
    
    GetAssemblyBaseName = strBaseName
    
    Set oNamedItem = Nothing
    Set oAsmParent = Nothing
    Set oAsmChild = Nothing
    
    Exit Function
    
ErrorHandler:
    GetAssemblyBaseName = ""
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description

End Function
 