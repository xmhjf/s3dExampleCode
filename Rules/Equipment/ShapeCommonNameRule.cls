VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShapeCommonNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project: EquipNamingRules
'
'  Abstract: The file contains naming rule implementation for Shape
'
'  Author: Samba
'
'   History:
'       21 Jul, 2005    Samba           Initial Creation
'**************************************************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "ShapeCommonNameRule: "
Private Const MODELDATABASE = "Model"
Private Const sCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTROLE = "part"

Dim m_oErrors As IJEditErrors
Dim m_oLocalizer As IJLocalizer
    
Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & App.EXEName
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oLocalizer = Nothing
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
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim sModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()

    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    sModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(sModelDBID)

    Dim oNameCounter As IJNameCounter
    Dim oChildNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim sChildName As String
    Dim sParentName As String
    Dim sNamedParentsString As String
    Dim nCount As Long
    Dim sPartname As String
    Dim sLocation As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oChildNamedItem = oObject
    
    sPartname = m_oLocalizer.GetString(IDS_STRING_SHAPE, "Shape")
    
    ' Use part Name for basename, remove blanks:
    sPartname = Join(Split(sPartname), "")

    If elements.Count > 0 Then
        For Each oParent In elements
            On Error Resume Next
            Set oParentNamedItem = oParent
            On Error GoTo ErrorHandler

            sParentName = oParentNamedItem.Name
            If (Len(sChildName) = 0) Then
                sChildName = sParentName
            Else
                sChildName = sChildName + "-" + sParentName
            End If

            Set oParentNamedItem = Nothing
        Next oParent

        sNamedParentsString = oActiveEntity.NamingParentsString

 'Check if New parent name string constructed and old parent name string existing are same
        If (sChildName + sPartname) <> sNamedParentsString Then
            oActiveEntity.NamingParentsString = sChildName + sPartname
        
            sLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, sChildName + sPartname, sLocation)
            
            sChildName = sChildName & "-" & sPartname & "-" & Format(nCount, sCountFormat)
            
            oChildNamedItem.Name = sChildName
        End If
    Else
 ' no parents - use just the strPartName as the base
        sNamedParentsString = oActiveEntity.NamingParentsString

 'Check if New parent name string constructed and old parent name string existing are same
        If sPartname <> sNamedParentsString Then
            oActiveEntity.NamingParentsString = sPartname

            sLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, sPartname, sLocation)
            
            If sLocation <> vbNullString Then
                sChildName = sPartname & "-" & sLocation & "-" & Format(nCount, sCountFormat)
            Else
                sChildName = sPartname & "-" & Format(nCount, sCountFormat)
            End If
            oChildNamedItem.Name = sChildName
        End If
    End If
   
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oChildNamedItem = Nothing
    Set oParent = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub

ErrorHandler:
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
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oUnkParentEqp As IUnknown
    Dim oShape As IJShape

    On Error Resume Next
    Set oShape = oEntity
    oShape.GetParent oUnkParentEqp

    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    If Not (oUnkParentEqp Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oUnkParentEqp)
    End If

    Set oUnkParentEqp = Nothing
    Set oShape = Nothing
Exit Function

ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
