VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniqueName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project: EquipCompNameRules
'
'  Abstract: The file contains Unique naming rule for Equipment Component
'
'  Author: Krishna Bathina
'
'   History:
'       17 Nov, 2004    Krishna Bathina     Initial Creation
'**************************************************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "EquipCompNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the Equipment Component passed in. The name is based on the parents
' name and object name. It is something like this: "Base Name" + "Object Name" + Index.
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
    Dim strModelDBID As String
    Dim strLocation As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    Dim oNameCounter As IJNameCounter
    Dim oChildNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim strChildName As String
    Dim strParentName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strPartname As String
    Dim oNamingRulesHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Dim oFilterFactory As GSCADMiddleFilters.MiddleFiltersFactory
    Dim oObjectTypeFilter As IJObjTypeFilter
    Dim oSimpleFilter As IJSimpleFilter
    Dim oModelPOM As IJDPOM
    Dim strCLSID As String
    Dim unkMoniker As IUnknown
    Dim bNameUnique As Boolean
    Dim strTempChildName As String
    Dim loopCount As Long
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If
    
    Set oFilterFactory = New GSCADMiddleFilters.MiddleFiltersFactory
    
    If Not oFilterFactory Is Nothing Then
        Set oObjectTypeFilter = oFilterFactory.CreateEntity(ObjTypeFilter, Nothing)
        If Not oObjectTypeFilter Is Nothing Then
            Set oModelPOM = oModelResourceMgr
            Set unkMoniker = oModelPOM.GetObjectMoniker(oObject)
            If Not unkMoniker Is Nothing Then
                strCLSID = oModelPOM.ClsidOf(unkMoniker)
                oObjectTypeFilter.clsid = strCLSID
                Set oSimpleFilter = oObjectTypeFilter
            End If
        End If
    End If
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    Set oNamingRulesHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Set oChildNamedItem = oObject
    bNameUnique = False
    ' Get name of part
    strPartname = GetPartName(oObject)
    If strPartname = "" Then
        ' if we cannot get the part name, then use the type string
        strPartname = oChildNamedItem.TypeString
    End If

    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    'Variable to maintain the loopcount , to avoid infinite loop situation
    loopCount = 1

    If elements.Count > 0 Then
        For Each oParent In elements
            On Error Resume Next
            Set oParentNamedItem = oParent
            On Error GoTo ErrorHandler

            strParentName = oParentNamedItem.Name

            If (Len(strChildName) = 0) Then
                strChildName = strParentName
            Else
                strChildName = strChildName + "-" + strParentName
            End If

            Set oParentNamedItem = Nothing
        Next oParent
        strTempChildName = strChildName
        strNamedParentsString = oActiveEntity.NamingParentsString
        
        'Check if New parent name string constructed and old parent name string existing are same
        If (strChildName + strPartname) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strChildName + strPartname
            Do While (Not bNameUnique)
                strChildName = strTempChildName
                
                strLocation = vbNullString
                nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName + strPartname, strLocation)
                
                If strLocation <> vbNullString Then
                    strChildName = strChildName & "-" & strPartname & "-" & strLocation & "-" & Format(nCount, strCountFormat)
                Else
                    strChildName = strChildName & "-" & strPartname & "-" & Format(nCount, strCountFormat)
                End If
                bNameUnique = oNamingRulesHlpr.IsGeneratedNameUnique(oObject, oSimpleFilter, strChildName)
                loopCount = loopCount + 1
                ' this is just to avoid infinite loop situation
                If loopCount > 50 Then
                    Exit Do
                End If
            Loop
                oChildNamedItem.Name = strChildName
        End If
    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strPartname <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname

            strLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocation)

            If strLocation <> vbNullString Then
                strChildName = strPartname & "-" & strLocation & "-" & Format(nCount, strCountFormat)
            Else
                strChildName = strPartname & "-" & Format(nCount, strCountFormat)
            End If
            oChildNamedItem.Name = strChildName
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
    Dim oEquipComponent As IJEquipmentComponent

    On Error Resume Next
    Set oEquipComponent = oEntity
    oEquipComponent.GetParent oUnkParentEqp
    
    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oUnkParentEqp Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oUnkParentEqp)
    End If

    Set oUnkParentEqp = Nothing
    Set oEquipComponent = Nothing
    Exit Function

ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

