VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a naming rule
'
'  Author: Kamrooz K. Hormoozi oss/kaho

'  History
'   Apr 21 2009     Kishore     DM-CP-163399  Performance problem when creating a new assembly
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "CommonNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "000000"       ' define fixed-width number field for name
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.
Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub
'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the  string "A" and an Index.
' The Index is unique for the Asembly.
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
    Dim strNamedParentsString As String
    Dim oNameCounter As IJNameCounter
    Dim oChildName As IJNamedItem
    Dim strChildName As String
    Dim strParentName As String
    Dim nCount As Long
    Dim strPartname As String
    Dim oParentName As IJNamedItem
    Dim strLocationID As String

    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

       
    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    
    Set oChildName = oObject
    strPartname = oChildName.TypeString
    
    strNamedParentsString = oActiveEntity.NamingParentsString
    
    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    Dim oEntity As Object
    
    If strPartname <> strNamedParentsString And oChildName.Name = "New Assembly" Then
       oActiveEntity.NamingParentsString = strPartname
       Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
       nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
       strChildName = "A" & strLocationID & Format(nCount, strCountFormat)
       oChildName.Name = strChildName
    ElseIf strPartname <> strNamedParentsString And oChildName.Name = "" Then
       oActiveEntity.NamingParentsString = strPartname
       Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
       nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
       strChildName = "A" & strLocationID & Format(nCount, strCountFormat)
       oChildName.Name = strChildName
    Else
       oActiveEntity.NamingParentsString = strPartname
    End If
   
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oEntity = Nothing

    Set oNameCounter = Nothing
    Set oChildName = Nothing
    Set oParentName = Nothing
    
Exit Sub
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parent added here is root of assembly hierarchy and is used in computing
' the name of the object in ComputeName() of the same interface. Both these methods are called from
' naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oGrandParent As Object
    Set oGrandParent = GetAssemblyTreeRoot()
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oGrandParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oGrandParent)
    End If

Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

' ***************************************************************************
'
' Function
'   AsIJNamedItem()
'
' Abstract
'   Returns the IJNamedItem interface for the given entity
'
' Usage
'   Use to quickly get access to the IJNamedItem interface without the
'   need for instanciating a variable.
'
' ***************************************************************************

Public Function AsIJNamedItem(ByVal Entity As Object) As IJNamedItem
    Set AsIJNamedItem = Entity
End Function

' ***************************************************************************
'
' Function
'   GetAssemblyName()
'
' Abstract
'   Returns the name of an assembly
'
' ***************************************************************************

Public Function GetAssemblyName(ByVal oAssembly As GSCADAsmHlpers.IJAssembly) As String
    GetAssemblyName = AsIJNamedItem(oAssembly).Name
End Function

' ***************************************************************************
'
' Function
'   PutAssemblyName()
'
' Abstract
'   Sets the name of an assembly entity
'
' ***************************************************************************

Public Sub PutAssemblyName(ByVal oAssembly As GSCADAsmHlpers.IJAssembly, ByVal strNewName As String)
    AsIJNamedItem(oAssembly).Name = strNewName
End Sub
 
