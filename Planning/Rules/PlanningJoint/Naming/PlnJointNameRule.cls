VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a naming rule
'
'  Author: Ole von Mortensen    oss/olmo
'******************************************************************

Option Explicit
Implements IJNameRule
                                                        
Private Const Module = "PlnJointNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000000"       ' define fixed-width number field for name
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
' Method
'   IJNameRule_ComputeName
'
' Description:
'   Creates a name for the PlanningJoint passed in. The name is based on the string "W" and
'   an Index. The Index is unique for the PlanningJoint. It is assumed that all Naming Parents
'   and the Object implement IJNamedItem. The Naming Parents are added in AddNamingParents()
'   of the same interface. Both these methods are called from naming rule semantic.
'
' Parameters:
'   [in] pEntity        Child object that needs to have the naming rule naming.
'   [in] pParents       Naming parents collection.
'   [in] pActiveEntity  Naming rules active entity on which the NamingParentsString is stored.
'
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oModelResourceMgr As IUnknown
    Dim oChildName As IJNamedItem
    Dim oNameCounter As IJNameCounter
    Dim oPhysConn As IJNamedItem
    
    Dim strModelDBID As String
    Dim strPartname As String
    Dim strChildName As String
    Dim nCount As Long
    Dim strLocationID As String
    
    Set oPhysConn = elements.Item(1)
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'First get resourcemanager frem modeldatabase where objects are to be stored
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
       
    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    
    'Get Typestring from IJNamedItem
    Set oChildName = oObject
    strPartname = oChildName.TypeString
    
    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    
    ' Get a string from the namecounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
    
    'Create name using the right format and set on object
    If Not strLocationID = vbNullString Then
        strChildName = oPhysConn.Name + "-W-" + strLocationID + "-" + Format(nCount, strCountFormat)
    Else
        strChildName = oPhysConn.Name + "-W" + Format(nCount, strCountFormat)
    End If
    
    oChildName.Name = strChildName
    
CleanUp:
    Set oPhysConn = Nothing
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing

    Set oNameCounter = Nothing
    Set oChildName = Nothing
    Exit Sub
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method
'   IJNameRule_GetNamingParents
'
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parent added here is root of assembly hierarchy and is used in computing
' the name of the object in ComputeName() of the same interface. Both these methods are called from
' naming rule semantic.
'
' Parameters:
'   [in] pEntity  Child object that needs to have the naming rule naming.
'   [out] IJNameRule_GetNamingParents, JObjectCollection with naming parents
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlnJoint As IJPlnJoint
    Set oPlnJoint = oEntity
    
    Dim oPhysConnection As Object
    Set oPhysConnection = oPlnJoint.GetPhysicalConnection
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oPhysConnection Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oPhysConnection)
    End If

CleanUp:
    Exit Function

ErrorHandler:
     m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Function
 
