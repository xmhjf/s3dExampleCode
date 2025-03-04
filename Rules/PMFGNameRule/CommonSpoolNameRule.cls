VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonSpoolNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2004-2006 Intergraph Corporation.  All rights reserved.
'
'  Project: PMFGNameRule
'  Class:   CommonSpoolNameRule
'
'  Abstract: The file contains the Implementation for naming rule interface for Spools
'
'  Author:
'
'  History:
'  17-Jan-2007 Sreelekha   DM113222  oSpool.Pipeline is returning invalid object
'******************************************************************
Option Explicit

Implements IJNameRule
Const vbInvalidArg = &H80070057

Private Const MODULE = "CommonSpoolNameRule: "
Private Const E_FAIL = -2147467259
Dim m_oErrors As IJEditErrors  ' Middle tier errors

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName "
    Dim oSpoolNamedItem As IJNamedItem
    Dim oSpool As IJDSpool
    Dim strName As String
    Dim strParentName As String
    Dim oPipeLine As Object
    On Error GoTo ErrorHandler
    
    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If
       
    Set oSpoolNamedItem = pEntity
    Set oSpool = pEntity
    
    Set oPipeLine = pParents.Item(1)
    
    strName = m_GetSpoolName(oPipeLine, oSpoolNamedItem.TypeString, oSpool.SpoolID)
    
    oSpoolNamedItem.Name = strName
    

    
    Set oSpool = Nothing
    Set oPipeLine = Nothing
    Set oSpoolNamedItem = Nothing
        
    Exit Sub
    
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrorHandler

    Dim oSpool As IJDSpool
    Dim oAssemblyChild As IJAssemblyChild
    Dim oNamingParent As Object
     
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
   
   'Get the naming parent
    Set oSpool = pEntity
    
    On Error Resume Next
    Set oNamingParent = oSpool.SpoolableObject
    On Error GoTo ErrorHandler
 
    IJNameRule_GetNamingParents.Add oNamingParent
    
    Set oAssemblyChild = Nothing
    Set oNamingParent = Nothing
    Set oSpool = Nothing

    Exit Function
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function m_GetSpoolName(oPipeLine As Object, strUserName As String, SpoolNumber As String) As String

Const METHOD = "m_GetSpoolName "
Dim strSpoolName As String
Dim oParentNamedItem As IJNamedItem
Dim ParentName As String
Dim oNameGeneratorService As New NameGeneratorService
Dim jContext As IJContext
Dim oModelResourceMgr As IUnknown
Dim oDBTypeConfig As IJDBTypeConfiguration
Dim oConnectMiddle As IJDAccessMiddle
Dim strModelDBID As String
    
On Error GoTo ErrorHandler
    
    Set oParentNamedItem = oPipeLine
    
    ParentName = oParentNamedItem.Name

   
  '  strUserName = "Spool"
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    'The Name for the Spool is presently created using the Parent BO of the
    'Spool, and the TypeString of the Spool Object which is sent as inputs to
    'the NameGeneratorService to retrieve a unique name.
    'As Part of V4 Enhancements to the Spool, there was a requirement to
    'add the generated number to some starting number incase of the
    'Block BO. But for this According to specifications the Starting Spool
    'number is retrieved from the interface IJBlockItem which is a part of
    'the Block BO...But since it is not been finalized as to where this
    'interface should be made available, presently to create the Name,
    'we are not using this, once finalized, this code can be modified to get
    'the number from the generated name , added to the Starting Spool Number
    'and again append back to the generated Name.
    '           ----- Satish N Kota (April 7, 2003)
    
    'Generate Spool Number using name generator service
    strSpoolName = oNameGeneratorService.GetNameFromBasename(oModelResourceMgr, ParentName & "_" & strUserName)
     
    
    'strSpoolName = ParentName & "_" & strUserName & SpoolNumber
    


    m_GetSpoolName = strSpoolName
    
   
    Set oParentNamedItem = Nothing
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oNameGeneratorService = Nothing
    
Exit Function
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function



