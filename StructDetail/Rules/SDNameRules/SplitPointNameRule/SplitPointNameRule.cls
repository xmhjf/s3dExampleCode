VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitPointNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 ICC.
'  All rights reserved.
'
'  Project: SDSplitPointRules
'
'  Abstract: The file contains name rules for Split Points
'
'   History:
'   Chaitanya       8/17/2011           Creation
'******************************************************************
Option Explicit

Implements IJNameRule
                                                        
Private Const MODULE = "SDSplitPointRules.SplitPointNameRule."
Private Const MODELDATABASE = "Model"

Private m_oErrors As IJEditErrors

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
   
    Set m_oErrors = New IMSErrorLog.JServerErrors

Exit Sub

ErrorHandler:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "Class_Initialize", Err.Description
    Err.Raise E_FAIL, METHOD, MODULE
End Sub

Private Sub Class_Terminate()
    'clear globals
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Function:
'   IJNameRule_ComputeName
' Description:
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, _
                                   ByVal oParents As IJElements, _
                                   ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim oNamedItem      As IJNamedItem
    Dim oSysParent      As Object
    Dim strParentName   As String
    Dim oDesignChild    As IJDesignChild

    If oEntity Is Nothing Then
        Err.Raise E_INVALIDARG, MODULE, METHOD
    End If
    
    Set oDesignChild = oEntity
    Set oSysParent = oDesignChild.GetParent
    Set oNamedItem = oSysParent
    
    strParentName = oNamedItem.Name
    
    Set oNamedItem = Nothing
    

    Dim jContext            As IJContext
    Dim oModelResourceMgr   As IUnknown
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim strModelDBID        As String
    Dim strPartName         As String
    Dim oNameCounter        As IJNameCounter
    Dim nCount              As Long

    'Get the middle context
    Set jContext = GetJContext()

    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")

    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
       
    Set oNamedItem = oEntity
    
    strPartName = oNamedItem.TypeString
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    nCount = oNameCounter.GetCount(oModelResourceMgr, strPartName)
    
    oNamedItem.Name = strParentName & "-SP-" & nCount
    
Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description, MODULE
    Err.Raise E_FAIL, METHOD, MODULE
End Sub

'********************************************************************
' Function:
'   IJNameRule_GetNamingParents
' Description
'   All the Naming Parents that need to participate in an objects naming are
'   added here to the IJElements collection. The parents added here are used
'   in computing the name of the object in ComputeName() of the same
'   interface. Both these methods are called from naming rule semantic.
'********************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description, MODULE
    Err.Raise E_FAIL, METHOD, MODULE
End Function


