VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipeClampNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2005, Intergraph Corporation.  All Rights Reserved.
'
' File: PipeCompNameRule.cls
'
' Author: Sreelekha
'
' Abstract: Implementation of Pipe Clamp naming rule
'
' History:
'
' 22-Feb-2005      Bharath      Removed the dependancy on the client tier component
'                               i.e. K:\CommonRoute\Client\Bin\RteErrorUtils.dll.
'                               included the reference to X:\Shared\Bin\ErrorLog.dll.
'
'*************************************************************************************************************
Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors 'middle tier errors
Private Const E_FAIL = -2147467259
Private Const Module = "ClampNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"
Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub
Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''IJNameRule''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''/////////////////////////////////////////////////////////////////////////////
'' METHOD:         IJNameRule_ComputeName
' Description:
'  Creates a name for the object passed in. The Naming Parents are added in AddNamingParents()
'  of the same interface. Both these methods are called from naming rule semantic.
'
' Arguments:
'  oObject       - Input.  Child object that needs to have the NamingRule naming.
'  oParents      - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'/////////////////////////////////////////////////////////////////////////////
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    
    On Error GoTo ErrHndler
    
    If oEntity Is Nothing Then Exit Sub
    
    Dim oJContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oNameCounter As IJNameCounter
    Dim strLocation As String
    Dim strSequenceNumber As String
    Dim lcount As Long
    Dim oConnItemNamedItem As IJNamedItem
    Dim strConnItemName As String
    Dim oConnItemOwnerNamedItem As IJNamedItem
    Dim strConnItemOwnerName As String
    Dim strName As String
    Dim strParent As String
    
    'Get the middle context
    Set oJContext = GetJContext()
    Set oDBTypeConfig = oJContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = oJContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    'Get the connection item owner part name
    Set oConnItemOwnerNamedItem = oParents.Item(1)
    strConnItemOwnerName = oConnItemOwnerNamedItem.Name
    Set oConnItemOwnerNamedItem = Nothing
    
    Set oConnItemNamedItem = oEntity
    strConnItemName = oConnItemNamedItem.TypeString
    
    strParent = oActiveEntity.NamingParentsString
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
   
    If strConnItemName <> strParent Then
        oActiveEntity.NamingParentsString = strConnItemName
        strLocation = vbNullString
        lcount = oNameCounter.GetCountEx(oModelResourceMgr, strConnItemName, strLocation)
        strSequenceNumber = Format(lcount, strCountFormat)
        strName = strConnItemOwnerName + "-" + "Clamp" + "-" + strSequenceNumber
        oConnItemNamedItem.Name = strName
    End If
   
    Set oJContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oNameCounter = Nothing
    Set oConnItemNamedItem = Nothing
    
Exit Sub
   
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, Module & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub


'///////////////////////////////////////////////////////////////////////////////////////
' METHOD:         IJNameRule_GetNamingParents
' Description:
'  All the Naming Parents that need to Participate in an Objects Naming are added here to the
' IJElements collection. The Parents added here are used in Computing the Name of the Object in
' ComputeName() of the same Interface. Both these methods are called from Naming Rule Semantic.
'
' Arguments:
'  oEntity        - Input.  Child object that needs to have the NamingRule naming.
'
'/////////////////////////////////////////////////////////////////////////////
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
  
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If oEntity Is Nothing Then Exit Function
    
    Set oDesignChild = oEntity
    Set oDesignParent = oDesignChild.GetParent
    
    Call IJNameRule_GetNamingParents.Add(oDesignParent)
    
    Set oDesignChild = Nothing
    Set oDesignParent = Nothing
    
  Exit Function
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, Module & METHOD, Err.Description
    Err.Raise E_FAIL
End Function





