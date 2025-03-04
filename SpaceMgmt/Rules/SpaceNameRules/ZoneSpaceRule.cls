VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZoneSpaceRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrSpaceNameRules
'
'  Abstract: The file contains an implementation of the default naming rule
'            for the Zone object in Space Management UE.
'
'  History:
'   Arnab           6th Aug 2001            Creation
'  Arunav           30th April 2002         Moved to present directory viz M:\CatalogData\Symbols\NameRules\SpaceNameRule
'                                           Dependency on CommonSpace removed
'       Sadish Sadasivam Oct/07/2003
'           TR CP49909: fixed the usage of LocationID
'  Raman            13th Oct 2003          TR:49909 Further changes are made
'                                          Removed the method "NextSpace_Name" and instead used "GetCountEx" of CommonApp for unique naming .
'                                          Removed the method "IsNameInNode" as we no longer bother about the new name if a space is moved from one folder to other folder
'  Arnab            24th Oct 2003          TR-CP·51106  The naming rules for Space management have references to client components
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Implements IJNameRule

Private Const MODULE = "ZoneSpaceRule: "
Private Const strCountFormat = "0000"   'define fixed-width number field

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'*********************************************************************************************
'  Author HRM
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object name. It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: ZoneName  = Zone description + Unique Index
'***************************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    Dim oChildNamedItem             As IJNamedItem
    Dim strChildName                As String
    Dim oParentNamedItem            As IJNamedItem
    Dim oSpaceChild                 As IJSpaceChild

    Dim JContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oModelResourceMgr As IUnknown
    Dim oNameCounter As IJNameCounter
    Dim strLocation As String
    Dim nCount As Long
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    Set oSpaceChild = pEntity
    Set oChildNamedItem = pEntity

    Set oParentNamedItem = oSpaceChild.GetParent
    
    'Get the parent's name
    strChildName = oParentNamedItem.Name
    
    strLocation = vbNullString
    'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName, strLocation)
    
    'Add LocationID, if available
    If strLocation <> vbNullString Then
        strChildName = strChildName & "-" & strLocation & "-" & Format(nCount, strCountFormat)
    Else
        strChildName = strChildName & "-" & Format(nCount, strCountFormat)
    End If
    
    oChildNamedItem.Name = strChildName
    
    Set oChildNamedItem = Nothing
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oSpaceChild = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, "ZoneSpaceRule::IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'****************************************************************************************************
'Author - HRM
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. Dummy function which does nothing
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Dim oSpaceChild         As IJSpaceChild
    Set oSpaceChild = pEntity
    IJNameRule_GetNamingParents.Add oSpaceChild.GetParent
    Set oSpaceChild = Nothing
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "ZoneSpaceRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
