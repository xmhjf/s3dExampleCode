VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultReportRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'***************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project: DrawingNameRules
'
'  Abstract: The file contains an implementation of the default report rule
'            for the CDrawingSheet class.
'
'  History:
'   08-Jan-2004         B. Covington * Code Review by:
'       CR54157  Copied and modified from SpaceRules project
'   01-Oct-2014         npottaba
'       TR-CP-259688    Recorded Exception minidump at DefaultSheetRule.cls @ 101
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Implements IJNameRule

Private Const MODULE = "DefaultReportRule"
Private Const strCountFormat = "0000"   'define fixed-width number field

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object name. It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: SheetName  = Parents Name + LocationID + Unique Index
'***************************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo errorhandler
    Dim oChildNamedItem     As IJNamedItem
    Dim strChildName        As String
    Dim oParentNamedItem    As IJNamedItem
    Dim oParents            As IJElements
    Dim JContext            As IJContext
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim strModelDBID        As String
    Dim oModelResourceMgr   As IUnknown
    Dim oModelPOM           As IJDPOM
    Dim strParentDBID       As String
    Dim unkMoniker          As IUnknown
    Dim oNameCounter        As IJNameCounter
    Dim strLocation         As String
    Dim lCount              As Long
    Dim oReportOcc          As IJReportsSheet
    Dim oReport             As IJReport
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set oModelPOM = oModelResourceMgr
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    On Error Resume Next
    Set oChildNamedItem = pEntity
    Set oReportOcc = pEntity
    If Not oReportOcc Is Nothing Then
        Set oReport = oReportOcc.TemplateReport
    End If
    Set oParents = IJNameRule_GetNamingParents(pEntity)
    On Error GoTo errorhandler
    
    If Not oParents Is Nothing Then
        If 0 < oParents.Count Then
            Set oParentNamedItem = oParents.Item(1)
        End If
    End If
    Debug.Print TypeName(pEntity)
    Set oParents = Nothing
    
    'Get the Report's name
    If Not oReport Is Nothing Then
        strChildName = oReport.Name
    Else
        ' use SnapIn name
        If oParentNamedItem Is Nothing Then
            GoTo errorhandler
        End If
        strChildName = oParentNamedItem.Name
    End If
    
    'Get the parent's moniker
    Set unkMoniker = oModelPOM.GetObjectMoniker(oParentNamedItem)
    
    'Get the database identifier for the parent
    strParentDBID = oModelPOM.DbIdentifierFromMoniker(unkMoniker)
    
    ' concatonate dbid, name, and module name for unique identifier in name generator
    ' VolumeParentRule uses name + DBID
    strParentDBID = strChildName & strParentDBID & MODULE
    
    strLocation = vbNullString
    'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
    lCount = oNameCounter.GetCountEx(oModelResourceMgr, strParentDBID, strLocation)
    
    'Add LocationID, if available
    If strLocation <> vbNullString Then
        strChildName = strChildName & "-" & strLocation & "-" & Format(lCount, strCountFormat)
    Else
        strChildName = strChildName & "-" & Format(lCount, strCountFormat)
    End If
    
    oChildNamedItem.Name = strChildName
    
    Set oChildNamedItem = Nothing
    Set oParentNamedItem = Nothing
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oModelPOM = Nothing
    Set unkMoniker = Nothing
    Set oNameCounter = Nothing
    Set oReport = Nothing
    
Exit Sub
errorhandler:
    m_oErrors.Add Err.Number, "DefaultSheetRule::IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'****************************************************************************************************
'Author - HRM
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo errorhandler
    
    ' create an empty collection to return
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Dim oDwgSheetSnapIn As IJDDwgSheetSnapIn
    
    On Error Resume Next
    Set oDwgSheetSnapIn = pEntity
    On Error GoTo errorhandler
    
    If Not oDwgSheetSnapIn Is Nothing Then
        ' if we get the DwgSheetSnapIn interface, then return the parent snapin(s)
        Set IJNameRule_GetNamingParents = oDwgSheetSnapIn.GetParentSnapIns
    End If
    
    Set oDwgSheetSnapIn = Nothing
    
Exit Function
errorhandler:
    m_oErrors.Add Err.Number, "DefaultSheetRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function



