VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VolumeNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  Project: DrawingNameRules
'
'  Abstract: The file contains an implementation of the Volume Name naming rule
'            for the CDrawingSheet class. It uses the associated volume of a Volume
'            Drawing Type to produce the name of the drawing.
'
'  History:
'   29-Aug-2006       B. Covington * Code Review by: J Haley
'       CR88705  Drawing naming rule for volume drawings should match drawing and volume names.
'           (Copied and modified the DefaultSheetRule class)
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Implements IJNameRule

Private Const MODULE = "VolumeNameSheetRule"

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
    Dim oNameCounter        As IJNameCounter
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set oModelPOM = oModelResourceMgr
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    Set oChildNamedItem = pEntity
    
    oChildNamedItem.Name = "Unspecified"
    
    Set oParents = IJNameRule_GetNamingParents(pEntity)
    
    If Not oParents Is Nothing Then
        If 0 < oParents.Count Then
            Set oParentNamedItem = oParents.Item(1)
        End If
    End If
    
    Set oParents = Nothing
    
    'Get the Volume's name
    If Not oParentNamedItem Is Nothing Then
        strChildName = oParentNamedItem.Name
    End If
    
    'Set Volume name as Drawing name
    If Not oChildNamedItem Is Nothing Then
        oChildNamedItem.Name = strChildName
    End If
    
    Set oParents = Nothing
    Set oModelPOM = Nothing
    Set oChildNamedItem = Nothing
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub
errorhandler:
    m_oErrors.Add Err.Number, "VolumeNameSheetRule::IJNameRule_ComputeName", Err.Description
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
    
    Dim oDwgSheet2 As IJDDwgSheet2
    Dim oDwgVolume As IJDwgVolume
    
    On Error Resume Next
    Set oDwgSheet2 = pEntity
    On Error GoTo errorhandler
    
    Dim oTargetObjectCol As IJDTargetObjectCol 'DCollectionHelper
    Dim oRelation As IJDAssocRelation 'DRelationHelper
    
    Set oRelation = oDwgSheet2
    
    If Not oRelation Is Nothing Then
        Set oTargetObjectCol = oRelation.CollectionRelations(CVar("IJDDwgSheet2"), "SheetHasViews_DEST") ' Xmadefrom relationship
        
        If Not oTargetObjectCol Is Nothing Then
            Set oRelation = oTargetObjectCol.Item(1)
        Else
            Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get collection of Drawing Views from Drawing Sheet."
        End If
    Else
        Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get Relation Helper from Drawing Sheet."
    End If
    
    If Not oRelation Is Nothing Then
        Set oTargetObjectCol = oRelation.CollectionRelations(CVar("IJDDwgView2"), "ViewHasBroker_DEST") ' Xmadefrom relationship
        
        If Not oTargetObjectCol Is Nothing Then
            Set oRelation = oTargetObjectCol.Item(1)
        Else
            Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get collection of Element Brokers from Drawing View."
        End If
    Else
        Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get Relation Helper from Drawing View."
    End If
    
    If Not oRelation Is Nothing Then
        Set oTargetObjectCol = oRelation.CollectionRelations(CVar("IJDElementBrokerRelations"), "BrokerHasViewInfo_DEST") ' Xmadefrom relationship
        
        If Not oTargetObjectCol Is Nothing Then
            Set oRelation = oTargetObjectCol.Item(1)
        Else
            Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get collection of Drawing View Infos from Element Broker."
        End If
    Else
        Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get Relation Helper from Element Broker."
    End If
    
    If Not oRelation Is Nothing Then
        Set oTargetObjectCol = oRelation.CollectionRelations(CVar("IJDwgViewInfo"), "ViewInfoHasVolumes_DEST") ' Xmadefrom relationship
        
        If Not oTargetObjectCol Is Nothing Then
            Set oDwgVolume = oTargetObjectCol.Item(1)
        Else
            Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get collection of Drawing Volumes from Drawing View Info."
        End If
    Else
        Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get Relation Helper from Drawing View Info."
    End If
    
    If Not oDwgVolume Is Nothing Then
        ' if we get the DwgSheetSnapIn interface, then return the parent snapin(s)
        Dim oElements As IJElements
        Set oElements = New Elements
        oElements.Add oDwgVolume
        Set IJNameRule_GetNamingParents = oElements
        Set oElements = Nothing
    Else
        Err.Raise vbObjectError, "VolumeNameSheetRule::IJNameRule_GetNamingParents", "Failed to get Drawing Volume associated to this Drawing."
    End If
    
   
Exit Function
errorhandler:
    m_oErrors.Add Err.Number, "VolumeNameSheetRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function


