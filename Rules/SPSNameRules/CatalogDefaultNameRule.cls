VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CatalogDefaultNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    CatalogDefaultNameRule.cls
'
'Author
'       Nov 2004        AP
'
'Description
'
'Notes
'
'History::
' Jul-31-2006 AS    Fixed TR101651. Check for smartocc before trying to set an object. Added standard header.
' Jun-26-2009 GG    TR#163724 NameRules are required for DM's driven by CanRules
'*******************************************************************

Implements IJNameRule
Private Const MODULE = "CatalogDefaultNameRule::"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
 
    Dim oNameCounter As IJNameCounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem
    Dim oChildNamedItem As IJNamedItem
  
    Dim strParentName As String
    Dim strChildName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strPartname As String
    Dim strlocation As String
    
    Dim iSO As IJSmartOccurrence
    Dim iSI As IJSmartItem
    'If the object is rule based Designed Member, use the rule's Part Description to compose the Name
    Dim oPart As IJDPart
    If TypeOf oObject Is ISPSDesignedMember Then
        Dim oMemPartCommon As ISPSMemberPartCommon
        Set oMemPartCommon = oObject
        Set oPart = GetPartFromRule(oMemPartCommon.Rule)
        If Not oPart Is Nothing Then
            strPartname = oPart.PartDescription
            strPartname = Join(Split(strPartname), "")
        End If
        Set oMemPartCommon = Nothing
    End If
    'Otherwise, the old way
    If oPart Is Nothing Then
        'Dont proceed if not smart occurrence
        If Not TypeOf oObject Is IJSmartOccurrence Then
            Exit Sub
        Else
            Set iSO = oObject
        End If
        
        Set iSI = iSO.ItemObject
          
        strPartname = iSI.Name
        Set iSI = Nothing
        Set iSO = Nothing
    End If
    Set oPart = Nothing
    Set oChildNamedItem = oObject
    
' Commented the code to get the NamingParent Name to be used to prefix for the name being generated
'   If elements.Count > 0 Then
'        For Each oParent In elements
'            On Error Resume Next
'            Set oParentNamedItem = oParent
'            On Error GoTo ErrHandler:
'
'            strParentName = oParentNamedItem.Name
'
'            If (Len(strChildName) = 0) Then
'                strChildName = strParentName
'            Else
'                strChildName = strChildName + "-" + strParentName
'            End If
'
'            Set oParentNamedItem = Nothing
'        Next oParent
'
'        strNamedParentsString = oActiveEntity.NamingParentsString
'
'        'Check if New parent name string constructed and old parent name string existing are same
'        If strChildName <> strNamedParentsString Then
'            oActiveEntity.NamingParentsString = strChildName
'
'            nCount = oNameCounter.GetCount(oModelResourceMgr, strPartname)
'
'            strChildName = strChildName + "-" + strPartname + "-" + Format(nCount, strCountFormat)
'            oChildNamedItem.Name = strChildName
'        End If
'    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strPartname <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname
            
            strlocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strlocation)

            If Not (strlocation = vbNullString) Then
                strChildName = strPartname + "-" + strlocation + "-" + Format(nCount, strCountFormat)
            Else
                strChildName = strPartname + "-" + Format(nCount, strCountFormat)
            End If
            
            oChildNamedItem.Name = strChildName
        End If
'    End If
            
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oChildNamedItem = Nothing
    Set oParent = Nothing
    Set oParentNamedItem = Nothing
    

Exit Sub

ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
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
    On Error GoTo ErrHandler
' Commented the code to get the NamingParent Name to be used to prefix for the name being generated
'    Dim oSysParent As IJSystem
'    Dim oSysChild As IJSystemChild

'    On Error Resume Next
'    Set oSysChild = oEntity
'    Set oSysParent = oSysChild.GetParent

'    On Error GoTo label

'    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

'    If Not (oSysParent Is Nothing) Then
'        Call IJNameRule_GetNamingParents.Add(oSysParent)
'    End If

    Set IJNameRule_GetNamingParents = Nothing
    
'    Set oSysParent = Nothing
'    Set oSysChild = Nothing
Exit Function
ErrHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'Used here to get IJDPart from SPSCanRule
Public Function GetPartFromRule(oRule As Object) As Object
    Const METHOD = "GetPartFromRule"
    On Error GoTo errorHandler
    
    Dim iAssocRelations As IJDAssocRelation
    Dim iRelationshipCol As IJDRelationshipCol
    Dim iRelationship As IJDRelationship
    Dim iIJDProxy As IJDProxy
    If oRule Is Nothing Then
        Exit Function
    End If
    Set iAssocRelations = oRule
    Set iRelationshipCol = iAssocRelations.CollectionRelations("IJGeometricConstruction", "Type")
    
    If Not iRelationshipCol Is Nothing And iRelationshipCol.Count > 0 Then
        Set iRelationship = iRelationshipCol.Item(1)
        Set iIJDProxy = iRelationship.Target
        Set GetPartFromRule = iIJDProxy.Source
    End If

    Exit Function

errorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
