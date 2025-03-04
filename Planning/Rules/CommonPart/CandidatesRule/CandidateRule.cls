VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CandidateRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonPartRules\ComparisonRule\
'  File        : IJCmnPartCandiate.cls
'
'  Description : Common Part Candiates Rule
'
'  History     :
'   14th June 2010      Triveni CH     Initial creation
'**************************************************************************************

Option Explicit

Implements IJCmnPartCandiate

Private Function GetAssembliesToBeRemoved() As IJElements
On Error GoTo ErrorHandler

    Set GetAssembliesToBeRemoved = New JObjectCollection

'' Assemblies not participating in Common Part process can be returned from this method

wrapup:
Exit Function
ErrorHandler:
    MsgBox Err.Description, , vbInformation, "GetAssembliesToBeRemoved"
    GoTo wrapup
End Function

Private Function IJCmnPartCandiate_GetAssemblies(ByVal pSelAssembly As Object, ByVal lGroupBy As Long) As Object
On Error GoTo ErrorHandler

    Dim oAllSubAssemblies                   As IJElements
    Dim oPlnIntHelper                       As IJDPlnIntHelper
    Dim oSelAssembly                        As IJAssemblyBase
    Dim oMoniker                            As IUnknown
    Dim oPOM                                As IJDPOM
    
    Set oPlnIntHelper = New CPlnIntHelper
    Set oSelAssembly = pSelAssembly
    
    Set oPOM = GetResMgrByDBType("Model")
    Set oMoniker = oPOM.GetObjectMoniker(oSelAssembly)
    
    If lGroupBy = 1 Then 'Each sub-assembly
        Set oAllSubAssemblies = oPlnIntHelper.GetStoredProcAssemblyChildren(oSelAssembly, "IJAssemblyBase", True)
        oAllSubAssemblies.Add oMoniker
    Else
        Set oAllSubAssemblies = New JMonikerCollection
        oAllSubAssemblies.Add oMoniker
    End If
    
    'Remove the read-only assemblies
    Dim lSubAsmCount As Long
    Dim lIndex       As Long
    Dim oEachMnkr    As IUnknown
    Dim oIJDObject   As IJDObject
    Dim oAsBase      As IJAssemblyBase
    
    lSubAsmCount = oAllSubAssemblies.Count
    
    If lSubAsmCount > 1 Then
    
        For lIndex = lSubAsmCount To 1 Step -1
            Set oIJDObject = oPOM.GetObject(oAllSubAssemblies.Item(lIndex))
            Set oAsBase = oIJDObject
            
            If oAsBase.Type = 16 Then
                oAllSubAssemblies.Remove lIndex
            ElseIf Not oIJDObject Is Nothing Then
                If ((oIJDObject.AccessControl And acUpdate) <> acUpdate) Then
                    oAllSubAssemblies.Remove lIndex
                End If
            End If
        Next
    
    End If
    
    Set IJCmnPartCandiate_GetAssemblies = oAllSubAssemblies
    
Exit Function
ErrorHandler:
'    MsgBox Err.Description
'    GoTo wrapup
End Function

Private Function IJCmnPartCandiate_GetCandidates(ByVal pSelAssembly As Object, ByVal pClassItem As Object) As Object
On Error GoTo ErrorHandler

    Dim oPlnCandidateHelper                 As IJPlnCandidateHelper
    Dim oAssyChildrenOfReqdType             As IJElements
    Dim lGroupBy                            As Long
    Dim sObjType                            As String
    
    Set oPlnCandidateHelper = New CPlnCandidateHelper
    
    sObjType = GetObjectType(pClassItem)
    
    lGroupBy = GetGroupBy(pClassItem)
    
    oPlnCandidateHelper.SetInputs pClassItem, pSelAssembly, lGroupBy
    
    Select Case sObjType

        Case "Plates"
            oPlnCandidateHelper.Include PlanarPlates
            oPlnCandidateHelper.Exclude StandAlonePlates
            oPlnCandidateHelper.Exclude BracketsByPlane

'The following instructions and examples are helpful in adding the Filter criteria.
' For "IncludeByFilterName/ExcludeByFilterName" , the argument is predefined Catalog Filter Name.
' For "IncludeByFilterQuery/ExcludeByFilterQuery" , the SQL query which returns the oids of Parts.
' For ExcludeByProperty first argument is 'PlateType' second is its value
' "1 = Decks", "2 = TBH", "3 = LBH" ....

'The following are some examples
'       oPlnCandidateHelper.IncludeByFilterName "Deck Plate Parts"
'       oPlnCandidateHelper.ExcludeByFilterName "Air Tight Transverse Bulkhead Plate Parts"
'       oPlnCandidateHelper.ExcludeByFilterQuery "select oid from JPlate where PlateType = 2 and oid in (select oid from JPlatPart)"
'       oPlnCandidateHelper.ExcludeByProperty "PlateType", "1"
             
        Case "Profiles"
            oPlnCandidateHelper.Include Stiffeners
            oPlnCandidateHelper.Exclude StandAloneProfiles
            oPlnCandidateHelper.Include EdgeReinProfiles
'            oPlnCandidateHelper.Exclude KnuckledProfiles
            oPlnCandidateHelper.Exclude NonLinearProfiles
        
        Case "Brackets"
            oPlnCandidateHelper.Include TrippingBrackets
            oPlnCandidateHelper.Include BracketsByPlane
        
        Case "Collars"
            oPlnCandidateHelper.Include CollarParts
        Case "Members"
            oPlnCandidateHelper.Include MemberParts
        
        Case "Designed Members"
            oPlnCandidateHelper.Include BuiltUpMemberAssemblies
    End Select
    
    oPlnCandidateHelper.Include UnAssignedParts
'    oPlnCandidateHelper.Include AllFailedParts
    oPlnCandidateHelper.Exclude PartsInApprovedGroups
    oPlnCandidateHelper.Exclude PartsInManualGroups
    
    Set oAssyChildrenOfReqdType = oPlnCandidateHelper.GetCandidates
   
    Set IJCmnPartCandiate_GetCandidates = oAssyChildrenOfReqdType

wrapup:
    Set oAssyChildrenOfReqdType = Nothing
    Set oPlnCandidateHelper = Nothing

Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, "IJCmnPartCandiate_GetCandidates"
    GoTo wrapup
End Function

Private Function GetObjectType(oClassItem As Object) As String
Const METHOD = "GetObjectType"
On Error GoTo ErrorHandler
        
    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle
    Dim varIID                  As Variant
    Dim oAttributeMetaData      As IJDAttributeMetaData
    Dim oAttributes             As IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Model")

    Set oAttributeMetaData = oClassItem
    varIID = oAttributeMetaData.IID("IJCommonPartDef")
    
    Set oAttributes = oClassItem
    Set oAttributesCol = oAttributes.CollectionOfAttributes(varIID)
    
    For Each oAttribute In oAttributesCol
        If oAttribute.AttributeInfo.Name = "ObjectTypes" Then
            GetObjectType = oIJDCodeListMetaData.ShortStringValue("ObjectTypes", oAttribute.Value)
            Exit For
        End If
    Next oAttribute

    Set oAttributes = Nothing
    Set oAttributesCol = Nothing
    Set oIJDCodeListMetaData = Nothing
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function


Private Function GetGroupBy(oClassItem As Object) As Long
Const METHOD = "GetGroupBy"
On Error GoTo ErrorHandler

    Dim varIID                  As Variant
    Dim oAttributeMetaData      As IJDAttributeMetaData
    Dim oAttributes             As IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute

    Set oAttributeMetaData = oClassItem
    varIID = oAttributeMetaData.IID("IJCommonPartDef")
    
    Set oAttributes = oClassItem
    Set oAttributesCol = oAttributes.CollectionOfAttributes(varIID)
    
    For Each oAttribute In oAttributesCol
        If oAttribute.AttributeInfo.Name = "GroupBy" Then
            GetGroupBy = oAttribute.Value
            Exit For
        End If
    Next oAttribute

    Set oAttributes = Nothing
    Set oAttributesCol = Nothing
    
Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, METHOD
End Function

Private Function GetResMgrByDBType(strDBType As String) As Object
On Error GoTo ErrorHandler

    Dim jContext            As IJContext
    Dim oResourceMgr        As IUnknown
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim strDBID             As String
    Dim oPOM                As IJDPOM
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strDBID = oDBTypeConfig.get_DataBaseFromDBType(strDBType)
    Set oResourceMgr = oConnectMiddle.GetResourceManager(strDBID)
    
    If Not oResourceMgr Is Nothing Then
        Set GetResMgrByDBType = oResourceMgr
    End If

Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, "GetResMgrByDBType"
End Function
