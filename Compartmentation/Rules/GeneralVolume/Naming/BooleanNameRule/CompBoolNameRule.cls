VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompBoolNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompBoolNameRule
'
'  Abstract:   The file contains an implementation of the boolean naming rule
'              for the Compartment object in Compartmentation UE.

'  Description:This name rule computes the compartment name based on the catalog part description and
'              adds '+' for the added boolean feature name and '-' for the subtracted boolean feature name if there are any features added to the compartment

'  Example:    If a compartment A is being created with B and C considerd as Boolean Add and Boolean subtract respectively.
'              Then, this name rule changes the name of B as '+ A' and C as '- A'
'
'  History:
'     Apparao 19th Feb 2010
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259
  
Implements IJNameRule

Private Const MODULE = "CompBoolNameRule: "
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
    
    Dim oChildNamedItem                  As IJNamedItem
    Dim strChildName                     As String
    Dim oSpacePart                       As IJDPart
    Dim oProxyPart                       As Object
    Dim oAssocRelation                   As IJDAssocRelation
    Dim oTargetObjCol                    As IJDTargetObjectCol
    Dim oRelationHelper                  As IMSRelation.DRelationHelper
    Dim oCollectionHelper                As IMSRelation.DCollectionHelper
    Dim oProxy                           As IJDProxy

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
    
    Set oChildNamedItem = pEntity
    
    If oChildNamedItem.Name <> vbNullString Then GoTo cleanup
            
    'Get Part from Relation
    Set oRelationHelper = pEntity
    Set oCollectionHelper = oRelationHelper.CollectionRelations("IJCompartEntity", "Part")
    
    If oCollectionHelper.Count = 1 Then
        Set oProxyPart = oCollectionHelper.Item(1)
    End If
        
    Set oProxy = oProxyPart
    Set oSpacePart = oProxy.Source

    strChildName = oSpacePart.PartDescription
    
    strLocation = vbNullString
    'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName, strLocation)
               
    'Add LocationID, if available
    If strLocation <> vbNullString Then
        strChildName = strChildName & "-" & strLocation & "-" & Format(nCount, strCountFormat)
    Else
        strChildName = strChildName & "-" & Format(nCount, strCountFormat)
    End If
    
    Set oChildNamedItem = pEntity
    oChildNamedItem.Name = strChildName
    
    'Set the name for boolean features
    SetNameForBoolFeatures pEntity, strChildName
    
    
cleanup:
    Set oChildNamedItem = Nothing
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oSpacePart = Nothing
    Set oProxy = Nothing
    
Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, "CompBoolNameRule::IJNameRule_ComputeName", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_NAMINGRULES_FAILED_PARTDESCRIPTIONRULE)
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
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompBoolNameRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

'This function adds '+' for the added feature and '-' for the subtracted feature
Private Sub SetNameForBoolFeatures(pEntity As Object, strParentName As String)
Const METHOD = "SetNameForBoolFeatures"
On Error GoTo ErrorHandler

    Dim oCompartCreation                          As IJCompartCreation
    Dim oAddFeatures                              As IJElements
    Dim oRemoveFeatures                           As IJElements
    Dim oNamedItem                                As IJNamedItem
    Dim i                                         As Long
    
    Set oCompartCreation = pEntity
    Set oAddFeatures = New JObjectCollection
    Set oRemoveFeatures = New JObjectCollection
    
    oCompartCreation.GetBoolFeatures oAddFeatures, oRemoveFeatures
    
    'Add '+' in the name of the added features
    For i = 1 To oAddFeatures.Count
         Set oNamedItem = oAddFeatures.Item(i)
        oNamedItem.Name = "+ " & strParentName
    Next i

    'Add '-' in the name of the subtracted features
    For i = 1 To oRemoveFeatures.Count
        Set oNamedItem = oRemoveFeatures.Item(i)
        oNamedItem.Name = "- " & strParentName
    Next i

    
    Set oNamedItem = Nothing
    Set oAddFeatures = Nothing
    Set oRemoveFeatures = Nothing
    Set oCompartCreation = Nothing
    
Exit Sub

ErrorHandler:
   
    m_oErrors.Add Err.Number, "CompBoolNameRule::IJNameRule_ComputeName", Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_NAMINGRULES_FAILED_PARTDESCRIPTIONRULE)
End Sub

