VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultByQueryRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project: DrawingNameRules
'
'  Abstract: The file contains an implementation of the default naming rule
'            for the Drawing Volume object in Space Management.
'
'  History:
'   25-Jan-2006         B. Covington * Code Review by: C Ramsey
'       TR85888  3D Model Data documents are created with no name if a Catalog filter is used
'   08-Jan-2004         B. Covington * Code Review by:
'       CR54157  Copied and modified from SpaceRules project
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Implements IJNameRule

Private m_3DModelDataRule As IJNameRule

Private Const MODULE = "ByQuerySheetRule: "
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
' Notes: VolumeName  = Parents Name + LocationID + Unique Index
'***************************************************************************
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo errorhandler
    Dim oChildNamedItem     As IJNamedItem
    Dim strChildName        As String
    Dim oParentNamedItem    As IJNamedItem
    Dim oParents            As IJElements
    
    Set oChildNamedItem = pEntity

    If Not m_3DModelDataRule Is Nothing Then
        ' Forward call to DefaultSheetRule which will compute and set the name
        m_3DModelDataRule.ComputeName pEntity, pParents, pActiveEntity
        
        ' now just exit
        GoTo Cleanup
    End If
    
    Set oParents = IJNameRule_GetNamingParents(pEntity)
    
    If Not oParents Is Nothing Then
        If 1 < oParents.Count Then
            Dim lCount As Long
            
            For lCount = 1 To oParents.Count - 1
                On Error Resume Next
                Set oParentNamedItem = oParents.Item(lCount)
                On Error GoTo errorhandler
                If Not oParentNamedItem Is Nothing Then
                    strChildName = strChildName & oParentNamedItem.Name & " - "
                End If
                Set oParentNamedItem = Nothing
            Next
            
            On Error Resume Next
            Set oParentNamedItem = oParents.Item(oParents.Count)
            On Error GoTo errorhandler
            If Not oParentNamedItem Is Nothing Then
                strChildName = strChildName & oParentNamedItem.Name
                Set oParentNamedItem = Nothing
            End If
        ElseIf 0 < oParents.Count Then
            On Error Resume Next
            Set oParentNamedItem = oParents.Item(1)
            On Error GoTo errorhandler
            If Not oParentNamedItem Is Nothing Then
                strChildName = oParentNamedItem.Name
                Set oParentNamedItem = Nothing
            End If
        End If
    End If

    Set oParents = Nothing
     
    oChildNamedItem.Name = strChildName
    
Cleanup:
    Set oParents = Nothing
    Set oChildNamedItem = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub
errorhandler:
    m_oErrors.Add Err.Number, "ByQuerySheetRule::IJNameRule_ComputeName", Err.Description
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
    
    Dim objSheet As IJDSheetRelations2
    Dim obj3DObjectsColl As IJElements
    Dim objFiltersColl As IJElements
    Dim objNamingParents As IJElements
    
    Set objNamingParents = New IMSCoreCollections.JObjectCollection
    
    On Error Resume Next
    Set objSheet = pEntity
    Set obj3DObjectsColl = objSheet.GetRelatedObject("3dobject")
    Set objFiltersColl = objSheet.GetRelatedObject("filter")
    On Error GoTo errorhandler
    
    If Not objFiltersColl Is Nothing Then
        If 0 < objFiltersColl.Count Then
            objNamingParents.Add objFiltersColl.Item(1)
        End If
    End If
    
    ' most Drawings-by-query types
    If Not obj3DObjectsColl Is Nothing Then
        If 0 < obj3DObjectsColl.Count Then
            objNamingParents.Add obj3DObjectsColl.Item(1)
        End If
    End If
    
    If objNamingParents.Count = 0 Then
        '3DModel Data type
        Set m_3DModelDataRule = New DefaultSheetRule
        
        ' forward call to DefaultSheetRule which will return the parent
        ' SnapIn for this Drawing
        Set objNamingParents = m_3DModelDataRule.GetNamingParents(pEntity)
    End If
    
    Set IJNameRule_GetNamingParents = objNamingParents
    
    Set obj3DObjectsColl = Nothing
    Set objFiltersColl = Nothing
    Set objSheet = Nothing
    
Exit Function
errorhandler:
    m_oErrors.Add Err.Number, "ByQuerySheetRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

