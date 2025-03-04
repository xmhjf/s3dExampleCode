VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldNameRuleWBSSeqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2004 Intergraph Corporation.  All rights reserved.
'
'  Project: PWeldNameRule
'  Class:   WeldNameRuleWBSSeqGen
'
'  Abstract: The file contains the Implementation for naming rule interface for Welds
'
'  Author: Vidya
'
'  History:
'
'******************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057
Private Const MODULE = "WeldNameRuleWBSSeqGen: "
Private Const E_FAIL = -2147467259
Private Const MODELDATABASE = "Model"
Dim m_oErrors As IJEditErrors  ' Middle tier errors

Private Sub Class_Initialize()
    'DI-CP-164764    Prevent DLL unloading for high usage DLLs
    If m_ForceDllInProcess = 0 Then
        m_ForceDllInProcess = ForceDllInProcess(AddressOf ForceDllInProcess)
    End If
    
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
'
' Creates a name for the object passed in. The name is based on the Pipeline
' to which the Weld belongs.  It is something like this: "WBS Item Name" + Index.
' The WBSItem is found in AddNamingParents() and is returned as the first
' item in an elements collection.  Then this sub is called and is passed that
' elements collection.  The naming rule semantic first calls AddNamingParents()
' and then this sub, ComputeName().
'
' Other naming rules may base the name on other parents and may construct a
' different form for the name.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    On Error GoTo ERR_SUB
    
    Const METHOD = "IJNameRule_ComputeName"
    Dim oWeldNamedItem As IJNamedItem
    Dim oWBSNamedItem As IJNamedItem
    Dim strWeldName As String
    Dim strWBSName As String
    Dim strSequenceId As String
    
    'GetNamingParents returns an elements collection.
    'GetNamingParents placed this Weld's Pipeline as the first item in the collection.
    'The elements argument that is passed to this sub is that elements collection.
    
    Set oWBSNamedItem = elements.Item(1) 'Get the IJNamedItem inteface for the Pipeline
    Set oWeldNamedItem = oObject 'This is the interface on which we change the Weld's name
    
    'Make sure that the name of the Pipeline is not NULL.  This will cause problems
    'for GetCount.
    If Not oWBSNamedItem Is Nothing Then
        strWBSName = oWBSNamedItem.Name
    End If
    
    If strWBSName = "" Then
        strWBSName = "WBS Item"
    End If
    
    strSequenceId = GetSequenceId(oObject)
        
'        '  If the weld does not already have a sequence id, set it to "0"
'        If (Len(.Value) = 0) Then
'            .Value = "0"
'        End If
'
        'If base part of name hasn't changed, don't compute a new name.
'MsgBox "strSequenceId = " & strSequenceId & "   pipeline = " & strPipelineName
    If oWeldNamedItem.Name = (strWBSName & strSequenceId) Then Exit Sub
        
    'Build the Weld's name and set it.
    oWeldNamedItem.Name = strWBSName + strSequenceId
    
    'Release object references.
    Set oWeldNamedItem = Nothing
    Set oWBSNamedItem = Nothing
    
    Exit Sub
ERR_SUB:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
    
End Sub
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
'
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
' In the case of Welds, we are using the WBS item as our naming parent. WBSItem->Part->Weld relation
' is used here
'
' The naming rule semantic will take the parents returned in the elements collection and
' create a relationship so that whenever their name changes, the rule is again fired and
' the name is updated by this rule.
' History:
' 28-Mar-2007   Vidya   CR 107382   Modified to check if the weld belongs to a WBS Item and
'                                   if so, add the WBS Item as its naming parent or else,
'                                   fall back to the default implementation of using the
'                                   pipeline as the naming parent.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    On Error GoTo ERR_SUB
    
    Const METHOD = "IJNameRule_GetNamingParents"
    
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    
    'Create elements collection in which to return parents on which the name
    'of the object are to be based.  In this case, only the Pipeline is needed.
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If oEntity Is Nothing Then Exit Function 'No Weld
    
    '  Weld is naming parent of self
    'Call IJNameRule_GetNamingParents.Add(oEntity)
    
    'Get Part, which is parent of Weld
    Set oDesignChild = oEntity
    Set oDesignParent = oDesignChild.GetParent
    
    Dim CollHlpr As IMSRelation.DCollectionHelper
    Dim TOColl As IJDTargetObjectCol
    Dim AssocRel As IJDAssocRelation
    Dim oObj As Object
    Dim oElems As IJElements
    Dim ii As Long
    
    On Error Resume Next
    Dim oChild As Object
    Set oChild = oDesignParent
    Set AssocRel = oChild
    Dim vIIDWBSItemChild As Variant
    vIIDWBSItemChild = GetGuidFromString("33F33972-39E3-412E-B509-88F4FFCF3776")
    Set CollHlpr = AssocRel.CollectionRelations(vIIDWBSItemChild, "WBSParent")
    If Not CollHlpr Is Nothing Then
        Set TOColl = CollHlpr
        If TOColl.Count = 0 Then
            ' either no parents, or more than one parent, so cannot be exclusive
            'Get Run, which is parent of Part
            Set oDesignChild = oDesignParent
            Set oDesignParent = oDesignChild.GetParent
            
            If Not oDesignParent Is Nothing Then
                'Get Pipeline, which is parent of Run
                Set oDesignChild = oDesignParent
                Set oDesignParent = oDesignChild.GetParent
                
                If Not oDesignParent Is Nothing Then
                    'Add Pipeline to elements collection
                    Call IJNameRule_GetNamingParents.Add(oDesignParent)
                End If
            End If
        Else
            Dim oWBSItem As IJWBSItem
            Set oWBSItem = TOColl.Item(1)
            If Not oWBSItem Is Nothing Then
                Call IJNameRule_GetNamingParents.Add(oWBSItem)
            End If
        End If
    End If
    
    'Release object references
    Set oDesignChild = Nothing
    Set oDesignParent = Nothing
    Set oWBSItem = Nothing
    Set oChild = Nothing
    Set AssocRel = Nothing
    
    Exit Function
ERR_SUB:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
    
End Function

Private Function GetGuidFromString(strGuid As String) As Variant
    Dim Guid(0 To 15) As Byte
    Dim strByte As String
    Dim ii As Integer
    Dim jj As Integer
    
    For ii = 0 To 15
        Guid(ii) = &H0
    Next ii
    jj = 3
    For ii = 0 To 3
        strByte = "&H" & Left$(strGuid, 2)
        strGuid = Right$(strGuid, Len(strGuid) - 2)
        Guid(jj) = CByte(strByte)
        jj = jj - 1
    Next ii
    strGuid = Right$(strGuid, Len(strGuid) - 1)
    jj = 5
    For ii = 0 To 1
        strByte = "&H" & Left$(strGuid, 2)
        strGuid = Right$(strGuid, Len(strGuid) - 2)
        Guid(jj) = CByte(strByte)
        jj = jj - 1
    Next ii
    strGuid = Right$(strGuid, Len(strGuid) - 1)
    jj = 7
    For ii = 0 To 1
        strByte = "&H" & Left$(strGuid, 2)
        strGuid = Right$(strGuid, Len(strGuid) - 2)
        Guid(jj) = CByte(strByte)
        jj = jj - 1
    Next ii
    strGuid = Right$(strGuid, Len(strGuid) - 1)

    For ii = 8 To 9
        strByte = "&H" & Left$(strGuid, 2)
        strGuid = Right$(strGuid, Len(strGuid) - 2)
        Guid(ii) = CByte(strByte)
    Next ii
    strGuid = Right$(strGuid, Len(strGuid) - 1)
    For ii = 10 To 15
        strByte = "&H" & Left$(strGuid, 2)
        strGuid = Right$(strGuid, Len(strGuid) - 2)
        Guid(ii) = CByte(strByte)
    Next ii
    
    GetGuidFromString = Guid
End Function



