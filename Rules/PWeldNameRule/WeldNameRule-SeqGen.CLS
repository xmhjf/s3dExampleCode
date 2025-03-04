VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldNameRuleSeqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2004 Intergraph Corporation.  All rights reserved.
'
'  Project: PWeldNameRule
'  Class:   WeldNameRuleSeqGen
'
'  Abstract: The file contains the Implementation for naming rule interface for Welds
'
'  Author: Furno
'
'  History:
'
'******************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057
Private Const MODULE = "WeldNameRuleSeqGen: "
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
' to which the Weld belongs.  It is something like this: "Pipeline Name" + Index.
' The Pipeline is found in AddNamingParents() and is returned as the first
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
    Dim oPipelineNamedItem As IJNamedItem
    Dim strWeldName As String
    Dim strPipelineName As String
    Dim strSequenceId As String
    
    'GetNamingParents returns an elements collection.
    'GetNamingParents placed this Weld's Pipeline as the first item in the collection.
    'The elements argument that is passed to this sub is that elements collection.
    
    Set oPipelineNamedItem = elements.Item(1) 'Get the IJNamedItem inteface for the Pipeline
    Set oWeldNamedItem = oObject 'This is the interface on which we change the Weld's name
    
    'Make sure that the name of the Pipeline is not NULL.  This will cause problems
    'for GetCount.
    If Not oPipelineNamedItem Is Nothing Then
        strPipelineName = oPipelineNamedItem.Name
    End If
    
    If strPipelineName = "" Then
        strPipelineName = "Pipeline"
    End If
    
    strSequenceId = GetSequenceId(oObject)
        
'        '  If the weld does not already have a sequence id, set it to "0"
'        If (Len(.Value) = 0) Then
'            .Value = "0"
'        End If
'
        'If base part of name hasn't changed, don't compute a new name.
'MsgBox "strSequenceId = " & strSequenceId & "   pipeline = " & strPipelineName
    If oWeldNamedItem.Name = (strPipelineName & strSequenceId) Then Exit Sub
        
    'Build the Weld's name and set it.
    oWeldNamedItem.Name = strPipelineName + strSequenceId
    
    'Release object references.
    Set oWeldNamedItem = Nothing
    Set oPipelineNamedItem = Nothing
    
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
' In the case of Welds, we are using the Pipeline as our naming parent. Pipeline->Run->Part->Weld relation
' is used here
'
' The naming rule semantic will take the parents returned in the elements collection and
' create a relationship so that whenever their name changes, the rule is again fired and
' the name is updated by this rule.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    On Error GoTo ERR_SUB
    
    Const METHOD = "IJNameRule_GetNamingParents"
    
    'Create elements collection in which to return parents on which the name
    'of the object are to be based.  In this case, only the Pipeline is needed.
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If oEntity Is Nothing Then Exit Function 'No Weld
    
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    
    'Get Part, which is parent of Weld
    Set oDesignChild = oEntity
    Set oDesignParent = oDesignChild.GetParent
    
    'accounting for orphaned welds
    If Not oDesignParent Is Nothing Then
        'Get Run, which is parent of Part
        Set oDesignChild = oDesignParent
        Set oDesignParent = oDesignChild.GetParent
        'account for orphaned parts
        If Not oDesignParent Is Nothing Then
            'Get Pipeline, which is parent of Run
            Set oDesignChild = oDesignParent
            Set oDesignParent = oDesignChild.GetParent
            'account for orphaned run
            If Not oDesignParent Is Nothing Then
                'Add Pipeline to elements collection
                Call IJNameRule_GetNamingParents.Add(oDesignParent)
            End If
        End If
    End If
    'Release object references
    Set oDesignChild = Nothing
    Set oDesignParent = Nothing
    
    Exit Function
ERR_SUB:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
    
End Function
