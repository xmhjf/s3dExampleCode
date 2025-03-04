VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUMemberAssyNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
' Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
' Project:
'   PlnNameRules
'
' Abstract:
'   The file contains the VB-namerule for Built-up member Assemblies
'
' Author:
'   Devi Kishore Adiraju
'
' History:
'   July 18 2012     Kishore     Creation
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "BUMemberAssemblyNameRule: "
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private Const IID_IJMfgParent As String = "{07AE45B9-D88F-4DF2-865D-FEBE67D82E32}"

Private m_oErrors           As IJEditErrors       ' To collect and propagate the errors.
Private m_oError            As IJEditError         ' The error to raise.


Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'********************************************************************
' Description:
' Assembly's name will match with name of the Built-up Member if any.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
        
    Dim oChild                              As IJNamedItem
    Dim oParent                             As IJNamedItem
          
    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    If elements.Count = 1 Then
        Set oParent = elements.Item(1)
    Else
        Exit Sub
    End If
    
    Set oChild = oObject
    
    If oChild.Name = "New Assembly" Or oChild.Name = "" Or oChild.Name <> oParent.Name Then
        oChild.Name = oParent.Name
    End If
        
    Set oChild = Nothing
    Set oParent = Nothing
    
Exit Sub
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' There will be only one naming parent: Built-up member, if any, related to the assembly
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Dim oAssocRelation          As IJDAssocRelation
    Dim oTargetObjCol           As IJDTargetObjectCol
    Dim oDesignParent           As Object
    Dim lCount                  As Long
    
    Set oAssocRelation = oEntity
    Set oTargetObjCol = oAssocRelation.CollectionRelations(IID_IJMfgParent, "Module")
    
    If Not oTargetObjCol Is Nothing Then
        If oTargetObjCol.Count = 1 Then
            Set oDesignParent = oTargetObjCol.Item(1)
        End If
    End If
    
    If Not oDesignParent Is Nothing Then
        If TypeOf oDesignParent Is ISPSDesignedMember Then
            IJNameRule_GetNamingParents.Add oDesignParent
        End If
    End If

Exit Function
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

 

