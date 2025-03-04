VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreatmentNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



'*******************************************************************
'  Copyright (C) 2005-2006 Global Research and Development, Inc.
'  All rights reserved.
'
'  Project: TreatmentNameRule
'
'  Abstract: The file contains name rules for Edge treatment object
'
'
'******************************************************************
Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors
Private m_oNameHelp As SDNameRuleHelper
Private m_oSDOHelper As StructDetailObjects.Helper
 

Private Sub Class_Initialize()
    On Error GoTo label
   
    'set globals
    Set m_oNameHelp = New SDNameRulesHelper.SDNameRuleHelper
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oSDOHelper = New StructDetailObjects.Helper
    
    Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "Class_Initialize", Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub Class_Terminate()
        'clear globals
    Set m_oErrors = Nothing
    Set m_oNameHelp = Nothing
    Set m_oSDOHelper = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
     Const METHOD = "IJNameRule_ComputeName"
On Error GoTo errLabel
    
    Dim oNamingEntity As IJNamedItem
    Dim oNamedItem As IJNamedItem
    Dim oResourceMgr As IUnknown
    Dim strNamedParentsString As String
    Dim strTypeString As String
    Dim strNamedObject As String
    
    Set oNamingEntity = pParents.Item(1)
    Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(pEntity)
    Set oNamedItem = pEntity
    Dim iPos As Long
    
    
    strTypeString = oNamingEntity.Name & "-" & "FreeEdge" & "<"
    strNamedObject = Mid(oNamedItem.Name, 1, Len(strTypeString))
    ' check the above string in the named object
    
    If UCase(strTypeString) <> UCase(strNamedObject) Then
        strNamedParentsString = oNamedItem.TypeString
        Call m_oNameHelp.SDIndexByCounter(strNamedParentsString, iPos, oResourceMgr)
        oNamedItem.Name = oNamingEntity.Name & "-" & "FreeEdge" & "<" & Format(iPos) & ">"
    End If
    
    Exit Sub
errLabel:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
 Const METHOD = "IJNameRule_GetNamingParents"
 
 On Error GoTo errLabel
     Dim oEdgeTreatment As IJShpStrEdgeTreatment
     Dim oPort As IJPort
     Dim oWorkColl As IJElements
     
     Set oWorkColl = New IMSCoreCollections.JObjectCollection
     Set oEdgeTreatment = pEntity
     Set oPort = oEdgeTreatment.GetEdge
     oWorkColl.Add oPort.Connectable
     Set IJNameRule_GetNamingParents = oWorkColl
     Exit Function

errLabel:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
