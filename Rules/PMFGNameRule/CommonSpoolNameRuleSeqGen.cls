VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonSpoolNameRuleSeqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2004-2006 Intergraph Corporation.  All rights reserved.
'
'  Project: PMFGNameRule
'  Class:   CommonSpoolNameRuleSeqGen
'
'  Abstract: The file contains the Implementation for naming rule interface for Spools
'
'  Author: Furno
'
'  History:
'  17-Jan-2007 Sreelekha   DM113222  oSpool.Pipeline is returning invalid object
'******************************************************************
Option Explicit

Implements IJNameRule
Const vbInvalidArg = &H80070057
Private Const MODULE = "CommonSpoolNameRuleSeqGen: "
Private Const E_FAIL = -2147467259
Dim m_oErrors As IJEditErrors  ' Middle tier errors

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub


Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName "
    Dim oNamedItem As IJNamedItem
    Dim strType As String
    Dim strParentName As String
    Dim oPipeLine As Object
    Dim strSeqId As String
    On Error GoTo ErrorHandler
    
    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If
       
    ' Get pipeline name
    Set oNamedItem = pParents.Item(1)
    strParentName = oNamedItem.Name
    
    ' Get weld type
    Set oNamedItem = pEntity
    strType = oNamedItem.TypeString
    
    strSeqId = GetSequenceId(pEntity)
    
    oNamedItem.Name = strParentName & "_" & strType & strSeqId
    Set oNamedItem = Nothing

    Exit Sub
    
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oSpool As IJDSpool
    Dim oAssemblyChild As IJAssemblyChild
    Dim oNamingParent As Object
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
   
      'Get the naming parent
    Set oSpool = pEntity
    
    On Error Resume Next
    Set oNamingParent = oSpool.SpoolableObject
    On Error GoTo ErrorHandler
    
    IJNameRule_GetNamingParents.Add oNamingParent
    
    Set oAssemblyChild = Nothing
    Set oNamingParent = Nothing
    Set oSpool = Nothing
    
    Exit Function
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function

