VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReferenceNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2003, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'  File:    SystemNameRule.cls
'
'  Abstract: The file contains a sample implementation of a naming rule
'       If the object is a reference object. This simply returns the
'    name of the given object. This is implemented separately instead of
' handling in the property page so that any future enhancements could be addressed easily.
'
'  Author: Yalla Kiran Kumar
'
'  23-Aug-2004  Created
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "ReferenceNameRule: "
Private Const MODELDATABASE = "Model"

Dim m_oErrors As IJEditErrors

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
' Just setting back its own name. Adding this for future requirements.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    
    Dim oChildNamedItem As IJNamedItem
    Dim strChildName As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oChildNamedItem = oObject
    strChildName = oChildNamedItem.Name
    
    'setting the same name back again. Nothing to do for now.
    oChildNamedItem.Name = strChildName
    
Exit Sub

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    
End Function
