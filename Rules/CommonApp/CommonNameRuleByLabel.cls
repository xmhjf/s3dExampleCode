VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonNameRuleByLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 1999, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a naming rule
'
'
'  Author: Murthy
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CommonNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTROLE = "part"
Private IID_IJPartOcc As Variant

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259



Private Sub Class_Initialize()
'{1146CF94-6B33-11D1-A300-080036409103}
IID_IJPartOcc = InitGuid(&H11, &H46, &HCF, &H94, &H6B, &H33, &H11, &HD1, &HA3, &H0, &H8, &H0, &H36, &H40, &H91, &H3)

Set m_oErrors = New IMSErrorLog.JServerErrors

End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "Base Name" + "Object Name" + Index.
' "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
' It is assumed that all Naming Parents and the Object implement IJPRodModelItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    Dim oLabel As IJReport
    Dim strName As String
    Dim oChildNamedItem As IJNamedItem
    
    Set oChildNamedItem = oObject
    Dim oMiddleTextLabelService As IJMiddleTextLabels
    'Get the Label Object from the Active entity
    If Not oActiveEntity Is Nothing Then
        Set oLabel = oActiveEntity.LabelObject
        If Not oLabel Is Nothing Then
            strName = oMiddleTextLabelService.GetFormattingResult
            If Len(strName) > 0 Then
                oChildNamedItem.Name = strName
            End If
        End If
    End If
    
        
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo label

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild

    On Error Resume Next
    Set oSysChild = oEntity
    Set oSysParent = oSysChild.GetParent
    
    On Error GoTo label

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oSysParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oSysParent)
    End If


    Set oSysParent = Nothing
    Set oSysChild = Nothing
Exit Function

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function



