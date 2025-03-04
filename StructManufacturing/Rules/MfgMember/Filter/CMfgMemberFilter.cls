VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgCustomFilter
' Module: MemberFilter
'
' Description:  To customize the filter criteria during the manufacture of members.
'  History:
'       Suma Mallena         May 06. 2009   created
'
'*******************************************************************************

Option Explicit

Implements IJSMCustomFilter

'******************************************************************************
' Routine: IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg
'
' Description:  This function is callled when an assembly is picked from the command.
'               Through this function, user can send the needed members of the assembly to be manufactured.
'
' NOTE:         Below implementation is default where we return all parts of the assembly as required

'******************************************************************************
Private Function IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg(ByVal pAssembly As GSCADMfgRulesDefinitions.IJAssembly) As IMSCoreCollections.IJElements

    Const METHOD = "IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg"
    On Error GoTo ErrorHandler

    'make a collection of parts with needed ones
    Dim oPartCollection As IJElements

    Set oPartCollection = New JObjectCollection
    'Get children of assembly

    Dim oChildren As IJDTargetObjectCol
    Set oChildren = pAssembly.GetChildren

    If oChildren.Count > 0 Then
        Dim Index As Long
        Dim oItem As Object

        'Get each assembly child and add child to elements list
        For Index = 1 To oChildren.Count
            'Get next item
            Set oItem = oChildren.Item(Index)

            If TypeOf oItem Is IJAssembly Then 'Assembly or block is child
                'get parts in this assembly
                oPartCollection.AddElements IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg(oItem)
            ElseIf TypeOf oItem Is ISPSMemberPartPrismatic Then ' member
                oPartCollection.Add oItem
            End If
        Next
    End If

    Set IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg = oPartCollection

    Exit Function
ErrorHandler:
    Err.Raise Err.Number

End Function

'******************************************************************************
' Routine: IJSMCustomFilter_MfgPartCustomFilterString
'
' Description:  This function is callled when the manufacturing member command is started.
'               Through this function, user can send the filter criteria for the members to be manufactured.
'
' NOTE:         Make sure that there are white spaces between the " AND NOT "

'******************************************************************************
Private Function IJSMCustomFilter_MfgPartCustomFilterString() As String
    
    IJSMCustomFilter_MfgPartCustomFilterString = " AND NOT [MfgCustomFilter.IsMemberIdentical,IsMemberPartIdentical]"

End Function


