VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilterFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2010 Intergraph Corporation.  All Rights Reserved.
'
' File: FilterFuncs.cls
'
' Author: Manasa Jaisetty
'
' Abstract: Filter functions
'
'  History:
'   2010-09-27  mjaisett     Creation date
'********************************************************************

Option Explicit

'Checks the given assemblies has connections with each other
Private Function IsConnectedAssemblies(obj1 As IJAssembly, obj2 As IJAssembly) As Boolean
            IsConnectedAssemblies = False
            
            'Get the subassemblies
            Dim oSubAssysOfSelAssy As IJElements
            Dim oSubAssysOfAdjAssy As IJElements
            
            Set oSubAssysOfSelAssy = GetSubAssemblies(obj1)
            Set oSubAssysOfAdjAssy = GetSubAssemblies(obj2)
                                                    
            'Make sure the assy collections are non empty
            If (oSubAssysOfSelAssy Is Nothing) Or (oSubAssysOfAdjAssy Is Nothing) Then Exit Function
            If (oSubAssysOfSelAssy.Count = 0) Or (oSubAssysOfAdjAssy.Count = 0) Then Exit Function
            
            Dim oMfgUtilAssyMargin As IJMfgUtilAssyMargin
            Set oMfgUtilAssyMargin = New MfgUtilAssyMargin
            
            'Get connections
            Dim oSelItem As IJAssembly
            Dim oAdjItem As IJAssembly
            Dim oConnections As IJElements
            Set oConnections = New JObjectCollection 
            
            For Each oSelItem In oSubAssysOfSelAssy
                For Each oAdjItem In oSubAssysOfAdjAssy
                    If Not oSelItem Is oAdjItem Then
                        oMfgUtilAssyMargin.GetConnectedChildren oSelItem, oAdjItem, oConnections
                        'If there is a connection, we got the right assembly
                        If oConnections.Count > 0 Then
                            IsConnectedAssemblies = True
                            Exit For
                        End If
                    End If
                Next
            Next
            
            'Clean
            Set oMfgUtilAssyMargin = Nothing
            Set oSubAssysOfSelAssy = Nothing
            Set oSubAssysOfAdjAssy = Nothing
            Set oConnections = Nothing

End Function

' *******************************************************************
' IsValidAssy: Returns 1 if the element is IJAssembly.
'              Otherwise, returns 0.
'                If another assembly is provided (eg StructFilter.UserArgument = oAnAssy),
'                returns 1 if its connected with other assembly
'                otherwise, retruns 0.
' *******************************************************************

Public Function IsValidAssy(Element As Object, oAdjAssy As Object) As Integer

    Dim oElem As Object
    On Error Resume Next
    Set oElem = Element
                                               
    IsValidAssy = 0
    If TypeOf oElem Is IJAssembly Then
        If oAdjAssy Is Nothing Then
            IsValidAssy = 1
        ElseIf IsConnectedAssemblies(oAdjAssy, oElem) Then
            IsValidAssy = 1
        End If
    End If
                
End Function

