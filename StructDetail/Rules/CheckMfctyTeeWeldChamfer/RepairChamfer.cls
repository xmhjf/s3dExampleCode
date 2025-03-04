VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepairChamfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2005 Intergraph Corporation.  All Rights Reserved.
'
' File: RepairChamfers.cls
'
' Author: D.A. Trent
'
' Abstract: Check Manufacturability repair tool that performs a recompute
' on an object.
'
'********************************************************************

Option Explicit

Private Const Module = "CheckMfctyTeeWeldChamfer.RepairChamfers"
Private Const strProgID = "CheckMfctyTeeWeldChamfer.RepairChamfers"
Private Const IID_IJFullObject As String = "{bcbfb3c0-98c2-11d1-93de-08003670a902}"

Implements IJCheckMfctyRepair

'

Private Sub Class_Initialize()
    Const Method As String = "Class_Initialize"
    On Error GoTo ErrorHandler
    GoTo Cleanup

ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number

Cleanup:

End Sub

'******************************************************************************
' Routine: IJCheckMfctyRepair_Repair
'
' Abstract:
'   Implementation of the IJCheckMfcty interface
'
' Description: Perform a repair on the specified object.
'
' Inputs:
'     vObject:  The object to be repaired.
'******************************************************************************
Public Sub IJCheckMfctyRepair_Repair(ByRef vObject As Variant)
Const Method As String = "IJCheckMfctyRepair_Repair"
    On Error GoTo ErrorHandler
    
    Dim sMessage As String
    Dim sObjectInfo As String
    
    Dim oObject As Object
    Dim oIJDObject As IJDObject
    Dim oNamedItem As IJNamedItem
    Dim oParentObject As Object
    
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools
    
    ' Current Object is an IJChamfer (a SmartItem)
    ' Get the owning Smart Object ( the Smart object that created it)
    sMessage = "...Retreiving Chamfer object"
    Set oObject = vObject
    
    sMessage = "...Retreiving Chamfer's parent AssemblyConnection"
    GetParentSmartObject oObject, oParentObject
    
    If TypeOf oObject Is IJNamedItem Then
        Set oNamedItem = oObject
        sObjectInfo = "Chamfer: " & oNamedItem.Name
        Set oNamedItem = Nothing
    Else
        sObjectInfo = "Chamfer: Is not valid (Name unknown)"
    End If
    
    If TypeOf oParentObject Is IJNamedItem Then
        Set oNamedItem = oParentObject
        sObjectInfo = sObjectInfo & vbCrLf & _
                      "AssemblyConnection: " & oNamedItem.Name
        Set oNamedItem = Nothing
    Else
        sObjectInfo = sObjectInfo & vbCrLf & _
                      "AssemblyConnection: Is not valid (Name unknown)"
    End If
    
    ' Check if Parent Smart Object is type IJAssemblyConnection
    If oParentObject Is Nothing Then
        sMessage = sObjectInfo & vbCrLf & _
                   "...AssemblyConnection Is Nothing"
        GoTo ErrorHandler
                   
    ElseIf Not TypeOf oParentObject Is IJAssemblyConnection Then
        sMessage = sObjectInfo & vbCrLf & _
                   "...Cahmfer Parentobject is NOT AssemblyConnection"
        GoTo ErrorHandler
    
    ElseIf TypeOf oParentObject Is IJDObject Then
        ' The CheckMfcty checking mechanism has set the status of the object
        ' to Rejected.  The Revision Manager cannot recompute if the object
        ' is not set to working.
        ' Therefore the status of the object must be changed.
        sMessage = sObjectInfo & vbCrLf & _
                   "...Setting Approval Status to: Working"
        Set oIJDObject = oParentObject
        oIJDObject.ApprovalStatus = Working
        Set oIJDObject = Nothing
        
        sMessage = sObjectInfo & vbCrLf & _
                   "...Forcing Recompute of AssemblyConnection"
        Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
        oStructAssocTools.UpdateObject oParentObject, IID_IJFullObject
        Set oStructAssocTools = Nothing
        
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, Method, sMessage).Number

Cleanup:

End Sub
