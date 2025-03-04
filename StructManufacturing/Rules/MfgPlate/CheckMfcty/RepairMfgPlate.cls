VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepairMfgPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: RepairMfgPlate.cls
'
' Author: J.Manasa
'
' Abstract: Check Manufacturability repair tool that deletes one of the manufacturing plates
'
'********************************************************************

Option Explicit

Private Const Module = "StrMfgCheckMfctyPlate.RepairMfgPlate"
Private Const strProgID = "StrMfgCheckMfctyPlate.RepairMfgPlate"

Private Const IID_IJStructGeometry = "{6034AD40-FA0B-11D1-B2FD-080036024603}"
Private Const IID_IJMfgPlateCreation_AE = "{C122691E-C84C-4d50-A04C-8BA7F6201CB9}"
Private Const DETAILPATEPARTINPUT_DEST = "DetailPlatePartInput_DEST"
Private Const MFGPLATEOUTPUT_ORIG = "MfgPlateOutput_ORIG"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"
Private Const IID_IJMfgPlatePart = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"

Implements IJCheckMfctyRepair

Private Sub Class_Initialize()
    Const METHOD As String = "Class_Initialize"
    On Error GoTo ErrorHandler
    GoTo Cleanup

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number

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
Const METHOD As String = "IJCheckMfctyRepair_Repair"
    On Error GoTo ErrorHandler
    Dim sMessage As String
        sMessage = "...Performing Repair"
        RepairObject vObject
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number

Cleanup:

End Sub
Private Sub RepairObject(ByRef vObject As Variant)
    
    Dim sMessage As String
    Dim Index As Integer
    Dim i As Integer
    Dim oMfgPlateColl As Collection
    Dim oAllMfgPlateParts As IJElements
    Dim oIJDObject As IJDObject
    Dim oPlatePart As IJDObject
    Dim oIJDOtherObject As IJDObject
    Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    
    Set oAllMfgPlateParts = oStructMfgGlobals.GetMfgPart(vObject, IID_IJMfgPlatePart)
    Set oMfgPlateColl = New Collection
    
    If oAllMfgPlateParts.Count > 1 Then
        'Get MfgPlate From AE
        sMessage = "...Retrieving manufacturing plate"
        For i = 1 To oAllMfgPlateParts.Count
            oMfgPlateColl.Add oAllMfgPlateParts.Item(i) 'Mfg plate
        Next
    End If
            
    sMessage = "...Finding the newest Manufacturing plate"
    'TO RETAIN THE NEWEST OBJECT
    Set oIJDObject = oMfgPlateColl.Item(1)
    For i = 2 To oMfgPlateColl.Count
        Set oIJDOtherObject = oMfgPlateColl.Item(i)
        If oIJDOtherObject.DateLastModified > oIJDObject.DateLastModified Then
            Set oIJDObject = oIJDOtherObject
            Set oIJDOtherObject = Nothing
        End If
    Next
        
    sMessage = "...Deleting other manufacturing plates"
    For i = 1 To oMfgPlateColl.Count
        Set oIJDOtherObject = oMfgPlateColl.Item(i)
        If Not oIJDOtherObject Is oIJDObject Then
            On Error Resume Next 'If there are no permissions on the Mfg object
            oIJDOtherObject.Remove
        End If
    Next
                                                    
    'The CheckMfcty command changes the status of the object.
    'After repair,we have to put it back to working status
    
    Set oPlatePart = vObject
    oPlatePart.ApprovalStatus = Working
    
End Sub

