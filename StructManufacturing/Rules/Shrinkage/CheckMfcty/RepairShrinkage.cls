VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepairShrinkage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: RepairShrinkage.cls
'
' Author: J.Manasa
'
' Abstract: Check Manufacturability repair tool that deletes one of the shrinkages
'
'********************************************************************

Option Explicit

Private Const Module = "StrMfgCheckMfctyShr.RepairShrinkage"
Private Const strProgID = "StrMfgCheckMfctyShr.RepairShrinkage"

Private Const IID_IJGeometry = "{96EB9676-6530-11D1-977F-080036754203}"
Private Const IID_IJShrinkage_AE = "{5E839CD8-EEEF-45E3-9339-06075272AA2E}"
Private Const INPUTGEOMETRY_AE = "InputGeometry_AE"
Private Const SHRINKAGEOUTPUT_SHRINKAGE = "ShrinkageOutput_Shrinkage"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"
Private Const IID_IJScalingShr = "{DE77050C-3300-11D5-BA1A-0090276F4279}"


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
    
    Dim oIJDAssocRelation As IJDAssocRelation
    Dim oIJDTargetObjectColl As IJDTargetObjectCol
    Dim oIJDObject As IJDObject
    Dim oShrinkageColl As Collection
    Dim oIJDOtherObject As IJDObject
    Dim oPart As IJDObject
    Dim Index As Integer
    Dim i As Integer
    Dim oShrinkage1 As IJScalingShr
    Dim oShrinkage2 As IJScalingShr
    
    Set oShrinkageColl = New Collection
    
    ' Current Object is an IJStructProfilePart or IJPlatePart
    sMessage = "...Retrieving the part"
    Set oIJDAssocRelation = vObject
    
    
    sMessage = "...Retrieving Shrinkage"
    Dim oAllShrinkages As IJElements
    Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    
    Set oAllShrinkages = oStructMfgGlobals.GetMfgPart(vObject, IID_IJScalingShr)
     
    If Not oAllShrinkages Is Nothing Then
         If oAllShrinkages.Count > 1 Then
             For i = 1 To oAllShrinkages.Count ' this loop checks whether shrinkage modes are same for different shrinkages
                  oShrinkageColl.Add oAllShrinkages.Item(i) 'Shrinkage
             Next
         End If
    End If
                                          
    sMessage = "...Finding the newest Shrinkage"
    'TO RETAIN THE NEWEST OBJECT
    Set oIJDObject = oShrinkageColl.Item(1)
    Set oShrinkage1 = oIJDObject
    For i = 2 To oShrinkageColl.Count
        Set oIJDOtherObject = oShrinkageColl.Item(i)
        Set oShrinkage2 = oIJDOtherObject
        If oShrinkage1.ShrinkageMode = oShrinkage2.ShrinkageMode Then
            If oIJDOtherObject.DateLastModified > oIJDObject.DateLastModified Then
                Set oIJDObject = oIJDOtherObject
                Set oIJDOtherObject = Nothing
            End If
        End If
    Next
        
    sMessage = "...Deleting other shrinkages"
    For i = 1 To oShrinkageColl.Count
        Set oIJDOtherObject = oShrinkageColl.Item(i)
        If Not oIJDOtherObject Is oIJDObject Then
            On Error Resume Next 'If there are no permissions on the object
            oIJDOtherObject.Remove
        End If
    Next
                                              
                                                    
                                                    
                                                    
    'The CheckMfcty command changes the status of the object.
    'After repair,we have to put it back to working status
    
    Set oPart = vObject
    oPart.ApprovalStatus = Working
    
    
End Sub

