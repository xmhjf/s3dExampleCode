VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepairProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: RepairMfgProfile.cls
'
' Author: J.Manasa
'
' Abstract: Check Manufacturability repair tool that deletes one of the manufacturing profiles
'
'********************************************************************

Option Explicit

Private Const Module = "StrMfgCheckMfctyProfile.RepairProfile"
Private Const strProgID = "StrMfgCheckMfctyProfile.RepairProfile"

Private Const IID_IJStructProfilePart = "{F933A05F-8820-4738-B743-E030EC9502AB}"
Private Const IID_IJMfgProfileCreation_AE = "{6C727AAA-F395-40EC-990A-36B7D7A73A64}"
Private Const DETAILPROFILEPARTINPUT_DEST = "DetailProfilePartInput_DEST"
Private Const MFGPROFILEOUTPUT_ORIG = "MfgProfileOutput_ORIG"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"


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
    Dim oMfgProfileColl As Collection
    Dim oIJDObject As IJDObject
    Dim oIJDOtherObject As IJDObject
    Dim oProfilePart As IJDObject
    Dim Index As Integer
    Dim i As Integer
    
    Set oMfgProfileColl = New Collection
    
    ' Current Object is an IJStructProfilePart
    sMessage = "...Retrieving profile part"
    Set oIJDAssocRelation = vObject
    Set oIJDTargetObjectColl = oIJDAssocRelation.CollectionRelations _
                                                (IID_IJStructProfilePart, DETAILPROFILEPARTINPUT_DEST)
    
    If oIJDTargetObjectColl.Count > 1 Then
        For Index = 1 To oIJDTargetObjectColl.Count
           
            'Get MfgProfile From AE
            sMessage = "...Retrieving Manufacturing profile"
             
            Dim oIJDAssocRel As IJDAssocRelation
            Dim oIJDTargetObjColl As IJDTargetObjectCol
             
            Set oIJDAssocRel = oIJDTargetObjectColl.Item(Index)
            Set oIJDTargetObjColl = oIJDAssocRel.CollectionRelations _
                                                        (IID_IJMfgProfileCreation_AE, MFGPROFILEOUTPUT_ORIG)
                                                        
             'TO DELETE THE OBJECT WHICH IS IN THE TO-DO-LIST
'            Set oIJDObject = oIJDTargetObjColl.Item(1) 'Mfg profile
'            Set oIJDAssocRel = oIJDObject
'            Set oIJDTargetObjColl = oIJDAssocRel.CollectionRelations(IID_IJDObject, "toErrorList")
'
'            If Not oIJDTargetObjColl Is Nothing Then 'The Mfg profile is in the To-Do-List
'                'Remove it
'                If oIJDTargetObjColl.Count > 0 Then
'                    sMessage = "...Deleting the Manufacturing profile"
'                    On Error Resume Next 'If there are no permissions on the Mfg object
'                    oIJDObject.Remove
'                End If
'            End If
            
            oMfgProfileColl.Add oIJDTargetObjColl.Item(1) 'Mfg profile
            
        Next
    End If
                                                    
    sMessage = "...Finding the newest Manufacturing profile"
    'TO RETAIN THE NEWEST OBJECT
    Set oIJDObject = oMfgProfileColl.Item(1)
    For i = 2 To oMfgProfileColl.Count
        Set oIJDOtherObject = oMfgProfileColl.Item(i)
        If oIJDOtherObject.DateLastModified > oIJDObject.DateLastModified Then
             Set oIJDObject = oIJDOtherObject
             Set oIJDOtherObject = Nothing
         End If
     Next
         
     sMessage = "...Deleting other manufacturing profiles"
     For i = 1 To oMfgProfileColl.Count
         Set oIJDOtherObject = oMfgProfileColl.Item(i)
         If Not oIJDOtherObject Is oIJDObject Then
             On Error Resume Next 'If there are no permissions on the Mfg object
             oIJDOtherObject.Remove
         End If
     Next
                                                    
                                                    
    'The CheckMfcty command changes the status of the object.
    'After repair,we have to put it back to working status
    
    Set oProfilePart = vObject
    oProfilePart.ApprovalStatus = Working
    
End Sub
