VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirstMeet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'Copyright (C) 2002, Intergraph Corporation. All rights reserved.
'
'Abstract:
'   Repair Rule for PlnJointCheckFirstMeet
'
'Description:
'   This class repairs planning joints to place them under appropriate locations
'   in hierarchy.
'Notes:
'
'History
'Kishore            30th July 2007            Creation
'Kishore            24th Apr 2009             TR-CP-161111    Planning: CheckManufacturability doesn't correctly work
'*******************************************************************

Option Explicit

Implements IJCheckMfctyRepair

Const Module = "PlnJointRepairRules.FirstMeet"

Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.


Private Sub IJCheckMfctyRepair_Repair(oObj As Variant)
    Const Method = "IJCheckMfctyRepair_Repair"
    On Error GoTo ErrorHandler

    Dim oPlnJoint           As IJPlnJoint
    Dim oSiblingPlnJoint    As IJPlnJoint
    Dim object              As IJDObject
    Dim oSmartOcc           As IJPlnJointSO
    Dim oWrapPhysConn       As PlanningObjects.PlnPhysConn
    Dim oRelatedJoints      As IJDObjectCollection
    Dim oNamedItem          As IJNamedItem
    Dim bAllUnderOneAssy    As Boolean
    Dim oRightAssy          As IJAssembly
    
    Set object = oObj
    object.ApprovalStatus = Working
    
    'Check permissions on Planning Joint
    If (object.AccessControl And acUpdate) <> acUpdate Then
        Set oNamedItem = oObj
        MsgBox "No write permissions on " + oNamedItem.Name + ", cannot repair"
        Exit Sub
    End If
    
    Set oPlnJoint = oObj
    Set oWrapPhysConn = New PlanningObjects.PlnPhysConn
    Set oWrapPhysConn.object = oPlnJoint.GetPhysicalConnection
    Set oSmartOcc = oPlnJoint.GetSmartOccurrence
    
    Set oRightAssy = oWrapPhysConn.FirstMeet
    Set object = oRightAssy
    Set oNamedItem = oRightAssy
    
    'Check permissions on Assembly
    If (object.AccessControl And acUpdate) <> acUpdate Then
        Set oNamedItem = oRightAssy
        MsgBox "No write permissions on " + oNamedItem.Name + ", cannot repair"
        Exit Sub
    End If
    
    'Repair planning joints which are not moved by the user
    If oPlnJoint.ManualAssign = False Then
        oSmartOcc.SetAssembly oRightAssy
    End If
    
    'When all the joints related to a planning joint smart occurrence are
    'under the same assembly the "ManualAssign" flag can be set to False.
    Set oRelatedJoints = oSmartOcc.GetPlanningJoints(True)
    
    bAllUnderOneAssy = True
    For Each oSiblingPlnJoint In oRelatedJoints
        
        'Get the other planning joints related to the planning joint
        'that needs to be repaired and check whether all of them
        'are under the same assembly.
        
        If Not oSiblingPlnJoint Is oPlnJoint Then
            If Not oSiblingPlnJoint.GetAssembly Is _
                oPlnJoint.GetAssembly Then
                
                bAllUnderOneAssy = False
                Exit For
                
            End If
        End If
    Next
    
    'Set the ManualAssign boolean to false when all the joints are
    'under the same assembly.
    If bAllUnderOneAssy Then
        For Each oPlnJoint In oRelatedJoints
            oPlnJoint.ManualAssign = False
        Next
    End If

CleanUp:
    Set oPlnJoint = Nothing
    Set oSiblingPlnJoint = Nothing
    Set object = Nothing
    Set oSmartOcc = Nothing
    Set oRelatedJoints = Nothing
    Set oWrapPhysConn = Nothing
    Set oNamedItem = Nothing
    Set oRightAssy = Nothing
    
    Exit Sub
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & Method)
    m_oError.Raise
    GoTo CleanUp
End Sub
