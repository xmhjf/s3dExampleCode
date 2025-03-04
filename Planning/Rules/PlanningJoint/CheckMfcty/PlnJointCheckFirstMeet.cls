VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointCheckFirstMeet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2004 Intergraph Corporation.  All Rights Reserved.
'
' File: PlnJointCheckFirstMeet.cls
'
' Author: Dick Swager
'
' Abstract: Check manufacturability rules for Planning Joints.
'
' Description:

' History:
'   Unknown                     Creation
'   13th Jul '07    Kishore     CR-CP·113082  Provide repair capability for “PlnJntCheckFirstMeet” check mfacturability rule
'
'
'********************************************************************

Option Explicit

Private Const Module = "PlnCheckMfcty.PlnJointCheckFirstMeet"
Private Const PROG_ID = "PlnJointRepairRules.FirstMeet"

Implements IJCheckMfcty

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description:
'   For planning joints, check manufacturability will need to check that the
' "FirstMeet" assembly is still valid for the associated planning joint smart
' occurrence.
'
' PseudoCode
'For each CPlnJointSO (the smart occurrence that creates the planning joints)
'    Get the Physical Connection (Call GetPhysicalConnection on IJPlnJointSO)
'    Create the Wrapper for the Physical Connection
'        Pass the PC
'        Call FirstMeet ()
'    Call GetAssembly() on IJPlnJointSO
'    Check that the Assemblies are both the same - Report an error otherwise
'End for
'
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck _
       (ByVal oCollection As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)

    Const Method = "IJCheckMfcty_ExecuteCheck"
    On Error GoTo ErrorHandler

    ' Setup the Check Manufacturability call back mechanism.
    Dim oCallBack As IJCheckMfctyCallback
    Dim lngDummy As Long
    Dim lngCancel As Long

    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' if the collection is empty, do not continue
    If oCollection.Count = 0 Then GoTo Cleanup
    
    ' get all the planning joints from the collection
    Dim oJoints As IJElements
    GatherPlanningJoints oCollection, oJoints
    
    ' Verify that each Planning Joint is in the Assembly that it thinks
    ' it is in.
    Dim oPlnJoint As IJPlnJoint
    For Each oPlnJoint In oJoints
    
        If oPlnJoint.ManualAssign = False Then
            ' Create the physical connection wrapper for the Planning Joint.
            Dim oPlnPhysConn As PlanningObjects.PlnPhysConn
            Set oPlnPhysConn = New PlanningObjects.PlnPhysConn
            Set oPlnPhysConn.Object = oPlnJoint.GetPhysicalConnection
            
            ' Get the FirstMeet assembly.  This is the assembly that contains
            ' both of the connected objects.
            Dim oAssyFirstMeet As IJAssembly
            Set oAssyFirstMeet = oPlnPhysConn.FirstMeet
            
            ' Get the assembly that contains the Planning Joint.
            Dim oAssyPlnJoint As IJAssembly
            Set oAssyPlnJoint = oPlnJoint.GetAssembly
            
            ' Verify that the FirstMeet and Planning Joint Smart Occurrence
            ' asseblies are the same.
            If Not oAssyFirstMeet Is oAssyPlnJoint Then
            
                ' An error condition exits.
                Dim oNamedItemPhysConn As IJNamedItem
                Set oNamedItemPhysConn = oPlnPhysConn.Object
                Dim oNamedItemFirstMeetAssy As IJNamedItem
                Set oNamedItemFirstMeetAssy = oAssyFirstMeet
                Dim oNamedItemAssyPlnJoint As IJNamedItem
                Set oNamedItemAssyPlnJoint = oAssyPlnJoint
                
                Dim strErrMsg As String
                strErrMsg = "FirstMeet assembly does not match the Planning Joint assembly" & vbNewLine _
                        & "Physical Connection: " + oNamedItemPhysConn.Name & vbNewLine _
                        & "First Meet Assembly: " + oNamedItemFirstMeetAssy.Name & vbNewLine _
                        & "Planning Joint Occurrence Assembly: " + oNamedItemAssyPlnJoint.Name
                            
                ' The arguments for IJCheckMfctyCallback.OnCheckError are:
                '   oPlnJoint              the object with error/warning
                '   PROG_ID                prog id of this rule
                '   ESeverity.siError      severity of the error (101 error or 102 warning)
                '   8                      code of this rule (found in spreadsheet
                '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
                '   strErrMsg              message for the error
                
                oCallBack.OnCheckError oPlnJoint, Module, _
                                       ESeverity.siError, 8, strErrMsg, "COM Repair Object", PROG_ID
                                       
    ' Note: cannot pass oPlnJointSO here because it does not support IJNamedItem.
    ' Perhaps the smart occurrence for Planning Joints should be modified to
    ' support IJNamedItem.
    '            oCallBack.OnCheckError oPlnJointSO, PROG_ID, _
    '                                   ESeverity.siError, 8, strErrMsg
                
            End If
        End If
        
    Next oPlnJoint
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oCallBack = Nothing
    Set oJoints = Nothing
    Set oPlnJoint = Nothing
    Set oPlnPhysConn = Nothing
    Set oAssyFirstMeet = Nothing
    Set oAssyPlnJoint = Nothing
    Set oNamedItemPhysConn = Nothing
    Set oNamedItemFirstMeetAssy = Nothing
    Set oNamedItemAssyPlnJoint = Nothing

End Sub ' IJCheckMfcty_ExecuteCheck



 
