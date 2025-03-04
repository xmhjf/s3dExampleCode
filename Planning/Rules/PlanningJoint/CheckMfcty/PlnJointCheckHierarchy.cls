VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointCheckHierarchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2004 Intergraph Corporation.  All Rights Reserved.
'
' File: PlnJointCheckHierarchy.cls
'
' Author: Dick Swager
'
' Abstract: check manufacturability rules for planning joints
'
' Description: check that the planning joints is located in an assembly
'              that contains both parts
'
'********************************************************************

Option Explicit

Implements IJCheckMfcty

Private Const PROG_ID = "PlnCheckMfcty.PlnJointCheckHierarchy"
Private Const Module = "PlnCheckMfcty.PlnJointCheckHierarchy"

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: implemented by IJCheckMfcty interface
'
' Description:
'  this check needs to ensure that the planning joints belong to an assembly
'  that contains both parts being welded in its tree
'
' PseudoCode:
'   for each planning joint
'       get the physical connection to get the two parts
'       get the assembly the planning joint is assigned to
'       check the two parts are underneath the assembly hierarchy
'   end for
'
' Note:
'   in the following statement
'            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
'                                   ESeverityIndex.siError, 9, strErrMsg
'       oPlnJoint              the object with error/warning
'       PROG_ID                prog id of this rule
'       ESeverity.siError      severity of the error (101 error or 102 warning)
'       9                      code of this rule (found in spreadsheet)
'       strErrMsg              message for the error
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck(ByVal oCollection As GSCADGlobalCheck.IJElements, _
                                      ByVal bStopOnError As Boolean, _
                                      varOptionCodes() As Variant, _
                                      ByVal oCallingObj As Object)
    Const Method = "IJCheckMfcty_ExecuteCheck"
    On Error GoTo ErrorHandler

    ' setup the check manufacturability call back mechanism
    Dim oCallBack As IJCheckMfctyCallback
    Dim lngDummy As Long
    Dim lngCancel As Long

    Set oCallBack = oCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' if the collection is empty, do not continue
    If oCollection.Count = 0 Then GoTo Cleanup
    
    ' get all the planning joints from the collection
    Dim oJoints As IJElements
    GatherPlanningJoints oCollection, oJoints
    
    Dim oAssembly As IJAssembly
    Dim oTestAssy As IJAssembly
    Dim oHoldAssy As IJAssembly
    Dim oNamedItem As IJNamedItem
    
    Dim oPhysConnWrap As PlanningObjects.PlnPhysConn
    Set oPhysConnWrap = New PlanningObjects.PlnPhysConn
    
    Dim nCount As Long
    Dim nMaxCount As Long
    nCount = 0
    nMaxCount = oJoints.Count
    
    ' look at all planning joints in the collection
    Dim oPlnJoint As IJPlnJoint
    For Each oPlnJoint In oJoints
        Set oAssembly = oPlnJoint.GetAssembly
        
        Set oPhysConnWrap.Object = oPlnJoint.GetPhysicalConnection
        Set oTestAssy = oPhysConnWrap.FirstMeet
        Do Until oTestAssy Is oAssembly
            Set oHoldAssy = oTestAssy
            Set oTestAssy = Nothing
            
            On Error Resume Next
            Set oTestAssy = oHoldAssy.GetParent
            On Error GoTo ErrorHandler
            Set oHoldAssy = Nothing
            
            If oTestAssy Is Nothing Then Exit Do
        Loop
        
        If Not oTestAssy Is oAssembly Then
            Set oNamedItem = oPlnJoint
            
            Dim strErrMsg As String
            strErrMsg = "Planning joint " + oNamedItem.Name + " is not in an assembly that contains the two joined parts"
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 9, strErrMsg, "", ""
            
            Set oNamedItem = Nothing
        End If
        
        Set oAssembly = Nothing
        Set oTestAssy = Nothing
        
        ' update the progress bar - first argument is the progress as a
        ' percentage (i.e. 0 - 100) - the second argument is a cancellation flag
        nCount = nCount + 1
        oCallBack.OnProgress nCount * 100 / nMaxCount, 0
    Next oPlnJoint
    
Cleanup:
    Set oCallBack = Nothing
    Set oJoints = Nothing
    Set oPlnJoint = Nothing
    Set oAssembly = Nothing
    Set oTestAssy = Nothing
    Set oHoldAssy = Nothing
    Set oPhysConnWrap = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
End Sub

 
