VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointCheckPhysConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2004 Intergraph Corporation.  All Rights Reserved.
'
' File: PlnJointCheckPhysConn.cls
'
' Author: Dick Swager
'
' Abstract: Check manufacturability rules for Planning Joints.
'
' Description:
'
'********************************************************************

Option Explicit

Private Const Module = "PlnCheckMfcty.PlnJointCheckPhysConn"
Private Const PROG_ID = "PlnCheckMfcty.PlnJointCheckPhysConn"

Implements IJCheckMfcty

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description:
'Originally, it was required to have the physical connection updated by the
'user modification of the planning joint. Due to the Assoc loop that would
'result in this, the physical connection will no longer be updated by the
'planning joint. The following diagram shows the unavoidable loop that could
'occur:
'
'The user splitting a planning joint and setting the vertical direction and
'accessibility as shown would trigger the physical connection to recompute.
'This could conceivably cause a change of bevel on the physical connection.
'The change of bevel could cause one of the planning joints to be deleted.
'
'To prevent any possibility of a loop to be made, the physical connection
'will not be updated by the planning joint. Instead, check manufacturability
'will be used to determine if the physical connection needs to be updated due
'to the settings on the planning joints.
'
'The Manufacturability Check has to examine each Planning Joint and check that
'there is no change in bevel required on the associated physical connection.
'Marcel to elaborate the checks to be made.
'
'
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck( _
        ByVal oCollection As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)

    Const Method = "IJCheckMfcty_Execute"
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
    
        Dim oPlnJointProps As IJPlnJointProps
        Dim oPlnJointProdProps As IJPlnJointProdProps
        Set oPlnJointProps = oPlnJoint
        Set oPlnJointProdProps = oPlnJoint
        
        ' Get the weld properties from the Planning Joint.
        Dim eWeldDirection As JointWeldDirection
        eWeldDirection = oPlnJointProdProps.WeldDirection
        Dim eWeldPosition As JointWeldPosition
        eWeldPosition = oPlnJointProps.WeldPosition
        Dim dWeldRatio As Double
        dWeldRatio = oPlnJoint.WeldRatio
        Dim eWeldShape As JointWeldShape
        eWeldShape = oPlnJointProps.WeldShape
        Dim eWeldSide As JointWeldSide
                
        eWeldSide = oPlnJointProps.WeldSide
               
        
        ' Create the physical connection wrapper for the Planning Joint.
        Dim oPlnPhysConn As PlanningObjects.PlnPhysConn
        Set oPlnPhysConn = New PlanningObjects.PlnPhysConn
        Set oPlnPhysConn.object = oPlnJoint.GetPhysicalConnection

        ' Temporary: do something with the physical connection.
        If oPlnPhysConn.WeldName = "Name Unknown" Then
        
            ' An error condition exits.
            Dim oNamedItemJoint As IJNamedItem
            Set oNamedItemJoint = oPlnJoint
            
            Dim strErrMsg As String
            strErrMsg = "Weld name is unknown" & vbNewLine _
                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine
                    
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   7                      code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 7, strErrMsg, "", ""
            
            Set oNamedItemJoint = Nothing
        
        ElseIf Not OrieintationMatchesAssembly(oPlnJoint, strErrMsg) Then
        
            ' An error condition exits.
            Set oNamedItemJoint = oPlnJoint
            
            strErrMsg = strErrMsg & vbNewLine _
                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine
                    
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   7                      code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 7, strErrMsg, "", ""
            
            Set oNamedItemJoint = Nothing
        End If
        
    Next oPlnJoint
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oCallBack = Nothing
    Set oJoints = Nothing
    Set oPlnJoint = Nothing
    Set oPlnJointProdProps = Nothing
    Set oPlnJointProps = Nothing
    Set oPlnPhysConn = Nothing
    Set oNamedItemJoint = Nothing

End Sub


'******************************************************************************
' Routine: OrieintationMatchesAssembly
'
' Abstract: Checks that the orientaion of a Planning Joint matches that
'   of the containing Assembly.
'
' Description:
' Compare PC Upside with orientation of Assembly.
' Example: PC upside = Reference = Molded.  Check the molded conventions
' for the Deck to find that Thickness Direction is UP, therefore Assembly
' orientation should be Orientation X Rotation = 180 Degrees.
'
'
'******************************************************************************
Private Function OrieintationMatchesAssembly _
        (oPlnJoint As IJPlnJoint, _
         strMessage As String) As Boolean
    OrieintationMatchesAssembly = False
    strMessage = "Orientation of the Physical Connection" _
            & " does not match the orientation of the Assembly."
End Function


 
