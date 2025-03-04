VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HoleDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurence\HoleRules
'
' File: HoleDefCM.cls
'
' Author: Hole Mgmt Team
'
' Abstract: definition custom methods for hole trace smart classes
'
' Description: These are the custom methods for the definitions for the hole
'              trace smart classes.
' History:
'   Date    By     Reason
' 08/11/2003     Added Code in ComputeHoleMatrix() for ConduitApp. TR 34661
' 04/26/2006     Appended Error Message to inform user that struct operation
'                are Not allowed.DI - 97082
' 06/08/06  CRS  Removed Hole Feature and Hole Fitting construction methods
'                (95% of the code).
' 06/12/06  CRS  Removed Hole Status references
'
'******************************************************************************

Option Explicit

Const MODULE = "HoleDefCM"

Private sERROR As String

'***********************************************************************
'***********************************************************************
'
' Start of conditional methods used by the member descriptions
'
'***********************************************************************
'***********************************************************************

'***********************************************************************
' METHOD:  CM_IsApproved
'
' DESCRIPTION:  Determine if the hole is "Approved"
'***********************************************************************
Public Sub CM_IsApproved(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    Const sMETHOD = "CM_IsApproved"
    
    bIsNeeded = False

    Dim oHoleTrace As IJHoleTraceAE
    Set oHoleTrace = GetHoleTrace(pMD)
    If (oHoleTrace Is Nothing) Then Exit Sub
    
Cleanup:
    Set oHoleTrace = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    GoTo Cleanup
End Sub


'***********************************************************************
'***********************************************************************
' Start of private methods used by the custom methods
'***********************************************************************
'***********************************************************************

'***********************************************************************
' METHOD:  GetHoleTrace
'
' DESCRIPTION:  get the hole trace from the member description
'***********************************************************************
Private Function GetHoleTrace(ByRef pMD As IJDMemberDescription) As IJHoleTraceAE
    On Error GoTo ErrorHandler
    Const sMETHOD = "GetHoleTrace"

    sERROR = "Getting hole trace from smart occurrence"
    Dim oHoleSmartOcc As IJHoleSmartOcc
    Set oHoleSmartOcc = pMD.CAO
    
    Dim oHoleTrace As IJHoleTraceAE
    Set oHoleTrace = oHoleSmartOcc.GetHoleTraceAE
    If (oHoleTrace Is Nothing) Then GoTo ErrorHandler
    
    Set GetHoleTrace = oHoleTrace
    
    Set oHoleTrace = Nothing
    Set oHoleSmartOcc = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function



 