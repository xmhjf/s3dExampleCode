Attribute VB_Name = "basPSteamTrapTy5"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   basPSteamTrapTy5.bas
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    Symbol Model No. is: I253AZ Page No. 367 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type


'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

