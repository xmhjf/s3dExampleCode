Attribute VB_Name = "basCIGloVwtDiaAct"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basCIGloVwtDiaAct.bas
'   Author:          KKC
'   Creation Date:  Monday, Jun 23 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-137825  Piping instrument catalog should be updated to current Intools Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    Name        As String
    Description As String
    Properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type TextInputType
    Name        As String
    Description As String
    Properties  As IMSDescriptionProperties
    Value       As String
End Type

Public Type OutputType
    Name            As String
    Description     As String
    Properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    Name                As String
    Description         As String
    Properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type


'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

