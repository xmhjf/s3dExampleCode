Attribute VB_Name = "basTeeStrainer"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basTeeStrainer.bas
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This symbol is created for T Strainer to support multiple dimensional bases as per CR-113397
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------     -----        ------------------

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

