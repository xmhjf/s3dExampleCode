Attribute VB_Name = "basCExhaustHeight"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   basCExhaustHead.bas
'   Author:          Babu Govindarajan
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
'    This symbol constructed as per the catalog available at URL http://www.nciweb.net/exhaust.htm
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                ------------------
'   09.Jul.2003     SymbolTeam(India)    Copyright Information, Header  is added.
'   27.Jan.2006     Sundar(svsmylav)     RI-28367: Revision history is updated with hyper link to Yardney's site.
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

