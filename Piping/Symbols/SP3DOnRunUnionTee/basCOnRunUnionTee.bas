Attribute VB_Name = "basCOnRunUnionTee"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basCOnRunUnionTee.bas
'   Author:         BG
'   Creation Date:  Monday, Aug 27 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No. is: F127 Page No. D-66 of PDS Piping Component Data Reference Guide.
'   Symbol is created with Ten Outputs
'   The Four physical aspect outputs are created as follows: 1.Union of shape Hexagon made up of
'   Line String and projecting it, 2.Nozzle-1 with length towards -ive X-axis, 3. Nozzle-2 along +ive X-axis
'   and 4. Nozzle-3 with length along +ive Y-axis.
'   Insulation aspect consist of 1. Insulation for body, 2. Nozzle-1, 3. Nozzle-2, 4. Insulation for Branch,
'   5. Insulation for Nozzle-3 and 6. Insulation for Union.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
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

