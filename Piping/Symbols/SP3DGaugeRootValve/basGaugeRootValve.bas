Attribute VB_Name = "basGaugeRootValve"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basGaugeRootValve.bas
'   Author:          MA
'   Creation Date:  Tuesday, May 2007
'   Description:
'   The following Part data basis cases are addressed for the parameters specified:
'   Case A (Part data Basis value -343): (Gauge Root Valve with Single Outlet)
'                                          FacetoFace,ValveWidth and Offset
'   Case B (Part data Basis value -345): (Gauge Root Valve with Multiple Outlet)
'                                 FacetoFace,ValveWidth,Offset,Port3Offset and Port4Offset
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----        ------------------
'   24.May.2007    MA       CR-113431: Implemented Part data basis for values 343 and 345.
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

