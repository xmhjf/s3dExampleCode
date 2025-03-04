Attribute VB_Name = "basVerticalRotatingEqpC"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   basVerticalRotatingEqpC.bas
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'   This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'   E410 Symbol in Page no 325.
'   Symbol is created using the following Outputs:
'       a) One Insulation aspect outputs,
'       b) One maintenance aspect output and
'       c) Two Physical aspect outputs: , top surface
'           ObjEquipment uses 'PlaceCylinder' and
'       d) Two ReferenceGeometry aspect outputs: A default Surface and a control point.
'       Other Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
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

