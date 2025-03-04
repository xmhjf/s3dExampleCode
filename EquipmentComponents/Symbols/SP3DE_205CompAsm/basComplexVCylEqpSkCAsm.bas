Attribute VB_Name = "basCCVCEqpSkirt"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   basCCVCEqpSkirt.bas
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This is Complex Vertical Cylindrical Equipment Skirt Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E205 Symbol in Page no 286.
'   Symbol is created using the following Outputs:
'   i)  4 standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: Vessel uses 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: a Default Surface and a Control point
'   ii) Variable Outputs:
'        a) Support
'        b) Surface for the support and
'        c) Intermediate dome for shell section 3
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

