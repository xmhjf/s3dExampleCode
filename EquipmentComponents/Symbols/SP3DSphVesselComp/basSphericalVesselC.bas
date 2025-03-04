Attribute VB_Name = "basSphericalVesC"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   basSphericalVesC.bas
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Spherical Vessel Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E230 Symbol in Page no 293.
'    Symbol is created using the following Outputs:
'    i)  Three standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: A Vessel uses 'PlaceSphere',
'       c) Two ReferenceGeometry aspect outputs: Default surface and a control point.
'    ii) Variable number of Supports (Can be either cylindrical or Cuboid) and
'       default surfaces (for cylindrical supportscase) are computed as per the user input.
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

