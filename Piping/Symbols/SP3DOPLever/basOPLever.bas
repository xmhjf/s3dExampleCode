Attribute VB_Name = "basOPLever"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOPLever.OPLever
'   Author:         RUK
'   Creation Date:  Wednesday, Sugust 27, 2008
'   Description: This symbol is graphical implementation of Lever operator
'   Source: E-141 section of the Design document)
'   This symbol implements following partdatabasis
'   Default
'   Lever, Type 1   (60)
'   Lever, inclined, Type 1 (61)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.08.2008      RUK     CR-148069  Provide a more realistic lever operator symbol (Source - E-141 section of the Design document)
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

