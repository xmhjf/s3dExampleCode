Attribute VB_Name = "basOffset"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   File:          basOffset.bas
'   Author:         VRK
'   Creation Date:  Wednesday, Jan 02 2008
'   Description:
'   Create offset symbol
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   02.Jan.2008     VRK        CR-131510:Create offset symbol to support the following options:
'                                          i.Offset, 45 degree
'                                         ii.Mechanical joint offset
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
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
Const E_FAIL = -2147467259
Err.Raise E_FAIL
End Sub

