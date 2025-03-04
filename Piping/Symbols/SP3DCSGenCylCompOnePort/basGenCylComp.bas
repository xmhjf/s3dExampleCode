Attribute VB_Name = "basGenCylComp"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basGenCylComp.bas
'   Author:          KKC
'   Creation Date:  Tuesday, Nov 06 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Nov.2007     KKC       CR-126830  Provide one port cylindrical & rectangular on-the-fly piping specialty symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type
Public Type TextInputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    Value       As String
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
