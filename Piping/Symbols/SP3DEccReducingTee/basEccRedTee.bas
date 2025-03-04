Attribute VB_Name = "basEccRedTee"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basC90Elbow.bas
'   Author:          KKC
'   Creation Date:  Wednesday, Jul 30, 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   30.Jul.2008     KKC      CR- 146404 Enhance Eccentric Reducing Tee symbol for seat-to-seat dimension per JIS G 5527
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

