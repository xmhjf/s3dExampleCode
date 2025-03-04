Attribute VB_Name = "basSpacer"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   basSpacer.bas
'   Author:         RUK
'   Creation Date:  Monday, Feb 04, 2008
'   Description:
'       This is Open/Blind spcer symbol. This is prepared based Appendex E-94 in Piping Design Document L57.
'        CR-134984  Provide symbol for open spacer and blind spacer set
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Feb.2008     RUK     CR-134984  Provide symbol for open spacer and blind spacer set
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basSpacer:" 'Used for error messages

Public Const OPEN_SPACER_INSTALLED = 5
Public Const BLIND_SPACER_INSTALLED = 10

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


