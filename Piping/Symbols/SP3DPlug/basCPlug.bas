Attribute VB_Name = "basCPlug"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   basCPlug.bas
'   Author:          NN
'   Creation Date:  Thursday, Jan 25 2001
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   16.JUL.2008     MP   CR-145604 assigned Constants for  Part data basis options 1028,1029,1030,1031 and 1032
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Const PLUG_DEFAULT = 1031
Public Const PLUG_SQUARE = 1032
Public Const PLUG_HEXAGON = 1028
Public Const PLUG_OCTAGON = 1029
Public Const PLUG_PENTAGON = 1030
Public Const PLUG_PLAIN = 1084
Public Const PLUG_ROUND = 1085
Public Const PLUG_COUNTERSUNK = 1086

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
