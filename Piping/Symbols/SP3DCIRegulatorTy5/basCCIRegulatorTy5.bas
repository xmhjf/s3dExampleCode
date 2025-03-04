Attribute VB_Name = "basCCIRegulatorTy5"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   basCCIRegulatorTy5.bas
'   Author:          BG
'   Creation Date:  Thusday, Dec 26 2002
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.May.2008     KKC  CR-135970  Provide ability to rotate actuator for on-the-fly control valves
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    Name        As String
    Description As String
    Properties  As IMSDescriptionProperties
    uomValue    As Variant
End Type

Public Type TextInputType
    Name        As String
    Description As String
    Properties  As IMSDescriptionProperties
    Value       As String
End Type

Public Type OutputType
    Name            As String
    Description     As String
    Properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    Name                As String
    Description         As String
    Properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type

'Core Trader Key
Public Const TKValueMgr = "ValueMgr"
'Nozzles Information Keys
Public Const VKFlangeDiam1 = "Flange Diameter1"
Public Const VKFlangeDiam2 = "Flange Diameter2"
Public Const VKPipeDiam1 = "Pipe Diameter1"
Public Const VKPipeDiam2 = "Pipe Diameter2"

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

