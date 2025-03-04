Attribute VB_Name = "basCPullPit"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basCPullPit.bas
'   Author:          RH
'   Creation Date:  01-May-08
'   Description:
'       This is Electrical Pull pit Assmebly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   13.June.2008     VRK     CR-134560:Provide pull-pit/manhole equipment symbol for use with duct banks
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    name As String
    Description As String
    properties As IMSDescriptionProperties
    uomValue As Double
End Type

Public Type OutputType
    name As String
    Description As String
    properties As IMSDescriptionProperties
    Aspect As SymbolRepIds
End Type

Public Type AspectType
    name As String
    Description As String
    properties As IMSDescriptionProperties
    AspectId As SymbolRepIds
End Type


'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

