Attribute VB_Name = "basHandwheelD"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basHandwheelD.cls.cls
'   Author:         RUK
'   Creation Date:  Wednesday, 2 April 2008
'   Description:
'   This Symbol details were taken from Appendix E-107 of the Design Document
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Apr.2008     RUK     CR-CP-133524  Sample data for Dimensional Basis for valve operators is required
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    name As String
    description As String
    properties As IMSDescriptionProperties
    uomValue As Double
End Type

Public Type OutputType
    name As String
    description As String
    properties As IMSDescriptionProperties
    Aspect As SymbolRepIds
End Type

Public Type AspectType
    name As String
    description As String
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

