Attribute VB_Name = "basOPSpurGear"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DOPSpurGear.OPSpurGear
'   File            :  basOPSpurGear.cls
'   Author          :  PK
'   Creation Date   :  Friday, Sept 10 2007
'   Description     :  Spur Gear type operator to be used with 3 way diverter combination valve
'                      of Tyco Flow Control
'   Reference       :  http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   10.Sep.2007     PK          CR-126718:Created the symbol.
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

