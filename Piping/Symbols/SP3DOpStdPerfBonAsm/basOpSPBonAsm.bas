Attribute VB_Name = "basOpSPBonAsm"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DOpStdPerfBonAsm.OpSPBonAsm
'   File            :  basOpSPBonAsm.cls
'   Author          :  PK
'   Creation Date   :  Thursday,20 Sep 2007
'   Description     :  Standard Performance Bonnet Assembly
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   20.SEP.2007      PK        CR-127653:Created the symbol.
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

