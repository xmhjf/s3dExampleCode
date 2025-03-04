Attribute VB_Name = "basOpSSSBonAsm2"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DOpSSSealedBonAsm2.SealedBonAsm2
'   File            :  basOpSSSBonAsm2.cls
'   Author          :  PK
'   Creation Date   :  Friday,19 Oct 2007
'   Description     :  Stainless Steel Sealed Bonnet Assembly,Type 2
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   19.OCT.2007      PK        CR-127653:Created the symbol.
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

