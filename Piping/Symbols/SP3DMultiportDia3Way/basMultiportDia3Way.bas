Attribute VB_Name = "basMultiportDia3Way"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basMultiportDia3Way.bas
'   Author:         RUK
'   Creation Date:  Thursday, Sep 27 2007
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is 72pdf. PDS symbol MC=VS3WD.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basMultiportDia3Way:" 'Used for error messages
Public Const MULTI_PORT_OPTIONS_3WAY = 461
Public Const STRAIGHT_INLET = 1
Public Const INLET_WITH_90DEG_ELBOW = 2
Public Const STRAIGHT_OUTLET = 1
Public Const OUTLET_WITH_90DEG_ELBOW = 2
Public Const OUTLET_WITH_OFFSET = 3

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


