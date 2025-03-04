Attribute VB_Name = "basCSteamTrapAssembly"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basCSteamTrapAssembly.cls
'   ProgID:         SP3DSteamTrapAssembly.SteamTA
'   Author:         MP
'   Creation Date:  Wednesday, Oct 15 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'    15.Oct.2008    MP     CR-151135  Provide steam trap fitting unit symbols per Yarway catalog
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Public Const STA_DEFAULT = 1093
Public Const STA_Cock = 1094
Public Const STP_Cock_Bypass = 1095

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
