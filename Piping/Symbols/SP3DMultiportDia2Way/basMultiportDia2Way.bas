Attribute VB_Name = "basMultiportDia2Way"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   basMultiportDia2Way.bas
'   Author:         RUK
'   Creation Date:  Thursday, Sep 27 2007
'   Description:
'       PDB: 460
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 2 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'       PDB: 459
'       This is a 2-way diverter valve symbol. This is prepared based on Gemu's catalog.
'       Source: Gemu  Multiport Valves M600-3-2C, Aseptic valve manifold machined from a single block.
'       M600 valve manifold designs, developed and produced according to customer requirements/specifications
'       The symbol has multiple operators. Each operator for each output port
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----   ------------------
'   27.Sep.07       RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols. (Implemented part data basis: Default, 460)
'   09-June-2008    MP      CR-141585  Multiport valve symbols need to be enhanced to address Gemu valve requirements. (Implemented part data basis: 459)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basMultiportDia2Way:" 'Used for error messages
Public Const MULTI_PORT_OPTIONS_2WAY = 460
Public Const MULTI_PORT_OPTIONS_2WAY_GEMU = 459
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

