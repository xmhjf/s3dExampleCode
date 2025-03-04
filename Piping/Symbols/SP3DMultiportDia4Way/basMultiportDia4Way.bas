Attribute VB_Name = "basMultiportDia4Way"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   basMultiportDia4Way.bas
'   Author:         RUK
'   Creation Date:  Thursday, Sep 27 2007
'   Description:
'       Default:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is 72pdf. PDS symbol MC=VS3WD.
'
'       PDB: 462
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 2 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'       PDB: 880
'       This is a 4-way diverver valve symbol. This is prepared based on Gemu's catalog.
'       Source: Gemu Multiport Valves M600-5-4B, Aseptic valve manifold machined from a single block.
'       M600 valve manifold designs, developed and produced according to customer requirements/specifications
'       The symbol has multiple operators. Each operator for each output port
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'  17.Oct.2007      RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'                           Added code for the new PDB value 461.
'   09-June-2008    MP      CR-141585  Multiport valve symbols need to be enhanced to address Gemu valve requirements. (Implemented part data basis: 459)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basMultiportDia4Way:"    'Used for error messages
Public Const MULTI_PORT_OPTIONS_4WAY = 462
Public Const MULTI_PORT_OPTIONS_4WAY_GEMU = 880
Public Const STRAIGHT_INLET = 1
Public Const INLET_WITH_90DEG_ELBOW = 2
Public Const STRAIGHT_OUTLET = 1
Public Const OUTLET_WITH_90DEG_ELBOW = 2
Public Const OUTLET_WITH_OFFSET = 3

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

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

Public Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function


Public Function Max5(x1 As Double, x2 As Double, x3 As Double, x4 As Double, x5 As Double) _
       As Double
    Dim dmax As Double
    dmax = x1
    If CmpDblGreaterthan(x2, dmax) Then
        dmax = x2
    ElseIf CmpDblGreaterthan(x3, dmax) Then
        dmax = x3
    ElseIf CmpDblGreaterthan(x4, dmax) Then
        dmax = x4
    ElseIf CmpDblGreaterthan(x5, dmax) Then
        dmax = x5
    End If
    Max5 = dmax
End Function

