Attribute VB_Name = "basSafetyShower07Asm"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   basCG4G_5421_02.bas
'   Author:         svsmylav
'   Creation Date:  Monday, Feb 24 2003
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    Name        As String
    Description As String
    Properties  As IMSDescriptionProperties
    Type As Variant
    StringValue As String
    uomType As UnitTypes
    uomValue As Double
'    PC As IJDParameterContent 'It could be DParameterContent
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
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
Const E_FAIL = -2147467259
Err.Raise E_FAIL
End Sub
