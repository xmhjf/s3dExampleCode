Attribute VB_Name = "basC5TCarryDeckCrane"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basC5TCarryDeckCrane.bas
'   Author:          CYW
'   Creation Date:  Thursday, Mar 27 2003
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
