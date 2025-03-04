Attribute VB_Name = "basCTestSupportLug"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basCTestSupportLug.bas
'   Author:          nka8411
'   Creation Date:  Tuesday, Feb 25 2003
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
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
