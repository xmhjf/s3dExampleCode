Attribute VB_Name = "basCOnBrUnionTee"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basCOnBrUnionTee.bas
'   Author:         MS
'   Creation Date:  Monday, Aug 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: F128 Page No. D-66 of PDS Piping Component Data Reference Guide.
'  Symbol is created with Ten Outputs
'   The Four physical aspect outputs are created as follows:
'   ObjUnionBody- Using 'PlaceProjection' function,
'   One ObjNozzle object by using 'CreateNozzle' function and another ObjNozzle by using CreateNozzleWithLength
' The Six Insulation aspect outputs are created as follows:
' ObjInsulatedBody ,ObjInsulatedPort1, ObjInsulatedPort2 , ObjInsulatedBranch, ObjInsulatedPort3 and
' ObjInsulatedUnion using PlaceCylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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

