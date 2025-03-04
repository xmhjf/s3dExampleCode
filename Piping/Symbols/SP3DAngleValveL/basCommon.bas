Attribute VB_Name = "basCommon"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   basCommon.bas
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'        Angle Valve
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Common Module

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
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)

Const METHOD = "ReportUnanticipatedError:"
'  Dim oTrader As Trader
'  Set oTrader = New Trader
'  Dim oErrService As IJErrorService
'  Set oErrService = oTrader.Service(TKErrorHandler, "")
'
'  If oErrService Is Nothing Then
'    MsgBox "Error Service is unavailable: Error was in module: " & InModule & " method: " & InMethod
'    Set oTrader = Nothing
'    Exit Sub
'  End If
'
'  Dim ern As JWellKnownErrorNumbers
'  ern = oErrService.ReportError(Err.Number, InModule & " " & InMethod, " " & Err.description, App)
'  Select Case ern
'    Case imsAbortApplication:
'      oErrService.TerminateApp
'    Case Else
'      'TODO: By default it stops the command.
'      '      Change it to take in account your errors
'  End Select
'
'  Set oErrService = Nothing
'  Set oTrader = Nothing
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
  
End Sub
