VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CableHoleCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 1998, 1999 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'       This class module contains the Custom Symbol Code required for processing
'       the CW hole Symbol
'
'Notes
'
'
'History
'
'    sypark@ship.samsung.co.kr    02/21/02                Creation.
'--------------------------------------------------------------------------------------------'

Option Explicit
Private Const MODULE = "CableHoleCM"
Private m_oErrors As New IMSErrorLog.JServerErrors

'********************************************************************
' Routine: ResymbCableWayHole
'
' Description:  Resymbolizes the Cable Way Hole
'
'********************************************************************
Public Sub ResymbCableWayHole _
        (ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
        ByRef ppArgument As Object)
    Const MT = "ResymbCableWayHole"
    On Error GoTo ErrorHandler
    
    ' QI the input for IJInput.
    Dim oInputDesc As IMSSymbolEntities.IJDInput
    Set oInputDesc = pInput
    
    ' Get the symbol definition.
    Dim oSymbolDef As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDef = oInputDesc.IJDInputDuringGame.Definition

    ' Get the HoleTraceAE argument
    Dim oSymbolDefinitionUtils As New CommonSymbolUtils.SymbolDefinitionUtils
    Set ppArgument = oSymbolDefinitionUtils.GetSymbolArgumentAtIndex(oInputDesc.Index, oSymbolDef)
    
'    ' What's the sense in this???  CRS
'    Dim oHoleTrace As Object
'    Set oHoleTrace = ppArgument
    
    GoTo Cleanup

ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
    
Cleanup:
    Set oInputDesc = Nothing
    Set oSymbolDef = Nothing
'    Set oHoleTrace = Nothing

End Sub

'********************************************************************
' Routine: ResymbParameterAsDimOrVar
'
' Description: Resymbolizes the input
'
'********************************************************************
Public Sub ResymbParameterAsDimOrVar(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                     ByVal pArgument As Object, ByRef ppResult As Object)
    Const MT = "ResymbParameterAsDimOrVar"
    On Error GoTo ErrorHandler

'$$$Debug   zResymbParameterAsDimOrVar "PenetrationCM::ResymbParameterAsDimOrVar", pInput, pArgument, ppResult
    
    ' use default ReSymbolization for Parameters (Dimensions and Variables)
    Dim oUtilsCommonCM As CommonSymbolUtils.CommonCM
    Set oUtilsCommonCM = New CommonCM
    oUtilsCommonCM.ResymbParameterAsDimOrVar pInput, pArgument, ppResult
    Set oUtilsCommonCM = Nothing
    
    Exit Sub
    
ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Sub

 