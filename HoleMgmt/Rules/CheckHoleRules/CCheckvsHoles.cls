VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckvsHoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: CCheckvsHoles.cls
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: Check the parents of structure and the minmun distance between Hole and Hole
'
' Description:
' IJCheckMfcty is implemented by GrobalCheck.tlb.
'********************************************************************

Option Explicit

Private Const Module = "GSCADHMRules.CCheckvsHoles:"
Private Const PROG_ID = "GSCADHMRules.CCheckvsHoles"

Implements IJCheckMfcty

Dim m_RuleMinDistance As Double   'For minimum distance between hole and hole.

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeUserDefined
'
'Interface: Private function
'
'Abstract: This method is to set the minimum distance between hole and hole
'          User can set the minimum distance
'
'Attention : The unit of measure should be mm
'
'   For instance
'   m_RuleMinDistance = 0.05 is that the minimum distance is 50 mm
'
'********************************************************************
Private Sub InitializeUserDefined()
    Const METHOD = "InitializeUserDefined"
    On Error GoTo ErrorHandler
    
    m_RuleMinDistance = 0.05             'If the nearest distance between two holes is less than 50 mm, should be checked.

Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description: 1. Set the rule minimum distance, Call "InitializeUserDefined"
'              2. Get the distance between hole and hole, Call "GetMinimumDistance"
'              3. Compare the minimum distance and distance between holes, call "CheckRules"
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck(ByVal pCollection As IMSCoreCollections.IJElements, _
                                      ByVal bStopOnError As Boolean, _
                                      pOptionCodes() As Variant, _
                                      ByVal pCallingObj As Object)
    Const METHOD = "IJCheckMfcty_Execute"
    On Error GoTo ErrorHandler

    Dim oCallBack As IJCheckMfctyCallback
    Dim oColHoles As IMSCoreCollections.IJElements

    Dim oCheckingTrace As IJHoleTraceAE
    Dim oCheckingCurves As IMSCoreCollections.IJElements

    Dim oComparedTrace As IJHoleTraceAE
    Dim oComparedCurves As IMSCoreCollections.IJElements

    Dim lngDummy As Long
    Dim lngCancel As Long
    Dim dblMinDistance As Double

    'Set the m_RuleMinDistance
    InitializeUserDefined

    'If workingset is nothing, exits.
    If pCollection.Count = 0 Then GoTo Cleanup

    'Get the holes from pCollection. This collection is workingset or selectset
    SetCollectionHoles pCollection, oColHoles

    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    'Check the parents of holeTrace.
    For Each oCheckingTrace In oColHoles
        CheckOverall PROG_ID, oCheckingTrace, oCallBack, 1
    Next oCheckingTrace

    'Those elements has a graphicentity like IJLine and IJARC from HoleTrace Symbol
    Set oCheckingCurves = New IMSCoreCollections.JObjectCollection
    Set oComparedCurves = New IMSCoreCollections.JObjectCollection

    'For loop and get the minimum distance, Between hole and hole
    For Each oCheckingTrace In oColHoles
        'Get the curves from HoleTrace
        Set oCheckingCurves = CanRetriveCurvesFromHole(PROG_ID, oCheckingTrace, oCallBack, 1)
        'if the hole is wrong, We don't have curves, display this hole information on the list.
        If Not oCheckingCurves Is Nothing Then
            If oCheckingCurves.Count <> 0 Then
                For Each oComparedTrace In oColHoles
                    If Not oCheckingTrace Is oComparedTrace Then
                        'Just compare the holes on the same structure
                        If IsSameParent(oCheckingTrace, oComparedTrace) Then
                            Set oComparedCurves = GetHoleTraceCurves(oComparedTrace)
                            
                            'if the hole is wrong, Skip to check.
                            If Not oComparedCurves Is Nothing Then
                                'Get the minimun distance
                                dblMinDistance = GetMinimumDistance(oCheckingCurves, _
                                                                    oComparedCurves)
                                If dblMinDistance >= 0 Then
                                    'Check the distance and return callback
                                    CheckRulesAndCallBack PROG_ID, oCheckingTrace, oComparedTrace, _
                                                          dblMinDistance, m_RuleMinDistance, oCallBack, 1
                                End If
                            End If
                        End If
                    End If
                Next oComparedTrace
            End If
        End If
    Next oCheckingTrace

Cleanup:
    Set oCallBack = Nothing
    Set oColHoles = Nothing
    Set oCheckingTrace = Nothing
    Set oCheckingCurves = Nothing
    Set oComparedTrace = Nothing
    Set oComparedCurves = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub
 
