VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBLStartAtBottom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileBendingLineRule2

Const MODULE = "MfgProfileProcess"

Private Sub IJDMfgProfileBendingLineRule2_GetBendingLineParameters(ByVal oProfilePart As Object, pBottomMargin As Double, pTopMargin As Double, pStartMargin As Double, pEndMargin As Double, pOverlapDist As Double, pMaxNoOfBendingLines As Long, pBendingLineOptions As enumBendingLineOptions, dOffset As Double)
    Const METHOD = "IBLDouble: IJDMfgProfileBendingLineRule2_GetBendingLineParameters"
    On Error GoTo ErrorHandler
    
    'Number of inverse bending lines to be placed should be two
    pBottomMargin = 0        ' Start IBL at the bottom of the flange
    pTopMargin = 0.01        ' The IBL should always be 10mm below the flange
    pStartMargin = 0         ' Start the IBL at the edge of a profile
    pEndMargin = 0           ' The end of the IBL can be at the other end of the profile
    pOverlapDist = 0.75      ' Maximum overlap distance between the IBLs is 750 mm
    pMaxNoOfBendingLines = 6 ' Maximum number of bending lines allowed
    pBendingLineOptions = StartAtBottom
    dOffset = 0#
    
    
    ' Use the profile part as input
    Dim oProfileClass As Object
    If TypeOf oProfilePart Is IJStiffenerPart Then
        Set oProfileClass = New StructDetailObjects.ProfilePart
    ElseIf TypeOf oProfilePart Is IJBeamPart Then
        Set oProfileClass = New StructDetailObjects.BeamPart
    End If
    Set oProfileClass.object = oProfilePart
    
    ' Check if the cross-section is a built-up
    Dim bIsBuiltUp As Boolean
    bIsBuiltUp = oProfileClass.IsCrossSectionABuiltUp
    
    If bIsBuiltUp = False Then
        ' If it's not a built-up, then we adjust the TopMargin to avoid painting where the flange is
        pTopMargin = pTopMargin + oProfileClass.FlangeThickness
    End If
    
    'adjust the active region for margins
    Dim oStructConnectable As IJStructConnectable

    If TypeOf oProfilePart Is IJStructConnectable Then
        Dim oBasePort As IJPort
        Dim oOffsetPort As IJPort
        Dim oLateralPortsCol As IJElements

        Set oStructConnectable = oProfilePart
        oStructConnectable.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol

        If Not oBasePort Is Nothing Then
            pStartMargin = -1 * GetMarginValueOnThePort(oBasePort)
        End If
        If Not oOffsetPort Is Nothing Then
            pEndMargin = -1 * GetMarginValueOnThePort(oOffsetPort)
        End If
    End If
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
