VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBLSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: IBLSingle
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileBendingLineRule
Const MODULE = "MfgMemberProcess"



Private Sub IJDMfgProfileBendingLineRule_GetBendingLineParameters(ByVal oMemberPart As Object, _
                                                                  pBottomMargin As Double, _
                                                                  pTopMargin As Double, _
                                                                  pStartMargin As Double, _
                                                                  pEndMargin As Double, _
                                                                  pOverlapDist As Double, _
                                                                  pMaxNoOfBendingLines As Long)
                                                                  
    Const METHOD = "IBLSingle: IJDMfgProfileBendingLineRule_GetBendingLineParameters"
    On Error GoTo ErrorHandler
    
    pBottomMargin = 0        ' Start IBL at the bottom of the flange
    pTopMargin = 0.01        ' The IBL should always be 10mm below the flange
    pStartMargin = 0         ' Start the IBL at the edge of a profile
    pEndMargin = 0           ' The end of the IBL can be at the other end of the profile
    pOverlapDist = 0.75      ' Maximum overlap distance between the IBLs is 750 mm
    pMaxNoOfBendingLines = 6 ' Maximum number of bending lines allowed
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
