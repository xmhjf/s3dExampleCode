VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBLDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: IBLDouble
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileBendingLineRule2

Const MODULE = "MfgMemberProcess"

Private Sub IJDMfgProfileBendingLineRule2_GetBendingLineParameters(ByVal oProfilePart As Object, _
                                                                   pBottomMargin As Double, _
                                                                   pTopMargin As Double, _
                                                                   pStartMargin As Double, _
                                                                   pEndMargin As Double, _
                                                                   pOverlapDist As Double, _
                                                                   pMaxNoOfBendingLines As Long, _
                                                                   pBendingLineOptions As enumBendingLineOptions, _
                                                                   dOffset As Double)
                                                                   
    Const METHOD = "IBLDouble: IJDMfgProfileBendingLineRule2_GetBendingLineParameters"
    On Error GoTo ErrorHandler
    
    'Number of inverse bending lines to be placed should be two
    pBottomMargin = 0        ' Start IBL at the bottom of the flange
    pTopMargin = 0.01        ' The IBL should always be 10mm below the flange
    pStartMargin = 0         ' Start the IBL at the edge of a profile
    pEndMargin = 0           ' The end of the IBL can be at the other end of the profile
    pOverlapDist = 0.75      ' Maximum overlap distance between the IBLs is 750 mm
    pMaxNoOfBendingLines = 6 ' Maximum number of bending lines allowed
    pBendingLineOptions = StartAtNeutralAxis
    dOffset = 0.5
    
        
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2028, , "RULES")
End Sub
