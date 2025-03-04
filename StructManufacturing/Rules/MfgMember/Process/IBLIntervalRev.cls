VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBLIntervalRev"
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
'   The measurement points are every 500mm apart with the length from the left end of the profile. However it will be controlled by the profile length
'   Profile Length l >1,000mm:  Every 500mm
'   Profile Length 500=< l<999mm:   Every 250mm
'   Profile Length 250=< l<499mm:?      Every 125mm
'   Profile Length l < 249mmNo IBL
'   Request to customize the divided interval individually at each profile to considerate a curvature of profiles from 500, 250 and 125.
'   IBL must maintain minimum of 15mm from the above landing curve edge
'   IBL also must maintain minimum of 40mm from profile flange or top
'   Maximum number of IBL's is 5
'   If multiple IBLs are needed, curves will overlap by 3 divided points. (In case of 500mm, 1,000mm)
'   Intersection points of IBL's must be the divisible location
'   If multiple IBLs are placed center of over lapped divided line should be marked. Also length of 3D shape measure bottom end to end in below red arrow line in drawing and also height of marked center of divided line normal to red arrow line in green line of below figure

'*******************************************************************************
Option Explicit

Implements IJDMfgProfileBendingLineRule
Const MODULE = "MfgProfileProcess"
Private Sub IJDMfgProfileBendingLineRule_GetBendingLineParameters(ByVal oMemberPart As Object, pBottomMargin As Double, pTopMargin As Double, pStartMargin As Double, pEndMargin As Double, pOverlapDist As Double, pMaxNoOfBendingLines As Long)
    Const METHOD = "IBLSingle: IJDMfgProfileBendingLineRule_GetBendingLineParameters"
    On Error GoTo ErrorHandler
    
      
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

