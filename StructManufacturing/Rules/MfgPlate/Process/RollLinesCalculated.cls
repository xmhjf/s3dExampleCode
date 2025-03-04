VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RollLinesCalculated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.RollLinesCalculated"

Implements IJDMfgProcessRollLines

Private Function IJDMfgProcessRollLines_ExecuteRule(ByVal pPart As Object) As Long
    Const METHOD = "RolllinesCalculated: IJDMfgProcessRollLines_ExecuteRule"
    On Error GoTo ErrorHandler

    Dim oSDPlatePart As StructDetailObjects.PlatePart
    Set oSDPlatePart = New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPart

    Dim oLength As Double
    oLength = oSDPlatePart.PlateLength * 1000 'Convert to mm

    ' KONI: Typecasting a double to an integer might not return the desired result
    ' Int (2.999999) will return 2 and not 3 !!!!
    IJDMfgProcessRollLines_ExecuteRule = Int(oLength / 300)
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1053, , "RULES")
End Function
 
