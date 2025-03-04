VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MostStiffener"
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

Private Const MODULE = "MfgPlateProcess.MostStiffener"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "MostStiffener: IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    ' Use the plate part as input
    Dim oPlatePart As New StructDetailObjects.PlatePart
    Set oPlatePart.object = pPlatePart
    
    Dim side As String
    side = oPlatePart.StiffenerSide
    
    If side = "Base" Then IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    
    If side = "Offset" Then
        IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    Else
        IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide 'pick one or new type has to be defined!!
    End If
    
CleanUp:
    Set oPlatePart = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1052, , "RULES")
    GoTo CleanUp
End Function
 
