VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AntiMoldedSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg template
'
' Author:
'
' Comments:
' 2006.01.11    Ram K.        creation
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgTemplateProcess.AntiMoldedSide"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    ' Use the plate part as input
    Dim oSDPlatePart As StructDetailObjects.PlatePart
    Set oSDPlatePart = New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPlatePart
      
    Dim side As String
    side = oSDPlatePart.AlternateMoldedSide
    
    If side = "Base" Then IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    If side = "Offset" Then IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    
Cleanup:
    Set oSDPlatePart = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
    GoTo Cleanup
End Function


 
