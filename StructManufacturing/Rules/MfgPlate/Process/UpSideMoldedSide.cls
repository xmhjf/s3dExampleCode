VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoldedSide"
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

Private Const MODULE = "MfgPlateProcess.MoldedSide"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "MoldedSide: IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    ' Use the plate part as input
    Dim oSDPlatePart As New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPlatePart
      
    Dim side As String
    
    If oSDPlatePart.plateType = Hull Then
        ' KLUDGE: We always want MoldedSide to be inside of Hull
        side = "Offset"
    Else
        side = oSDPlatePart.AlternateMoldedSide
    End If
    
    If side = "Base" Then IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    If side = "Offset" Then IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    
CleanUp:
    Set oSDPlatePart = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1050, , "RULES")
    GoTo CleanUp
End Function

 
