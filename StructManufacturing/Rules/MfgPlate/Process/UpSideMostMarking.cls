VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MostMarking"
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

Private Const Module = "MfgPlateProcess.MostMarking"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "MostMarking: IJDMfgPlateUpSideRule_GetPlateUpSide"
    
    On Error GoTo ErrorHandler
    
    'Use the Plate Part as a reference
    Dim oSDPlatePart As StructDetailObjects.PlatePart
    Set oSDPlatePart = New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPlatePart
    
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = pPlatePart

    Dim count_base As Integer
    Dim count_offset As Integer
    
    count_base = 0
    count_offset = 0
        
    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = oSDPlatePart.ConnectedObjects
    
    Dim oConnectionData As ConnectionData
    Dim i As Long
    For i = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(i)
        
        Dim oPort As IJPort
        Set oPort = oConnectionData.ConnectingPort
        
        If oPlateWrapper.GetSide(oPort) = BaseSide Then
            count_base = count_base + 1
        End If
        
        If oPlateWrapper.GetSide(oPort) = OffsetSide Then
            count_offset = count_offset + 1
        End If
        
        Set oPort = Nothing
    Next i
    
    If count_base >= count_offset Then
        IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    Else
        IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    End If
    
CleanUp:
    Set oSDPlatePart = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1051, , "RULES")
    GoTo CleanUp
End Function

 
