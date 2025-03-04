VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideOuter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingPlate
' Module: UpsideOuter
'
' Description:  Determines the marking side for the mfg plate
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Implements IJDMfgPlateUpSideRule

Option Explicit

Private Const MODULE = "MfgMarkingPlate.UpsideOuter"


Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "UpsideOuter: IJDMfgPlateUpSideRule_GetPlateUpSide"

    On Error GoTo ErrorHandler
        
    IJDMfgPlateUpSideRule_GetPlateUpSide = UnDefinedSide
    
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = pPlatePart
    
    If Not oSDPlateWrapper.plateType = Hull Or oSDPlateWrapper.plateType = LBulkheadPlate Then
         Exit Function
    End If

    Dim oGeomUtils As IJTopologyLocate
    Set oGeomUtils = New TopologyLocate
        
    Dim oStructConnectable  As IJStructConnectable
    Dim oEnumPorts          As IJElements
    Set oEnumPorts = New JObjectCollection
    
    Set oStructConnectable = pPlatePart
    oEnumPorts.AddElements oStructConnectable.enumAllPorts

    Dim j As Integer
    Dim oPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oCenter As IJDPosition
    Dim oNormal As IJDVector

    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
        If oStructPort.ContextID = CTX_BASE Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
             
            If oCenter.Y > 0 Then
                If oNormal.Y > 0 Then
                    IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
                    Exit For
                End If
            ElseIf oCenter.Y < 0 Then
                If oNormal.Y < 0 Then
                    IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
                    Exit For
                End If
            End If
            
         ElseIf oStructPort.ContextID = CTX_OFFSET Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
             
            If oCenter.Y > 0 Then
                If oNormal.Y > 0 Then
                    IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
                    Exit For
                End If
            ElseIf oCenter.Y < 0 Then
                If oNormal.Y < 0 Then
                    IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
                    Exit For
                End If
            End If
            
        End If
        
    Next j
    
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oGeomUtils = Nothing
    Set oEnumPorts = Nothing
    Set oStructConnectable = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4011, , "RULES")
    GoTo CleanUp
End Function
