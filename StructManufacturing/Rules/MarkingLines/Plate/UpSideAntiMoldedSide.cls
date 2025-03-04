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
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingPlate
' Module: AntimoldedSide
'
' Description:  Determines the marking side for the mfg plate
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Implements IJDMfgPlateUpSideRule

Option Explicit

Private Const MODULE = "MfgMarkingPlate.AntiMoldedSide"

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "AntimoldedSide: IJDMfgPlateUpSideRule_GetPlateUpSide"

    On Error GoTo ErrorHandler

    ' Use the plate part as input
    Dim oSDPlateWrapper As New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = pPlatePart

    Dim side As String
    If oSDPlateWrapper.PlateType = Hull Then
        ' KLUDGE: We always want MoldedSide to be inside of Hull
        side = "Offset"
        
    Else
        side = oSDPlateWrapper.AlternateMoldedSide
    End If
 
    If side = "Base" Then IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    If side = "Offset" Then IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide

CleanUp:
    Set oSDPlateWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4005, , "RULES")
    GoTo CleanUp
End Function
 
