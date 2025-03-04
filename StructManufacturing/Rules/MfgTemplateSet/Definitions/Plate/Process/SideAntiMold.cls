VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideAntiMold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Anti-Molded Side Process rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

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
    If oSDPlatePart.PlateType = Hull Then
        ' KLUDGE: We always want MoldedSide to be inside of Hull
        side = "Offset"
        
    Else
        side = oSDPlatePart.AlternateMoldedSide
    End If
    
    If side = "Base" Then IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    If side = "Offset" Then IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    
Cleanup:
    Set oSDPlatePart = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
    GoTo Cleanup
End Function


 
