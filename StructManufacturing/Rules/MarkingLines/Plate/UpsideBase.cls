VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingPlate
' Module: UpsideBase
'
' Description:  Determines the marking side for the mfg plate
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Implements IJDMfgPlateUpSideRule

Option Explicit

Private Const MODULE = "MfgMarkingPlate.UpsideBase"


Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "UpsideBase: IJDMfgPlateUpSideRule_GetPlateUpSide"

    On Error GoTo ErrorHandler
    
    IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    
    Exit Function
    
ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4006, , "RULES")
    
End Function
