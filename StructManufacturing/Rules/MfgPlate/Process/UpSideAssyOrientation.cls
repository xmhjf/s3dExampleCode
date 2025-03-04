VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssemblyOrientation"
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

Implements IJDMfgPlateUpSideRule

Private Const Module = "MfgPlateProcess.AssemblyOrientation"
Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
'    On Error GoTo ErrorHandler
    Const METHOD = "AssemblyOrientation: IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error Resume Next
   
    'Use the assembly to which the parts belongs as an input
    Dim oAssyObj As MfgAssemblyHelper
    Set oAssyObj = New MfgAssemblyHelper
    
    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = pPlatePart
    Set oAssyObj.object = oAssemblyChild.Parent
    
    Dim oAssyBasePlateOrientation As eBasePlateSide
    oAssyBasePlateOrientation = oAssyObj.BasePlateOrientationInAssembly
    
    If oAssyBasePlateOrientation = BasePlateSide_Base Then
        IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    Else
        If oAssyBasePlateOrientation = BasePlateSide_Offset Then
            IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
        Else
            'No base plate orientation is defined either since the baseplate is not defined yet
            IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
        End If
    End If
    

CleanUp:
    Set oAssemblyChild = Nothing
    Set oAssyObj = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1039, , "RULES")
    GoTo CleanUp
End Function

 
