Attribute VB_Name = "CommonFuncs"
 '/*******************************************************************
'Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructManufacturing\Middle\Symbols\ProcessAndMarkingSymbols\StructMfgSymbols.vbp
'
'File: S:\StructManufacturing\Middle\Symbols\ProcessAndMarkingSymbols\CommonFuncs.bas
'
'Revision:
'     02/07/01 GDreybus
'
' Description:
' The functions and subs in this file are commonly used by slot
' symbols.
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "StructMfgSymbols.CommonFuncs(CommonFunc.bas)"

' There is one reference input for the plate symbols
Sub SetPlateCommonInputs(pDef As IJDSymbolDefinition)
  
  Dim pInput As IJDInput
  Set pInput = New DInput

  pInput.Name = "PlatePart"
  pInput.Description = "Plate Part"
  pInput.Index = igPlateIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset

End Sub
' There is one reference input for the plate symbols
Sub SetProfileCommonInputs(pDef As IJDSymbolDefinition)
  
  Dim pInput As IJDInput
  Set pInput = New DInput

  pInput.Name = "ProfilePart"
  pInput.Description = "Profile Part"
  pInput.Index = igProfileIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset

End Sub

' There is  one reference input(PlatePart) for the template symbols
Sub SetTemplateCommonInputs(pDef As IJDSymbolDefinition)
    Dim pInput As IJDInput
    Set pInput = New DInput

    pInput.Name = "PlatePart"
    pInput.Description = "Plate Part"
    pInput.Index = igPlateIndex
    pDef.IJDInputs.Add pInput
    pInput.Reset
  
End Sub

 