VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrainHoleParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : DrainHoleParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.DrainHoleParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.DrainHoleParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\DrainHoleParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    oIH.SetInput INPUT_PORT1FACE
    oIH.SetInput INPUT_PORT2EDGE
    oIH.SetInput INPUT_PORT3EDGE

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
 
End Sub
'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
    oOH.SetOutput "OffsetFromWeb"
    oOH.SetOutput "OffsetFromBase"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
 
End Sub
'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(oPLH As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    ' Currently drain holes are only applied to connected clips for SlotA and SlotC:
    ' Primary clips:  Appliable to SlotA and SlotC
    '    ClipAAC --- "FB", "BUT", "BUTL2", "BUTL3",  "EA", "T_XType", "TSType", "UA",
    '    ClipBAC --- "B"

    ' Secondary clips:  Appliable to SlotA only
    '    ClipABC --- "FB", "BUT", "BUTL2", "BUTL3",  "EA", "T_XType", "TSType", "UA",
    '    ClipBBC --- "B"
    '
    '    dClipWebToSlotEdge is determined by cross section and clip type
    '    In following examples,for primary clips:
    '      If cross section type is "FB"
    '          dClipWebToSlotEdge = dSlotFlangeClearance - dClipWebClearance
    '      If cross section type is "UA"
    '          dClipWebToSlotEdge = dProfileFlangeLength + dSlotFlangeClearance - dClipWebClearance
    '
    '                                                                                              |---   dProfileFlangeLength
    '                            | ---- dSlotFlangeClearance                            |              | ---  dSlotFlangeClearance
    '                  |<-------------------->|                                      |<-------------------->|<->|
    '        xxxxxxxxxxxxxxxxxxxx                               xxxxxxxxxxxxxxxxxxxxxx
    '       x                                 x                              x                                    x
    '       x      ==                       x                              x     =============     x
    '         x    |  |                        x                              x    | |                             x
    '           x  |  |                        x                              x    | |                             x
    '            x |  | ******************************                 x    | | **********************************
    '            x |  | *                      x            *                x    | | *                           x            *
    '            x |  | *                      x            *                x    | | *                           x            *
    '            x |  | *                      x            *                x    | | *                           x            *
    '            x |  | *                      x            *                x    | | *                           x            *
    '            x |  | *                      x            *                x    | | *                           x            *
    ' ----------------------------------------------------------------------------------------------------------------------------------------
    '                   |<-------------------->|                                      |<------------------------->|
    '                               | ---------- dClipWebToSlotEdge                      | ---------- dClipWebToSlotEdge
    '
    ' For secondary clips:
    '    If cross section type is "T"
    '        dClipWebToSlotEdge = (dProfileFlangeLength - dProfileWebThickness)/2 + dSlotFlangeClearance - dClipWebClearance
    '    If cross section type is "F"
    '        dClipWebToSlotEdge = dSlotBackClearance - dClipWebClearance
    '
    '                  | -- dSlotFlangeClearance
    '                  |        | -- dProfileFlangeLength                | -- dSlotBackClearance
    '              |<----->|<------------->|                               |<-------------->|
    '              xxxxxxxxxxxxxxxxxxxxxx                    xxxxxxxxxxxxxxxxxxxxxx
    '              x                                   x                    x                                   x
    '              x        =========         x                    x                 ====           x
    '              x                ||                  x                    x                ||                 x
    '              x                ||                  x                    x                ||                 x
    '    ********************   ||                  x         *******************     ||                 x
    '    *         x           *    ||                  x         *         x           *    ||                 x
    '    *         x           *    ||                  x         *         x           *    ||                 x
    '    *         x           *    ||                  x         *         x           *    ||                 x
    '    *         x           *    ||                  x         *         x           *    ||                 x
    ' -------------------------------------------------------------------------------------------------------------------------------------
    '               |<-------->|                                  |<------------------->|
    '                      | ---------- dClipWebToSlotEdge        | ---------- dClipWebToSlotEdge
    
    '    Current logic for setting drain hole offset from web is:
    '      If dClipWebToSlotEdge < 2 * dDiameter then
    '          dOffsetFromWeb = dClipWebToSlotEdge/2
    '      else
    '          dOffsetFromWeb = 65mm
    '
    '      Note: dOffsetFromBase is always set to 45mm
    '
    Dim oSysChild As IJSystemChild
    Dim oCollar As Object
    
    Set oSysChild = oPLH.SmartOccurrence
    Set oCollar = oSysChild.GetParent()
    Set oSysChild = Nothing
    
    Dim oCollarWrapper As New StructDetailObjects.Collar
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Dim oSlot As Object
    
    Set oCollarWrapper.object = oCollar
    Set oSlot = oCollarWrapper.Slot
    Set oSlotWrapper.object = oSlot
    
    Dim strAnswerToCreationOrder As String
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    GetSelectorAnswer oCollar, "CollarCreationOrder", strAnswerToCreationOrder
    
    Set oCommonHelper = Nothing
    
    Dim strXSectionType As String
    Dim dSectionFlangeLength As Double
    Dim dSectionFlangeOffset As Double
    Dim dSectionWebTh As Double
    Dim strSlotType As String

    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfileWrapper As New StructDetailObjects.ProfilePart
        Set oProfileWrapper.object = oCollarWrapper.Penetrating
        
        strXSectionType = oProfileWrapper.sectionType
        dSectionFlangeLength = oProfileWrapper.FlangeLength
        dSectionFlangeOffset = oProfileWrapper.FlangeOffset
        dSectionWebTh = oProfileWrapper.WebThickness
        
        Set oProfileWrapper = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, strXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        
        oSlotMappingRule.GetSectionWidth oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dSectionFlangeLength
        oSlotMappingRule.GetSectionFlangeOffset oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dSectionFlangeOffset
        oSlotMappingRule.GetSectionWebThickness oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dSectionWebTh
    End If

    strSlotType = oSlotWrapper.ClassName
    
    Dim dClipWebClearance As Double
    Dim dClipWebToSlotEdge As Double
    
    Dim dSlotFlangeClearance As Double
    Dim dSlotBackClearance As Double
    
    Dim dProfileFlangeLength As Double
    Dim eResultType As eResultType
    
    Dim oSDObjectHelper As New StructDetailObjects.Helper
    Dim strTableName As String
    Dim lCodeListValue As Long
    Dim strShortDesc As String
    Dim strLongDesc As String
    
    oSDObjectHelper.GetCustomAttributeParamValue oCollar, _
                                                                              "IJUASmartClip", _
                                                                              "WebClearance", _
                                                                              dClipWebClearance, _
                                                                              eResultType, _
                                                                              strTableName, _
                                                                              lCodeListValue, _
                                                                              strShortDesc, _
                                                                              strLongDesc

    oSDObjectHelper.GetCustomAttributeParamValue oSlot, _
                                                                              "IJUASlot", _
                                                                              "FlangeClearance", _
                                                                              dSlotFlangeClearance, _
                                                                              eResultType, _
                                                                              strTableName, _
                                                                              lCodeListValue, _
                                                                              strShortDesc, _
                                                                              strLongDesc
                                                                              
    Select Case strXSectionType
        Case "B", "EA", "UA", "BUTL3"
            If strAnswerToCreationOrder = "Primary" Then
                dClipWebToSlotEdge = dSectionFlangeLength + dSlotFlangeClearance - dClipWebClearance
            Else
                oSDObjectHelper.GetCustomAttributeParamValue oSlot, _
                                                                                          "IJUASlotA", _
                                                                                          "BackClearance", _
                                                                                          dSlotBackClearance, _
                                                                                          eResultType, _
                                                                                          strTableName, _
                                                                                          lCodeListValue, _
                                                                                          strShortDesc, _
                                                                                          strLongDesc
                dClipWebToSlotEdge = dSlotBackClearance - dClipWebClearance
            End If
            
        Case "BUT", "TSType", "T_XType"
            dClipWebToSlotEdge = (dSectionFlangeLength - dSectionWebTh) / 2 + _
                                               dSlotFlangeClearance - dClipWebClearance
        Case "BUTL2"
            If strAnswerToCreationOrder = "Primary" Then
                dClipWebToSlotEdge = dSectionFlangeLength + dSlotFlangeClearance - _
                                                   dSectionFlangeOffset - dSectionWebTh - dClipWebClearance
            Else
                dClipWebToSlotEdge = dSectionFlangeOffset + dSlotFlangeClearance - dClipWebClearance
            End If
            
        Case "FB"
            If strAnswerToCreationOrder = "Primary" Then
                dClipWebToSlotEdge = dSlotFlangeClearance - dClipWebClearance
            Else
                oSDObjectHelper.GetCustomAttributeParamValue oSlot, _
                                                                                          "IJUASlotA", _
                                                                                          "BackClearance", _
                                                                                          dSlotBackClearance, _
                                                                                          eResultType, _
                                                                                          strTableName, _
                                                                                          lCodeListValue, _
                                                                                          strShortDesc, _
                                                                                          strLongDesc
            
                dClipWebToSlotEdge = dSlotBackClearance - dClipWebClearance
            End If

    End Select
        
    Set oCollar = Nothing
    Set oSlot = Nothing
    Set oCollarWrapper = Nothing
    Set oSlotWrapper = Nothing

    Dim dDiameter As Double
    
    oSDObjectHelper.GetCustomAttributeParamValue oPLH.SmartOccurrence, _
                                                                              "IJUASmartDrainHole", _
                                                                              "Diameter", _
                                                                              dDiameter, _
                                                                              eResultType, _
                                                                              strTableName, _
                                                                              lCodeListValue, _
                                                                              strShortDesc, _
                                                                              strLongDesc
    Set oSDObjectHelper = Nothing

    Dim dOffsetFromWeb As Double
    Dim dOffsetFromBase As Double
    
    If dClipWebToSlotEdge < 2 * dDiameter Then
        dOffsetFromWeb = dClipWebToSlotEdge / 2
    Else
        dOffsetFromWeb = 0.065
    End If
       
    dOffsetFromBase = 0.045
        
    oPLH.Add "OffsetFromWeb", dOffsetFromWeb
    oPLH.Add "OffsetFromBase", dOffsetFromBase
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymbolDefinition As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oSymbolDefinition.IJDInputs.RemoveAllInput
    oSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    
    Dim oDefFactory As New DefinitionFactory
  
    oDefFactory.InitAbstractParameterRule oSymbolDefinition
    Set oDefFactory = Nothing
    
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.definition = oSymbolDefinition
    oIH.InitAs m_FamilyProgid
    ParameterRuleInputs oIH
    Set oIH = Nothing
    
   '----------- Define Output and Add it to the representation ------------------
    Dim oOH As IJDOutputsHelper
    
    Set oOH = New OutputHelper
    oOH.Representation = oSymbolDefinition.IJDRepresentations.Item(1)
    oOH.InitAs m_FamilyProgid
    ParameterRuleOutputs oOH
    Set oOH = Nothing

End Sub
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal oRM As Object) As Object
    Dim oDefFactory As New DefinitionFactory
  
    Set IJDUserSymbolServices_InstanciateDefinition = oDefFactory.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), oRM)
    Set oDefFactory = Nothing

End Function
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal pRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
End Sub
'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(oRep As IJDRepresentation)

    Dim oPLH As IJDParameterLogic
  
    Set oPLH = New ParameterLogic
    oPLH.Representation = oRep
    ParameterRuleLogic oPLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
