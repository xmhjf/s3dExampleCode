VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const sMODULE As String = m_sProjectPath + "SlotSelCM.cls"

Public Sub GenerateClearance(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    
    Set pSL = New SelectorLogic
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)

    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = pSL.SmartOccurrence

    'Get the Cross Section Type
    Dim sSectionType As String
    If TypeOf oSlot.Penetrating Is IJProfile Then
        'Penetrating Object is a Profile
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlot.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
    ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
        'Penetrating Object is a Plate
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    Dim oClearanceInput As IMSSymbolEntities.IJDInput
    Set oClearanceInput = oSymbolDefinition.IJDInputs.Item("Clearance")
  
    'Parameter content of input
    Dim pPC As IJDParameterContent
    Set pPC = New DParameterContent
    pPC.Type = igValue
    
    Select Case sSectionType
        Case "EA", "UA"
            pPC.UomValue = 0.01
        Case "F", "SB", "ST", "FB"
            pPC.UomValue = 0.02
        Case "T_XType", "TSType", "BUT", "BUTL2"
            pPC.UomValue = 0.03
        Case Else
            pPC.UomValue = 0.99
    End Select
    oClearanceInput.DefaultParameterValue = pPC
    
    'Set return argument for input
    oInputDG.Argument = pPC
    Set ppArgument = oInputDG.Argument
    
    'Release object variables
    Set pPC = Nothing
    Set oClearanceInput = Nothing
    Set oSymbolDefinition = Nothing
    Set oInputDG = Nothing
    Set oSlot = Nothing

End Sub

