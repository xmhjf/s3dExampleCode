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

Const sMODULE As String = "S:\StructDetail\Data\SmartOccurrence\SlotRules\SlotSelCM.cls"

Public Sub GenerateClearance(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    
  Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
  Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
  Dim oInputs As IMSSymbolEntities.IJDInputs
  Dim oClearanceInput As IMSSymbolEntities.IJDInput
  
  Set oInputDG = pInput
  Set oSymbolDefinition = oInputDG.Definition
  Set oInputs = oSymbolDefinition.IJDInputs
  Set oClearanceInput = oSymbolDefinition.IJDInputs.Item("Clearance")
  
  'Get Penetrating object
  Dim oProfilePart As StructDetailObjects.ProfilePart
  Set oProfilePart = New StructDetailObjects.ProfilePart
  'GetInputAtIndex(1) is the penetrating graphic input defined on the symbol
  Set oProfilePart.object = oInputs.GetInputAtIndex(1).IJDInputDuringGame.Argument
  
  'Parameter content of input
  Dim pPC As IJDParameterContent
  Set pPC = New DParameterContent
  pPC.Type = igValue

  Select Case oProfilePart.SectionType
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
  Set oProfilePart = Nothing
  Set pPC = Nothing
  Set oClearanceInput = Nothing
  Set oInputs = Nothing
  Set oSymbolDefinition = Nothing
  Set oInputDG = Nothing

End Sub
