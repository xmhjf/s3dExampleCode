VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndCutSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const sMODULE As String = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\EndCutSelCM.cls"

Public Sub UpdateEndCondition(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                                ByRef ppArgument As Object)
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Set oInputDG = pInput
    
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oInputDG.Definition
    
    Dim oEndCondition As IMSSymbolEntities.IJDInput
    Set oEndCondition = oSymbolDefinition.IJDInputs.Item(CL_ENDCONDITION)
    
    'Get Penetrating object
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = oSymbolDefinition.IJDInputs
    
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oInputs.Item(INPUT_BOUNDED).IJDInputDuringGame.Argument
    
    'Parameter content of input
    Dim pPC As IJDParameterContent
    Set pPC = New DParameterContent
    pPC.Type = igValue

'  select on AC fixity
'  Select Case BoundedProfileSystem.Fixity
'     Case "gsFixed"
'        pPC.UomValue = "gsW"
''
'     Case "gsFree"
'        pPC.UomValue = "gsS", "gsF", "gsFV"
'     Case "gsFlangeFree"
'        pPC.UomValue = "gsC",
'        If oBounded.SectionType = "gsFB" Then
'            add "gsR", "gsRV"
'  End Select

' if no fixity from parent, default end condition to "C"

    oEndCondition.DefaultParameterValue = pPC
    
    'Set return argument for input
    oInputDG.Argument = pPC
    Set ppArgument = oInputDG.Argument
    
    'Set new answer on IAnswer interface of Smart Occurrence
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    
    pSLH.Representation = oSymbolDefinition.IJDRepresentations.Item(1)
    pSLH.Answer(CL_ENDCONDITION) = pPC.UomValue
    
    'Release object variables
    Set oProfilePart = Nothing
    Set pPC = Nothing
    Set oEndCondition = Nothing
    Set oInputs = Nothing
    Set oSymbolDefinition = Nothing
    Set oInputDG = Nothing
    Set pSLH = Nothing
End Sub
