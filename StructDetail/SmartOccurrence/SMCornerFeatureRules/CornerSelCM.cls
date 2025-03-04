VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "CornerSelCM"

Private sError As String
Private sMETHOD As String
Const PI = 3.14159265358979
'   pnalugol - Nov 1 2011 -202656 -  Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature


Public Sub SetPlacement(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer(gsPlacement) = "Manual"
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub SetCrackArrest(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    Dim oCF As IJSDOCornerFeature
    Set oCF = New StructDetailObjectsex.CornerFeature
    Set oCF.object = pSL.SmartOccurrence
    
    Dim oPart As Object
    Set oPart = oCF.GetPartObject
    
    If TypeOf oPart Is IJPlate Then
        If IsBracket(oPart) Then
            pSL.Answer(gsCrackArrest) = "No"
        Else
            pSL.Answer(gsCrackArrest) = "Yes"
        End If
    Else
        pSL.Answer(gsCrackArrest) = "Yes"
    End If
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub SetDrainage(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer(gsDrainage) = "Yes"
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub SetFlip(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer(gsCornerFlip) = "NoFlip"
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub SetApplyTreatment(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer(gsApplyTreatment) = "No"
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub SetEnforceSeamAdjustment(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("EnforceSeamAdjustment") = "Yes"
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

