VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketSelCM"


Private sError As String
Private sMETHOD As String

Public Sub CM_SetBracketContourTypeSup1(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
     If TypeOf pSL.SmartOccurrence Is IJBracketPlateSystem Then
        Dim oBracketPlate As IJBracketPlateSystem
        Set oBracketPlate = pSL.SmartOccurrence
        Dim eConnectType As ShpStrBktSupportConnectionType
        eConnectType = oBracketPlate.GetSupportsConnectionType(1)
        If eConnectType = ConnType_None Then
            pSL.Answer(QUESTION_BracketSupport1Type) = "Trimmed"
             oBracketPlate.SetSupportsConnectionType 1, Trimmed
        Else
             pSL.Answer(QUESTION_BracketSupport1Type) = IIf(eConnectType = Trimmed, "Trimmed", IIf(eConnectType = Lapped, "Lapped", "Reference"))
        End If
    End If

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


Public Sub CM_SetBracketContourTypeSup2(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    If TypeOf pSL.SmartOccurrence Is IJBracketPlateSystem Then
        Dim oBracketPlate As IJBracketPlateSystem
        Set oBracketPlate = pSL.SmartOccurrence
        Dim eConnectType As ShpStrBktSupportConnectionType
        eConnectType = oBracketPlate.GetSupportsConnectionType(2)
        If eConnectType = ConnType_None Then
            pSL.Answer(QUESTION_BracketSupport2Type) = "Trimmed"
            oBracketPlate.SetSupportsConnectionType 2, Trimmed
        Else
             pSL.Answer(QUESTION_BracketSupport2Type) = IIf(eConnectType = Trimmed, "Trimmed", IIf(eConnectType = Lapped, "Lapped", "Reference"))
        End If
    End If
    
    Exit Sub

    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


Public Sub CM_SetBracketContourTypeSup3(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    If TypeOf pSL.SmartOccurrence Is IJBracketPlateSystem Then
        Dim oBracketPlate As IJBracketPlateSystem
        Set oBracketPlate = pSL.SmartOccurrence
        Dim eConnectType As ShpStrBktSupportConnectionType
        If oBracketPlate.NumberOfSupports < 3 Then
              oBracketPlate.SetSupportsConnectionType 3, ConnType_None
              Exit Sub
        End If
        eConnectType = oBracketPlate.GetSupportsConnectionType(3)
        If eConnectType = ConnType_None Then
            pSL.Answer(QUESTION_BracketSupport3Type) = "Trimmed"
            oBracketPlate.SetSupportsConnectionType 3, Trimmed
        Else
             pSL.Answer(QUESTION_BracketSupport3Type) = IIf(eConnectType = Trimmed, "Trimmed", IIf(eConnectType = Lapped, "Lapped", "Reference"))
        End If
    End If

    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


Public Sub CM_SetBracketContourTypeSup4(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
   If TypeOf pSL.SmartOccurrence Is IJBracketPlateSystem Then
        Dim oBracketPlate As IJBracketPlateSystem
        Set oBracketPlate = pSL.SmartOccurrence
        Dim eConnectType As ShpStrBktSupportConnectionType
        
        If oBracketPlate.NumberOfSupports < 4 Then
              oBracketPlate.SetSupportsConnectionType 4, ConnType_None
              Exit Sub
        End If
        eConnectType = oBracketPlate.GetSupportsConnectionType(4)
        If eConnectType = ConnType_None Then
            pSL.Answer(QUESTION_BracketSupport4Type) = "Trimmed"
            oBracketPlate.SetSupportsConnectionType 4, Trimmed
        Else
             pSL.Answer(QUESTION_BracketSupport4Type) = IIf(eConnectType = Trimmed, "Trimmed", IIf(eConnectType = Lapped, "Lapped", "Reference"))
        End If
    End If

    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


Public Sub CM_SetBracketContourTypeSup5(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
   If TypeOf pSL.SmartOccurrence Is IJBracketPlateSystem Then
        Dim oBracketPlate As IJBracketPlateSystem
        Set oBracketPlate = pSL.SmartOccurrence
         If oBracketPlate.NumberOfSupports < 5 Then
              oBracketPlate.SetSupportsConnectionType 5, ConnType_None
              Exit Sub
        End If
        Dim eConnectType As ShpStrBktSupportConnectionType
        eConnectType = oBracketPlate.GetSupportsConnectionType(5)
        If eConnectType = ConnType_None Then
            pSL.Answer(QUESTION_BracketSupport5Type) = "Trimmed"
            oBracketPlate.SetSupportsConnectionType 5, Trimmed
        Else
             pSL.Answer(QUESTION_BracketSupport5Type) = IIf(eConnectType = Trimmed, "Trimmed", IIf(eConnectType = Lapped, "Lapped", "Reference"))
        End If
    End If
    

    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub






