VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotRootSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotRootSel"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
'  'Define allowed answers to questions
'  Dim codeCol As New Collection
'  codeCol.Add gsDrop
'  codeCol.Add gsSlide
'  pQH.DefineCodeList "AssyMethodCol", codeCol
    
  'Define questions
    pQH.SetQuestion "AssyMethod", gsDrop, "AssyMethodCol"
    pQH.SetQuestion "ApplyTreatment", "No", "BooleanCol"
    pQH.SetQuestion "SlotType", _
                    "Default", _
                    "SlotTypeCol", _
                    "CMSetAnswerToSlotType"
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  Dim sERROR As String
  
  sERROR = "Picking Slot"
  
  
  'Get Slot object
  Dim oSlot As StructDetailObjects.Slot
  Set oSlot = New Slot
  Set oSlot.object = pSLH.SmartOccurrence

'CHECK HERE TO SEE IF PROFILE OR NOT...
'IF PEN OB A PLATE, GET XSEC FROM SLOTMAPPING
'IF PEN OB A PROFILE, JUST GET XSEC OF OBJECT

    'Get the Penetrating Section
    Dim sXSectionType As String
    If TypeOf pSLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
        'Penetrating Object is a Profile
        Dim oProfilePart As StructDetailObjects.ProfilePart
        Set oProfilePart = New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pSLH.InputObject(INPUT_PENETRATING)
        sXSectionType = oProfilePart.sectionType
    ElseIf TypeOf pSLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
        'Penetrating Object is a Plate
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
        'Unknown Penetrating Object Type
        Exit Sub
    End If
  
  ' Get the Question answers
  Dim sAssyMethod As String
  sAssyMethod = pSLH.Answer("AssyMethod")

   Dim oPlate As IJPlate
   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim ePlateTightness As StructPlateTightness
   
   On Error Resume Next
   Set oPlate = oSlot.Penetrated
   
   On Error GoTo ErrorHandler
   If oPlate Is Nothing Then
      Set oPlatePartWrapper.object = oSlot.BasePlate
   Else
      Set oPlatePartWrapper.object = oPlate
   End If
   ePlateTightness = oPlatePartWrapper.Tightness

  ' ********************* Selection code - TODO  ****************
  ' Currently following Xsections are supported:
  '    Type          Open     Connected
  '                            L  T  R
  ' Angle: EA,UA      x
  '                            x  x  x (L, LT,LR,LTR)
  ' B                 x
  '                            x  x  x (L, LT,LR,LTR)
  ' FB                x
  '                            x  x  x (L, LT,LR,LTR)
  ' BUTL2             x
  '                            x  x  x (L, T, LT,LR,LTR)
  ' BUT               x
  '                            x  x  x (L, T, LT,LR,LTR)
  
  If ePlateTightness = NonTight Or _
     ePlateTightness = UnSpecifiedTightness Then
     pSLH.Add "SlotC"
     pSLH.Add "SlotA"
     
    If sAssyMethod = gsSlide Then
        'also add SlotI
        pSLH.Add "SlotI"
    End If
     
  Else
     If sXSectionType = "BUT" Or _
        sXSectionType = "BUTL2" Then
        pSLH.Add "SlotC2"
     End If
     
     pSLH.Add "SlotC"
     pSLH.Add "SlotA"
     
     If oSlot.IsPerpendicular = True And sAssyMethod = gsSlide Then
        pSLH.Add "SlotI"
     End If
  End If
    
'  Select Case oProfilePart.SectionType
'     Case "FB", "EA", "UA", "BUTL3"
'     Case "T_XType", "TSType", "BUT", "BUTL2"
'     Case "B"
'     Case "C_SS", "CSType", "I", "ISType"
'
''     Case "HalfR", "P", "R"
''        sERROR = "No Slot defined for this profile type ('" & oProfilePart.SectionType & "'"
''        GoTo ErrorHandler
'     Case Else
'        sERROR = "Invalid cross section type ('" & oProfilePart.SectionType & "') specified for " & m_sRuleProgID
'        GoTo ErrorHandler
'  End Select
  
  ' *********************************************************
  
  Set oProfilePart = Nothing
  Set oSlot = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic", sERROR).Number
End Sub



' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pSelector.IJDInputs.RemoveAllInput
  pSelector.IJDRepresentations.RemoveAllRepresentation
  
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_sFamilyProgID
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_sRuleProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSLH As IJDSelectorLogic
  Set pSLH = New SelectorLogic
  pSLH.Representation = pRep
  SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Public Sub CMSetAnswerToSlotType( _
               ByVal oInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
               ByRef oArgument As Object)
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymDef As IMSSymbolEntities.IJDSymbolDefinition
    Dim oSL As IJDSelectorLogic
    Dim sAnswer As String
    
    Set oInputDG = oInput
    Set oSymDef = oInputDG.definition
    Set oInputDG = Nothing
    
    Set oSL = New SelectorLogic
    oSL.Representation = oSymDef.IJDRepresentations(1)
    Set oSymDef = Nothing
    
    ' Modify the answer as needed
    sAnswer = "Default"
    oSL.Answer("SlotType") = sAnswer
    Set oSL = Nothing
    
End Sub

