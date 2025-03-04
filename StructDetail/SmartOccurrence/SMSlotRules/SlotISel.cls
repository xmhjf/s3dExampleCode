VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotISel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'pnalugol - Nov1 2011 - Changes done to support 202656, added
'       new selector questions to create Corner features as child of Slots
'Removed height restrictions which restricts selection of smart item
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "SlotRules.SlotISel"
Const m_SelectorName As String = CUSTOMERID + "SlotRules.SlotISel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "SlotRules\SlotISel.cls"

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
 
 'Define questions
  pQH.SetQuestion "Clearance", 0.1, , "GenerateClearance", LIBRARY_SOURCE_ID
  
  pQH.SetQuestion "BaseCorners", "No", "BooleanCol", "SetAnswerToBaseCorners", m_SelectorProgid
  pQH.SetQuestion "InsideCorners", "No", "BooleanCol", "SetAnswerToInsideCorners", m_SelectorProgid
  pQH.SetQuestion "OutsideCorners", "No", "BooleanCol", "SetAnswerToOutsideCorners", m_SelectorProgid
   
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  
  ' ************* Collect the Arguments ************
  
   Dim oSlotWrapper As New StructDetailObjects.Slot
   Set oSlotWrapper.object = pSLH.SmartOccurrence
  
  'Get the Penetrating Section
    Dim sXSectionType As String
    Dim dProfilePartHeight As Double
    Dim bIsPenetratingObjPlate As Boolean
    If TypeOf pSLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
        'Penetrating Object is a Profile

        Dim oProfilePart As StructDetailObjects.ProfilePart
        Set oProfilePart = New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pSLH.InputObject(INPUT_PENETRATING)
        sXSectionType = oProfilePart.sectionType
        dProfilePartHeight = oProfilePart.Height
        dProfilePartHeight = Round(dProfilePartHeight, 3)
        bIsPenetratingObjPlate = False
        Set oProfilePart = Nothing
    ElseIf TypeOf pSLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
        'Penetrating Object is a Plate
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    
        oSlotMappingRule.GetSectionDepth oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dProfilePartHeight
        dProfilePartHeight = Round(dProfilePartHeight, 3)
        bIsPenetratingObjPlate = True
    Else
        'Unknown Penetrating Object Type
        Exit Sub
    End If

    ' Get the Question answers
   Dim clearance As Double
   clearance = pSLH.Answer("Clearance")
   
  
   '***********************************************************
   'DOES NOT APPEAR TO SERVE ANY PURPOSE
   'MARK BLOCK FOR REMOVAL? CLR
   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim ePlateTightness As StructPlateTightness
   
   If TypeOf oSlotWrapper.Penetrated Is IJPlatePart Then
      Set oPlatePartWrapper.object = oSlotWrapper.Penetrated
   Else
      Set oPlatePartWrapper.object = oSlotWrapper.BasePlate
   End If
   
   ePlateTightness = oPlatePartWrapper.Tightness
   '**********************************************************

  ' ********************* Selection code  ****************
  
  'Select Case oProfilePart.sectionType
  Select Case sXSectionType
     Case "EA", "UA"
        ' Left Right Connected
        If dProfilePartHeight > 0.125 Then
            pSLH.Add "SlotAI_LR_PAAP_STR"
        End If
        
        ' If a Plate/Stiffener Combo is used to form a EA or UA section alias
        ' then there is no fillet radius on the top flange right top corner.
        ' SlotAI_LTR_PAAP_STR expects a radius at the corner.  Therefore, if the
        ' penetrated part is a plate then we cannot add SlotAI_LTR_PAAP_STR.
        If Not bIsPenetratingObjPlate Then
            ' Left Top Right Connected
            pSLH.Add "SlotAI_LTR_PAAP_STR"
        Else
            ' A new symbol will have to be created that does not include a radius.
            ' For now just add the LR Connected if it wasn't already added so there
            ' is something
            
                pSLH.Add "SlotAI_LR_PAAP_STR"
            
        End If
        
     Case "B"
        ' Left Right Connected
        If dProfilePartHeight >= 0.18 Then
            pSLH.Add "SlotBI_LR_PAAP_STR"
        End If
        
        ' Left Top Right Connected
        pSLH.Add "SlotBI_LTR_PAAP_STR"
        
     Case "FB"
        ' Left Right Connected
        pSLH.Add "SlotFI_LR_PAAP_STR"
        
        ' Left Top Right Connected
        pSLH.Add "SlotFI_LTR_PAAP_STR"
        
     Case "BUTL2"
        ' Left Right Connected
        If dProfilePartHeight > 0.15 Then
           pSLH.Add "SlotL2I_LR_PAAP_STR"
           
        End If
        
        ' Left Top Right Connected
        pSLH.Add "SlotL2I_LTR_PAAP_STR"
        
     Case "BUT"
        ' Left Right Connected
           pSLH.Add "SlotTI_LR_PAAP_STR"
        ' Left Top Right Connected
        pSLH.Add "SlotTI_LTR_PAAP_STR"

     'Case Else
     '  sERROR = "Invalid cross section type ('" & oProfilePart.SectionType & "') specified for " & m_SelectorProgid
     '   GoTo ErrorHandler
  End Select
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub
  


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
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
  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
Public Sub SetAnswerToBaseCorners( _
               ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
               ByRef ppArgument As Object)
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
    
    pSL.Answer("BaseCorners") = "No"
    
    Exit Sub


ErrorHandler:

End Sub

Public Sub SetAnswerToOutsideCorners( _
               ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
               ByRef ppArgument As Object)
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
    
    pSL.Answer("OutsideCorners") = "No"
    
    Exit Sub

ErrorHandler:

End Sub
Public Sub SetAnswerToInsideCorners( _
               ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
               ByRef ppArgument As Object)
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
    
    pSL.Answer("InsideCorners") = "No"
    
    Exit Sub

ErrorHandler:

End Sub

