VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'pnalugol - Nov1 2011 - Changes done to support 202656, added
'       new selector questions to create Corner features as child of Slots

Option Explicit
'Removed height restrictions which restricts selection of smart item

Const m_SelectorProgid As String = CUSTOMERID + "SlotRules.SlotCSel"
Const m_SelectorName As String = CUSTOMERID + "SlotRules.SlotCSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "SlotRules\SlotCSel.cls"

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
  pQH.SetQuestion "OutsideCorners", "No", "BooleanCol", "SetAnswerToOutsideCorners", m_SelectorProgid
   
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
    Dim sERROR As String

    'Get the Slot
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = pSLH.SmartOccurrence
  
    'Get the Penetrating Section and Height
    Dim sXSectionType As String
    Dim dProfilePartHeight As Double
    If TypeOf pSLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
        'Penetrating Object is a Profile
        Dim oProfilePart As StructDetailObjects.ProfilePart
        Set oProfilePart = New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pSLH.InputObject(INPUT_PENETRATING)
        dProfilePartHeight = oProfilePart.Height
        dProfilePartHeight = Round(dProfilePartHeight, 3)
        sXSectionType = oProfilePart.sectionType
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
    Else
        'Unknown Penetrating Object Type
        Exit Sub
    End If

    ' Get the Question answers
   Dim clearance As Double
   
   clearance = pSLH.Answer("Clearance")


   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim ePlateTightness As StructPlateTightness
   
   If TypeOf oSlotWrapper.Penetrated Is IJPlatePart Then
      Set oPlatePartWrapper.object = oSlotWrapper.Penetrated
   Else
      Set oPlatePartWrapper.object = oSlotWrapper.BasePlate
   End If
   
   ePlateTightness = oPlatePartWrapper.Tightness
   
   Dim bNonTightSlot As Boolean
   Dim bTightSlot As Boolean
   Dim sSlotType As String
   Dim oCommonHelper As New CommonHelper
   Dim oSO As IJSmartOccurrence
   Dim oSC As IJSmartClass
   
   bNonTightSlot = False
   bTightSlot = False

   Set oSO = pSLH.SmartOccurrence
   Set oSC = oSO.RootSelectionObject
   
   sSlotType = oCommonHelper.GetAnswer( _
                                    oSO, _
                                    oSC.SelectionRuleDef, _
                                    "SlotType")
                                       
   If LCase(sSlotType) = LCase("Default") Then
      If ePlateTightness = NonTight Or _
         ePlateTightness = UnSpecifiedTightness Then
         bNonTightSlot = True
         bTightSlot = False
      Else
         bNonTightSlot = False
         bTightSlot = True
      End If
      
   ElseIf LCase(sSlotType) = LCase("NonTight") Then
      bNonTightSlot = True
      bTightSlot = False
      
   ElseIf LCase(sSlotType) = LCase("Tight") Then
      bNonTightSlot = False
      bTightSlot = True
      
   End If
   
   Dim sRuleChoice As String ' This variable controlls which rule will be used
   
   sRuleChoice = "Other"
  ' ********************* Selection code ****************

  Select Case sXSectionType
     Case "EA", "UA"
       
           If bNonTightSlot = True Then
              ' Left connected
              pSLH.Add "SlotAC_L_PAT_DTR"
              ' Left & Top Connected
              pSLH.Add "SlotAC_LT_PAA"
              pSLH.Add "SlotAC_LT_PAT"
           End If
              
           If bTightSlot = True Then
              ' Left Top connected
              pSLH.Add "SlotAC_LT_PAT"
              pSLH.Add "SlotAC_LT_PAT2"
           
              ' Left connected
              pSLH.Add "SlotAC_L_PAT_STR"
           End If
       

     Case "BUTL2"
        If dProfilePartHeight > 0.1 Then
           If bNonTightSlot = True Then
              ' Left connected
              pSLH.Add "SlotL2C_L_PAT_DTR"      'Default selection
           End If
              
           If bTightSlot = True Then
              ' Left connected
              pSLH.Add "SlotL2C_L_PAT_DTR2"
           End If
        End If

     Case "B"
        If sRuleChoice = "Default" Then
           If bNonTightSlot = True Then
              If dProfilePartHeight >= 0.18 Then
                 ' Left connected
                 pSLH.Add "SlotBC_L_PAT_DTR"
              End If
           End If
              
           If bTightSlot = True Then
              ' Left Top Connected
              pSLH.Add "SlotBC_LT_PAT"
              pSLH.Add "SlotBC_LT_PAT2"
            
              ' Left Connected
              pSLH.Add "SlotBC_L_LTT_STR"
              pSLH.Add "SlotBC_L_PAT_STR"
           End If
        Else
           If bNonTightSlot = True Then
              If dProfilePartHeight > 0.12 Then
                 pSLH.Add "SlotBC_L_PAT_DTR"
              Else
                 pSLH.Add "SlotBC_L_PAT_STR"
              End If
           End If
              
           If bTightSlot = True Then
              ' Left Connected
              pSLH.Add "SlotBC_L_LTT_STR"
              pSLH.Add "SlotBC_L_PAT_STR"
              
              ' Left Top Connected
              pSLH.Add "SlotBC_LT_PAT"
              pSLH.Add "SlotBC_LT_PAT2"
           End If
        End If
        
     Case "BUT"
        If dProfilePartHeight > 0.1 Then
           If bNonTightSlot = True Then
              ' Left Connected
              pSLH.Add "SlotTC_L_PAT_DTR"
              'Top Connected
              pSLH.Add "SlotTC_T_PAA_STR"
           End If
              
           If bTightSlot = True Then
              ' Left Connected
              pSLH.Add "SlotTC_L_PAT_STR"
           End If
        End If
                
     Case "FB"
        If bNonTightSlot = True Then
           ' Left Connected
           pSLH.Add "SlotFC_L_PAT_DTR"
        End If
              
        If bTightSlot = True Then
           ' Left Top connected
           pSLH.Add "SlotFC_LT_PAT"
           pSLH.Add "SlotFC_LT_PAT2"
           
           ' Left Connected
           pSLH.Add "SlotFC_L_PAT_STR"
        End If

     Case Else
        'sERROR = "Invalid cross section type ('" & oProfilePart.sectionType & "') specified for " & m_SelectorProgid
        sERROR = "Invalid cross section type ('" & sXSectionType & "') specified for " & m_SelectorProgid
        GoTo ErrorHandler
  End Select
  
  ' *********************************************************
  
  Set oProfilePart = Nothing
 
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

