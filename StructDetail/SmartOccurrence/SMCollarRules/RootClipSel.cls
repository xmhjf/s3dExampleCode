VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootClipSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of avaible Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parametrize the selection
'   Removed height restrictions which restricts selection of smart item
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "CollarRules.RootClipSel"
Const m_SelectorName As String = CUSTOMERID + "CollarRules.RootClipSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CollarRules\RootClipSel.cls"

Private Const m_sQuestion_UseTightCollar As String = "UseTightCollar"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_BOUNDINGPLATE
  pIH.SetInput INPUT_SLOT
      
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
'  Dim codeStressLevelCol As New Collection
'  codeStressLevelCol.Add gsHigh
'  codeStressLevelCol.Add gsMedium
'  codeStressLevelCol.Add gsLow
'
'  pQH.DefineCodeList "StressLevelCol", codeStressLevelCol
  
'  Dim codeCollarSideCol As New Collection
'  codeCollarSideCol.Add gsNoFlip
'  codeCollarSideCol.Add gsFlip
'  codeCollarSideCol.Add gsCentered
'
'  pQH.DefineCodeList "CollarSideCol", codeCollarSideCol
  
  'Define questions
  pQH.SetQuestion gsStressLevel, gsHigh, "StressLevelCol" ' NOTE: this default is only used if not created by Assembly connection
   
  pQH.SetQuestion gsCollarCreationOrder, "Primary" ' NOTE: This question should NEVER be changed by user (property page)
  
  pQH.SetQuestion gsCollarSideOfPart, gsNoFlip, "CollarSideCol"
  
  pQH.SetQuestion m_sQuestion_UseTightCollar, "Yes", "BooleanCol", "SetAnswerToUseTightCollar", m_SelectorProgid
  
  Exit Sub

ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  Dim strError As String
  
  strError = "Selecting Correct Collar"
  'Get Slot object from Collar inputs
  Dim oSlot As New StructDetailObjects.Slot
  Set oSlot.object = pSL.InputObject(INPUT_SLOT)
  
  'Get the Cross Section Type and Section Depth
  Dim sXSectionType As String
  Dim dProfilePartHeight As Double
  If TypeOf oSlot.Penetrating Is IJProfile Then
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlot.Penetrating
    
    sXSectionType = oProfilePart.sectionType
    dProfilePartHeight = oProfilePart.Height
  ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange

    oSlotMappingRule.GetSectionDepth oSlot.Penetrating, oSlot.Penetrated, dProfilePartHeight
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  
  dProfilePartHeight = Round(dProfilePartHeight, 3)

  'Get answers
  Dim strStressLevel As String
  Dim strCollarOrder As String
  Dim sSlotClassName As String
  
  strStressLevel = pSL.Answer(gsStressLevel)
  strCollarOrder = pSL.Answer(gsCollarCreationOrder)
  sSlotClassName = oSlot.ClassName
  
  Dim sUseTightCollar As String
  Dim bUseTightCollar As Boolean
  
  bUseTightCollar = True
  sUseTightCollar = pSL.Answer(m_sQuestion_UseTightCollar)
  
  If sUseTightCollar = "Yes" Then
     bUseTightCollar = True
  Else
     bUseTightCollar = False
  End If
  
  Dim bIsClipCNeeded As Boolean
  Dim bIsClipANeeded As Boolean
  
  bIsClipCNeeded = False
  bIsClipANeeded = False
  
  ' ********************* Selection code - TODO  ****************
  
  If (strCollarOrder = "Primary") Then
     If bUseTightCollar = True Then
        ' Use tight collar
        Select Case sSlotClassName
           Case "SlotI"
              pSL.Add "CollarNoSnipe"
           Case "SlotC", SLOTC_TWO_COLLARS
              pSL.Add "CollarOneSnipe"
           Case "SlotA"
              pSL.Add "CollarTwoSnipes"
        End Select
        
     Else
        ' Use non-tight collar
        Dim dAngle As Double
        Dim bProfileOnHullPenetratingLBH As Boolean
        Dim bIsSpecialCase As Boolean
         
        bIsSpecialCase = False
        bProfileOnHullPenetratingLBH = CheckProfileOnHullPenetratingLBH(oSlot.object)
        If bProfileOnHullPenetratingLBH = True Then
           GetSlotAngles oSlot.object, JXSEC_WEB_RIGHT, dAngle
           If dAngle > HALF_PI Then
              bIsSpecialCase = True
           End If
        End If
         
        If bIsSpecialCase Then
           Dim dClipRightLength As Double
           Dim nOption As Integer
            
           dClipRightLength = EstimateClipRightSideLength(oSlot.object, dAngle)
           If dClipRightLength >= 0.09 Then
              nOption = 1 ' Normal
           ElseIf dClipRightLength > 0.065 And dClipRightLength < 0.09 Then
              nOption = 2 ' Clip A with Bottom snipe
           ElseIf dClipRightLength > 0.01 And dClipRightLength < 0.065 Then
              nOption = 3 ' Clip C, Minium 10mm
           Else
              nOption = 0
           End If
            
           Select Case sSlotClassName
              Case "SlotC", SLOTC_TWO_COLLARS
                 If strStressLevel = gsHigh Then
                    If nOption = 1 Or nOption = 2 Then
                       pSL.Add "ClipC"
                       pSL.Add "ClipA"
                    ElseIf nOption = 3 Then
                       pSL.Add "ClipC"
                    End If
                 End If
                  
              Case "SlotA"
                 If strStressLevel = gsMedium Or _
                    strStressLevel = gsHigh Then
                    If nOption = 1 Or nOption = 3 Then
                       pSL.Add "ClipC"
                       pSL.Add "ClipA"
                    ElseIf nOption = 2 Then
                       pSL.Add "ClipA"
                       pSL.Add "ClipC"
                    End If
                 End If
           End Select
            
        Else
           Select Case sSlotClassName
              Case "SlotC", SLOTC_TWO_COLLARS
                 If strStressLevel = gsHigh Then
                    Select Case sXSectionType
                       Case "EA", "UA"
                          If dProfilePartHeight >= 0.1 Then
                             bIsClipCNeeded = True
                          End If
                          If dProfilePartHeight >= 0.2 Then
                             bIsClipANeeded = True
                          End If
                           
                       Case "B"
                          If dProfilePartHeight >= 0.18 Then
                             bIsClipANeeded = True
                             bIsClipCNeeded = True
                          End If
                        
                       Case "FB"
                          bIsClipCNeeded = True
                          If dProfilePartHeight >= 0.2 Then
                             bIsClipANeeded = True
                          End If
                        
                       Case "BUTL2", "BUT"
                         
                             bIsClipCNeeded = True
                          
                          If dProfilePartHeight >= 0.2 Then
                             bIsClipANeeded = True
                          End If
                        
                       Case Else
                        
                    End Select
                     
                    If bIsClipCNeeded Then
                       pSL.Add "ClipC"
                    End If
                    If bIsClipANeeded Then
                       pSL.Add "ClipA"
                    End If
                 End If
              Case "SlotA"
                 If strStressLevel = gsMedium Or _
                    strStressLevel = gsHigh Then
                    pSL.Add "ClipC"
                    pSL.Add "ClipA"
                 End If
           End Select
        End If
     End If
     
  ElseIf (strCollarOrder = "Secondary") Then ' Secondary
     If bUseTightCollar = True Then
        ' Add tight collar
        If oSlot.ClassName = SLOTC_TWO_COLLARS Then
           ' Second collar is needed
           pSL.Add "CollarOneSnipe"
        End If
     Else
        ' Add non tight collar
        
        bIsSpecialCase = False
        bProfileOnHullPenetratingLBH = CheckProfileOnHullPenetratingLBH(oSlot.object)
        If bProfileOnHullPenetratingLBH = True Then
           GetSlotAngles oSlot.object, JXSEC_WEB_RIGHT, dAngle
           If dAngle > HALF_PI Then
              bIsSpecialCase = True
           End If
        End If
         
        If bIsSpecialCase Then
           
           dClipRightLength = EstimateClipRightSideLength(oSlot.object, dAngle)
           If dClipRightLength >= 0.09 Then
              nOption = 1 ' Normal
           ElseIf dClipRightLength > 0.065 And dClipRightLength < 0.09 Then
              nOption = 2 ' Clip A with Bottom snipe
           ElseIf dClipRightLength > 0.01 And dClipRightLength < 0.065 Then
              nOption = 3 ' Clip C, Minium 10mm
           Else
              nOption = 0
           End If
            
           Select Case sSlotClassName
              Case "SlotA"
                 If strStressLevel = gsMedium Or _
                    strStressLevel = gsHigh Then
                    If nOption = 1 Or nOption = 3 Then
                       pSL.Add "ClipC"
                       pSL.Add "ClipB"
                    ElseIf nOption = 2 Then
                       pSL.Add "ClipB"
                       pSL.Add "ClipC"
                    End If
                 End If
              Case SLOTC_TWO_COLLARS
                pSL.Add "ClipC"
           End Select
        Else
           Select Case sSlotClassName
              Case "SlotA"
                 If strStressLevel = gsMedium Or _
                    strStressLevel = gsHigh Then
                    pSL.Add "ClipC"
                    pSL.Add "ClipB"
                 End If
              Case SLOTC_TWO_COLLARS
                pSL.Add "ClipC"
           End Select
        End If
     End If
  Else
        strError = "Invalid Collar Creation Order: " & strCollarOrder
        GoTo ErrorHandler
  End If
   
  ' *********************************************************
  
  Set oSlot = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
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
  
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
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
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSL As IJDSelectorLogic
  Set pSL = New SelectorLogic
  pSL.Representation = pRep
  SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Public Sub SetAnswerToUseTightCollar( _
               ByVal oInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
               ByRef oArgument As Object)
    
   Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
   Dim oSymDef As IMSSymbolEntities.IJDSymbolDefinition
   Dim oSL As IJDSelectorLogic
    
   Set oInputDG = oInput
   Set oSymDef = oInputDG.definition
   Set oInputDG = Nothing
   Set oSL = New SelectorLogic
   oSL.Representation = oSymDef.IJDRepresentations(1)
   Set oSymDef = Nothing

   Dim oSlot As New StructDetailObjects.Slot
   
   Set oSlot.object = oSL.InputObject(INPUT_SLOT)
  
   On Error Resume Next
   Dim oPlate As IJPlate
   Dim oPlatePart As New StructDetailObjects.PlatePart
   
   Set oPlate = oSlot.Penetrated
   If oPlate Is Nothing Then
      Set oPlatePart.object = oSlot.BasePlate
   Else
      Set oPlatePart.object = oSlot.Penetrated
      Set oPlate = Nothing
   End If
   
   Dim sSlotType As String
   Dim oCommonHelper As New CommonHelper
   Dim oSO As IJSmartOccurrence
   Dim oSC As IJSmartClass
   
   Set oSO = oSlot.object
   Set oSC = oSO.RootSelectionObject
   sSlotType = oCommonHelper.GetAnswer( _
                             oSO, _
                             oSC.SelectionRuleDef, _
                             "SlotType")
   Set oSO = Nothing
   Set oSC = Nothing
   Set oCommonHelper = Nothing
   
   If LCase(sSlotType) = LCase("Default") Then
      ' Defalut selection is based on plate tightness
      ' For tight plate,place tight collar,for non-tight plate, place non-tight collar
      ' Customer has the final control whether tight or non-tight collar is placed
      If oPlatePart.Tightness = AirTight Or _
         oPlatePart.Tightness = OilTight Or _
         oPlatePart.Tightness = WaterTight Then
         oSL.Answer(m_sQuestion_UseTightCollar) = "Yes"
      Else
         oSL.Answer(m_sQuestion_UseTightCollar) = "No"
      End If
      
   ElseIf LCase(sSlotType) = LCase("Tight") Then
      oSL.Answer(m_sQuestion_UseTightCollar) = "Yes"
   Else
      oSL.Answer(m_sQuestion_UseTightCollar) = "No"
   End If
   
   Set oPlatePart = Nothing
   Set oSL = Nothing
   
   Exit Sub
End Sub

