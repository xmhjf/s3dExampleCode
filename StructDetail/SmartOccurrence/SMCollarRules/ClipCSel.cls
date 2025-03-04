VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipCSel"
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

Const m_SelectorProgid As String = CUSTOMERID + "CollarRules.ClipCSel"
Const m_SelectorName As String = CUSTOMERID + "CollarRules.ClipCSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CollarRules\ClipCSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    oIH.SetInput INPUT_PENETRATING
    oIH.SetInput INPUT_BOUNDINGPLATE
    oIH.SetInput INPUT_SLOT
  
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(oQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
  
    oQH.SetQuestion gsAddCornerSnipe, "No", "BooleanCol"
    oQH.SetQuestion gsAddDrainHole, "No", "BooleanCol"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
  
    Dim sERROR As String
    sERROR = ""
  
  'Get Slot object from Collar inputs
  Dim oSlot As New StructDetailObjects.Slot
  Set oSlot.object = oSL.InputObject(INPUT_SLOT)
  
  Dim dProfilePartHeight As Double
  Dim sXSectionType As String
  Dim strCollarOrder As String
  Dim dSlotAngle As Double
  
  strCollarOrder = oSL.SelectorAnswer(CUSTOMERID + "CollarRules.RootClipSel", gsCollarCreationOrder)
  dSlotAngle = GetSlotAngle(oSL.SmartOccurrence)
  
  If TypeOf oSlot.Penetrating Is IJProfile Then
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlot.Penetrating
    dProfilePartHeight = oProfilePart.Height
    sXSectionType = oProfilePart.sectionType
    Set oProfilePart = Nothing
  ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange

    oSlotMappingRule.GetSectionDepth oSlot.Penetrating, oSlot.Penetrated, dProfilePartHeight
    Set oSlotMappingRule = Nothing
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  
  Set oSlot = Nothing
  dProfilePartHeight = Round(dProfilePartHeight, 3)
    
    ' ********************* Selection code - TODO  ****************
    If strCollarOrder = "Primary" Then
        Select Case sXSectionType
            Case "EA", "UA"
               If dProfilePartHeight >= 0.1 Then
                  If dSlotAngle <= 0.00174532925 Then
                     oSL.Add CLIP_AAC_ST
                  ElseIf dSlotAngle > 0.00174532925 And dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                     oSL.Add CLIP_AAC_ST_S
                  Else
                     oSL.Add CLIP_AAC_BT
                  End If
               End If
                
            Case "B"
               If dProfilePartHeight >= 0.18 Then
                  If dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                     oSL.Add CLIP_AAC_ST
                  Else
                     oSL.Add CLIP_AAC_BT
                  End If
               End If
               
            Case "FB"
               If dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                  oSL.Add CLIP_AAC_ST
               Else
                  oSL.Add CLIP_AAC_BT
               End If
               
            Case "BUTL2", "BUT"
                If dSlotAngle <= 0.00174532925 Then
                   oSL.Add CLIP_AAC_ST
                ElseIf dSlotAngle > 0.00174532925 And dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                   oSL.Add CLIP_AAC_ST_S
                Else
                   oSL.Add CLIP_AAC_BT
                End If
                            
            Case Else
               sERROR = sXSectionType & " is not handled in " & m_SelectorProgid
               MsgBox sERROR
               
               GoTo ErrorHandler
        End Select
        
    ElseIf strCollarOrder = "Secondary" Then
    ' Not specified yet
        Select Case sXSectionType
            Case "EA", "UA"
               If dProfilePartHeight >= 0.1 Then
                  If dSlotAngle <= 0.00174532925 Then
                     oSL.Add CLIP_AAC_ST_Left
                  ElseIf dSlotAngle > 0.00174532925 And dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                     oSL.Add CLIP_AAC_ST_S_Left
                  Else
                     oSL.Add CLIP_AAC_BT_Left
                  End If
               End If
                
            Case "B"
               If dProfilePartHeight >= 0.18 Then
                  If dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                     oSL.Add CLIP_AAC_ST_Left
                  Else
                     oSL.Add CLIP_AAC_BT_Left
                  End If
               End If
               
            Case "FB"
               If dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                  oSL.Add CLIP_AAC_ST_Left
               Else
                  oSL.Add CLIP_AAC_BT_Left
               End If
               
            Case "BUTL2", "BUT"
                If dSlotAngle <= 0.00174532925 Then
                   oSL.Add CLIP_AAC_ST_Left
                ElseIf dSlotAngle > 0.00174532925 And dSlotAngle <= SLOT_ANGLE_THRESHOLD_FOR_BENDED_CLIPC Then
                   oSL.Add CLIP_AAC_ST_S_Left
                Else
                   oSL.Add CLIP_AAC_BT_Left
                End If
                            
            Case Else
               sERROR = sXSectionType & " is not handled in " & m_SelectorProgid
               MsgBox sERROR
               
               GoTo ErrorHandler
        End Select
    
    ' *********************************************************
  
  End If
  
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  oSelector.IJDInputs.RemoveAllInput
  oSelector.IJDRepresentations.RemoveAllRepresentation
  
  Dim oDFact As New DefinitionFactory
  Dim oIH As IJDInputsHelper
  
  oDFact.InitAbstractSelector oSelector
  
  Set oIH = New InputHelper
  oIH.definition = oSelector
  oIH.InitAs m_FamilyProgid
  SelectorInputs oIH
  Set oIH = Nothing
  
  Dim oQH As IJDQuestionsHelper
  
  Set oQH = New QuestionHelper
  oQH.Selector = oSelector
  SelectorQuestions oQH
  Set oQH = Nothing
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal oRM As Object) As Object
  Dim oDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = oDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), oRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal pRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(oRep As IJDRepresentation)
  Dim oSL As IJDSelectorLogic
  
  Set oSL = New SelectorLogic
  oSL.Representation = oRep
  SelectorLogic oSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



