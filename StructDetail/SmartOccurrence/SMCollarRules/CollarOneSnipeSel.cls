VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOneSnipeSel"
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
'
Option Explicit

' Modify class name to particular rule
Private Const m_sClassName As String = "CollarOneSnipeSel"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

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
  
    oQH.SetQuestion gsAddCornerSnipe, "No", "BooleanCol", _
                    "SetAnswerToAddCornerSnipe", _
                    m_sRuleProgID
  
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
  
  Dim sXSectionType As String
  Dim sSlotItemName As String
  Dim sSlotClassName As String
  sSlotClassName = oSlot.ClassName
  sSlotItemName = oSlot.ItemName
  
  Dim bIsPenetratingObjPlate As Boolean
  If TypeOf oSlot.Penetrating Is IJProfile Then
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlot.Penetrating
    sXSectionType = oProfilePart.sectionType
    Set oProfilePart = Nothing
    bIsPenetratingObjPlate = False
  ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Set oSlotMappingRule = Nothing
    bIsPenetratingObjPlate = True
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  
  Set oSlot = Nothing
    
    ' Tsuneishi has 3 types of one snipe collars:
    ' 1.Regular collar,connected to Web Left,Top,Web Right
    ' 2.Right side collar,connected to Web right
    ' 3.Left side collar,connected to Web Left
    '
    ' A slot can have:
    ' One collar,regular collar or right side collar (A)
    ' Two collars,one Right side collar (A), and one Left side collar(B)

    ' ********************* Selection code - TODO  ****************
    If sSlotClassName = "SlotC" Then
       If sSlotItemName = "SlotAC_LT_PAT" Then
        
          ' If a Plate/Stiffener Combo is used to form a EA or UA section alias
          ' then there is no fillet radius on the top flange right top corner.
          ' CollarACT_A expects a radius at the corner.  Therefore, if the
          ' penetrated part is a plate then we cannot add CollarACT_A.
          ' Use CollarTCT_A instead, which does not have a radius
          If Not bIsPenetratingObjPlate Then
            oSL.Add "CollarACT_A"
          Else
            oSL.Add "CollarTCT_A"
          End If
          
       ElseIf sSlotItemName = "SlotAC_LT_PAT2" Then
          ' If a Plate/Stiffener Combo is used to form a EA or UA section alias
          ' then there is no fillet radius on the top flange right top corner.
          ' CollarACT_A2 expects a radius at the corner.  Therefore, if the
          ' penetrated part is a plate then we cannot add CollarACT_A2.
          ' Use CollarTCT_A2 instead, which does not have a radius
          If Not bIsPenetratingObjPlate Then
            oSL.Add "CollarACT_A2"
          Else
            oSL.Add "CollarTCT_A2"
          End If
          
       ElseIf sSlotItemName = "SlotBC_LT_PAT" Then
          oSL.Add "CollarBCT_A"
          
       ElseIf sSlotItemName = "SlotBC_LT_PAT2" Then
          oSL.Add "CollarBCT_A2"
          
       ElseIf sSlotItemName = "SlotFC_LT_PAT" Then
          oSL.Add "CollarFCT_A"
          
       ElseIf sSlotItemName = "SlotFC_LT_PAT2" Then
          oSL.Add "CollarFCT_A2"
          
       ElseIf sSlotItemName = "SlotTC_T_PAA_STR" Or _
              sSlotItemName = "SlotAC_LT_AAA" Then
          ' No collar specified by TSuneishi
          ' If user selects these items, when they change this to other item,they must validate it!
       ElseIf sSlotItemName = "SlotAC_LT_PAA" Then
            'If Plate/Stiffener Combo then use the Collar for BUT
            If bIsPenetratingObjPlate Then
                oSL.Add "CollarTCT_A3"
            Else
                'No Equivalent Collar for EA/UA Profile
            End If
       Else
          Select Case sXSectionType
             Case "EA", "UA"
             
                ' If a Plate/Stiffener Combo is used to form a EA or UA section alias
                ' then there is no fillet radius on the top flange right top corner.
                If Not bIsPenetratingObjPlate Then
                    oSL.Add "CollarACT_SM"
                Else
                    ' A new symbol will have to be created that does not include a radius.
                End If
                
             Case "B"
                If sSlotItemName = "SlotBC_L_LTT_STR" Then
                   oSL.Add "CollarBCT_A3"
                End If
                oSL.Add "CollarBCT_SM"
                
             Case "FB"
                oSL.Add "CollarFCT_SM"
                            
             Case "BUT", "BUTL2"
                oSL.Add "CollarTCT_SM"
             
             Case Else
                sERROR = "Invalid cross section type"
                GoTo ErrorHandler
          End Select
       End If
    ElseIf sSlotClassName = SLOTC_TWO_COLLARS Then
      Dim strCollarOrder As String
      strCollarOrder = oSL.SelectorAnswer(m_sProjectName + ".RootClipSel", gsCollarCreationOrder)
      If sSlotItemName = "SlotL2C2_T_PTT_STR" Or _
         sSlotItemName = "SlotTC2_T_PTT_STR" Then
      
         If strCollarOrder = "Primary" Then
            oSL.Add "CollarTCT_A"
         Else
            oSL.Add "CollarTCT_B"
         End If
         
      ElseIf sSlotItemName = "SlotL2C2_T_PTT_STR2" Or _
             sSlotItemName = "SlotTC2_T_PTT_STR2" Then

         If strCollarOrder = "Primary" Then
            oSL.Add "CollarTCT_A2"
         Else
            oSL.Add "CollarTCT_B2"
         End If
             
      ElseIf sSlotItemName = "SlotTC2_T_PAA_STR" Then
         If strCollarOrder = "Primary" Then
            oSL.Add "CollarTCT_A3"
         Else
            oSL.Add "CollarTCT_B3"
         End If
      
      Else
         sERROR = "Invalid slot type"
         GoTo ErrorHandler
      End If
    Else
        sERROR = "Invalid slot type"
        GoTo ErrorHandler
    End If

    ' *********************************************************
  
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
    IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oSelector.IJDInputs.RemoveAllInput
    oSelector.IJDRepresentations.RemoveAllRepresentation

    
    Dim oDFact As New DefinitionFactory
    
    oDFact.InitAbstractSelector oSelector
    Set oDFact = Nothing
    
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.definition = oSelector
    oIH.InitAs m_sFamilyProgID
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
  
    Set IJDUserSymbolServices_InstanciateDefinition = oDFact.InstanciateSelector(m_sRuleProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), oRM)
    Set oDFact = Nothing
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
    Set oSL = Nothing
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
' This method sets answer to "AssyMethod" question based on planning assembly info.
Public Sub SetAnswerToAddCornerSnipe( _
                  ByVal oInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                  ByRef oArgument As Object)
    On Error GoTo ErrorHandler
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymDef As IMSSymbolEntities.IJDSymbolDefinition
    Dim oSL As IJDSelectorLogic
    
    Set oInputDG = oInput
    Set oSymDef = oInputDG.definition
    
    Set oSL = New SelectorLogic
    oSL.Representation = oSymDef.IJDRepresentations(1)
    oSL.Answer("AddCornerSnipe") = "Yes"

    Exit Sub
 
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetAnswerToAddCornerSnipe").Number

End Sub


