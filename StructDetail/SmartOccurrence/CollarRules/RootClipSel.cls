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
'
Option Explicit

Const m_SelectorProgid As String = "CollarRules.RootClipSel"
Const m_SelectorName As String = "CollarRules.RootClipSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\RootClipSel.cls"

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
  
  'Get Penetrated Object from Slot inputs.  This will be used to get the tightness.  The base plate of
  'the slot is used if the penetrated object is a profile
  On Error Resume Next
  Dim oPlate As IJPlate
  Dim oPlatePart As New StructDetailObjects.PlatePart
  Set oPlate = oSlot.Penetrated
  If oPlate Is Nothing Then
    Set oPlatePart.object = oSlot.BasePlate
  Else
    Set oPlatePart.object = oSlot.Penetrated
  End If
  
  'Get answers
  Dim strStressLevel As String
  strStressLevel = pSL.Answer(gsStressLevel)
  Dim strCollarOrder As String
  strCollarOrder = pSL.Answer(gsCollarCreationOrder)
    
  ' ********************* Selection code - TODO  ****************
  
  If (strCollarOrder = "Primary") Then
        Select Case oPlatePart.Tightness
           Case AirTight, OilTight, WaterTight
               Select Case oSlot.ClassName
                   Case "SlotI"
                       pSL.Add "CollarNoSnipe"
                   Case "SlotC"
                       pSL.Add "CollarOneSnipe"
                   Case "SlotA"
                       pSL.Add "CollarTwoSnipes"
               End Select
               
           Case NonTight, UnSpecifiedTightness
                 If oSlot.ClassName = "SlotC" Then
                     If strStressLevel = gsHigh Then
                         pSL.Add "ClipA"
                         pSL.Add "ClipC"
                     End If
                 ElseIf oSlot.ClassName = "SlotA" Then
                     If strStressLevel = gsMedium Or strStressLevel = gsHigh Then
                         pSL.Add "ClipA"
                         pSL.Add "ClipC"
                     End If
                 End If
                 
        End Select
  ElseIf (strCollarOrder = "Secondary") Then ' Secondary
          Select Case oPlatePart.Tightness
           Case AirTight, OilTight, WaterTight
                 ' No clip required
           Case NonTight, UnSpecifiedTightness
                 If oSlot.ClassName = "SlotA" And strStressLevel = gsHigh Then
                     pSL.Add "ClipB"
                     pSL.Add "ClipC"
                 End If
        End Select
     
  Else
        strError = "Invalid Collar Creation Order: " & strCollarOrder
        GoTo ErrorHandler
  End If
   
  ' *********************************************************
  
  Set oPlatePart = Nothing
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
  pIH.Definition = pSelector
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


