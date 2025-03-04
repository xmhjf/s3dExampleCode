VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = "CollarRules.CollarParm"
Const m_ParameterRuleName As String = "CollarRules.CollarParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_BOUNDINGPLATE
  pIH.SetInput INPUT_SLOT
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

    pOH.SetOutput "SideOfPart"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

  'Get slot object
  Dim oSlot As StructDetailObjects.Slot
  Set oSlot = New Slot
  Set oSlot.object = pPRL.InputObject(INPUT_SLOT)
    
   'Check to see if penetrated is a plate or profile
    Dim sStiffenrSide As String
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
   
    Select Case pHelper.ObjectType(oSlot.Penetrated)
        Case SDOBJECT_PLATE
            Dim oPlate As New StructDetailObjects.PlatePart
            Set oPlate.object = oSlot.Penetrated
            sStiffenrSide = oPlate.StiffenerSide
        Case Else
            sStiffenrSide = "Base"
    End Select
  
  'Get Collar SideOfPart Answer
  Dim strCollarSide As String
  strCollarSide = pPRL.SelectorAnswer("CollarRules.RootClipSel", gsCollarSideOfPart)
  
  ' Parameter rule:  Set Collar SideOfPart based on Answer
  ' StiffenerSide Base and Offset are internally set and not exposed to the user.
  ' IJDParameterLogic calls PenetrationSymbol, where:
  ' 0 = PLATE_MOLDED (offset)
  ' 1 = PLATE_ANTI-MOLDED (base)
  ' 2 = PLATE_CENTERED
  Select Case strCollarSide
    Case gsNoFlip
        If sStiffenrSide = "Base" Then
            pPRL.Add "SideOfPart", 0
        ElseIf sStiffenrSide = "Offset" Then
            pPRL.Add "SideOfPart", 1
        End If
    Case gsFlip
        If sStiffenrSide = "Base" Then
            pPRL.Add "SideOfPart", 1
        ElseIf sStiffenrSide = "Offset" Then
            pPRL.Add "SideOfPart", 0
        End If
    Case gsCentered
        pPRL.Add "SideOfPart", 2
 End Select

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **





' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation
  
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

