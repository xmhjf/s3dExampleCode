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

' Modify class name to particular rule
Private Const m_sClassName As String = "CollarParm"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
'
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

Public Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

    oOH.SetOutput "SideOfPart"
    oOH.SetOutput PARAM_LAP_DISTANCE

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
    Dim dLapDistance As Double
    Dim dPlateThickness As Double
    Dim pHelper As StructDetailObjects.Helper
    
    Set pHelper = New StructDetailObjects.Helper
   
    dLapDistance = 0.05 ' Default to 0.05
    Select Case pHelper.ObjectType(oSlot.Penetrated)
        Case SDOBJECT_PLATE
            Dim oPlate As New StructDetailObjects.PlatePart
            Dim sCollarItemName As String
                        
            Set oPlate.object = oSlot.Penetrated
            sStiffenrSide = oPlate.StiffenerSide
            
            dPlateThickness = oPlate.PlateThickness
            If dPlateThickness < 0.0255 Then
               dLapDistance = 0.05
            Else
               dLapDistance = 0.075
            End If
        Case Else
            sStiffenrSide = "Base"
    End Select
  
  'Get Collar SideOfPart Answer
  Dim strCollarSide As String
  'TR-169311 Trying to retrieve the question's answer
    'without harcoding the path
    GetSelectorAnswer pPRL, gsCollarSideOfPart, strCollarSide
    
  ' Parameter rule:  Set Collar SideOfPart based on Answer
  ' StiffenerSide Base and Offset are internally set and not exposed to the user.
  ' IJDParameterLogic calls PenetrationSymbol, where:
  ' 0 = BASE
  ' 1 = OFFSET
  ' 2 = PLATE_CENTERED
  Select Case strCollarSide
    Case gsNoFlip
        If sStiffenrSide = "Base" Then
            pPRL.Add "SideOfPart", 1
        ElseIf sStiffenrSide = "Offset" Then
            pPRL.Add "SideOfPart", 0
        End If
    Case gsFlip
        If sStiffenrSide = "Base" Then
            pPRL.Add "SideOfPart", 0
        ElseIf sStiffenrSide = "Offset" Then
            pPRL.Add "SideOfPart", 1
        End If
    Case gsCentered
        pPRL.Add "SideOfPart", 2
  End Select
   
  pPRL.Add PARAM_LAP_DISTANCE, dLapDistance
  
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
  IJDUserSymbolServices_GetDefinitionName = m_sRuleName
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
  pIH.InitAs m_sFamilyProgID
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_sFamilyProgID
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_sRuleProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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

