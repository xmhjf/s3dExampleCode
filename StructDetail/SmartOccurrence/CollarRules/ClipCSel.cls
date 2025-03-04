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
'
Option Explicit

Const m_SelectorProgid As String = "CollarRules.ClipCSel"
Const m_SelectorName As String = "CollarRules.ClipCSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\ClipCSel.cls"

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
  
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim strCollarOrder As String

    Set oProfilePart.object = oSL.InputObject(INPUT_PENETRATING)
    strCollarOrder = oSL.SelectorAnswer("CollarRules.RootClipSel", gsCollarCreationOrder)

    ' ********************* Selection code - TODO  ****************
    If strCollarOrder = "Primary" Then
        Select Case oProfilePart.SectionType
            Case "EA", "UA", "T_XType", "TSType", "BUT", "BUTL2", "BUTL3", "I", "ISType", "C", "C_SS", "CSType", "FB"
                oSL.Add "ClipAAC"
                
            Case "B"
                oSL.Add "ClipBAC"
                
            Case Else
                sERROR = "Invalid cross section type in " & m_SelectorProgid
                GoTo ErrorHandler
        End Select
    ElseIf strCollarOrder = "Secondary" Then
        Select Case oProfilePart.SectionType
            Case "EA", "UA", "T_XType", "TSType", "BUT", "BUTL2", "BUTL3", "I", "ISType", "C", "C_SS", "CSType", "FB"
                oSL.Add "ClipABC"
                
            Case "B"
                oSL.Add "ClipBBC"
                
            Case Else
                sERROR = "Invalid cross section type in " & m_SelectorProgid
                GoTo ErrorHandler
        End Select
    Else
        sERROR = "Invalid collar order in " & m_SelectorProgid
    End If
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  oSelector.IJDInputs.RemoveAllInput
  oSelector.IJDRepresentations.RemoveAllRepresentation
  
  Dim oDFact As New DefinitionFactory
  Dim oIH As IJDInputsHelper
  
  oDFact.InitAbstractSelector oSelector
  
  Set oIH = New InputHelper
  oIH.Definition = oSelector
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



