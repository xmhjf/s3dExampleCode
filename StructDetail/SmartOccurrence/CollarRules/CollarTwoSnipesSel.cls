VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarTwoSnipesSel"
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

Const m_SelectorProgid As String = "CollarRules.CollarTwoSnipesSel"
Const m_SelectorName As String = "CollarRules.CollarTwoSnipesSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarTwoSnipesSel.cls"

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
  
    oQH.SetQuestion gsAddCornerSnipes, gsNone, "CornerSnipeCol"
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
  
    Dim sERROR As String
    Dim oSlot As New StructDetailObjects.Slot
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    
    sERROR = ""
    'Get Slot object from Collar inputs
    Set oSlot.object = oSL.InputObject(INPUT_SLOT)
  
    'Get Penetrating object
    Set oProfilePart.object = oSL.InputObject(INPUT_PENETRATING)

    ' ********************* Selection code - TODO  ****************
    If oSlot.ClassName = "SlotA" Then
        Select Case oProfilePart.SectionType
            Case "EA", "UA"
                oSL.Add "CollarAAT"
                
            Case "T_XType", "TSType"
                oSL.Add "CollarTsAT"
                
            Case "BUT", "BUTL2"
                oSL.Add "CollarTAT"
                
            Case "BUTL3"
                oSL.Add "CollarLAT"
                
            Case "FB"
                oSL.Add "CollarFAT"
                   
            Case "B"
                oSL.Add "CollarBAT"
            
            Case Else
                sERROR = "Invalid cross section type"
                GoTo ErrorHandler
        End Select
    Else
        sERROR = "Invalid slot type"
        GoTo ErrorHandler
    End If
  
    ' *********************************************************
  
    Set oProfilePart = Nothing
    Set oSlot = Nothing
  
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oSelector.IJDInputs.RemoveAllInput
    oSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim oDFact As New DefinitionFactory
    
    oDFact.InitAbstractSelector oSelector
    
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.Definition = oSelector
    oIH.InitAs m_FamilyProgid
    SelectorInputs oIH
    
    Dim oQH As IJDQuestionsHelper
    
    Set oQH = New QuestionHelper
    oQH.Selector = oSelector
    SelectorQuestions oQH
    
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
