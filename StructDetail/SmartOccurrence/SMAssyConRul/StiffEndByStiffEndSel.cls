VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByStiffEndSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StiffEndByStiffEndSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StiffEndByStiffEndSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndByStiffEndSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
zMsgBox "StiffEndByStiffEndSel::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions
    strError = "Setting questions."
    pQH.SetQuestion "EndCutType", gsW, "EndCutTypeCodeList", "SetEndCutType", CUSTOMERID + "AssyConRul.AssyConnSelCM"
    
    ''' Can NOT Delete (Remove) a Question from the Selector
    ''' (Can Not change the order of Questions in a Selector)
    ''' to remove a Question, set Name and Default value to NULL
    ''' Question will not be displayed on Property Pages
    ''' the Question Name can not be re-used
    pQH.SetQuestion "", ""
''''    pQH.SetQuestion "EndToEndCase", "NoAngle", , "SetEndToEndDefault", CMLIBRARY_ASSYCONNSEL
    
    pQH.SetQuestion "WeldPart", "First"
       
    ' Add new Questions after all existing Questions
    pQH.SetQuestion "SplitEndToEndCase", _
                    "NoAngle", _
                    "SplitEndCutTypes", _
                    "SetSplitEndToEndCase"
                    
    pQH.SetQuestion "FlipEndToEndAngle", "NoFlip", "FlipSplitEndCutAngle"
    pQH.SetQuestion "ChamferType", "None", , "ChamferQuestionOnER", CMLIBRARY_ASSYCONNSEL
       
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    ' There MUST always be an "EndCut" applied to the Profile'e Bounded  End
    pSL.Add "StiffenerEndToStiffenerEnd"
zMsgBox "StiffEndByStiffEndSel::SelectorLogic" & vbCrLf & _
        "   pSL.Add StiffenerEndToStiffenerEnd"
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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

Public Sub SetSplitEndToEndCase( _
                      ByVal oInputStdCM As IJDInputStdCustomMethod, _
                      ByRef oArgument As Object)
                      
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymDef As IMSSymbolEntities.IJDSymbolDefinition
    Dim oSL As IJDSelectorLogic
    
    Set oInputDG = oInputStdCM
    Set oSymDef = oInputDG.definition
    Set oInputDG = Nothing
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Set oSL = New SelectorLogic
    oSL.Representation = oSymDef.IJDRepresentations(1)
    Set oSymDef = Nothing
    
    Dim oACWrapper As New StructDetailObjects.AssemblyConn
    
    Set oACWrapper.object = oSL.SmartOccurrence
    If oACWrapper.FromPlanningSeam = True Then
       oSL.Answer("SplitEndToEndCase") = "AlongGlobalAxis"
    Else
       oSL.Answer("SplitEndToEndCase") = "NoAngle"
    End If
    
    Set oACWrapper = Nothing
    
    Set oSL = Nothing
    
    Exit Sub
End Sub

