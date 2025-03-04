VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bracket4SSel"
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

Const m_SelectorProgid As String = "BracketRules.Bracket4SSel"
Const m_SelectorName As String = "BracketRules.Bracket4SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\Bracket4SSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add input description below
    ' inputs:
    '   Plane that bracket is on
    '   Group of supports
    '   Thickness direction
    
    pIH.SetInput INPUT_SUPPORT1
    pIH.SetInput INPUT_SUPPORT2
    pIH.SetInput INPUT_SUPPORT3
    pIH.SetInput INPUT_SUPPORT4
    pIH.SetInput INPUT_PLANE
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    'Get Bracket
    Dim oBracket As StructDetailObjects.Bracket
    Set oBracket = New Bracket
    Set oBracket.object = pSL.SmartOccurrence
    
    'Get Supports
    Dim nSupports As Long
    Dim pRefPlane As IJPlane
    Dim pSupport1 As IJPort
    Dim pSupport2 As IJPort
    Dim pSupport3 As Object
    Dim pSupport4 As Object
    Dim pSupport5 As Object

    oBracket.GetInputs nSupports, pRefPlane, pSupport1, pSupport2, pSupport3, pSupport4, pSupport5
    
    'Check support  4 and 5 is plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
   
    Select Case pHelper.ObjectType(pSupport3)
        Case SDOBJECT_PLATE
            Select Case pHelper.ObjectType(pSupport4)
                Case SDOBJECT_PLATE
                    pSL.Add "4SBracketPlates"
                Case SDOBJECT_STIFFENER
                    pSL.Add "4SBracketPlateProfile"
            End Select
        Case SDOBJECT_STIFFENER
            Select Case pHelper.ObjectType(pSupport4)
                Case SDOBJECT_PLATE
                    pSL.Add "4SBracketProfilePlate"
                Case SDOBJECT_STIFFENER
                    pSL.Add "4SBracketProfileProfile"
            End Select
        Case Else
            'Support 3 or 4 is not a plate or profile
    End Select

    
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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

