VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bracket3SSel"
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

Const m_SelectorProgid As String = "BracketRules.Bracket3SSel"
Const m_SelectorName As String = "BracketRules.Bracket3SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\Bracket3SSel.cls"

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
    
    'Check if support 3 is plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
   
    Select Case pHelper.ObjectType(pSupport3)
        Case SDOBJECT_PLATE
        
            'Get the lengths of the supports
            Dim dSupportLength1 As Double
            Dim dSupportLength2 As Double
            Dim dSupportLength3 As Double
            Dim dSupportLength4 As Double
            Dim dSupportLength5 As Double
            oBracket.GetSupportLengths dSupportLength1, _
                                       dSupportLength2, _
                                       dSupportLength3, _
                                       dSupportLength4, _
                                       dSupportLength5
            
            ' check the support lengths.  If length3 < 500 mm, call a bracket that
            ' dimensions from the width of support3.  Otherwise, if length2 < 500 mm,
            ' call a bracket that dimensions from the width of support2.
            If dSupportLength3 <= 0.5 Then
                pSL.Add "3SBracketShortSupport3"
            ElseIf dSupportLength2 <= 0.5 Then
                pSL.Add "3SBracketShortSupport2"
            Else
                pSL.Add "3SBracketPlates"
                'pSL.Add "3SBracket2TrimSupports"
            End If
        Case SDOBJECT_STIFFENER
            pSL.Add "3SBracketProfile"
            'pSL.Add "3SBracket2TrimSupports"
        Case Else
            'Support is not a plate or profile
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

