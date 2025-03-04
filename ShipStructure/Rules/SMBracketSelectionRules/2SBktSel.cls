VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt2SSel"
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

Const m_SelectorProgid As String = "BktSelRules.Bkt2SSel"
Const m_SelectorName As String = "BktSelRules.Bkt2SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktSelRules\2SBktSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    
    pIH.SetInput INPUT_SUPPORTS

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler

    pQH.SetQuestion QUESTION_BracketContourType, _
                    DEFAULT_BracketContourType_LINEAR, _
                    CODELIST_BracketContourType
  
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketName As String
    Dim sBracketContourType As String
    Dim sBracketSupport1Type As ShpStrBktSupportConnectionType
    Dim sBracketSupport2Type As ShpStrBktSupportConnectionType
    
    '1. Get inputs for deciding bracket selection...
    sBracketSupport1Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 1)
    sBracketSupport2Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 2)
    sBracketContourType = pSL.SelectorAnswer("BktSelRules.Bkt2SSel", QUESTION_BracketContourType)
    
    '2. Define the bracket prefix for naming...
    sBracketName = "BPS_"
    
    '3. Append the bracket contour type...
    Select Case sBracketContourType
        Case "Linear"
            sBracketName = sBracketName & "2SL"
            
        Case "Curved"
            sBracketName = sBracketName & "2SC"
        
    End Select
 
    '4. Append the bracket lapped conditions (if required)...
    If (sBracketSupport1Type = Lapped) And (sBracketSupport2Type = Lapped) Then
        sBracketName = sBracketName & "_Lapped"
    ElseIf (sBracketSupport1Type = Lapped) Or (sBracketSupport2Type = Lapped) Then
        sBracketName = sBracketName & "_Lap1"
   
       
    End If
    
    '5. Add Items to the list of available brackets...
    Call Add2SBrackets(pSL, sBracketName)
    
    If (sBracketSupport1Type = Lapped) And (sBracketSupport2Type = Lapped) Then
        Select Case sBracketContourType
        Case "Linear"
                 pSL.Add "BPS_2SLL2Sup 2000x2000x150x150x150x150"
                
            Case "Curved"
                 pSL.Add "BPS_2SCL2Sup 2000x2000x150x150x150x150x1500"
            
        End Select
    ElseIf (sBracketSupport1Type = Lapped) Or (sBracketSupport2Type = Lapped) Then
        Select Case sBracketContourType
            Case "Linear"
                pSL.Add "BPS_2SLL1Sup HxWxE1xE2xL1"
                
            Case "Curved"
                pSL.Add "BPS_2SL_Lap1 2000x2000x150x150"
        
        End Select
    Else
        Select Case sBracketContourType
            Case "Linear"
                pSL.Add "BPS_2SLT HxWxN1xN2"
                
            Case "Curved"
                pSL.Add "BPS_2SCT HxWxN1xN2"
            
        End Select
       pSL.Add "BPS_2SLTMem1"
    End If
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



