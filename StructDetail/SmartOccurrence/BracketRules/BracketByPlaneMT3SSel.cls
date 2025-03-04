VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlaneMT3SSel"
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

Const m_SelectorProgid As String = "BracketRules.BracketByPlaneMT3SSel"
Const m_SelectorName As String = "BracketRules.BracketByPlaneMT3SSel"
Const m_FamilyProgid As String = ""

Const QUESTION_BracketByPlaneType As String = "BracketByPlaneType"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketByPlaneMT3SSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' inputs:
    '   Plane that bracket is on
    '   Bracket Plate System
    '   U Point
    '   V Point
    '   Group of supports
    
    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
    pIH.SetInput INPUT_UPOINT
    pIH.SetInput INPUT_VPOINT
    pIH.SetInput INPUT_SUPPORTS
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    pQH.SetQuestion QUESTION_BracketByPlaneType, _
                    "3SSymmetrical", _
                    "BracketByPlane3SCol", _
                    "BracketType"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketByPlaneType As String
    sBracketByPlaneType = pSL.Answer(QUESTION_BracketByPlaneType)
    
    Select Case sBracketByPlaneType
        Case "3SSymmetrical"
            pSL.Add "BPMT_B3S H1xH2xW1xW2xN1xN2xR1"
            pSL.Add "BPMT_B3S 500x500x250x300x50x50x100"
            pSL.Add "BPMT_B3S 800x800x300x350x75x75x150"
            pSL.Add "BPMT_B3S 1200x1200x500x600x100x100x200"
            pSL.Add "BPMT_B3S 1500x1500x500x600x100x100x200"
        Case "3SUnSymParametric"
            pSL.Add "BPMT_B3S W/4xH2xW2xN1xR1"
            pSL.Add "BPMT_B3S W/4x300x400x50x100"
            pSL.Add "BPMT_B3S W/4x500x700x100x200"
            pSL.Add "BPMT_B3S W/4x800x900x150x300"
            pSL.Add "BPMT_B3S W/4x1000x1200x200x400"
        Case "3SUnSymmetrical"
            pSL.Add "BPMT_B3S H1xH2xW2xN1xR1"
            pSL.Add "BPMT_B3S 500x300x400x50x100"
            pSL.Add "BPMT_B3S 800x500x700x100x200"
            pSL.Add "BPMT_B3S 1200x800x900x150x300"
            pSL.Add "BPMT_B3S 1500x1000x1200x200x400"
        Case "3STrimS1S2Lap"
            pSL.Add "BPMT_B3S1S2L H1xH2xW1xW2xN1xN2xR1"
            pSL.Add "BPMT_B3S1S2L 400x400x200x200x50x50x100"
            pSL.Add "BPMT_B3S1S2L 800x800x400x400x100x100x200"
            pSL.Add "BPMT_B3S1S2L 1200x1200x600x600x200x200x400"
            pSL.Add "BPMT_B3S1S2L 1500x1500x750x750x250x250x500"
        Case "3STrimS1Lap"
            pSL.Add "BPMT_B3S1L H1xH2xW1xW2xN1xN2xR1"
            pSL.Add "BPMT_B3S1L 400x400x200x200x50x50x100"
            pSL.Add "BPMT_B3S1L 800x800x400x400x100x100x200"
            pSL.Add "BPMT_B3S1L 1200x1200x600x600x200x200x400"
            pSL.Add "BPMT_B3S1L 1500x1500x750x750x250x250x500"
        Case "3SLinearLap"
            pSL.Add "BPMT_B3SL  H1xH2xW1xW2xN1xN2xLD"
            pSL.Add "BPMT_B3S 500x500x200x200x50x50x50"
            pSL.Add "BPMT_B3S 800x800x400x400x75x75x75"
            pSL.Add "BPMT_B3S 1200x1200x600x600x75x75x75"
            pSL.Add "BPMT_B3S 1500x1500x800x800x100x100x100"
         Case "3SFillShortPt"
            pSL.Add "BPMT_B3FSp  H1xH2"
            pSL.Add "BPMT_B3FSp 25x25"
            pSL.Add "BPMT_B3FSp 50x50"
            pSL.Add "BPMT_B3FSp 75x75"
            pSL.Add "BPMT_B3FSp 100x100"

        Case "3SSetBackShortPt"
            pSL.Add "BPMT_B3SBSp  H1xH2xN1xSBxSc"
            pSL.Add "BPMT_B3SBSp 200x50x50x50x25"
            pSL.Add "BPMT_B3SBSp 300x50x50x75x25"
            pSL.Add "BPMT_B3SBSp 400x75x50x100x50"
            pSL.Add "BPMT_B3SBSp 800x150x50x100x50"
        Case "NGBe1"
             pSL.Add "BPMT_GBe1_1"
              
    End Select
                
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub
Public Sub BracketType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.Definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
    pSL.Answer(QUESTION_BracketByPlaneType) = "3SSymmetrical"
    
              
    Exit Sub
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



