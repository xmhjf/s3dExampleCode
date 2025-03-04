VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlane2SSel"
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

Const m_SelectorProgid As String = "BracketRules.BracketByPlane2SSel"
Const m_SelectorName As String = "BracketRules.BracketByPlane2SSel"
Const m_FamilyProgid As String = ""

Const QUESTION_BracketByPlaneType As String = "BracketByPlaneType"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketByPlane2SSel.cls"

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
                    "2SLinearTrim", _
                    "BracketByPlane2SCol", _
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
        Case "2SLinearTrim"
            pSL.Add "BP_B2SLT HxWxE1xE2xR1"
            pSL.Add "BP_B2SLT 100x100x10x10"
            pSL.Add "BP_B2SLT 400x400x20x20"
            pSL.Add "BP_B2SLT 800x800x30x30"
            pSL.Add "BP_B2SLT 1200x1200x40x40"
            pSL.Add "BP_B2SLT 2000x2000x50x50"
        Case "2SLinearTrimLap"
            pSL.Add "BP_B2STL HxWxE1xE2xL1"
            pSL.Add "BP_B2STL 300x300x50x50x20"
            pSL.Add "BP_B2STL 400x400x50x50x50"
            pSL.Add "BP_B2STL 800x400x50x50x50"
            pSL.Add "BP_B2STL 800x800x50x50x50"
            pSL.Add "BP_B2STL 1200x1200x50x50x75"
            pSL.Add "BP_B2STL 2000x2000x150x150x75"
        Case "2SLinearTrimWithFlange"
            pSL.Add "BP_B2SLT_F HxWxE1xE2xR1"
        Case "2SCurvedTrim"
            pSL.Add "BP_B2ST HxWxE1xE2xR1"
            pSL.Add "BP_B2ST 400x400x50x50x500"
            pSL.Add "BP_B2ST 600x600x75x75x700"
            pSL.Add "BP_B2ST 800x800x75x75x1000"
            pSL.Add "BP_B2ST 1000x1000x100x100x1100"
            pSL.Add "BP_B2ST 2000x2000x150x150x2100"
        Case "2SLinearLap"
            pSL.Add "BP_B2SL HxWxE1xE2xL1"
            pSL.Add "BP_B2SL 300x300x50x50x20"
            pSL.Add "BP_B2SL 400x400x50x50x50"
            pSL.Add "BP_B2SL 800x800x75x75x100"
            pSL.Add "BP_B2SL 1200x1200x100x100x150"
            pSL.Add "BP_B2SL 2000x2000x150x150x200"
        Case "2SCurvedLinearTrim"
            pSL.Add "BP_B2SCT HxWxE1xE2xR1"
            pSL.Add "BP_B2SCT 300x300x50x50x200"
            pSL.Add "BP_B2SCT 400x400x50x50x250"
            pSL.Add "BP_B2SCT 800x800x75x75x350"
            pSL.Add "BP_B2SCT 1200x1200x100x100x500"
            pSL.Add "BP_B2SCT 2000x2000x150x150x1000"
        Case "2SLinearTrimScallop"
            pSL.Add "BP_B2STS HxWxE1xE2"
            pSL.Add "BP_B2STS 300x300x50x50"
            pSL.Add "BP_B2STS 400x400x50x50"
            pSL.Add "BP_B2STS 800x800x100x100"
            pSL.Add "BP_B2STS 1200x1200x150x150"
            pSL.Add "BP_B2STS 2000x2000x250x250"
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
        
    pSL.Answer(QUESTION_BracketByPlaneType) = "2SLinearTrim"
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



