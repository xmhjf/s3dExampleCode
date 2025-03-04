VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt1SSel"
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

Const m_SelectorProgid As String = "BktUnifiedRules.Bkt1SSel"
Const m_SelectorName As String = "BktUnifiedRules.Bkt1SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktUnifiedRules\Bkt1SSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
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
    
    Dim sBracketContourType As String
    sBracketContourType = pSL.Answer(QUESTION_BracketContourType)
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = pSL.SmartOccurrence
    
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
   
    'Get supports from BracketContourSO.
    Dim oSupports As IJElements
    Dim nSupports As Long
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupports, nSupports
    Dim oProfileObject As StructDetailObjects.ProfileSystem
    Set oProfileObject = New StructDetailObjects.ProfileSystem
    Dim oDCrossSection As IJDProfileSection
    Dim flangewidth As Double
    Dim webheight As Double
    With oProfileObject
        ' use the affected leaf system as the cross section can change
        Set .object = oSupports(1)
        Set oDCrossSection = .object
        'if the profile has a flange, it can be used for a "1S"
        If (oDCrossSection.SectionType And Tee) > 0 Then
            flangewidth = .FlangeLength * 0.4 'want the bracket to be 40% of the flange
        Else
            flangewidth = .FlangeLength * 0.8 'want the bracket to be 80% of the flange
        End If
    
        webheight = .WebLength - .FlangeThickness
    End With

    Select Case sBracketContourType
        Case "Linear"
            pSL.Add "BKT_1SLT HxWxN1xN2"
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BKT_1SLT 400x400x50x50"
            End If
            If ((flangewidth > 0.6) And (webheight > 0.6)) Then
                pSL.Add "BKT_1SLT 600x600x75x75"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BKT_1SLT 800x800x75x75"
            End If
            If ((flangewidth > 1#) And (webheight > 1#)) Then
                pSL.Add "BKT_1S_B2SLT 1000x1000x100x100"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "BKT_1SLT 2000x2000x150x150"
            End If
         
         Case "Curved"
            pSL.Add "BKT_1SCT HxWxN1xN2"
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BKT_1SCT 400x400x50x50"
            End If
            If ((flangewidth > 0.6) And (webheight > 0.6)) Then
                pSL.Add "BKT_1SCT 600x600x75x75"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BKT_1SCT 800x800x75x75"
            End If
            If ((flangewidth > 1#) And (webheight > 1#)) Then
                pSL.Add "BKT_1SCT 1000x1000x100x100"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "BKT_1SCT 2000x2000x150x150"
            End If
             
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





