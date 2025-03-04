VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlaneMT1SSel"
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

Const m_SelectorProgid As String = "BracketRules.BracketByPlaneMT1SSel"
Const m_SelectorName As String = "BracketRules.BracketByPlaneMT1SSel"
Const m_FamilyProgid As String = ""

Const QUESTION_BracketByPlaneType As String = "BracketByPlaneType"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketByPlaneMT1SSel.cls"

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
    'SetQuestion(name As String, defaultVal, [CodeListName As String], [CMName As String], [LibID As String], [customOcx As String]) As IJDInput
    
    pQH.SetQuestion QUESTION_BracketByPlaneType, _
                    "2SLinearTrim", _
                    "BracketByPlane2SCol", _
                    "BracketType"
    ' 3S brackets on a single support are not supported for this release
'    pQH.SetQuestion QUESTION_BracketByPlaneType, _
'                    "3SSymmetrical", _
'                    "BracketByPlane3SCol", _
'                    "BracketType"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketByPlaneType As String
    sBracketByPlaneType = pSL.Answer(QUESTION_BracketByPlaneType)
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
            flangeWidth = .FlangeLength * 0.4 'want the bracket to be 40% of the flange
        Else
            flangeWidth = .FlangeLength * 0.8 'want the bracket to be 80% of the flange
        End If
    
        webheight = .WebLength - .FlangeThickness
    End With

    

    Select Case sBracketByPlaneType
        Case "2SLinearTrim"
            pSL.Add "BPMT_1S_B2SLT HxWxE1xE2xR1"
            If ((flangewidth > 0.1) And (webheight > 0.1)) Then
                pSL.Add "BPMT_1S_B2SLT 100x100x10x10"
            End If
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BPMT_1S_B2SLT 400x400x20x20"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BPMT_1S_B2SLT 800x800x30x30"
            End If
            If ((flangewidth > 1.2) And (webheight > 1.2)) Then
                pSL.Add "BPMT_1S_B2SLT 1200x1200x40x40"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "BPMT_1S_B2SLT 2000x2000x50x50"
            End If
        Case "2SLinearTrimWithFlange"
            pSL.Add "BPMT_1S_B2SLT_F HxWxE1xE2xR1"
        Case "2SCurvedTrim"
            pSL.Add "BPMT_1S_B2ST HxWxE1xE2xR1"
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BPMT_1S_B2ST 400x400x50x50x500"
            End If
            If ((flangewidth > 0.6) And (webheight > 0.6)) Then
                pSL.Add "BPMT_1S_B2ST 800x800x75x75x1000"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BPMT_1S_B2ST 800x800x75x75x1000"
            End If
            If ((flangewidth > 1#) And (webheight > 1#)) Then
                pSL.Add "BPMT_1S_B2ST 1000x1000x100x100x1100"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "2000x2000x150x150x2100"
            End If
        Case "2SCurvedLinearTrim"
            pSL.Add "BPMT_1S_B2SCT HxWxE1xE2xR1"
            If ((flangewidth > 0.3) And (webheight > 0.3)) Then
                pSL.Add "BPMT_1S_B2SCT 300x300x50x50x200"
            End If
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BPMT_1S_B2SCT 400x400x50x50x250"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BPMT_1S_B2ST 800x800x75x75x1000"
            End If
            If ((flangewidth > 1.2) And (webheight > 1.2)) Then
                pSL.Add "BPMT_1S_B2SCT 1200x1200x100x100x500"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "BPMT_1S_B2SCT 2000x2000x150x150x1000"
            End If
        Case "2SLinearTrimScallop"
            pSL.Add "BPMT_1S_B2STS HxWxE1xE2"
            If ((flangewidth > 0.3) And (webheight > 0.3)) Then
                pSL.Add "BPMT_1S_B2STS 300x300x50x50"
            End If
            If ((flangewidth > 0.4) And (webheight > 0.4)) Then
                pSL.Add "BPMT_1S_B2STS 400x400x50x50"
            End If
            If ((flangewidth > 0.8) And (webheight > 0.8)) Then
                pSL.Add "BPMT_1S_B2STS 800x800x100x100"
            End If
            If ((flangewidth > 1.2) And (webheight > 1.2)) Then
                pSL.Add "BPMT_1S_B2STS 1200x1200x150x150"
            End If
            If ((flangewidth > 2#) And (webheight > 2#)) Then
                pSL.Add "BPMT_1S_B2STS 2000x2000x250x250"
            End If
    ' 3S brackets on a single support are not supported for this release
            
'            Case "3SSymmetrical"
'            pSL.Add "BPMT_1S_B3S H1xH2xW1xW2xN1xN2xR1"
'            pSL.Add "BPMT_1S_B3S 500x500x250x300x50x50x100"
'            pSL.Add "BPMT_1S_B3S 800x800x300x350x75x75x150"
'            pSL.Add "BPMT_1S_B3S 1200x1200x500x600x100x100x200"
'            pSL.Add "BPMT_1S_B3S 1500x1500x500x600x100x100x200"
'
'        Case "3SUnSymParametric"
'            pSL.Add "BPMT_1S_B3S W/4xH2xW2xN1xR1"
'            pSL.Add "BPMT_1S_B3S W/4x300x400x50x100"
'            pSL.Add "BPMT_1S_B3S W/4x500x700x100x200"
'            pSL.Add "BPMT_1S_B3S W/4x800x900x150x300"
'            pSL.Add "BPMT_1S_B3S W/4x1000x1200x200x400"
'        Case "3SUnSymmetrical"
'            pSL.Add "BPMT_1S_B3S H1xH2xW2xN1xR1"
'            pSL.Add "BPMT_1S_B3S 500x300x400x50x100"
'            pSL.Add "BPMT_1S_B3S 800x500x700x100x200"
'            pSL.Add "BPMT_1S_B3S 1200x800x900x150x300"
'            pSL.Add "BPMT_1S_B3S 1500x1000x1200x200x400"
        
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





