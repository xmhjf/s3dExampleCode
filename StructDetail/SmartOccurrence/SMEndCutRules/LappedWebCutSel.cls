VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LappedWebCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgId
Const m_SelectorProgid As String = CUSTOMERID + "EndCutRules.LappedWebCutSel"
Const m_SelectorName As String = CUSTOMERID + "EndCutRules.LappedWebCutSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\LappedWebCutSel.cls"

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
zMsgBox "EndToEndWebCutSel::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions
    ''' Can NOT Delete (Remove) a Question from the Selector
    ''' (Can Not change the order of Questions in a Selector)
    ''' to remove a Question, set Name and Default value to NULL
    ''' Question will not be displayed on Property Pages
    ''' the Question Name can not be re-used
     
    strError = "Setting questions."
    SetQuestionEndCutType pQH
     
    pQH.SetQuestion gsAddCornerFeature, "No", "BooleanCol", "AddCornerFeatureCM", m_SelectorProgid
     
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    strError = "Error selecting lapped endcut"
    
    Dim oWebCut As New Structdetailobjects.WebCut
    Set oWebCut.object = pSLH.SmartOccurrence
    Dim oBoundedPart As Object
    Set oBoundedPart = oWebCut.Bounded
    
    ' get the Bounded Part's CrossSection type
    Dim phelper As New Structdetailobjects.Helper
    Dim sSectionType As String
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oProfile As New Structdetailobjects.ProfilePart
        Set oProfile.object = oBoundedPart
        sSectionType = oProfile.SectionType
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBeam As New Structdetailobjects.BeamPart
        Set oBeam.object = oBoundedPart
        sSectionType = oBeam.SectionType
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_MEMBER Then
        Dim oMember As New Structdetailobjects.MemberPart
        Set oMember.object = oBoundedPart
        sSectionType = oMember.SectionType
    Else

        Exit Sub
    End If

    Select Case sSectionType
        Case "FB"
            pSLH.Add "WebCut_LappedWOWA"
        Case Else
            pSLH.Add "WebCut_LappedWOWAFA"
    End Select
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
End Sub



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
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    pSLH.Representation = pRep
    SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


' ********************
' Method sets default answer for whether or not the webcut has a child corner feature
' ***********************
Public Sub AddCornerFeatureCM(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
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
    
    If IsBoundingStiffPenetratingSupport(pSL.SmartOccurrence) Then
        pSL.Answer(gsAddCornerFeature) = "Yes"
    Else
        pSL.Answer(gsAddCornerFeature) = "No"
    End If
    
    Exit Sub
End Sub


