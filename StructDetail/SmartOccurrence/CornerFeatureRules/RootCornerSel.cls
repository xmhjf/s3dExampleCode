VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootCornerSel"
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

Const m_SelectorProgid As String = "CornerFeatureRules.RootCornerSel"
Const m_SelectorName As String = "CornerFeatureRules.RootCornerSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CornerFeatureRules\RootCornerSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    ' Define questions
    pQH.SetQuestion gsPlacement, "Manual", "PlacementCol"   ' NOTE: This default is only used if not created by Assembly connection
    pQH.SetQuestion gsCrackArrest, "Yes", "BooleanCol"
    pQH.SetQuestion gsDrainage, "Yes", "BooleanCol"
    pQH.SetQuestion gsCornerFlip, "NoFlip", "FlipCol"      'Will use answer to this question in parameter rule, not selector rule
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
     
    On Error GoTo ErrorHandler
    
' Get corner feature
Dim oCornerFeature As New Structdetailobjects.CornerFeature
Set oCornerFeature.object = pSL.SmartOccurrence

'Get answers
Dim strPlacement As String
strPlacement = pSL.Answer(gsPlacement)

Dim strDrainage As String
strDrainage = pSL.Answer(gsDrainage)

Dim strCrackArrest As String
strCrackArrest = pSL.Answer(gsCrackArrest)

Dim strCornerFlip As String
strCornerFlip = pSL.Answer(gsCornerFlip)

'Get corner feature part object
Dim oPart As Object
Set oPart = oCornerFeature.GetPartObject

'Selection code
If TypeOf oPart Is IJPlate Then
    Dim oPlate As New Structdetailobjects.PlatePart
    Set oPlate.object = oPart
    Select Case oPlate.Tightness
        Case NonTight
            Select Case strCrackArrest
                Case "Yes"
                    pSL.Add "SmartGeneralCorner"
                    pSL.Add "SmartScallop"
                    pSL.Add "SmartSnipe"
                Case "No"
                    Select Case strDrainage
                        Case "Yes"
                            pSL.Add "SmartScallop"
                            pSL.Add "SmartGeneralCorner"
                            pSL.Add "SmartSnipe"
                        Case "No"
                            Call CornerPlacement(pSL)
                    End Select
            End Select
        Case Else   'Plate is Watertight, airtight, or oiltight
            Call CornerPlacement(pSL)
    End Select

ElseIf TypeOf oPart Is IJProfile Or TypeOf oPart Is ISPSMemberPartPrismatic Then  'Part is a profile or member
    Select Case strDrainage
        Case "Yes"
        
            '**********************************************
            'new code for checking the corner gap
            If oCornerFeature.CornerGapExists(0.001) Then
                pSL.Add "SmartGeneralCorner"
            Else 'no gap, take the defaults
                pSL.Add "SmartScallop"
                pSL.Add "SmartSnipe"
                pSL.Add "SmartGeneralCorner"
            End If
            
            '************************************************
            
            
        Case "No"
            Call CornerPlacement(pSL)
    End Select
End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub


'  ------------------------------------------------------------
' This selection rule selects a corner feature based on whether it has been
' placed manually or by an assembly connection

Private Sub CornerPlacement(pSL As IJDSelectorLogic)
On Error GoTo ErrorHandler

' Get corner feature
Dim oCornerFeature As New Structdetailobjects.CornerFeature
Set oCornerFeature.object = pSL.SmartOccurrence

'Get answers
Dim strPlacement As String
strPlacement = pSL.Answer(gsPlacement)

'Get corner part object
On Error Resume Next
Dim oPart As IJPlate
Set oPart.object = oCornerFeature.GetPartObject
    
Select Case strPlacement
    Case "Manual"   'Always have all corner feature items available for manual placement
        'After bug on using parms in multiple selectors is fixed (03001000?), a separate
        'Manual class will be created, containing a complete list of items (such as
        '"ManualScallop") pointing to the same symbols files and same parms as the
        'Smart items.
        pSL.Add "SmartScallop"
        pSL.Add "SmartGeneralCorner"
        pSL.Add "SmartSnipe"
    Case "AssyConn"
        'No Corner Feature
End Select

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CornerPlacement").Number
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

